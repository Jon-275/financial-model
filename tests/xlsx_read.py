"""Read cached Excel cells without opening or modifying the workbook."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath
import re
from typing import Iterator
from xml.etree import ElementTree as ET
import zipfile


MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

ERRS = frozenset({"#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!", "#SPILL!", "#CALC!"})


def q(name: str) -> str:
    return f"{{{MAIN}}}{name}"


def sheet_parts(source: zipfile.ZipFile) -> dict[str, str]:
    book = ET.fromstring(source.read("xl/workbook.xml"))
    rels = ET.fromstring(source.read("xl/_rels/workbook.xml.rels"))
    targets = {rel.get("Id"): rel.get("Target") for rel in rels}
    return {
        node.get("name"): str(PurePosixPath(target.lstrip("/")) if target.startswith("/") else PurePosixPath("xl", target))
        for node in book.findall(f"{q('sheets')}/{q('sheet')}")
        if (target := targets.get(node.get(f"{{{REL}}}id")))
    }


def address_parts(address: str) -> tuple[int, int]:
    match = re.fullmatch(r"([A-Z]{1,3})([1-9][0-9]*)", address)
    if not match:
        raise ValueError(f"Invalid Excel cell address: {address!r}")
    col, row = match.groups()
    index = 0
    for char in col:
        index = 26 * index + ord(char) - 64
    return int(row), index


def col_letter(index: int) -> str:
    col = ""
    while index:
        index, ch = divmod(index - 1, 26)
        col = chr(65 + ch) + col
    return col


def text_content(element: ET.Element) -> str:
    """Ignore the optional pronunciation runs in shared and inline strings."""
    direct = element.findtext(q("t"))
    return direct if direct is not None else "".join((r.findtext(q("t")) or "") for r in element.findall(q("r")))


@dataclass(frozen=True)
class Cell:
    coordinate: str
    value: str | int | float | bool | None
    data_type: str
    formula: str | None = None


@dataclass
class Sheet:
    title: str
    rows: dict[int, dict[int, Cell]] = field(default_factory=dict)

    @property
    def max_row(self) -> int:
        return max(self.rows, default=0)

    def __getitem__(self, address: str) -> Cell:
        row, column = address_parts(address)
        return self.rows.get(row, {}).get(column, Cell(address, None, "n"))

    def iter_rows(self) -> Iterator[tuple[Cell, ...]]:
        for cells in self.rows.values():
            yield tuple(cells.values())


class Workbook:
    def __init__(self, path: Path) -> None:
        self._archive = zipfile.ZipFile(path)
        try:
            self.parts = sheet_parts(self._archive)
            self._sheets: dict[str, Sheet] = {}
            try:
                raw = self._archive.read("xl/sharedStrings.xml")
            except KeyError:
                self._shared = []
            else:
                self._shared = [text_content(s) for s in ET.fromstring(raw).findall(q("si"))]
        except BaseException:
            self._archive.close()
            raise

    def __enter__(self) -> Workbook:
        return self

    def __exit__(self, *args) -> None:
        self.close()

    def close(self) -> None:
        self._archive.close()

    def __iter__(self) -> Iterator[Sheet]:
        for name in self.parts:
            yield self[name]

    def __getitem__(self, name: str) -> Sheet:
        if name not in self._sheets:
            self._sheets[name] = self._read_sheet(name)
        return self._sheets[name]

    def _read_sheet(self, name: str) -> Sheet:
        return Sheet(name, dict(self.iter_sheet_rows(name)))

    def iter_sheet_rows(self, name: str) -> Iterator[tuple[int, dict[int, Cell]]]:
        """Read saved rows in order without retaining the entire worksheet."""
        with self._archive.open(self.parts[name]) as stream:
            for _, element in ET.iterparse(stream, events=("end",)):
                if element.tag != q("row"):
                    continue
                row: dict[int, Cell] = {}
                for cell in element.findall(q("c")):
                    address = cell.get("r")
                    kind = cell.get("t", "n")
                    raw = cell.findtext(q("v"))
                    formula = cell.find(q("f"))
                    if kind == "inlineStr":
                        value = text_content(cell.find(q("is")))
                    elif raw in (None, ""):
                        value = None
                    elif kind == "s":
                        value = self._shared[int(raw)]
                    elif kind in ("str", "e", "d"):
                        value = raw
                    elif kind == "b":
                        value = raw == "1"
                    else:
                        value = float(raw) if re.search(r"[.Ee]", raw) else int(raw)
                    if value is not None or formula is not None:
                        row[address_parts(address)[1]] = Cell(address, value, kind, (formula.text or "") if formula is not None else None)
                if row:
                    yield int(element.get("r")), row
                element.clear()


def scan_errors(wb) -> None:
    failures = []
    count = 0
    for sheet in wb:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.data_type == "e" or (isinstance(cell.value, str) and cell.value in ERRS):
                    count += 1
                    if len(failures) < 12:
                        failures.append(f"{sheet.title}!{cell.coordinate}={cell.value}")
    if failures:
        raise AssertionError(f"Found {count} spreadsheet error cells: {', '.join(failures)}")

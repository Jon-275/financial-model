"""Cover the chart age labels and table cells in the tagged example PDF."""

from contextlib import nullcontext
from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import patch

from pypdf import PdfWriter

from prepare_example_pdf import attach, data_table, element, kids, verify_age_axis


AGES = ('35', '50', '65', '80', '95')


class ExamplePdfTests(unittest.TestCase):
    def check_axis(self, labels: tuple[str, ...], *, wrapped: bool = False) -> None:
        """Supply the glyph geometry reported by the PDF parser."""
        glyphs = []
        for marker, label in enumerate(labels):
            for offset, digit in enumerate(label):
                glyphs.append({'text': digit, 'mcid': marker,
                               'top': 405 + (10 * offset if wrapped else 0),
                               'x0': 105 + 60 * marker + (0 if wrapped else 4 * offset)})
        page = SimpleNamespace(width=612, lines=[{'width': 400, 'height': 0, 'top': 400}], chars=glyphs)
        with patch('prepare_example_pdf.pdfplumber.open', return_value=nullcontext(SimpleNamespace(pages=[page]))):
            verify_age_axis(Path('example.pdf'), AGES)

    def test_readable_age_labels(self) -> None:
        """Accept whole labels at the expected chart ages."""
        self.check_axis(AGES)

    def test_dense_age_labels(self) -> None:
        """Reject the extra categories emitted by a dense chart."""
        with self.assertRaisesRegex(ValueError, 'unexpected age labels'):
            self.check_axis(('35', '36', *AGES[1:]))

    def test_wrapped_age_labels(self) -> None:
        """Reject whole label text when its digits are stacked vertically."""
        with self.assertRaisesRegex(ValueError, 'wrapped age label'):
            self.check_axis(AGES, wrapped=True)

    def test_missing_trailing_risk_table_cell(self) -> None:
        """Restore an empty unpaid amount without shifting existing cells."""
        writer = PdfWriter()
        page = writer.add_blank_page(width=612, height=792).indirect_reference
        for columns in (4, 5):
            with self.subTest(columns=columns):
                header = element(writer, '/TR', page)
                attach(header, [element(writer, '/TD', page).indirect_reference for _ in range(columns)])
                row = element(writer, '/TR', page)
                existing = [element(writer, '/TD', page).indirect_reference for _ in range(columns - 1)]
                attach(row, existing)

                table = data_table(writer, page, [header.indirect_reference, row.indirect_reference], columns)

                cells = kids(row)
                self.assertEqual(len(cells), columns)
                self.assertEqual(cells[:-1], existing)
                blank = cells[-1].get_object()
                self.assertEqual(blank['/S'], '/TD')
                self.assertEqual(blank.raw_get('/P'), row.indirect_reference)
                self.assertNotIn('/K', blank)
                self.assertEqual(row.raw_get('/P'), table)
                self.assertTrue(all(ref.get_object()['/S'] == '/TH' and
                                    ref.get_object()['/A']['/Scope'] == '/Column' for ref in kids(header)))


if __name__ == '__main__':
    unittest.main()

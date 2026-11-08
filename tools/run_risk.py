#!/usr/bin/env python3
"""Run the retirement workbook's funding policy against seeded market paths."""

import argparse
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
from tempfile import NamedTemporaryFile, TemporaryDirectory

from risk_engine.report import MAX_TRIALS, run_analysis
from risk_engine.scenario import load_scenario
from workbook_recalc import clear_formula_results, recalculate


def _load(source, cached, office_bin, timeout):
    if cached:
        scenario = load_scenario(source)
        return scenario, scenario.workbook_sha256, 'Saved workbook values'
    with TemporaryDirectory(prefix='retirement-risk-') as temp:
        temp = Path(temp)
        snapshot = temp / 'source.xlsx'
        data = source.read_bytes()
        source_hash = hashlib.sha256(data).hexdigest()
        snapshot.write_bytes(data)
        fresh, output = temp / 'recalculate.xlsx', temp / 'out'
        output.mkdir()
        clear_formula_results(snapshot, fresh)
        calculated = recalculate(fresh, output, temp / 'office-profile', office_bin, timeout)
        version = subprocess.run([office_bin, '--version'], capture_output=True, text=True,
                                 check=False, timeout=15)
        label = version.stdout.strip() if version.returncode == 0 else ''
        return load_scenario(calculated), source_hash, label or 'LibreOffice'


def _publish(report, output):
    output.parent.mkdir(parents=True, exist_ok=True)
    stage = None
    try:
        with NamedTemporaryFile(mode='w', encoding='utf-8', dir=output.parent, delete=False) as temp:
            stage = Path(temp.name)
            json.dump(report, temp, indent=2, allow_nan=False)
            temp.write('\n')
        os.link(stage, output)
    finally:
        if stage is not None:
            stage.unlink(missing_ok=True)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('workbook', type=Path, help='Phase 2 or later compatible XLSX file')
    parser.add_argument('--output', required=True, type=Path, help='New JSON report path')
    parser.add_argument('--trials', type=int, default=1000, help='Paired paths, default 1000')
    parser.add_argument('--seed', type=int, default=2026, help='Unsigned 64-bit seed, default 2026')
    parser.add_argument('--volatility', type=float, default=1.0,
                        help='Multiplier on asset volatility, default 1. Set 0 for no shocks.')
    parser.add_argument('--cached', action='store_true',
                        help='Use saved formula results only if they pass independent reconciliation')
    parser.add_argument('--soffice', default=os.environ.get('SOFFICE_BIN', 'soffice'),
                        help='LibreOffice executable')
    parser.add_argument('--timeout', type=int, default=240, help='LibreOffice timeout in seconds')
    args = parser.parse_args(argv)
    if not 1 <= args.trials <= MAX_TRIALS:
        parser.error(f'--trials must be from 1 to {MAX_TRIALS:,}.')
    if not 0 <= args.seed < 1 << 64:
        parser.error('--seed must be an unsigned 64-bit integer.')
    if not 0 <= args.volatility < float('inf'):
        parser.error('--volatility must be finite and nonnegative.')
    if args.timeout < 1:
        parser.error('--timeout must be at least 1 second.')
    if args.output.suffix.lower() != '.json' or args.output.exists():
        parser.error('--output must name a new .json file.')
    try:
        def progress(line):
            print(line, file=sys.stderr, flush=True)

        progress('Checking saved workbook values.' if args.cached else 'Recalculating a temporary workbook copy.')
        scenario, source_hash, calculation = _load(args.workbook, args.cached, args.soffice, args.timeout)
        report = run_analysis(scenario, trials=args.trials, seed=args.seed, volatility=args.volatility,
                              source_sha256=source_hash, workbook_calculation=calculation,
                              progress=progress)
        _publish(report, args.output)
    except (OSError, RuntimeError, ValueError, subprocess.TimeoutExpired) as exc:
        parser.exit(1, f'Risk analysis stopped. {exc}\n')
    print(args.output.resolve())


if __name__ == '__main__':
    main()

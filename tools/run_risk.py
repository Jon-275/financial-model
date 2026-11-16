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

import risk_workbook
from risk_engine.report import MAX_TRIALS, run_analysis
from risk_engine.scenario import load_scenario
from workbook_recalc import clear_formula_results, recalculate


def _load(source, cached, office_bin, timeout, work_dir):
    if cached:
        return load_scenario(source), source, 'Saved workbook values'
    fresh, output = work_dir / 'recalculate.xlsx', work_dir / 'out'
    output.mkdir()
    clear_formula_results(source, fresh)
    calculated = recalculate(fresh, output, work_dir / 'office-profile', office_bin, timeout)
    version = subprocess.run([office_bin, '--version'], capture_output=True, text=True,
                             check=False, timeout=15)
    label = version.stdout.strip() if version.returncode == 0 else ''
    return load_scenario(calculated), calculated, label or 'LibreOffice'


def _publish(data, output):
    output.parent.mkdir(parents=True, exist_ok=True)
    stage = None
    try:
        with NamedTemporaryFile(dir=output.parent, delete=False) as temp:
            stage = Path(temp.name)
            temp.write(data)
        os.link(stage, output)
    finally:
        if stage is not None:
            stage.unlink(missing_ok=True)


def _check_closed(source):
    locks = (source.with_name(f'~${source.name}'), source.with_name(f'.~lock.{source.name}#'))
    if any(path.exists() for path in locks):
        raise ValueError('Save and close the source workbook before refreshing risk.')


def _updated_workbook(source, calculated, scenario, report, work_dir, office_bin, timeout, workbook_name):
    enriched = work_dir / 'risk-package.xlsx'
    risk_workbook.write_snapshot(source, enriched, calculated, scenario, report, workbook_name=workbook_name)
    fresh, output = work_dir / 'recalculate-risk.xlsx', work_dir / 'risk-out'
    output.mkdir()
    clear_formula_results(enriched, fresh)
    recalculated = recalculate(fresh, output, work_dir / 'office-profile', office_bin, timeout)
    risk_workbook.verify_snapshot(recalculated, report)

    cached = work_dir / 'risk-final.xlsx'
    cache_script = Path(__file__).with_name('cache_model.py')
    result = subprocess.run([sys.executable, str(cache_script), str(enriched), str(recalculated), str(cached)],
                            capture_output=True, timeout=timeout, check=False)
    if result.returncode or not cached.is_file():
        raise ValueError('The refreshed workbook failed its formula or chart cache checks.')
    risk_workbook.verify_snapshot(cached, report)
    return cached.read_bytes()


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('workbook', type=Path, help='Saved retirement workbook (.xlsx)')
    parser.add_argument('--output', required=True, type=Path, help='New .xlsx workbook or .json report path')
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
    suffix = args.output.suffix.lower()
    if suffix not in ('.xlsx', '.json') or args.output.exists():
        parser.error('--output must name a new .xlsx or .json file.')
    if suffix == '.xlsx' and args.workbook.suffix.lower() != '.xlsx':
        parser.error('Refreshing an .xlsx workbook requires an .xlsx source file.')
    if suffix == '.xlsx' and args.cached:
        parser.error('Refreshing an .xlsx workbook requires a fresh calculation. Omit --cached.')
    try:
        def progress(line):
            print(line, file=sys.stderr, flush=True)

        workbook = args.workbook.resolve()
        if suffix == '.xlsx':
            _check_closed(workbook)
        with TemporaryDirectory(prefix='retirement-risk-') as directory:
            work_dir = Path(directory)
            source = work_dir / 'source.xlsx'
            data = workbook.read_bytes()
            source_hash = hashlib.sha256(data).hexdigest()
            source.write_bytes(data)
            if suffix == '.xlsx':
                risk_workbook.validate_source(source)
            progress('Checking saved workbook values.' if args.cached else 'Recalculating a temporary workbook copy.')
            scenario, calculated, calculation = _load(source, args.cached, args.soffice, args.timeout, work_dir)
            report = run_analysis(scenario, trials=args.trials, seed=args.seed, volatility=args.volatility,
                                  source_sha256=source_hash, workbook_calculation=calculation,
                                  progress=progress)
            if suffix == '.xlsx':
                progress('Updating and checking the risk results in Excel format.')
                data = _updated_workbook(source, calculated, scenario, report, work_dir, args.soffice,
                                         args.timeout, args.output.name)
                _check_closed(workbook)
                if hashlib.sha256(workbook.read_bytes()).hexdigest() != source_hash:
                    raise ValueError('The source workbook changed during the run. Refresh again from the latest copy.')
            else:
                data = (json.dumps(report, indent=2, allow_nan=False) + '\n').encode('utf-8')
            _publish(data, args.output)
    except (OSError, RuntimeError, ValueError, subprocess.TimeoutExpired) as exc:
        parser.exit(1, f'Risk analysis stopped. {exc}\n')
    print(args.output.resolve())


if __name__ == '__main__':
    main()

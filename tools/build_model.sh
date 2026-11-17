#!/usr/bin/env bash
set -euo pipefail

if [[ $# -ne 2 ]]; then
  printf 'Usage: %s BASELINE.xlsx OUTPUT.xlsm\n' "$0" >&2
  exit 2
fi

tool_dir=$(cd -- "$(dirname -- "$0")" && pwd)
input_dir=$(cd -- "$(dirname -- "$1")" && pwd)
input="$input_dir/$(basename -- "$1")"
output=$2
[[ "$output" == *.xlsm ]] || { printf 'Choose an .xlsm output for the workbook with the market risk button.\n' >&2; exit 2; }
[[ -f "$input" && ! -e "$output" ]] || { printf 'Use an existing baseline and a new output path.\n' >&2; exit 2; }
python_bin=${PYTHON_BIN:-python3}
office_bin=${SOFFICE_BIN:-soffice}
temp_dir=$(mktemp -d "${TMPDIR:-/tmp}/retirement-model.XXXXXXXX")
trap 'rm -rf -- "$temp_dir"' EXIT
mkdir -- "$temp_dir/draft"

"$python_bin" "$tool_dir/model_plan.py" "$input" "$temp_dir/plan.json"
"$python_bin" "$tool_dir/author_model.py" "$temp_dir/plan.json" "$temp_dir/authored.xlsx"
"$python_bin" "$tool_dir/package_model.py" "$input" "$temp_dir/plan.json" "$temp_dir/authored.xlsx" "$temp_dir/draft/retirement-model.xlsx"
"$python_bin" "$tool_dir/native_bootstrap.py" "$temp_dir/draft/retirement-model.xlsx" \
  "$temp_dir/example/retirement-model.xlsx" --soffice "$office_bin"
"$python_bin" "$tool_dir/../tests/model_acceptance.py" --release "$temp_dir/example/retirement-model.xlsx"
mkdir -p -- "$(dirname -- "$output")"
"$python_bin" "$tool_dir/package_macro.py" "$temp_dir/example/retirement-model.xlsx" "$output"

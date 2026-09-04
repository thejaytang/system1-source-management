#!/bin/zsh
set -u
script_dir="${0:A:h}"
program_dir="${script_dir}/.."
cd "${program_dir}"
if [[ ! -x .venv/bin/python ]]; then
  echo "Initial setup has not been completed. Run Code/deployment/setup_macos.command first."
  read -k 1 "?Press any key to close..."
  exit 2
fi
.venv/bin/python src/source_updater.py --config config/config.json --audit-only
result=$?
echo
if [[ $result -eq 0 ]]; then
  echo "Consistency audit passed."
else
  echo "Consistency audit found an issue. Correct it before the next update cycle."
fi
read -k 1 "?Press any key to close..."
exit $result

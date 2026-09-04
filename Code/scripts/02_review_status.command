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
.venv/bin/python src/leader_orchestrator.py --config config/config.json --force-all
result=$?
echo
if [[ $result -eq 0 ]]; then
  echo "Review cycle completed. Open the Excel Dashboard and the latest report under Code/runtime/logs/leader_reports."
else
  echo "The review cycle failed before producing a complete report. Check the terminal output and runtime logs."
fi
read -k 1 "?Press any key to close..."
exit $result

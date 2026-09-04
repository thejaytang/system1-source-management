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
.venv/bin/python src/manual_intake.py --config config/config.json
result=$?
if [[ $result -eq 0 ]]; then
  .venv/bin/python src/human_operations.py --config config/config.json
  result=$?
fi
echo
if [[ $result -eq 0 ]]; then
  echo "Manual intake scanned. Open Human Operation Desktop to complete and review the candidates."
else
  echo "Manual intake processing failed. Check the terminal output and leave the original files in place."
fi
read -k 1 "?Press any key to close..."
exit $result

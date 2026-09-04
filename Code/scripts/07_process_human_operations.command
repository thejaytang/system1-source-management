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
.venv/bin/python src/human_operations.py --config config/config.json
result=$?
echo
if [[ $result -eq 0 ]]; then
  echo "Human operations processed. Reopen Excel and review Human Operation Desktop and Dashboard."
else
  echo "Human-operation processing failed. Check the visible program note and terminal output."
fi
read -k 1 "?Press any key to close..."
exit $result

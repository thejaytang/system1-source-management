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
.venv/bin/python -m unittest discover -s tests -v
result=$?
echo
if [[ $result -eq 0 ]]; then
  echo "All automated scenario tests passed."
else
  echo "Tests failed. Do not enable scheduled runs."
fi
read -k 1 "?Press any key to close..."
exit $result

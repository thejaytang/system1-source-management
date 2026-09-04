#!/bin/zsh
set -u
script_dir="${0:A:h}"
code_dir="${script_dir}/.."
cd "${code_dir}"
if [[ ! -x .venv/bin/python ]]; then
  echo "The project environment is missing. Run Code/deployment/setup_macos.command first."
  read -k 1 "?Press any key to close..."
  exit 2
fi
PYTHONPATH=src .venv/bin/python -m system1 menu --config config/config.json --schedule config/schedule.json
result=$?
echo
read -k 1 "?Press any key to close..."
exit $result


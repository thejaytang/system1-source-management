#!/bin/zsh
set -eu
script_dir="${0:A:h}"
code_dir="${script_dir}/.."
cd "${code_dir}"
PYTHONPATH=src .venv/bin/python -m system1 doctor --config config/config.json --schedule config/schedule.json
PYTHONPATH=src .venv/bin/python -m system1 register-schedule --config config/config.json --schedule config/schedule.json
echo "Schedule registered for the current macOS user."
read -k 1 "?Press any key to close..."


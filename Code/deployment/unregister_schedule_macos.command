#!/bin/zsh
set -eu
script_dir="${0:A:h}"
code_dir="${script_dir}/.."
cd "${code_dir}"
PYTHONPATH=src .venv/bin/python -m system1 unregister-schedule
echo "Schedule unregistered for the current macOS user."
read -k 1 "?Press any key to close..."


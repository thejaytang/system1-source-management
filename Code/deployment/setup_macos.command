#!/bin/zsh
set -eu
script_dir="${0:A:h}"
code_dir="${script_dir}/.."
cd "${code_dir}"
if ! command -v python3 >/dev/null 2>&1; then
  echo "Python 3.11 or newer is required. Install Python, then run this setup again."
  read -k 1 "?Press any key to close..."
  exit 2
fi
python_version="$(python3 -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")')"
if ! python3 -c 'import sys; raise SystemExit(0 if sys.version_info >= (3, 11) else 1)'; then
  echo "Python ${python_version} is too old. Python 3.11 or newer is required."
  read -k 1 "?Press any key to close..."
  exit 2
fi
python3 -m venv --prompt SmarterComplianceSystem1 .venv
.venv/bin/python -m pip install --upgrade pip
.venv/bin/python -m pip install -r deployment/requirements.txt
[[ -f config/config.json ]] || cp config/config.example.json config/config.json
[[ -f config/schedule.json ]] || cp config/schedule.example.json config/schedule.json
PYTHONPATH=src .venv/bin/python -m system1 doctor --config config/config.json --schedule config/schedule.json || true
echo
echo "Environment created. Review Code/config/config.json and schedule.json, then run the doctor again."
read -k 1 "?Press any key to close..."

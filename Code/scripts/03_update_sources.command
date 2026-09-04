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
echo "Save and close Requirement_Source_Registry.xlsx before continuing."
read -k 1 "?Press any key to start the controlled update..."
echo
.venv/bin/python src/leader_orchestrator.py --config config/config.json --force-all --execute
result=$?
echo
if [[ $result -eq 0 ]]; then
  echo "All collectable current sources were checked. Reopen Excel and review the Dashboard."
else
  echo "Update cycle stopped. The current files were preserved; check runtime logs and the Excel failure fields."
fi
read -k 1 "?Press any key to close..."
exit $result

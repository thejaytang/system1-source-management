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
read "source_id?Enter the source_id to retry, for example PA010: "
source_id="${source_id:u}"
if [[ ! "${source_id}" =~ '^[A-Z]{2}[0-9]{3}$' ]]; then
  echo "Invalid source_id. Expected two letters followed by three digits."
  read -k 1 "?Press any key to close..."
  exit 2
fi
.venv/bin/python src/source_updater.py --config config/config.json --source-id "${source_id}" --include-pending
result=$?
echo
if [[ $result -eq 0 ]]; then
  echo "Retry completed. Reopen Excel and review the source row and Dashboard."
else
  echo "Retry did not produce a valid download. The previous valid file, if any, was preserved."
fi
read -k 1 "?Press any key to close..."
exit $result

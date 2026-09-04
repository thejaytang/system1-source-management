# Data

`00_Human_Intake` is the operator drop folder for authorised PDF, HTML, XLSX or ZIP files that are not yet registered. After adding files, save and close Excel, use the platform launcher in the System1 root, and choose `Routine Cycle`. The program proposes metadata in `Human Operation Desktop`; it does not create a registered source until a human chooses `ACCEPT`.

This directory contains the current original-format source files managed by the Requirement Source Management System. Current valid files are stored directly in the `folder_code` directory mapped from `source_family`. Every filename must begin with `snapshot_id_`.

The folder represents the organisation responsible for the content, not the channel through which the team obtained the copy. For example, an OEM manual supplied by a client still belongs in the manufacturer / supplier folder; the acquisition channel is recorded in `acquisition_channel`.

When content changes, the updater archives the previous file under:

`<folder_code>/_archive/<source_id>/`

Preserve the original published format. Supported formats are `.pdf`, `.html`, `.xlsx` and `.zip`. Do not convert HTML into PDF for convenience.

For a manual replacement, place the original file in `00_Human_Intake`, run `Routine Cycle` from the platform launcher, and complete the generated `MANUAL_FILE_REPLACEMENT` row in `Human Operation Desktop`. The program updates the register, snapshot sequence, stored filename, provenance and audit timestamps after `ACCEPT`; do not edit `Source Register` directly. Numbered scripts under `Code/scripts` are advanced diagnostic and compatibility entry points, not the normal operator path.

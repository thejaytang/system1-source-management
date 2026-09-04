# Requirement Workstream Workspace | System1 Source Management

This folder is the portable, shareable working area for the Requirement Workstream. It contains the current System1 source-management implementation, its human-governed workbook, controlled source files, platform launchers, and supporting presentation material.

Start here according to your role:

- Human operators and recipients: read this `README.md`.
- Developers and technical reviewers: continue with `Code/README.md`.
- Codex or another coding agent: read `AGENTS.md` before inspecting or changing the workspace.
- Agents working specifically inside `Code`: also follow `Code/AGENTS.md`.

## Purpose

System1 maintains the controlled source library used to identify Requirements. It records source identity, original-format files, retrieval state, selection decisions, human review, and monthly quality checks.

System1 does not provide legal interpretation, approve client applicability, bypass paywalls, or accept newly discovered sources without a human decision.

System2 Requirement extraction is not delivered as an integrated runtime in this folder. Material under `presentation` and `Others` may describe a retained or future System2 technical design and must not be treated as evidence that the complete pipeline is operational.

## Data flow and outputs

```text
Official sources / registered retrieval targets / authorised manual files
                               +
            Human decisions in Human Operation Desktop
                               |
                               v
                     System1 Leader
                               |
          +--------------------+--------------------+
          |                    |                    |
 Discovery & Intake   Retrieval & Monitoring   Governance & QA
          +--------------------+--------------------+
                               |
                               v
     Workbook state + controlled source library + human task queue
                  + backups, logs and Leader reports
```

The primary business output is a maintained, traceable source package. It includes governed source identity, original-format files, version and hash evidence, retrieval state, selection state, and human-review status. It is the controlled input contract for future Requirement extraction, not a final Requirement dataset by itself.

## Folder structure

```text
05_Working area of requirements side/
├── Requirement_Source_Registry.xlsx
├── README.md
├── AGENTS.md
├── Data/
│   ├── 00_Human_Intake/
│   ├── A_Public_Authority/
│   ├── B_Standards_Body/
│   ├── C_Certification_Scheme/
│   ├── D_Manufacturer_Supplier/
│   ├── E_Project_Engineering/
│   ├── F_Client_Internal/
│   └── Z_Pending_Classification/
├── Code/
├── presentation/
├── Others/
├── Run System - macOS.command
└── Run System - Windows.cmd
```

Operational rule:

- `05_Working area of requirements side/Requirement_Source_Registry.xlsx` is the only workbook that System1 reads and writes.
- `Code/runtime/backups/` and `Code/runtime/dashboard_staging_archive_20260904/` are historical artifacts only and are not operational sources.
- Do not use `Others/Legacy_Requirement_Workbook.xlsx` or other historical files as current System1 input.

Keep the workbook, `Data`, and `Code` together. All shared paths are relative to this folder.

`presentation` contains the current shareable architecture visual. `Others` contains supporting material and historical presentation variants; verify dates and labels before reusing them as current-state evidence.

## Requirements

- Python 3.11 or later, installed by the user
- Microsoft Excel 365 or Excel 2021
- Network access for dependency installation and source retrieval
- Read and write access to the complete System1 folder

System1 does not install Python, change operating-system security settings, provide an `.app` or `.exe`, or require an API key.

The project environment is local to `Code/.venv`. Each recipient should create or recreate it with the supplied setup script. Do not depend on another computer's copied `.venv`, because Python executables and installed paths are platform- and device-specific.

## First-time setup

### macOS

1. Install Python 3.11 or later.
2. Double-click `Code/deployment/setup_macos.command`.
3. Copy the configuration examples if the active files are absent.
4. Review both configuration files without adding device-specific shared paths.
5. Run `Run System - macOS.command` and choose `Validate Environment`.

### Windows

1. Install Python 3.11 or later and make either `py` or `python` available.
2. Double-click `Code\deployment\setup_windows.cmd`.
3. Copy the configuration examples if the active files are absent.
4. Review both configuration files without adding device-specific shared paths.
5. Run `Run System - Windows.cmd` and choose `Validate Environment`.

Do not enable unattended automation until the environment doctor passes.

## Workbook use

The workbook contains five sheets:

- `Instructions`: short operator guidance and recovery steps.
- `Categories`: controlled classification and dropdown values.
- `Source Register`: program-generated source state. Do not edit it directly.
- `Dashboard`: read-only, single-screen management overview for source coverage, selection, retrieval health, human workload, and QA. It is not an editing surface.
- `Human Operation Desktop`: the only normal human editing surface.

The default desktop shows ten fields only:

`source_id`, `source_title`, `task_type`, `issue_summary`, `requested_action`, `operator_action`, `operator`, `operator_note`, `program_status`, and `program_result`.

Correction, candidate, and program-support fields are hidden by default. Unhide correction or candidate fields only when `requested_action` or `program_result` asks for them. Do not edit program-support fields.

The Dashboard summarizes the current registry through formula-backed KPI cards, source-portfolio and selection charts, retrieval exceptions, governance and QA status, and the five highest-priority open tasks. Priority items are ordered by `NEEDS_REPLAN`, then `WAITING_FOR_HUMAN`, then `PENDING`; older tasks are shown first within the same status. Complete decisions only in `Human Operation Desktop`.

## Human operation workflow

1. Open `Human Operation Desktop`.
2. Inspect every visible task.
3. Complete requested correction fields only when required.
4. Choose `operator_action` from the row-specific dropdown.
5. Enter `operator`. Add `operator_note` when returning, rejecting, or reporting an incorrect QA result.
6. Save and close Excel.
7. Run `Routine Cycle` from the platform launcher.
8. Reopen the workbook and inspect `program_status` and `program_result`.

`checked_at` and `program_operated_at` are filled by the program. Completed tasks remain in the audit trail and are hidden. Before applying an old task, the program checks the source fingerprint and refuses to overwrite newer source data.

### Supported human cases

- `Existing source review`: inspect the requested change, then choose `APPLY` or `RETURN`.
- `Selection review`: complete `operator_selection_decision` as `INCLUDE`, `PENDING`, or `EXCLUDE`, then choose `APPLY` or `RETURN`.
- `New source candidate`: complete missing candidate fields, then choose `ACCEPT` or `REJECT`. A new source never enters the register automatically.
- `Manual file intake`: place an authorised original file in `Data/00_Human_Intake`; Routine Cycle creates the appropriate human task.
- `Manual replacement/paywall`: provide only an authorised copy. A failed replacement preserves the current valid file.
- `Random QA`: inspect the sampled record and file, then choose `CORRECT` or `INCORRECT`.

## Source, retrieval, and selection rules

- Prefer a complete official HTML source over an equivalent official PDF.
- Preserve the original retrieved format. Never convert HTML to PDF or PDF to HTML for storage.
- Prefer the authoritative language. Among equally authoritative official versions, prefer English, then Norwegian, then Other.
- `official_url` identifies the official source or landing page.
- `retrieval_url` points to the actual collected HTML, PDF, XLSX, or ZIP resource.
- `SUCCESS`, `FAIL`, and `PAYWALL_BLOCKED` describe retrieval. System1 never bypasses a paywall.
- Retrieval and selection are independent. System1 may attempt to collect any collectable `CURRENT` source.
- `INCLUDE`, `PENDING`, and `EXCLUDE` control formal source-list membership. `INCLUDE` requires a valid current snapshot and completed governance gates.

## Running System1

Use the root launcher for normal work:

1. `Routine Cycle`: processes human decisions, scans manual intake, merges open tasks, performs consistency checks, creates due monthly QA, and updates the workbook. It does not perform a full download sweep.
2. `Full Source Check`: performs Routine Cycle plus retrieval and validation for every collectable `CURRENT` source.
3. `Status Only`: displays the latest Leader report without changing the workbook.
4. `Validate Environment`: runs the environment doctor.
5. `View Last Run`: displays the latest run result.

Always save and close Excel before a write-enabled run.

## Optional automation

Automation is opt-in. The template defaults to Routine Cycle Monday to Friday at 09:00, Full Source Check Sunday at 09:00, timezone `Europe/Oslo`, and catch-up enabled.

To enable it:

1. Review `Code/config/schedule.json`.
2. Set `enabled` to `true`.
3. Run the appropriate registration script in `Code/deployment`.

The registered runner checks whether a job is due and exits when it is not. Missed runs and month-end QA are caught up. If Excel is open or another System1 process is running, the run is deferred without writing.

Moving the System1 folder invalidates the device-specific scheduler record. Unregister the old schedule and register it again. Without registration, System1 remains completely manual.

## Safety and recovery

- System1 creates a backup before controlled workbook updates.
- Workbook updates use a temporary file and validated replacement.
- Excel-open and process locks prevent concurrent writers.
- Failed downloads and failed manual replacements preserve the current valid source file.
- Logs and backups are stored under `Code/runtime`.

If Excel reports workbook damage, choose `No` before automatic repair, keep Excel closed, restore the latest backup from `Code/runtime/backups`, and run `Validate Environment` before retrying.

For `DEFERRED_WORKBOOK_OPEN`, close Excel and run again. For `DEFERRED_ALREADY_RUNNING`, wait for the current run and use `View Last Run`. For `FAIL` or `PAYWALL_BLOCKED`, use `Human Operation Desktop` to provide a corrected route or authorised manual file.

## Sharing this workspace

Share the complete `05_Working area of requirements side` folder so that the workbook, `Data`, `Code`, launchers, and relative paths remain together.

Before sharing:

1. Close Excel and confirm no System1 process is running.
2. Do not include customer-confidential material from outside this workspace.
3. Treat `Code/runtime/logs` and `Code/runtime/backups` as operational evidence; review whether the recipient needs them before distribution.
4. The recipient should create a fresh project-local environment with the appropriate setup script and run `Validate Environment`.
5. Keep automation disabled until the recipient has reviewed `Code/config/schedule.json`, passed Environment Doctor, and explicitly chosen to register it on that computer.
6. Treat time-specific Dashboard values and historical diagrams as snapshots, not permanent project claims.

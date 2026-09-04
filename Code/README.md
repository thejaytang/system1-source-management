# System1 Code | Engineering Guide

## Architecture

System1 uses one deterministic Leader and three bounded modules:

1. `Discovery & Intake`
2. `Retrieval & Monitoring`
3. `Governance & QA`

The Leader owns orchestration, the complete-run lock, persisted state, task consolidation, retry handoff, notification, and run reporting. Human-facing work is consolidated to one current task per source.

No search provider or AI API is configured. Discovery therefore processes `Data/00_Human_Intake`, the structured candidate inbox, duplicate checks, and coverage checks only. It must not claim autonomous regulatory discovery.

An optional provider boundary exists in `src/system1/provider_extensions.py`. A future search provider may propose candidates and an assessment provider may propose human-governed field updates. Provider output has no direct write authority: candidates and suggestions are validated and routed to `Human Operation Desktop`. The shipped configuration still runs in `NO_API` mode and includes no provider adapter, credentials, or external API call.

## Engineering structure

```text
Code/
├── config/
│   ├── config.example.json
│   ├── config.json
│   ├── schedule.example.json
│   └── schedule.json
├── deployment/
│   ├── requirements.txt
│   ├── register_schedule_macos.command
│   ├── register_schedule_windows.cmd
│   ├── setup_macos.command
│   ├── setup_windows.cmd
│   ├── unregister_schedule_macos.command
│   └── unregister_schedule_windows.cmd
├── scripts/
│   ├── run_macos.command
│   └── run_windows.cmd
├── src/
│   ├── system1/
│   │   ├── provider_extensions.py
│   │   └── ...
│   ├── leader_orchestrator.py
│   ├── human_operations.py
│   ├── manual_intake.py
│   ├── source_updater.py
│   ├── staged_pipeline.py
│   └── sync_run_state.py
├── tests/
├── runtime/
├── AGENTS.md
└── README.md
```

The retained numbered scripts are advanced compatibility and diagnostic entry points. Normal users should use the root platform launcher.

## Environment contract

- Python 3.11 or later
- `openpyxl==3.1.5`
- `portalocker==4.3.0`
- `tzdata==2026.3`

Setup scripts create the project-local `Code/.venv`, displayed as `SmarterComplianceSystem1`, and install only pinned Python dependencies there. They do not install project packages into the system or default Python environment, install Python itself, alter security settings, register a schedule, or add credentials.

```bash
# macOS
./deployment/setup_macos.command
PYTHONPATH=src .venv/bin/python -m system1 doctor
```

```bat
rem Windows
deployment\setup_windows.cmd
set PYTHONPATH=src
.venv\Scripts\python.exe -m system1 doctor
```

## CLI

Run from `Code` with `PYTHONPATH=src`:

```text
python -m system1 menu
python -m system1 routine
python -m system1 full
python -m system1 status
python -m system1 doctor
python -m system1 scheduled
python -m system1 register-schedule
python -m system1 unregister-schedule
```

`routine` uses `execute_collection = false`. `full` forces all modules and performs collection. `scheduled` checks persisted due keys and runs at most one due job, with Full Source Check taking precedence.

## Configuration

`config.json` contains only shared relative paths and portable operational settings: workbook and `Data` paths, runtime paths, timeout and byte limits, timezone, HTML-first format preference, authoritative-first language policy, eligible source states, folder prefixes, Random QA, and module intervals.

`schedule.json` contains schedule intent: `enabled`, routine and full-check times, timezone, polling interval, catch-up behavior, and notification preference.

Device-specific absolute paths exist only in the local `launchd` or Windows Task Scheduler record. They are not written to the workbook or shared source code.

## Scheduling adapters

macOS registration creates a current-user LaunchAgent. Windows registration creates a current-user Task Scheduler task. Both launch a lightweight periodic runner that exits when no job is due, records successful due keys, catches up missed runs, avoids repeating a completed period, gives Full Source Check precedence, and defers on an open workbook or active System1 run.

Registration requires an existing `.venv` and `schedule.json` with `enabled = true`. Moving the folder requires unregistering and registering again.

## Module responsibilities

### Leader

Runs the human cycle, wakes due modules, re-reads findings, consolidates one current task per source, persists state, writes one JSON report, and produces a user-facing status. `REVIEW` is a completed business outcome, not a process crash.

### Discovery & Intake

Scans manual intake and the structured candidate inbox, checks duplicate identities and coverage gaps, and sends every candidate to the human gate. It does not assign a formal source ID before `ACCEPT`.

### Retrieval & Monitoring

Evaluates collectable `CURRENT` records, prefers complete official HTML over equivalent official PDF, downloads original formats, validates content, compares hashes, archives changed snapshots, and preserves the current valid file after failure. Paywalls become `PAYWALL_BLOCKED` and are never bypassed.

### Governance & QA

Evaluates selection fields, audits workbook-file consistency, creates at most one monthly QA batch, catches up missed month-end creation, and creates correction tasks for incorrect QA results.

## Workbook contract

The first four sheets are generated views. `Human Operation Desktop` is the only normal human write surface. Its ten-field operator area is visible; corrections and candidate details are hidden unless required; program-support fields remain hidden and program-owned.

`human_operations.py` applies task-type-specific validation, fills timestamps, checks source fingerprints, processes decisions, refreshes the Dashboard's five priority items, preserves completed history, and maintains one current task per source. Dashboard priority follows `NEEDS_REPLAN`, `WAITING_FOR_HUMAN`, and `PENDING`, with the oldest task first inside each status.

## File and workbook safety

- `exclusive_process_lock` protects the complete Leader cycle.
- `system_lock` protects each workbook write and checks the Excel temporary lock.
- Controlled writes create a backup and use temporary-file replacement.
- Windows file occupation stops the run safely.
- Failure paths preserve current valid source files.
- Filename validation covers Windows reserved names, invalid characters, case-insensitive collisions, path length, and paths containing spaces.

## Environment doctor

`python -m system1 doctor` checks Python and dependency versions, configuration and relative paths, workbook and `Data`, Excel lock, timezone, runtime writability, Windows-compatible paths, schedule configuration, and scheduler registration state.

## Testing

The regression suite is local and does not perform a full network download:

```bash
PYTHONPATH=src .venv/bin/python -m unittest discover -s tests -v
```

Coverage includes selection, failure preservation, human decisions, stale fingerprints, consolidation, monthly QA, schedule catch-up and idempotence, Windows path rules, workbook contracts, and atomic safety behavior.

`tests/test_operator_journeys.py` adds end-to-end simulations for every supported operator journey, manual intake and replacement, paywall recovery, Random QA, Excel-open deferral, deterministic `NO_API` operation, and fake-provider `API_CONNECTED` success and failure boundaries. See `tests/OPERATOR_SCENARIO_MATRIX.md` for the acceptance mapping. Fake-provider tests do not claim that a real API adapter, credential, or external service has been validated.

Platform acceptance additionally requires fresh `.venv` tests on macOS and Windows, a path containing spaces, Excel-open and concurrent-run deferral, schedule registration/unregistration, and folder-move re-registration. A full network sweep is an explicit operational run, not a routine regression test.

## Runtime state

- `runtime/backups`: recovery workbooks
- `runtime/inbox`: structured candidate handoff
- `runtime/logs`: updater events, Leader reports, scheduler state, and scheduler output

Runtime files are operational evidence. Do not edit them to force a business result.

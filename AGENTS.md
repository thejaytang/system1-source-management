# Requirement Workstream Agent Contract

## Purpose and read order

This file governs agent work inside `05_Working area of requirements side` when the workspace is shared independently from the repository root.

Read in this order:

1. `README.md` for the human-facing purpose, setup, workflow, and sharing model.
2. `Code/README.md` for System1 architecture and engineering commands.
3. `Code/AGENTS.md` before changing anything under `Code`.
4. The current workbook, configuration, and runtime evidence before making a current-state claim.

## Scope and authority

- Treat this folder as the active Requirement Workstream workspace.
- Treat `Requirement_Source_Registry.xlsx` as the human-governed System1 business-state store.
- Treat `Source Register` as program-managed. Normal human changes belong in `Human Operation Desktop`.
- System1 governs source identity, retrieval, snapshots, selection, provenance, human review, and QA.
- System1 does not perform legal interpretation and does not represent a completed System2 Requirement-extraction pipeline.
- Treat `presentation` as the current shareable visual area. Treat variants under `Others` as supporting or historical until their dates and labels are verified.

## Current System1 model

System1 uses one deterministic Leader, three bounded modules, and one human gate:

1. `Discovery & Intake`
2. `Retrieval & Monitoring`
3. `Governance & QA`
4. `Human Operation Desktop` as the normal human decision and correction surface

Without a configured provider, do not claim autonomous regulatory discovery, semantic reasoning, or complete source coverage. API-connected capabilities are optional extensions and never replace Python validation or human authority.

## Data and output contract

Inputs may include registered official sources, `retrieval_url` targets, authorised files placed in `Data/00_Human_Intake`, structured candidates, configuration, schedule intent, and operator decisions.

System1 produces and maintains:

- governed workbook state;
- original-format source snapshots under `Data`;
- one current human-operation task per source where action remains unresolved;
- backups, locks, persisted state, logs, and Leader reports under `Code/runtime`;
- a maintained eligible source package for downstream Requirement extraction.

Do not describe this maintained source package as an already extracted Requirement dataset.

## Non-negotiable operating rules

- Resolve all shared paths relative to this workspace. Never write device-specific absolute paths into the workbook or shared configuration.
- Use only the project-local environment at `Code/.venv`. Never install project dependencies into the system or default Python environment.
- Keep `official_url` as source identity or landing page and `retrieval_url` as the actual retrievable resource.
- Keep retrieval independent of selection.
- Preserve authoritative language and original published format.
- Never bypass a paywall or silently substitute a convenience conversion for an authoritative original.
- Preserve the current valid snapshot after a failed retrieval or replacement.
- Require a named human decision before accepting a new candidate or applying a governed judgment.
- Do not write while Excel is open or another System1 process holds the lock.
- Preserve backups, logs, source snapshots, failure evidence, human-review history, and Random QA history.
- Do not add API keys, credentials, customer-confidential content, or device-specific scheduler records to shared files.
- Do not commit, push, publish, email, enable automation, or run a Full Source Check without explicit user authorization.

## Working method and checkpoints

Use adaptive, evidence-based work:

1. Inspect the relevant source, workbook structure, configuration, code, or runtime evidence.
2. Record the observed symptom and distinguish it from plausible root causes.
3. Run the smallest check that can discriminate between those explanations.
4. Modify only after diagnosis.
5. Validate in proportion to the change.
6. At the checkpoint, choose `CONTINUE`, `ADJUST`, `BACKTRACK`, `PIVOT`, or `STOP`.

Validation expectations:

- Documentation-only change: verify file paths, commands, terminology, and cross-file consistency.
- Local code change: run targeted tests.
- Cross-module change: run the complete local regression suite.
- Environment or schedule change: run Environment Doctor before enabling or registering automation.
- Workbook or source-state change: verify workbook integrity, Leader/workbook task agreement, retained completed history, IDs, filenames, hashes, and current files.
- Full Source Check: treat as a network-enabled operational run requiring explicit intent, not as a normal test.

## Sharing and handoff

- Keep `README.md` human-facing and `AGENTS.md` agent-facing. Do not merge implementation-only instructions into the operator workflow.
- Do not assume a copied `Code/.venv` is portable. Recipients should recreate the project-local environment with the supplied setup script.
- Before reporting completion, state what was verified, what remains time-specific or unresolved, and whether external acceptance is still required.
- Historical reports or diagrams may retain obsolete architecture or metrics. Label them as historical rather than rewriting audit evidence.

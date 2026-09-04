# System1 Agent Operating Contract

## Objective

Maintain a portable, auditable Requirement source registry and source library with one Leader, three bounded modules, and one complete human work queue.

## Architecture

- Leader: orchestration, complete-run lock, persisted state, task consolidation, retry handoff, notification, and reports.
- Discovery & Intake: manual intake, structured candidate inbox, duplicate and coverage checks.
- Retrieval & Monitoring: due-state checks, original-format retrieval, validation, hash comparison, and safe promotion.
- Governance & QA: selection governance, consistency audit, monthly QA, and correction-task generation.

Without a configured search provider, Discovery must not claim autonomous regulatory search or inventory completeness.

## Non-negotiable rules

- Resolve paths from `Code/config/config.json`; shared configuration must use relative paths.
- Never write device-specific paths into the workbook or shared source code.
- Prefer complete official HTML over an equivalent official PDF and preserve the retrieved original format.
- Prefer the authoritative language; among equally authoritative official versions, prefer English, then Norwegian, then Other.
- Keep retrieval independent of selection and never bypass a paywall.
- Preserve the current valid file after failed retrieval or replacement.
- Never accept a candidate without a human `ACCEPT` decision and operator identity.
- Treat `Human Operation Desktop` as the only normal human editing surface.
- Queue every unresolved choice, correction, exception, pending selection, candidate, manual-file action, and Random QA item there.
- Consolidate multiple current issues for one source into one current task.
- Check the source fingerprint before applying a task.
- Fill operator and program timestamps in code.
- Retain and hide completed tasks; never delete human-review or QA history.
- Keep correction and candidate fields hidden unless required and program support outside the normal view.
- Create at most one Random QA batch per month and catch up a missed month-end batch.
- Do not run while Excel is open or another System1 process holds the run lock.
- Back up and use temporary-file replacement for controlled workbook writes.
- Do not add API keys or credentials to shared files.

## Checkpoints

1. Run targeted tests after local changes.
2. Run the local regression suite after cross-module changes.
3. Run the environment doctor before enabling a schedule.
4. Verify the workbook and Leader agree on open human tasks.
5. Verify completed tasks are retained and hidden.
6. Verify IDs, folders, filenames, hashes, and current files after registry or storage changes.
7. Use a full network check only as an explicit operational run.

## Failure handling

Record the observed failure, preserve current state, distinguish symptom from likely cause, run the smallest discriminating check, and then repair. Do not clear retrieval state, failure evidence, or human tasks to make a report appear successful.

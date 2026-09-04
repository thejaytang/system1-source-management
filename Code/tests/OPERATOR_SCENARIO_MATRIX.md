# System1 Operator Scenario Matrix

This matrix defines the local, network-free acceptance coverage for the human-facing System1 workflow. All scenarios run against isolated temporary workbooks and `Data` folders. They do not modify the production registry or run a Full Source Check.

| Operator journey | Expected system outcome | Automated evidence |
|---|---|---|
| Existing source review, `APPLY`, with note | Governed result is written to the source record; operator and timestamps are recorded; completed task remains retained and hidden | `test_existing_source_review_applies_note_and_return_requires_replan` |
| Existing source review, `RETURN`, with note | Source is not overwritten; task becomes `NEEDS_REPLAN` and remains visible | `test_existing_source_review_applies_note_and_return_requires_replan` |
| Selection review, `PENDING` | A note is required; the human intent is recorded; effective selection remains `PENDING` | `test_selection_review_handles_pending_note_and_effective_include_gate` |
| Selection review, `INCLUDE` or `EXCLUDE` | Human intent is recorded; effective `INCLUDE` still requires valid retrieval and governance gates | `test_selection_review_can_record_exclude_then_restore_include_intent` |
| New candidate, `ACCEPT` | No formal source exists before acceptance; acceptance assigns the next ID and creates the register row | `test_new_source_candidate_accept_and_reject_are_human_gated` |
| New candidate, `REJECT` | No register row is created; decision remains in hidden audit history | `test_new_source_candidate_accept_and_reject_are_human_gated` |
| Human drops an authorised original file | Routine intake creates a candidate task; acceptance stores the original format, records hash/provenance, and archives the intake copy | `test_manual_file_intake_creates_candidate_then_preserves_original_html` |
| Manual replacement or paywall recovery, `REJECT` | Current valid file is preserved | `test_manual_replacement_reject_preserves_current_then_accept_archives_it` |
| Manual replacement or paywall recovery, `ACCEPT` | New snapshot is promoted; old current file is archived; status returns to `SUCCESS` | `test_manual_replacement_reject_preserves_current_then_accept_archives_it` |
| Random QA, `CORRECT` | Item-level QA result is retained and hidden after processing | `test_random_qa_correct_and_incorrect_retains_history_and_creates_followup` |
| Random QA, `INCORRECT` | QA result is retained and a visible correction task is created | `test_random_qa_correct_and_incorrect_retains_history_and_creates_followup` |
| Missing operator or stale task fingerprint | Program refuses the update and returns the row to `WAITING_FOR_HUMAN` | `test_missing_operator_and_stale_fingerprint_are_safely_returned` |
| Duplicate intake scan | The same file creates at most one candidate | `test_manual_intake_is_idempotent_and_excel_open_defers_without_writing` |
| Excel is open | Routine Cycle returns `DEFERRED_WORKBOOK_OPEN` and workbook bytes remain unchanged | `test_manual_intake_is_idempotent_and_excel_open_defers_without_writing` |
| No API configured | Deterministic Leader runs all local modules without calling the downloader during Routine Cycle | `test_no_api_routine_cycle_is_deterministic_and_never_calls_downloader` |
| Simulated API returns a candidate and assessment suggestion | Both proposals enter `Human Operation Desktop`; neither changes `Source Register` before human approval | `test_simulated_api_proposals_are_human_gated_before_any_register_change` |
| Simulated API times out or proposes a program-owned field | Optional extension is `DEGRADED`; unsafe output is rejected; deterministic cycle completes and current state is preserved | `test_simulated_api_failure_degrades_extension_but_core_cycle_completes` |

The API-connected tests use in-memory fake providers. They validate the provider contract and governance boundary, not external connectivity, credentials, model quality, or vendor availability.

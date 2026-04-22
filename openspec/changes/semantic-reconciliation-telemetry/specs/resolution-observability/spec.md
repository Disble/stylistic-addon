# Resolution Observability Specification

## Purpose

Define how suggestion resolution MUST preserve semantic truth and emit observability data when Word host behavior is partial, late-invalidating, or otherwise non-atomic.

## Requirements

### Requirement: Semantic reconciliation after resolution failure

The system MUST re-observe a suggestion after execution, cleanup, or final-inspection failure before concluding a retryable error.

#### Scenario: Accept resolves semantically before cleanup fails

- GIVEN a replace suggestion whose tracked changes were fully accepted in Word
- WHEN comment cleanup or final inspection throws `ItemNotFound`
- THEN the result MUST remain terminal `accepted`
- AND the result MUST include a non-terminal warning about cleanup/inspection failure

#### Scenario: Accept only resolves one side before failure

- GIVEN a replace suggestion whose added side was accepted but deleted side remains pending
- WHEN execution fails before reconciliation proves full resolution
- THEN the result MUST remain retryable
- AND the taskpane MUST NOT discard the card

### Requirement: Resolution telemetry events

The system MUST emit structured telemetry events for resolution phases: `observe-before`, `execute`, `reconcile`, `cleanup`, and `inspect-after`.

#### Scenario: Resolution emits correlated phase evidence

- GIVEN one accept or reject attempt
- WHEN the workflow advances through its phases
- THEN every emitted event MUST include `workflowAttemptId`, `suggestionId`, `action`, `phase`, and `outcome`
- AND warning/error events MUST include enough metadata to identify the failing phase

### Requirement: Telemetry is non-interfering

The system MUST NOT let telemetry emission failures alter resolution semantics.

#### Scenario: Telemetry sink throws during reconciliation

- GIVEN a telemetry adapter failure during any resolution phase
- WHEN the suggestion is otherwise semantically resolved
- THEN the final result MUST remain unchanged
- AND the telemetry failure MUST be swallowed or downgraded to local warning behavior

### Requirement: Terminal warnings in the taskpane

The system MUST distinguish terminal semantic success with warnings from retryable failure.

#### Scenario: Taskpane keeps terminal accepted state after reconciled success

- GIVEN the adapter returns `accepted` plus cleanup warnings
- WHEN the taskpane updates the card
- THEN the card MUST become terminal/non-retryable
- AND the UI SHOULD surface the warning without restoring accept/reject actions

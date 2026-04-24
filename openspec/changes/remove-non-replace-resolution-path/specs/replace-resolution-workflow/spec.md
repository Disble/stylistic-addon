# Spec: Remove Non-Replace Resolution Path

## MODIFIED Requirements

### Requirement: Tracked-change resolution SHALL assume replace semantics

The adapter SHALL resolve every valid `track-change` suggestion through the replace workflow.

#### Scenario: Accept tracked-change enters replace workflow directly
- **GIVEN** a valid `track-change` suggestion with non-empty `anchor` and `suggestedText`
- **WHEN** the user accepts the suggestion
- **THEN** the adapter SHALL enter replace resolution directly
- **AND** it SHALL NOT branch into a separate tracked-change non-replace path

#### Scenario: Reject tracked-change enters replace workflow directly
- **GIVEN** a valid `track-change` suggestion with non-empty `anchor` and `suggestedText`
- **WHEN** the user rejects the suggestion
- **THEN** the adapter SHALL enter replace resolution directly
- **AND** it SHALL NOT branch into a separate tracked-change non-replace path

### Requirement: Invalid tracked-change suggestions SHALL fail fast

The adapter SHALL reject a `track-change` suggestion that violates the contract instead of attempting recovery.

#### Scenario: Missing suggested text fails fast
- **GIVEN** a suggestion with `type = "track-change"`
- **AND** `suggestedText` is empty or missing
- **WHEN** resolution starts
- **THEN** the adapter SHALL return an error result describing an invalid tracked-change contract
- **AND** it SHALL NOT attempt observation or execution recovery

#### Scenario: Missing anchor fails fast
- **GIVEN** a suggestion with `type = "track-change"`
- **AND** `anchor` is empty
- **WHEN** resolution starts
- **THEN** the adapter SHALL return an error result describing an invalid tracked-change contract
- **AND** it SHALL NOT attempt observation or execution recovery

## REMOVED Requirements

### Requirement: Non-replace tracked-change recovery path

(Reason: this path models an impossible product case under the current contract and hides invalid input as recoverable workflow variance.)

#### Scenario: Non-replace tracked-change retry branch is removed
- **GIVEN** resolution receives a `track-change` suggestion
- **WHEN** the adapter evaluates execution strategy
- **THEN** it SHALL choose only between replace workflow or fail-fast invalid-contract handling
- **AND** it SHALL NOT keep a separate non-replace tracked-change retry branch

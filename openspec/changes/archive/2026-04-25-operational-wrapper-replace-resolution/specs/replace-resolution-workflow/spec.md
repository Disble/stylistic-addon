# Delta for Replace Resolution Workflow

## ADDED Requirements

### Requirement: Operational Wrapper Identity Validation
The system SHALL use operational wrapper metadata as the explicit resolution identity. The system MUST NOT rely on fallback heuristics, full-body candidate matching, or duplicate-CC scoring.

#### Scenario: Ambiguous Location Abort
- GIVEN a suggestion resolution is triggered
- WHEN the location is ambiguous or lacks strict operational wrapper identity
- THEN the system SHALL abort with an `ambiguous-location` state
- AND the system SHALL NOT mutate the document or fall back to legacy discovery heuristics

### Requirement: Contiguous Suggestion Grouping
The system SHALL process contiguous or adjacent replace suggestions as explicit, unified groups for an all-or-nothing resolution.

#### Scenario: All-or-nothing group resolution
- GIVEN multiple contiguous replace suggestions forming a semantic block
- WHEN the user accepts or rejects the group
- THEN the system SHALL validate contiguous `acceptAll` or `rejectAll` for the entire cluster
- AND the system SHALL NOT use legacy candidate ranking for individual parts

### Requirement: Mixed Decisions Degradation
The system SHALL degrade unobservable or mixed states explicitly rather than attempting automated resolution.

#### Scenario: Mixed user decisions within a group
- GIVEN a contiguous suggestion group
- WHEN the user has made mixed decisions (e.g., accepts one part, rejects another) inside the group
- THEN the system SHALL explicitly degrade the resolution to an unobservable state
- AND the system SHALL NOT attempt to auto-resolve or route to a preserved legacy fallback

### Requirement: Comments-Only Cleanup
The system SHALL restrict post-resolution cleanup strictly to comment residues related to the replace operation, preserving wrapper and inserted-side CC metadata.

#### Scenario: Preserving unrelated changes
- GIVEN a resolved replace suggestion area with surrounding non-Stylistic tracked changes
- WHEN the cleanup policy executes
- THEN the system SHALL remove only Stylistic-inserted comments and related metadata
- AND the system SHALL leave non-Stylistic tracked changes entirely untouched

### Requirement: Mandatory No-Legacy Validation
The system SHALL NOT contain any dormant, reachable, or hidden legacy accept/reject workflow code. The implementation MUST demonstrably delete legacy paths rather than bypassing them.

#### Scenario: Static and runtime validation of legacy removal
- GIVEN the operational wrapper workflow is implemented
- WHEN the codebase and tests are reviewed
- THEN no legacy full-body fallback branches or compatibility shims SHALL remain
- AND tests asserting legacy behavior MUST be deleted or rewritten to assert the new wrapper model
- AND unsupported legacy states SHALL degrade explicitly rather than routing to a preserved fallback path

## REMOVED Requirements

### Requirement: Strategy extraction SHALL preserve the current workflow skeleton
(Reason: The current workflow skeleton is conceptually flawed and built on ad-hoc observation and duplicate-CC scoring. We are aggressively tearing out the legacy workflow in favor of a clean architectural replacement based on first-class operational lineage, treating legacy code as technical debt.)

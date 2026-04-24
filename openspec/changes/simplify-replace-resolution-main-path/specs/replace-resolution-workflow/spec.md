# Spec: Replace Resolution Workflow Simplification

## ADDED Requirements

### Requirement: Replace resolution SHALL use one primary semantic workflow

The adapter SHALL resolve replace suggestions through semantic side execution with fresh re-observation as the primary orchestration strategy.

#### Scenario: Accept replace follows semantic main path
- **GIVEN** a replace suggestion whose observation exposes a confirmed pending semantic pair
- **WHEN** the user accepts the suggestion
- **THEN** the workflow SHALL execute the semantic primary side first
- **AND** it SHALL re-observe the remaining side using fresh Word proxies
- **AND** it SHALL execute the remaining side before certifying success

#### Scenario: Reject replace follows semantic main path
- **GIVEN** a replace suggestion whose observation exposes a confirmed pending semantic pair
- **WHEN** the user rejects the suggestion
- **THEN** the workflow SHALL execute the semantic primary side first for reject semantics
- **AND** it SHALL re-observe the remaining side using fresh Word proxies
- **AND** it SHALL execute the remaining side before certifying success

### Requirement: Atomic batching SHALL be bounded to post-execute verification

The adapter SHALL keep at most one atomic retry after a fresh post-execute observation still exposes the full replace pair.

#### Scenario: Fresh post-execute pair triggers bounded atomic retry
- **GIVEN** semantic execution has completed
- **AND** a fresh post-execute observation still exposes the full replace pair for the same suggestion
- **WHEN** the adapter attempts recovery
- **THEN** it MAY queue one atomic retry
- **AND** it SHALL require fresh re-observation before certifying completion

#### Scenario: Reject does not gain a new early atomic workflow
- **GIVEN** a reject replace flow enters semantic execution
- **WHEN** pre-verification recovery is needed
- **THEN** the workflow SHALL stay within the retained semantic verification path
- **AND** this change SHALL NOT introduce a new early atomic retry path for reject

### Requirement: Early atomic fallback branches SHALL NOT remain in the workflow

The adapter SHALL NOT attempt atomic replace recovery before the post-execute verification gate.

#### Scenario: Added-side failure before post-execute verification
- **GIVEN** the primary semantic side cannot be certified before post-execute verification
- **WHEN** the workflow handles that failure
- **THEN** it SHALL continue through the retained semantic verification path or fail closed
- **AND** it SHALL NOT invoke an earlier atomic fallback branch dedicated to that pre-verification state

### Requirement: Tests SHALL protect only the retained workflow contract

The focused adapter suite SHALL include only tests directly related to the retained workflow contract.

#### Scenario: Removed fallback-specific tests
- **GIVEN** a test exists only to validate an early atomic fallback branch that no longer exists
- **WHEN** the refactor is applied
- **THEN** that test SHALL be removed from the focused suite
- **AND** retained tests SHALL still cover semantic execution, bounded post-execute retry, and fail-closed behavior

#### Scenario: Focused suites keep only representative retained-path coverage
- **GIVEN** the replace workflow still depends on specific observation surfaces and fresh proxy recovery
- **WHEN** the focused adapter suites are pruned
- **THEN** they SHALL retain coverage for semantic ordering, fresh re-observation, fail-closed verification, and the bounded post-execute atomic retry
- **AND** they SHALL keep representative coverage for retained evidence sources and multi-candidate `compound-v2` selection
- **AND** they SHALL omit unrelated comment-only, telemetry, and generic taskpane-state tests from this change scope


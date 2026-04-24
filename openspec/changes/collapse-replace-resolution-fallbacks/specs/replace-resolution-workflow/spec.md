# Spec: Collapse Replace Resolution Fallbacks

## MODIFIED Requirements

### Requirement: Replace resolution SHALL use one semantic workflow

The adapter SHALL resolve replace suggestions through semantic side execution with fresh re-observation, and SHALL NOT introduce alternate execution modes once the workflow has entered replace resolution.

#### Scenario: Accept replace follows one semantic path
- **GIVEN** a replace suggestion with a confirmed pending semantic pair
- **WHEN** the user accepts the suggestion
- **THEN** the adapter SHALL execute `Added` first
- **AND** it SHALL re-observe the remaining state before attempting `Deleted`
- **AND** it SHALL certify completion from fresh observation before cleanup

#### Scenario: Reject replace follows one semantic path
- **GIVEN** a replace suggestion with a confirmed pending semantic pair
- **WHEN** the user rejects the suggestion
- **THEN** the adapter SHALL execute `Deleted` first
- **AND** it SHALL re-observe the remaining state before attempting `Added`
- **AND** it SHALL certify completion from fresh observation before cleanup

### Requirement: Recovery SHALL use one fresh-proxy retry policy

The adapter SHALL recover uncertified mutation only by re-locating the suggestion, re-observing fresh Word proxies, retrying once, and then failing closed when certification still fails.

#### Scenario: Non-replace uncertified mutation retries once
- **GIVEN** a non-replace tracked-change resolution cannot be certified after `apply`
- **WHEN** the adapter attempts recovery
- **THEN** it SHALL re-locate and re-observe the suggestion
- **AND** it SHALL retry once with fresh proxies
- **AND** it SHALL fail closed if fresh observation still cannot certify completion

#### Scenario: Replace side uncertified mutation retries once
- **GIVEN** one semantic replace side cannot be certified after `apply`
- **WHEN** the adapter attempts recovery
- **THEN** it SHALL re-observe that side with fresh proxies
- **AND** it SHALL retry once
- **AND** it SHALL fail closed if the side remains pending after the final re-observation

## REMOVED Requirements

### Requirement: Post-execute atomic retry for replace

(Reason: this branch creates a second orchestration model and then depends on semantic recovery anyway.)

#### Scenario: Fresh post-execute full pair no longer triggers atomic batching
- **GIVEN** a replace workflow reaches post-execute observation
- **WHEN** the same suggestion still appears pending
- **THEN** the adapter SHALL certify through the retained semantic workflow only
- **AND** it SHALL NOT invoke `applyAtomically` as a fallback branch

### Requirement: Body-text silent-no-op recovery

(Reason: text matching escapes the observed suggestion identity boundary and weakens the `compound-v2` contract.)

#### Scenario: Silent no-op does not trigger body-text matching
- **GIVEN** a replace side mutation cannot be certified
- **WHEN** the adapter handles that uncertainty
- **THEN** it SHALL retry only through fresh suggestion re-observation
- **AND** it SHALL NOT search `body.getTrackedChanges()` by text to continue resolution

### Requirement: Same-click non-replace recovery branch

(Reason: non-replace recovery is absorbed into the shared fresh-proxy retry policy.)

#### Scenario: Non-replace no longer has dedicated same-click fallback choreography
- **GIVEN** a non-replace `apply` call fails partially or ambiguously
- **WHEN** the adapter attempts recovery
- **THEN** it SHALL use the shared re-locate and fresh-proxy retry policy
- **AND** it SHALL NOT keep a separate same-click fallback branch with distinct behavior

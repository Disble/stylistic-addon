# Spec: Extract Replace Resolution Strategy

## MODIFIED Requirements

### Requirement: Replace semantic policy SHALL come from one shared source

The adapter SHALL define replace action policy from one shared strategy source so `accept` and `reject` keep the same external behavior without duplicating semantic-order decisions across command and executor layers.

#### Scenario: Accept replace still resolves inserted side first
- **GIVEN** a replace suggestion with a confirmed pending semantic pair
- **WHEN** the user accepts the suggestion
- **THEN** the adapter SHALL execute `Added` before `Deleted`
- **AND** command and executor SHALL derive that order from the same replace policy source

#### Scenario: Reject replace still resolves deleted side first
- **GIVEN** a replace suggestion with a confirmed pending semantic pair
- **WHEN** the user rejects the suggestion
- **THEN** the adapter SHALL execute `Deleted` before `Added`
- **AND** command and executor SHALL derive that order from the same replace policy source

### Requirement: Strategy extraction SHALL preserve the current workflow skeleton

The adapter SHALL preserve the current locate, observe, execute, re-observe, and cleanup workflow while extracting only the action-specific replace policy.

#### Scenario: Replace strategy does not introduce a second workflow
- **GIVEN** the adapter resolves a replace suggestion
- **WHEN** the replace policy is consulted
- **THEN** the workflow SHALL still run through the existing command orchestration
- **AND** the strategy SHALL provide policy only, not alternate execution choreography

## REMOVED Requirements

### Requirement: Inline action branching for replace policy in multiple classes

(Reason: the same accept/reject policy is currently duplicated and should be centralized.)

#### Scenario: Command and executor no longer hardcode separate replace order rules
- **GIVEN** the codebase after the refactor
- **WHEN** replace semantic policy is needed
- **THEN** the adapter SHALL read it from the shared strategy contract
- **AND** it SHALL NOT keep separate hardcoded order rules in both command and executor

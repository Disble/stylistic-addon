# Apply Snapshot Specification

## Purpose

Define how batch suggestion application SHALL track the real Word document as it mutates, without re-reading the full document after every change.

## Requirements

### Requirement: Real-Word Apply Patch

The system MUST treat Word as the source of truth and SHALL update any local snapshot only after a real apply operation succeeds in Word.

#### Scenario: Successful apply returns a localized patch

- GIVEN a pending track-change suggestion
- WHEN the system applies it successfully in Word
- THEN the apply layer MUST return enough localized state to update the live snapshot
- AND it MUST NOT require a full-document reread

### Requirement: Incremental Snapshot Rebase

The system MUST maintain a live apply snapshot during a batch and SHALL rebase pending suggestions against the newest snapshot version after each successful apply.

#### Scenario: Pending suggestions shift after a nearby replacement

- GIVEN a batch with multiple pending suggestions
- WHEN one suggestion changes the text length in a paragraph
- THEN later suggestions MUST be rebased against the updated snapshot before their apply attempt

#### Scenario: Distant suggestions avoid full recomputation

- GIVEN a successful apply in one paragraph
- WHEN another pending suggestion is outside the affected zone
- THEN the system SHOULD update it through offset/delta adjustment without re-reading unrelated document regions

### Requirement: Hot Paragraph Re-read

The system MUST support localized rereads for affected paragraphs or equivalent local containers when nearby suggestions can no longer trust stale backend context.

#### Scenario: Same paragraph mutates repeatedly

- GIVEN multiple pending suggestions in the same paragraph
- WHEN one suggestion is applied
- THEN the system MUST re-read the affected paragraph or local container before applying the next nearby suggestion

### Requirement: Ordering Based on Real Position Signals

The system MUST NOT equate reverse backend array order with real document position.

#### Scenario: Backend order differs from document order

- GIVEN suggestions returned in an order that does not match their actual location in Word
- WHEN the batch execution order is prepared
- THEN the system MUST use real position signals from the live snapshot or an explicit ranking strategy
- AND it MUST NOT rely solely on `reverse()` as a proxy for document position

### Requirement: Heuristic Fallback Remains Secondary

The system MAY use global text-location heuristics only when snapshot rebase and localized rereads cannot confidently relocate a suggestion.

#### Scenario: Local rebase cannot confirm the target

- GIVEN a pending suggestion whose local paragraph rebase still cannot identify a safe target
- WHEN the system escalates
- THEN it MAY fall back to existing text-location heuristics
- AND it MUST preserve Word as the source of truth for the final apply decision

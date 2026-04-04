# Review Domain and Track Changes Lifecycle

This document records the current shared understanding of the **stylistic-addon frontend domain**, the **ubiquitous language** that should guide future refactors, and the **requirement change** for Track Changes lifecycle management.

It is intentionally detailed. The goal is not just to capture the final answer, but also to preserve the reasoning that prevents the team from falling back into accidental designs.

---

## 1. Why this document exists

We started from a concrete bug-shaped product problem:

- the add-in receives style suggestions from the backend,
- the frontend applies them as native Word tracked changes,
- the current implementation toggles `changeTrackingMode` per suggestion,
- Track Changes are global at the document level,
- Word UI flickers across instances because the mode is turned on/off repeatedly.

That problem exposed a deeper issue: the project had patterns, adapters, state machines, and ports, but it **did not yet have an explicit, shared domain definition for the frontend add-in itself**.

Without that shared domain definition, every architectural discussion drifts toward one of these traps:

1. treating the frontend as “just a client for the backend”,
2. treating an internal state (`DocumentReviewState`) as if it were the full domain,
3. treating Word implementation details (`changeTrackingMode`, content controls, comments) as if they were domain language,
4. treating workflow orchestration and domain concepts as the same thing.

This document fixes that.

---

## 2. Scope: what domain are we defining here?

This document defines the domain of the **frontend Word add-in** (`stylistic-addon`).

It does **not** define the domain of the backend.

That distinction matters.

The backend owns the editorial intelligence that decides:

- which style issues exist,
- which wording is preferable,
- what justification is returned,
- how a future debate/conversation is interpreted.

The frontend owns a different responsibility:

- presenting suggestions to the user,
- materializing them inside Word,
- letting the user accept, reject, or debate them,
- keeping the review state consistent inside the document,
- coordinating Word-specific behavior such as Track Changes lifecycle.

So, whenever this document says “domain”, it means:

> the domain of the **frontend add-in**, not the domain of the full product and not the domain of the backend.

---

## 3. Official domain definition

The official domain definition for the add-in is:

> **Presentar al usuario sugerencias de estilo provenientes del backend y permitirle aceptarlas, rechazarlas o debatirlas dentro de Word.**

This is the definition that should anchor future architecture, documentation, and naming decisions.

### Why this definition is correct

Because it is centered on the real user-facing responsibility of the add-in:

- **present** suggestions,
- **accept** suggestions,
- **reject** suggestions,
- **debate** suggestions.

This definition avoids several category mistakes:

- the domain is **not** “Track Changes”,
- the domain is **not** “the pipeline”,
- the domain is **not** “the document review state” alone,
- the domain is **not** “a Word client for a backend”.

Those are all important pieces, but they are **supporting mechanisms** or **internal models**, not the complete domain definition.

---

## 4. Core capabilities of the frontend domain

Inside that official domain definition, the add-in owns these core capabilities:

1. **Read text from Word**
   - resolve selection vs full document,
   - prepare the text that will be analyzed.

2. **Orchestrate the analysis process**
   - call the backend through the analysis port,
   - coordinate chunking, retries, and partial success behavior.

3. **Materialize suggestions inside Word**
   - convert backend output into document-visible artifacts,
   - apply tracked changes and comments where appropriate.

4. **Manage the lifecycle of applied suggestions**
   - pending,
   - accepted,
   - rejected,
   - future debate evolution.

5. **Maintain consistency of the document’s review state**
   - avoid duplicate re-application,
   - identify what is still pending,
   - decide when review-related UI affordances should appear.

6. **Coordinate Track Changes as part of that review lifecycle**
   - enable when needed,
   - keep active while relevant Stylistic review artifacts remain,
   - stop delegating that responsibility to per-suggestion commands.

These capabilities are why the frontend add-in is **not merely presentation** and **not merely transport**. It owns real workflow and consistency logic, even if the editorial intelligence itself lives in the backend.

---

## 5. Ubiquitous language (minimum viable vocabulary)

The project needs explicit language. The following terms are now the minimum vocabulary to use when discussing the add-in.

### 5.1 `ReviewSuggestion`

**Definition:**

> A style suggestion already contextualized in Word, with frontend-owned state and interactions.

This is the central operational unit of the current add-in.

It is **not** only a raw backend DTO.
It is **not** just a content control.
It is **not** just a UI card.

It is the suggestion as the frontend understands it once the suggestion becomes part of the Word review experience.

### 5.2 `DocumentReviewState`

**Definition:**

> The logical picture of the current Stylistic review state as it exists in the Word document.

This includes, conceptually:

- which `ReviewSuggestion` artifacts are materialized,
- which remain pending,
- which have been resolved,
- what global document review conditions follow from that state.

This is an important internal model, but it is **not the entire domain**.

Current implementation note:

- `DocumentReviewState` remains the authoritative snapshot derived from Word.
- `DocumentReviewStateMachine` is the explicit frontend/application state model
  built on top of that snapshot to keep taskpane review UI semantics consistent.
- `ReviewSessionMediator` is the explicit coordinator that combines document
  resolution outcomes, cleanup visibility, and review-state UI consequences for
  the taskpane.
- This means the document is still the source of truth, while the state machine
  is the single interpreter of review-state UI rules and the mediator is the
  single coordinator of taskpane-facing review policy.

### 5.3 `ReviewProcessState`

**Definition:**

> The temporary workflow state of the frontend while it is reading, analyzing, applying, or finishing.

Examples:

- `reading`
- `connecting`
- `analyzing`
- `applying`
- `done`
- `error`

This is distinct from `DocumentReviewState`.

It is also distinct from `DocumentReviewStateMachine`:

- `ReviewProcessState` describes **what the frontend is doing right now**.
- `DocumentReviewStateMachine` describes **what review state the taskpane must expose based on the document reality**.

### 5.4 `Debate`

**Definition:**

> A future first-class capability in which the user responds to Stylistic-originated Word comments, the backend processes that exchange, and the frontend renders the returned response back into Word/taskpane context.

Important: debate is **not** equivalent to the current fire-and-forget feedback payload.

### 5.5 “Pending suggestion”

For the Track Changes requirement, a pending suggestion is defined from the **document’s artifacts**, not from taskpane visibility or the current pipeline run.

Operationally:

> A suggestion is pending while its Stylistic artifacts remain active in Word and therefore still represent unresolved Stylistic review work.

---

## 6. What is central vs what is supporting?

This distinction matters because several earlier discussions drifted into over-promoting support models into the domain definition.

### Central to the current domain

- `ReviewSuggestion`
- presenting the suggestion
- allowing accept / reject / debate
- user interaction with suggestion state inside Word

### Supporting but important

- `DocumentReviewState`
- `ReviewProcessState`
- Track Changes lifecycle rules
- content controls
- comment cleanup logic
- backend workflow orchestration

### Infrastructure details, not domain language

- `Word.run`
- `context.sync`
- `changeTrackingMode`
- content control tag serialization
- SDK polling shapes

This hierarchy prevents a common mistake:

> if something is important to implementation, that does **not** automatically make it the domain.

---

## 7. Current architecture reality vs corrected conceptual model

The codebase already contains strong patterns, but some concepts are mixed together.

### What is already solid

- Ports and adapters are real.
- The analysis pipeline already uses Chain of Responsibility.
- State machines already exist.
- Suggestion resolution already has meaningful lifecycle semantics.
- Document-review UI semantics now have an explicit state machine instead of scattered boolean checks.
- Taskpane-facing review coordination now has an explicit mediator instead of procedural glue in `taskpane.ts`.

### What was missing

- an explicit frontend domain definition,
- explicit ubiquitous language,
- a clean distinction between process state and document review state,
- a clean ownership model for global document state like Track Changes,
- an explicit recognition that accept/reject is a workflow, not just a direct UI-to-adapter call.

### Corrected conceptual model

The frontend should be understood as:

> a system that presents backend-originated style suggestions in Word and lets the user interact with them through accept/reject/debate flows, while maintaining a consistent document review state.

---

## 8. Requirement change: Track Changes lifecycle

### 8.1 The problem

Current behavior (at the time of this document):

- each suggestion application loads the document’s current `changeTrackingMode`,
- sets it to `trackAll`,
- applies one suggestion,
- restores the previous mode in a `finally` block,
- repeats this per suggestion.

This happens inside `ApplySuggestionCommand.ts`.

### Why this is a problem

Because `changeTrackingMode` is global at the document level.

As a result:

- the UI flickers,
- Word review state is toggled repeatedly for a concern that is broader than a single suggestion,
- the implementation assigns ownership of a document-global state to the wrong abstraction,
- the command that should apply one suggestion now also governs review lifecycle policy.

That is exactly the kind of architectural leakage that creates brittle behavior.

---

## 9. Decisions taken for the requirement change

The following decisions were made during the requirement grill-me and should be treated as the agreed basis for implementation.

### 9.1 Source of truth for “pending”

The source of truth is the **document**, not the taskpane.

That means:

- not “what is currently visible in the panel”,
- not “what belongs only to the latest run”,
- but “what Stylistic artifacts currently remain active in Word”.

### 9.2 Operational criterion for pending work

For this requirement:

> There are pending suggestions while **any Stylistic artifact remains active in Word**.

This intentionally includes artifacts that may predate the current panel render or current analysis run.

### 9.3 Track Changes activation policy

Track Changes must be enabled **lazily**.

That means:

- it is **not** enabled when the user merely clicks analyze,
- it is **not** enabled when suggestions arrive from the backend,
- it is enabled only when the first `track-change` suggestion is actually about to be applied.

If a run contains only `comment-only` suggestions, Track Changes must never be enabled.

### 9.4 Track Changes persistence policy

Once enabled for Stylistic review application, Track Changes must remain enabled while pending Stylistic artifacts still exist in the document.

This is the core behavior change.

### 9.5 No automatic deactivation

When no Stylistic pending artifacts remain:

- the add-in must **not** silently disable Track Changes,
- the add-in must **not** assume ownership over all document review activity,
- the user must remain in control.

### 9.6 Final CTA behavior

The final CTA must:

- appear **immediately after** the accept/reject action that transitions the document to zero pending Stylistic artifacts,
- be a **single explicit action**:
  - `Desactivar control de cambios`

If the user does nothing, Track Changes remains active.

### 9.7 Why this CTA shape was chosen

The team explicitly preferred a lighter interaction.

So the decision was:

- do not auto-disable,
- do not show a two-button modal-style choice,
- simply offer one explicit deactivation CTA when the condition becomes true.

---

## 10. Architectural consequences of the requirement change

The requirement change is not just a one-line bug fix.

It changes where responsibility belongs.

### 10.1 `ApplySuggestionCommand` must lose Track Changes lifecycle ownership

The command may still remain the unit that applies a single suggestion.

But it should no longer be the abstraction that decides:

- when Track Changes turns on,
- when it stays on,
- when the user should be offered the final CTA.

That responsibility is above the command level.

### 10.2 The analysis flow must continue to use Chain of Responsibility

The project already chose Chain of Responsibility for the analysis workflow.

Therefore, the requirement change should follow the same architectural language.

That means adding the Track Changes enablement concern to the workflow level, not sneaking it into one command instance at a time.

### 10.3 Accept/reject must evolve toward its own Chain of Responsibility

The current accept/reject flow is too procedural and too distributed across taskpane + adapter.

Because the resolution flow now needs to also:

- resolve the suggestion,
- inspect the document’s pending state,
- decide whether the final Track Changes CTA should be offered,
- preserve non-blocking feedback semantics,

it is better modeled as a workflow, not a single direct adapter call.

### 10.4 Accept/reject variation should use Strategy inside the resolution workflow

The shared flow is mostly the same:

- resolve in document,
- inspect pending state,
- decide CTA,
- build/send feedback,
- return rich result.

The variable part is the action:

- accept vs reject,
- positive vs negative feedback payload.

So the design direction is:

- **CoR for the workflow**,
- **Strategy for per-action variation**.

### 10.5 Feedback remains non-blocking

The current system already treats feedback as fire-and-forget.

That policy remains correct.

Even if feedback becomes part of the resolution workflow, it must still be:

- non-blocking,
- swallowed on error,
- invisible to the user as failure,
- but observable in logs and explicit workflow result metadata.

### 10.6 `feedbackStatus` must become explicit in the resolution result

Even though feedback does not block UX, the workflow should return explicit semantic status such as:

- `sent`
- `failed`
- `skipped`

This improves testability and observability.

---

## 11. Present vs future scope

It is important not to pollute the current implementation scope with future debate mechanics.

### In scope now

- analysis flow,
- suggestion application,
- accept/reject lifecycle,
- Track Changes lifecycle,
- final deactivation CTA,
- non-blocking feedback semantics.

### Explicitly out of scope for this requirement

- full conversational debate threads,
- backend conversation interpretation logic,
- thread-oriented UI evolution,
- richer discussion state machines.

### But the model must leave room for that future

Today, `ReviewSuggestion` is enough.

In the future, the domain may evolve toward a thread-shaped concept if discussion becomes a real back-and-forth conversation.

That future possibility is acknowledged, but it should **not** force premature complexity into the current implementation.

---

## 12. Layer map derived from the agreed domain

This layer map is the corrected interpretation for the add-in.

### Domain

- `ReviewSuggestion`
- `DocumentReviewState`
- `ReviewProcessState`
- state transition rules
- policies around pending Stylistic review state

### Application

- analysis Chain of Responsibility
- future resolution Chain of Responsibility
- action Strategy for accept/reject
- orchestration decisions about when to query/update document review state
- policy flow that determines when Track Changes should be enabled or when CTA should be offered

### Ports

- `IDocumentPort`
- `IAnalysisPort`
- `IFeedbackPort`

### Adapters / Infrastructure

- `WordAdapter`
- `MastraAdapter`
- `FeedbackAdapter`
- Office.js document operations
- backend SDK transport details

### UI / Presentation

- taskpane rendering
- buttons
- status text
- CTA rendering
- user event binding

---

## 13. Identity policy for `ReviewSuggestion`

The project discovered an important modeling problem:

- some behaviors currently rely on frontend-generated IDs,
- others rely on `context + anchor`,
- others rely only on `anchor`,
- persisted Word tags also carry their own operational identifier form.

That is an identity policy spread across accidental implementation details.

The corrected direction is:

1. `ReviewSuggestion` should have **one domain identity**,
2. Word-specific handles should be treated as **operational references**, not second identities.

This document does not finalize the exact identity serialization yet, but it does record the principle:

> One suggestion = one domain identity. Word artifact handles are references, not competing identities.

### Important extension for replace suggestions

The project now has a stronger architectural conclusion for native Word replace
suggestions:

- a replace suggestion is a **composed review unit**, not a single inserted-side
  artifact,
- the inserted-side `ContentControl` is an operational reference, not the full
  identity,
- the deleted/original side is semantically part of the same suggestion even if
  Word exposes it asymmetrically,
- therefore, failing to observe tracked changes around the inserted-side anchor
  is **not enough** to conclude `already-resolved`.

This matters because repeated regressions have shown that host observability may
be incomplete even while the suggestion is still pending in Word.

The stronger proposal is documented here:

- [`replace-suggestion-identity-proposal.md`](./replace-suggestion-identity-proposal.md)

Recorded rule:

> Never upgrade observability failure into terminal resolution.

### Compatibility rollout for replace identity

Current rollout policy for replace suggestions:

1. new replace suggestions persist `compound-v2` metadata with inserted-side,
   deleted-side, and operational-anchor refs,
2. replace resolution no longer supports legacy bare-ID artifacts as a valid
   model path,
3. zero observed tracked changes downgrade to `unobservable`,
4. corrupt/incomplete v2 metadata maps to `identity-lost`,
5. neither `unobservable` nor `identity-lost` may emit feedback.

---

## 14. What must change in documentation and skills

When future contributors read the repo, they should not infer the wrong architecture from stale docs.

Therefore, documentation and skill guidance must reflect these corrections:

1. the add-in domain definition must be user-centered,
2. `DocumentReviewState` must be documented as an internal support concept, not the whole domain,
3. Track Changes lifecycle must be documented as a workflow-level concern,
4. Chain of Responsibility must be recognized not only for analysis, but as the intended direction for resolution as well,
5. current docs must stop presenting preserve-and-restore Track Changes behavior as the long-term intended pattern.

---

## 15. Practical implementation checklist derived from the architecture

This section is not the final task plan, but a practical bridge between the architecture and implementation work.

### Immediate direction

1. Remove Track Changes lifecycle responsibility from per-suggestion command ownership.
2. Move Track Changes enablement to the workflow level.
3. Detect pending Stylistic artifacts from the document, not the panel.
4. Keep Track Changes enabled while pending artifacts exist.
5. Detect the zero-pending transition immediately after final accept/reject.
6. Render a one-action CTA: `Desactivar control de cambios`.
7. Do not auto-disable Track Changes.
8. Preserve fire-and-forget feedback semantics.
9. Prepare accept/reject to become a true workflow with explicit application-layer orchestration.

### Immediate anti-goals

1. Do not bolt more policy into `ApplySuggestionCommand`.
2. Do not let taskpane procedural code become the long-term owner of resolution workflow logic.
3. Do not use taskpane visibility as the source of truth for pending state.
4. Do not model debate as current feedback.

---

## 16. Final summary

The stylistic-addon frontend domain is not the backend’s editorial engine, and it is not reducible to Track Changes, pipelines, or internal state machines.

Its real responsibility is simpler and more important:

> present backend-originated style suggestions in Word and let the user accept, reject, or debate them.

From that definition, everything else falls into place:

- `ReviewSuggestion` is the current operational core,
- `DocumentReviewState` is the important support model that reflects the Word document’s review reality,
- `ReviewProcessState` describes temporary workflow progress,
- Track Changes lifecycle belongs to workflow-level policy, not to individual commands,
- and the zero-pending moment is where the frontend must offer the user a clean, explicit deactivation CTA.

That is the architecture this requirement change is really asking for.

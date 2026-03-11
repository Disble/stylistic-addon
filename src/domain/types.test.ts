import type { SuggestionState, SuggestionActionResult } from './types'
import type { IDocumentPort } from './ports'
import type { Suggestion } from './types'

// ---------------------------------------------------------------------------
// SuggestionState — compile-time assignability checks
// ---------------------------------------------------------------------------

const _state1: SuggestionState = "pending"
const _state2: SuggestionState = "accepted"
const _state3: SuggestionState = "rejected"
const _state4: SuggestionState = "already-resolved"

// Suppress unused variable warnings
void _state1; void _state2; void _state3; void _state4

// ---------------------------------------------------------------------------
// SuggestionActionResult — shape checks
// ---------------------------------------------------------------------------

const _result: SuggestionActionResult = {
  status: "accepted",
  trackedChangesAffected: 2,
  commentDeleted: true,
}

const _resultWithError: SuggestionActionResult = {
  status: "error",
  trackedChangesAffected: 0,
  commentDeleted: false,
  error: "something went wrong",
}

void _result; void _resultWithError

// ---------------------------------------------------------------------------
// IDocumentPort — acceptSuggestion and rejectSuggestion method presence
// ---------------------------------------------------------------------------

type _HasAccept = IDocumentPort extends { acceptSuggestion(s: Suggestion): Promise<SuggestionActionResult> } ? true : false
type _HasReject = IDocumentPort extends { rejectSuggestion(s: Suggestion): Promise<SuggestionActionResult> } ? true : false
const _checkAccept: _HasAccept = true
const _checkReject: _HasReject = true

void _checkAccept; void _checkReject

// ---------------------------------------------------------------------------
// Runtime no-op test so Vitest registers the file
// ---------------------------------------------------------------------------

import { describe, it } from 'vitest'

describe('domain types (compile-time checks)', () => {
  it('passes if the file compiles without errors', () => {
    // All checks are compile-time — if this file compiles, the types are correct.
  })
})

export interface Suggestion {
  id: string;
  originalText: string;
  suggestedText: string;
  justification: string;
  paragraphIndex?: number;
}

export interface InsertionResult {
  successCount: number;
  failedSuggestions: Suggestion[];
}

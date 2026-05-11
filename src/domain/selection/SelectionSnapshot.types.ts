/**
 * Reactive snapshot of the user's current Word selection.
 * Emitted by `IDocumentPort.subscribeSelectionChanges` whenever the host
 * selection changes, so the taskpane can preview what will be analyzed.
 */
export interface SelectionSnapshot {
  /** Whether the current selection contains non-whitespace text. */
  hasSelection: boolean;

  /** Number of characters in the selected text (0 when no selection). */
  charCount: number;

  /** First fragment of the selected text for the UI preview (already truncated). */
  preview: string;
}

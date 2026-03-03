/* global document, Office */

import { getDocumentText, insertSuggestionsAsTrackedChanges } from "../lib/wordApi";
import { Suggestion } from "../lib/types";

const PREVIEW_MAX_CHARS = 200;
const STATUS_DISPLAY_MS = 3000;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg")!.style.display = "none";
    document.getElementById("app-body")!.style.display = "flex";
    document.getElementById("btn-read")!.onclick = handleReadDocument;
    document.getElementById("btn-test-suggestion")!.onclick = handleTestSuggestion;
  }
});

function showStatus(message: string, type: "success" | "error"): void {
  const bar = document.getElementById("status-bar")!;
  bar.textContent = message;
  bar.className = `stylistic-status ${type}`;
  bar.style.display = "block";
  setTimeout(() => {
    bar.style.display = "none";
  }, STATUS_DISPLAY_MS);
}

function showPreview(text: string): void {
  const preview = document.getElementById("doc-preview")!;
  const truncated = text.length > PREVIEW_MAX_CHARS ? text.substring(0, PREVIEW_MAX_CHARS) + "..." : text;
  preview.textContent = truncated;
  preview.style.display = "block";
}

function extractFirstWord(text: string): string | null {
  const match = text.trim().match(/\S+/);
  return match ? match[0] : null;
}

function buildTestSuggestion(firstWord: string): Suggestion {
  return {
    id: "test-1",
    originalText: firstWord,
    suggestedText: `[REPLACED:${firstWord}]`,
    justification: "Test suggestion — verifying Track Changes integration.",
  };
}

async function handleReadDocument(): Promise<void> {
  try {
    const text = await getDocumentText();

    if (!text || text.trim().length === 0) {
      showStatus("The document is empty.", "error");
      return;
    }

    showPreview(text);
    showStatus(`Document read: ${text.length} characters total.`, "success");
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    showStatus(`Error reading document: ${message}`, "error");
  }
}

async function handleTestSuggestion(): Promise<void> {
  try {
    const text = await getDocumentText();

    if (!text || text.trim().length === 0) {
      showStatus("The document is empty. Write some text first.", "error");
      return;
    }

    const firstWord = extractFirstWord(text);
    if (!firstWord) {
      showStatus("Could not find a word in the document.", "error");
      return;
    }

    const suggestion = buildTestSuggestion(firstWord);
    const result = await insertSuggestionsAsTrackedChanges([suggestion]);

    if (result.successCount > 0) {
      showStatus(
        `Tracked change inserted: "${firstWord}" → "${suggestion.suggestedText}"`,
        "success"
      );
    } else {
      showStatus(`Could not find "${firstWord}" in the document.`, "error");
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    showStatus(`Error: ${message}`, "error");
  }
}

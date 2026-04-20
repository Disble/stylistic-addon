/**
 * Shared contracts for Word text location collaborators.
 *
 * These types define the adapter-facing boundary without forcing downstream
 * consumers to import the concrete `WordTextLocatorAdapter` implementation.
 * This module also acts as the authorized composition point for the default
 * locator singleton used by the current Word-side workflows.
 *
 * @module WordTextLocatorContext
 */

import { WordTextLocatorAdapter } from "./WordTextLocatorAdapter";

/** Minimal Word search surface shared by body and range containers. */
export type WordSearchContainer = {
  search(text: string, options: Record<string, boolean>): Word.RangeCollection;
  load(property: "text"): void;
  text: string;
};

/** Input contract for one Word text-location request. */
export type WordTextLocationRequest = {
  context: Word.RequestContext;
  container: WordSearchContainer;
  searchText: string;
};

/** Narrow contract used by commands/orchestrators that locate Word ranges. */
export type TextLocator = {
  locate(request: WordTextLocationRequest): Promise<Word.Range | null>;
};

let defaultTextLocator: TextLocator | null = null;

/** Returns the authorized singleton locator for Word adapter workflows. */
export function getDefaultTextLocator(): TextLocator {
  if (defaultTextLocator === null) {
    defaultTextLocator = new WordTextLocatorAdapter();
  }

  return defaultTextLocator;
}

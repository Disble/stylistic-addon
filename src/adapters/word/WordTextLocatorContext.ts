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
import type { TextLocator } from "./WordTextLocatorContext.types";

let defaultTextLocator: TextLocator | null = null;

/** Returns the authorized singleton locator for Word adapter workflows. */
export function getDefaultTextLocator(): TextLocator {
  if (defaultTextLocator === null) {
    defaultTextLocator = new WordTextLocatorAdapter();
  }

  return defaultTextLocator;
}

/* global vi */

import type { Suggestion } from "../../domain/suggestion/Suggestion.types";

/** Mocked parent content-control shape used by ApplySuggestionCommand tests. */
export type ParentCC = {
  tag: string;
  isNullObject: boolean;
  load: ReturnType<typeof vi.fn>;
  delete: ReturnType<typeof vi.fn>;
  title?: string;
  getRange?: ReturnType<typeof vi.fn>;
};

/** Mocked Word range shape used by ApplySuggestionCommand tests. */
export type MockRange = {
  text: string;
  font: { italic: boolean; bold: boolean };
  load: ReturnType<typeof vi.fn>;
  search: ReturnType<typeof vi.fn>;
  insertText: ReturnType<typeof vi.fn>;
  getReviewedText: ReturnType<typeof vi.fn>;
  insertComment: ReturnType<typeof vi.fn>;
  insertContentControl: ReturnType<typeof vi.fn>;
  parentContentControlOrNullObject: ParentCC;
  paragraphs: {
    getFirst: ReturnType<typeof vi.fn>;
  };
};

/** Mocked Word range collection returned from search helpers. */
export type RangeCollection = {
  items: MockRange[];
  load: ReturnType<typeof vi.fn>;
};

/** Installed fake Word context and key ranges for ApplySuggestionCommand tests. */
export type ApplyCommandTestContext = {
  context: {
    document: {
      body: MockRange & { search: ReturnType<typeof vi.fn>; text: string };
      load: ReturnType<typeof vi.fn>;
      changeTrackingMode: string;
    };
    sync: ReturnType<typeof vi.fn>;
  };
  bodyRange: MockRange;
  anchorRange: MockRange;
  insertedRange: {
    text: string;
    font: { italic: boolean; bold: boolean };
    getReviewedText: ReturnType<typeof vi.fn>;
    paragraphs: {
      getFirst: ReturnType<typeof vi.fn>;
    };
    search: ReturnType<typeof vi.fn>;
    insertContentControl: ReturnType<typeof vi.fn>;
    insertComment: ReturnType<typeof vi.fn>;
  };
  cc: {
    tag: string;
    title: string;
    appearance: string;
    cannotDelete: boolean;
  };
  operationalWrapper: {
    tag: string;
    title: string;
    appearance: string;
    cannotDelete: boolean;
    getRange: ReturnType<typeof vi.fn>;
  };
  operationalWrapperRange: MockRange;
};

/** Partial suggestion fixture override accepted by test helpers. */
export type ApplySuggestionFixtureOverride = Partial<Suggestion>;

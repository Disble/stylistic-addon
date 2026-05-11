/* global vi */

/** Mock tracked-change shape used by WordAdapter action tests. */
export type MockTrackedChange = {
  id?: string;
  text?: string;
  type?: string;
  accept?: ReturnType<typeof vi.fn>;
  reject?: ReturnType<typeof vi.fn>;
  getRange?: ReturnType<typeof vi.fn>;
};

/** Mock Word range that exposes tracked-change lookup APIs. */
export type MockRangeWithTrackedChanges = {
  compareLocationWith: ReturnType<typeof vi.fn>;
  getTrackedChanges: ReturnType<typeof vi.fn>;
};

/** Mock tracked-change collection used by Office.js fakes. */
export type MockTrackedChangeCollection = {
  items: MockTrackedChange[];
  load: ReturnType<typeof vi.fn>;
  acceptAll: ReturnType<typeof vi.fn>;
  rejectAll: ReturnType<typeof vi.fn>;
};

/** Mock Word comment shape used by colocated-comment tests. */
export type MockComment = {
  authorName?: string;
  content?: string;
  getRange: ReturnType<typeof vi.fn>;
  delete: ReturnType<typeof vi.fn>;
};

/** Mock comment range capable of comparison and tracked-change lookup. */
export type MockCommentRange = {
  compareLocationWith: ReturnType<typeof vi.fn>;
  getTrackedChanges: ReturnType<typeof vi.fn>;
};

/** Fake Word context installed for resolve-suggestion action tests. */
export type ResolveSuggestionContext = {
  document: {
    contentControls: {
      getByTag: ReturnType<typeof vi.fn>;
      load: ReturnType<typeof vi.fn>;
      items: Array<{ tag: string }>;
    };
    load: ReturnType<typeof vi.fn>;
    changeTrackingMode: string;
    body: {
      search: ReturnType<typeof vi.fn>;
      load: ReturnType<typeof vi.fn>;
      text: string;
      getComments: ReturnType<typeof vi.fn>;
      getTrackedChanges: ReturnType<typeof vi.fn>;
    };
  };
  sync: ReturnType<typeof vi.fn>;
  _ccsCollection: { items: unknown[]; load: ReturnType<typeof vi.fn> };
  _commentsCollection: { items: MockComment[]; load: ReturnType<typeof vi.fn> };
  _bodyTCCollection: { items: MockTrackedChange[]; load: ReturnType<typeof vi.fn> };
  _rangeTCCollection: { items: MockTrackedChange[]; load: ReturnType<typeof vi.fn> };
  _commentRangeTCCollections: Array<{ items: MockTrackedChange[]; load: ReturnType<typeof vi.fn> }>;
  _ccItems: Array<{
    title: string;
    tag: string;
    load: ReturnType<typeof vi.fn>;
    getTrackedChanges: ReturnType<typeof vi.fn>;
    getRange: ReturnType<typeof vi.fn>;
    delete: ReturnType<typeof vi.fn>;
  }>;
  _cc: {
    title: string;
    tag: string;
    load: ReturnType<typeof vi.fn>;
    getTrackedChanges: ReturnType<typeof vi.fn>;
    getRange: ReturnType<typeof vi.fn>;
    delete: ReturnType<typeof vi.fn>;
  };
};

/** Minimal fake Word global installed by action tests. */
export type MockWordGlobal = {
  run: ReturnType<typeof vi.fn>;
};

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

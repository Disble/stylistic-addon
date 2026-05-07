/** Parent content control shape used while reusing operational wrappers. */
export type ParentOperationalContentControl = Word.ContentControl & {
  tag?: string;
  title?: string;
  getRange?: () => Word.Range;
};

/** Result of resolving or creating an operational wrapper. */
export type ApplySuggestionOperationalWrapperResolution =
  | {
      /** Successful resolution branch. */
      kind: "success";

      /** Reusable or newly created wrapper. */
      wrapper: Word.ContentControl;
    }
  | {
      /** Fail-closed resolution branch. */
      kind: "error";

      /** Stable fail-closed error message. */
      error: string;
    };

/** Parent content control shape used while reusing operational wrappers. */
export type ParentOperationalContentControl = Word.ContentControl & {
  tag?: string;
  title?: string;
  getRange?: () => Word.Range;
};

/** Result of resolving or creating an operational wrapper. */
export type ApplySuggestionOperationalWrapperResolution =
  | {
      /** Reusable or newly created wrapper. */
      wrapper: Word.ContentControl;

      /** No error when wrapper resolution succeeded. */
      error?: undefined;
    }
  | {
      /** No wrapper is available when resolution fails closed. */
      wrapper?: undefined;

      /** Stable fail-closed error message. */
      error: string;
    };

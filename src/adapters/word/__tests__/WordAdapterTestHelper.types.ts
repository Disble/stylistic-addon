/** Callback shape accepted by the fake `Word.run` harness. */
export type WordRunCallback<T> = (context: any) => Promise<T> | T;

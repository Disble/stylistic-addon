/** Backend bridge response returned to the Office auth dialog. */
export type AuthBridgeSessionResponse = Readonly<{
  session?: unknown;
  error?: string;
}>;

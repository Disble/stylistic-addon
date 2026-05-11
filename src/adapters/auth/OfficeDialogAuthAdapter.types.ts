import type { AuthSession } from "../../domain/auth/AuthSession.types";

/** Cross-window message contract exchanged with the Office auth dialog. */
export type DialogMessage = Readonly<
  | { type: "stylistic-auth-success"; session: AuthSession }
  | { type: "stylistic-auth-error"; message: string }
>;

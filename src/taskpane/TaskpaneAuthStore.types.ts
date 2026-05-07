import type { AuthSession } from "../domain/auth/AuthSession.types";

/** Authentication lifecycle exposed to taskpane presentation components. */
export type TaskpaneAuthStatus = "loading" | "authenticated" | "unauthenticated";

/** Root Zustand state for taskpane authentication UI. */
export type TaskpaneAuthState = Readonly<{
  status: TaskpaneAuthStatus;
  session?: AuthSession;
  error?: string;
  isSigningIn: boolean;
  isSigningOut: boolean;
}>;

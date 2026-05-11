import type { TaskpaneAuthState } from "./TaskpaneAuthStore.types";

/** Initial auth presentation state before session bootstrap completes. */
export const INITIAL_TASKPANE_AUTH_STATE: TaskpaneAuthState = {
  status: "loading",
  isSigningIn: false,
  isSigningOut: false,
};

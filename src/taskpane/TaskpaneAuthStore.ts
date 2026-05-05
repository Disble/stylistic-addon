import { create } from "zustand";
import type { AuthSession } from "../domain/auth/AuthSession.types";

export type TaskpaneAuthStatus = "loading" | "authenticated" | "unauthenticated";

export type TaskpaneAuthState = Readonly<{
  status: TaskpaneAuthStatus;
  session?: AuthSession;
  error?: string;
  isSigningIn: boolean;
  isSigningOut: boolean;
}>;

const INITIAL_AUTH_STATE: TaskpaneAuthState = {
  status: "loading",
  isSigningIn: false,
  isSigningOut: false,
};

/** Zustand store for React-owned auth presentation state. */
export const useTaskpaneAuthStore = create<TaskpaneAuthState>()(() => INITIAL_AUTH_STATE);

/** Returns the current bearer token used by authenticated Mastra adapters. */
export function getTaskpaneAuthToken(): string | undefined {
  return useTaskpaneAuthStore.getState().session?.token;
}

/** Marks auth bootstrap as in progress. */
export function setTaskpaneAuthLoading(): void {
  useTaskpaneAuthStore.setState({ status: "loading", error: undefined });
}

/** Stores the authenticated session in presentation state. */
export function setTaskpaneAuthenticated(session: AuthSession): void {
  useTaskpaneAuthStore.setState({
    status: "authenticated",
    session,
    error: undefined,
    isSigningIn: false,
    isSigningOut: false,
  });
}

/** Clears the current auth state and exposes an optional user-facing error. */
export function setTaskpaneUnauthenticated(error?: string): void {
  useTaskpaneAuthStore.setState({
    status: "unauthenticated",
    session: undefined,
    error,
    isSigningIn: false,
    isSigningOut: false,
  });
}

/** Sets the login button loading state. */
export function setTaskpaneSigningIn(isSigningIn: boolean): void {
  useTaskpaneAuthStore.setState({ isSigningIn });
}

/** Sets the logout button loading state. */
export function setTaskpaneSigningOut(isSigningOut: boolean): void {
  useTaskpaneAuthStore.setState({ isSigningOut });
}

/** Resets auth state for deterministic tests. */
export function resetTaskpaneAuthState(): void {
  useTaskpaneAuthStore.setState(INITIAL_AUTH_STATE, true);
}

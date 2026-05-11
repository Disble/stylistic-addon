import type { TaskpaneAuthStatus } from "../../TaskpaneAuthStore.types";

/** Props required to render the taskpane authentication section. */
export type AuthSectionProps = Readonly<{
  error?: string;
  isSigningIn: boolean;
  onSignIn: () => Promise<void> | void;
  status: TaskpaneAuthStatus;
}>;

/** Griffel class slots consumed by the auth section. */
export type AuthSectionClasses = Readonly<{
  root: string;
  content: string;
  title: string;
  description: string;
  button: string;
}>;

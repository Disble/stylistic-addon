import type { AuthSession } from "../../../domain/auth/AuthSession.types";
import type { TaskpaneAuthStatus } from "../../TaskpaneAuthStore";

export type AuthSectionProps = Readonly<{
  error?: string;
  isSigningIn: boolean;
  isSigningOut: boolean;
  onSignIn: () => Promise<void> | void;
  onSignOut: () => Promise<void> | void;
  session?: AuthSession;
  status: TaskpaneAuthStatus;
}>;

export type AuthSectionClasses = Readonly<{
  root: string;
  content: string;
  title: string;
  description: string;
  button: string;
}>;

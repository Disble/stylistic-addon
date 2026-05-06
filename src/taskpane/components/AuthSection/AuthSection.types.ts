import type { TaskpaneAuthStatus } from "../../TaskpaneAuthStore";

export type AuthSectionProps = Readonly<{
  error?: string;
  isSigningIn: boolean;
  onSignIn: () => Promise<void> | void;
  status: TaskpaneAuthStatus;
}>;

export type AuthSectionClasses = Readonly<{
  root: string;
  content: string;
  title: string;
  description: string;
  button: string;
}>;

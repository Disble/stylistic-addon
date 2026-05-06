import type { AuthSession } from "../../../domain/auth/AuthSession.types";

export type AccountSettingsProps = Readonly<{
  isSigningOut: boolean;
  onSignOut: () => Promise<void> | void;
  session?: AuthSession;
}>;

export type AccountSettingsClasses = Readonly<{
  root: string;
  header: string;
  title: string;
  row: string;
  email: string;
  logoutButton: string;
}>;

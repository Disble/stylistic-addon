import type { AuthSession } from "../../../domain/auth/AuthSession.types";

/** Props required to render the account settings panel. */
export type AccountSettingsProps = Readonly<{
  isSigningOut: boolean;
  onSignOut: () => Promise<void> | void;
  session?: AuthSession;
}>;

/** Griffel class slots consumed by the account settings panel. */
export type AccountSettingsClasses = Readonly<{
  root: string;
  header: string;
  title: string;
  row: string;
  email: string;
  logoutButton: string;
}>;

import type { AuthSession } from "../../../domain/auth/AuthSession.types";

export type SettingsViewProps = Readonly<{
  isSigningOut: boolean;
  onBack: () => void;
  onSignOut: () => Promise<void> | void;
  session?: AuthSession;
}>;

export type SettingsViewClasses = Readonly<{
  root: string;
  header: string;
  backButton: string;
  title: string;
  body: string;
}>;

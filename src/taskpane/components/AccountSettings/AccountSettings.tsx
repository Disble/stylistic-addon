import * as React from "react";
import { Button, Spinner } from "@fluentui/react-components";
import { ArrowExitRegular, PersonRegular } from "@fluentui/react-icons";
import type { AccountSettingsProps } from "./AccountSettings.types";
import { useAccountSettings } from "./useAccountSettings";

/** Renders the account row with the active session email and a logout action. */
export function AccountSettings({
  isSigningOut,
  onSignOut,
  session,
}: AccountSettingsProps): React.JSX.Element {
  const classes = useAccountSettings();
  const identity = session?.user.email ?? session?.user.name ?? "Usuario autenticado";

  return (
    <section className={classes.root} aria-label="Cuenta">
      <div className={classes.header}>
        <PersonRegular aria-hidden="true" />
        <span className={classes.title}>Cuenta</span>
      </div>
      <div className={classes.row}>
        <span className={classes.email} title={identity} data-testid="account-email">
          {identity}
        </span>
        <Button
          appearance="subtle"
          aria-label="Cerrar sesión"
          className={classes.logoutButton}
          data-testid="logout-button"
          disabled={isSigningOut}
          icon={isSigningOut ? <Spinner size="tiny" /> : <ArrowExitRegular aria-hidden="true" />}
          onClick={() => {
            void onSignOut();
          }}
          type="button"
        >
          {isSigningOut ? "Cerrando..." : "Cerrar sesión"}
        </Button>
      </div>
    </section>
  );
}

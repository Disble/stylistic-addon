import * as React from "react";
import { Button, MessageBar, MessageBarBody, Spinner } from "@fluentui/react-components";
import { PersonRegular } from "@fluentui/react-icons";
import type { AuthSectionProps } from "./AuthSection.types";
import { useAuthSection } from "./useAuthSection";

/**
 * Renders the pre-app gate: a loading indicator while the persisted session is verified,
 * or a login screen when the user has no active session. Once authenticated, the App
 * shell takes over and this section is no longer rendered.
 */
export function AuthSection({
  error,
  isSigningIn,
  onSignIn,
  status,
}: AuthSectionProps): React.JSX.Element {
  const classes = useAuthSection();

  if (status === "loading") {
    return (
      <section className={classes.root}>
        <div className={classes.content}>
          <Spinner label="Verificando sesión..." size="small" />
        </div>
      </section>
    );
  }

  return (
    <section className={classes.root}>
      <div className={classes.content}>
        <div className={classes.title}>Iniciá sesión para usar Stylistic</div>
        <div className={classes.description}>Conectamos con Google para proteger tus análisis y tu historial.</div>
        {error ? (
          <MessageBar intent="error">
            <MessageBarBody>{error}</MessageBarBody>
          </MessageBar>
        ) : null}
        <Button
          appearance="primary"
          className={classes.button}
          disabled={isSigningIn}
          icon={isSigningIn ? <Spinner size="tiny" /> : <PersonRegular aria-hidden="true" />}
          onClick={() => {
            void onSignIn();
          }}
          type="button"
        >
          {isSigningIn ? "Abriendo Google..." : "Continuar con Google"}
        </Button>
      </div>
    </section>
  );
}

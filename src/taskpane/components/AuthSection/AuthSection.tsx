import * as React from "react";
import { Button, MessageBar, MessageBarBody, Spinner } from "@fluentui/react-components";
import { ArrowExitRegular, PersonRegular } from "@fluentui/react-icons";
import type { AuthSectionProps } from "./AuthSection.types";
import { useAuthSection } from "./useAuthSection";

/** Renders login/logout controls for the taskpane. */
export function AuthSection({
  error,
  isSigningIn,
  isSigningOut,
  onSignIn,
  onSignOut,
  session,
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

  if (status === "authenticated") {
    return (
      <section className={classes.root}>
        <div className={classes.content}>
          <div className={classes.title}>Sesión activa</div>
          <div className={classes.description}>{session?.user.email ?? session?.user.name ?? "Usuario autenticado"}</div>
          <Button
            appearance="secondary"
            className={classes.button}
            disabled={isSigningOut}
            icon={isSigningOut ? <Spinner size="tiny" /> : <ArrowExitRegular aria-hidden="true" />}
            onClick={() => {
              void onSignOut();
            }}
            type="button"
          >
            {isSigningOut ? "Cerrando sesión..." : "Cerrar sesión"}
          </Button>
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

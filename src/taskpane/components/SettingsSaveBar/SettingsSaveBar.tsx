import * as React from "react";
import { Button, Spinner } from "@fluentui/react-components";
import type { SettingsSaveBarProps } from "./SettingsSaveBar.types";
import { useSettingsSaveBar } from "./SettingsSaveBar.hooks";

/**
 * Renders the persistent save bar at the bottom of the settings page.
 *
 * Disables the save button while there are no pending changes (`isDirty=false`)
 * or while a save is in flight. Shows a Spinner inside the button as the active
 * indicator and surfaces a save error inline above the actions row.
 */
export function SettingsSaveBar({
  isDirty,
  isSaving,
  saveError,
  onSave,
}: SettingsSaveBarProps): React.JSX.Element {
  const classes = useSettingsSaveBar();
  const isDisabled = !isDirty || isSaving;
  return (
    <div className={classes.root}>
      {saveError && (
        <span className={classes.error} data-testid="settings-save-error" role="alert">
          {saveError}
        </span>
      )}
      <div className={classes.actions}>
        <Button
          appearance="primary"
          aria-label="Guardar configuración"
          className={classes.saveButton}
          data-testid="settings-save-button"
          disabled={isDisabled}
          icon={isSaving ? <Spinner size="tiny" /> : undefined}
          onClick={onSave}
          type="button"
        >
          {isSaving ? "Guardando..." : "Guardar"}
        </Button>
      </div>
    </div>
  );
}

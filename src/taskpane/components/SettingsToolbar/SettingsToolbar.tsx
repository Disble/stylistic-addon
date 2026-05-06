import * as React from "react";
import { Button, Tooltip } from "@fluentui/react-components";
import { SettingsRegular } from "@fluentui/react-icons";
import type { SettingsToolbarProps } from "./SettingsToolbar.types";
import { useSettingsToolbar } from "./useSettingsToolbar";

/** Persistent footer toolbar exposing access to the settings view. */
export function SettingsToolbar({
  isDisabled,
  onOpenSettings,
}: SettingsToolbarProps): React.JSX.Element {
  const classes = useSettingsToolbar();

  return (
    <footer className={classes.root} aria-label="Acciones secundarias">
      <Tooltip content="Settings" relationship="label">
        <Button
          appearance="subtle"
          aria-label="Settings"
          className={classes.button}
          data-testid="open-settings-button"
          disabled={isDisabled}
          icon={<SettingsRegular aria-hidden="true" />}
          onClick={onOpenSettings}
          type="button"
        />
      </Tooltip>
    </footer>
  );
}

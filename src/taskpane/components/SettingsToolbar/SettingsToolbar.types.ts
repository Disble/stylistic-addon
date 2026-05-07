/** Props required to render the settings toolbar action. */
export type SettingsToolbarProps = Readonly<{
  isDisabled?: boolean;
  onOpenSettings: () => void;
}>;

/** Griffel class slots consumed by the settings toolbar. */
export type SettingsToolbarClasses = Readonly<{
  root: string;
  button: string;
}>;

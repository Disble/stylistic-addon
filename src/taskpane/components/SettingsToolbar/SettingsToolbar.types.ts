export type SettingsToolbarProps = Readonly<{
  isDisabled?: boolean;
  onOpenSettings: () => void;
}>;

export type SettingsToolbarClasses = Readonly<{
  root: string;
  button: string;
}>;

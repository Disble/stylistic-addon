/** Props consumed by the settings save bar. */
export type SettingsSaveBarProps = Readonly<{
  isDirty: boolean;
  isSaving: boolean;
  saveError?: string;
  onSave: () => void;
}>;

/** Griffel class slots consumed by the settings save bar. */
export type SettingsSaveBarClasses = Readonly<{
  root: string;
  error: string;
  actions: string;
  saveButton: string;
}>;

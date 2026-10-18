/** Canonical identifier shared by persisted analysis profiles and the analysis genre. */
export type AnalysisProfileId =
  | "narrativa-literaria"
  | "ensayo-academico"
  | "periodismo-cultural"
  | "general";

/** An analysis profile option shown in the UI dropdown. */
export interface Profile {
  /** Machine-readable identifier sent to the workflow. */
  id: AnalysisProfileId;

  /** Human-readable label displayed in the dropdown. */
  label: string;
}

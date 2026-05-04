import type { AnalysisProfileOption } from "./App.types";

export const DEFAULT_APP_TITLE = "Stylistic";

export const DEFAULT_APP_SUBTITLE = "Sugerencias editoriales con Track Changes";

export const ANALYSIS_PROFILE_OPTIONS: readonly AnalysisProfileOption[] = [
  {
    value: "narrativa-literaria",
    label: "Literatura de ficción",
  },
  {
    value: "general",
    label: "General",
  },
  {
    value: "ensayo-academico",
    label: "Ensayo académico",
  },
  {
    value: "periodismo-cultural",
    label: "Periodismo cultural",
  },
];

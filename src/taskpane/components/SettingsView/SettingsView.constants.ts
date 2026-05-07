import { DEFAULT_PROFILES } from "../../../infrastructure/config";
import type { AnalysisProfileOption } from "../AnalysisProfileSection";

/** Selectable analysis-profile options derived from the canonical profile list. */
export const ANALYSIS_PROFILE_OPTIONS: readonly AnalysisProfileOption[] = DEFAULT_PROFILES.map(
  (profile) => ({
    value: profile.id,
    label: profile.label,
  })
);

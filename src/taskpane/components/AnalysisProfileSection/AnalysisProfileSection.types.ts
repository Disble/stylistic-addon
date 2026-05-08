import type { AnalysisProfileId } from "../../../domain/Profile.types";

/** Selectable editorial profile option shown in the analyze section. */
export type AnalysisProfileOption = Readonly<{
  value: AnalysisProfileId;
  label: string;
}>;

/** Props required to render the analysis-profile selector section. */
export type AnalysisProfileSectionProps = Readonly<{
  isDisabled: boolean;
  onGeneroChange: (value: AnalysisProfileId) => void;
  options: readonly AnalysisProfileOption[];
  selectedGenero: AnalysisProfileId;
}>;

/** Griffel class slots consumed by the analysis-profile section. */
export type AnalysisProfileSectionClasses = Readonly<{
  root: string;
  field: string;
  dropdown: string;
}>;

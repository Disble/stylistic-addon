import type { AnalysisProfileOption } from "../App";

export type AnalysisProfileSectionProps = Readonly<{
  isDisabled: boolean;
  options: readonly AnalysisProfileOption[];
}>;

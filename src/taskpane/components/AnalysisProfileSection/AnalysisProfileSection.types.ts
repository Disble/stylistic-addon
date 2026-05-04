import type * as React from "react";
import type { AnalysisProfileOption } from "../App";

export type AnalysisProfileSectionProps = Readonly<{
  isDisabled: boolean;
  onGeneroChange: React.ChangeEventHandler<HTMLSelectElement>;
  options: readonly AnalysisProfileOption[];
  selectedGenero: string;
}>;

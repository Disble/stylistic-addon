export type AnalysisProfileOption = Readonly<{
  value: string;
  label: string;
}>;

export type AnalysisProfileSectionProps = Readonly<{
  isDisabled: boolean;
  onGeneroChange: (value: string) => void;
  options: readonly AnalysisProfileOption[];
  selectedGenero: string;
}>;

export type AnalysisProfileSectionClasses = Readonly<{
  root: string;
  field: string;
  dropdown: string;
}>;

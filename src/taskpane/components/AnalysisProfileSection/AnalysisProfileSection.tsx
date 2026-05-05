import * as React from "react";
import { Dropdown, Field, Option } from "@fluentui/react-components";
import type { AnalysisProfileSectionProps } from "./AnalysisProfileSection.types";
import { useAnalysisProfileSection } from "./useAnalysisProfileSection";

/** Renders the analysis-profile selector using a Fluent Dropdown inside a labelled Field. */
export function AnalysisProfileSection({
  isDisabled,
  onGeneroChange,
  options,
  selectedGenero,
}: AnalysisProfileSectionProps): React.JSX.Element {
  const classes = useAnalysisProfileSection();
  const selectedOption = options.find((option) => option.value === selectedGenero);

  return (
    <div className={classes.root}>
      <Field label="Perfil de análisis" className={classes.field}>
        <Dropdown
          aria-label="Perfil de análisis"
          className={classes.dropdown}
          data-testid="profile-dropdown"
          disabled={isDisabled}
          onOptionSelect={(_event, data) => {
            if (typeof data.optionValue === "string") {
              onGeneroChange(data.optionValue);
            }
          }}
          selectedOptions={[selectedGenero]}
          value={selectedOption?.label ?? ""}
        >
          {options.map((option) => (
            <Option key={option.value} value={option.value}>
              {option.label}
            </Option>
          ))}
        </Dropdown>
      </Field>
    </div>
  );
}

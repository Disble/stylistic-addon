import * as React from "react";
import { Tooltip } from "@fluentui/react-components";
import { TextEditStyleRegular } from "@fluentui/react-icons";
import type { SelectionPreviewProps } from "./SelectionPreview.types";

/**
 * Single-line status strip showing the active Word selection char count.
 * The full truncated quote is exposed via Tooltip on hover/focus to keep
 * vertical real estate for the primary content (suggestion cards).
 */
export function SelectionPreview({
  isVisible,
  charCount,
  preview,
  classes,
}: SelectionPreviewProps): React.JSX.Element | null {
  if (!isVisible) {
    return null;
  }

  return (
    <Tooltip
      content={`“${preview}”`}
      relationship="description"
      positioning="below-start"
      withArrow
    >
      <div
        id="selection-preview"
        className={classes.root}
        role="status"
        aria-live="polite"
        tabIndex={0}
      >
        <TextEditStyleRegular className={classes.icon} aria-hidden="true" />
        <span className={classes.label}>
          Selección activa · {charCount} caracteres
        </span>
      </div>
    </Tooltip>
  );
}

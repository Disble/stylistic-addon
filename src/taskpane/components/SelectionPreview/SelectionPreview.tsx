import * as React from "react";
import {
  Body1,
  Caption1Strong,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
} from "@fluentui/react-components";
import { TextEditStyleRegular } from "@fluentui/react-icons";
import type { SelectionPreviewProps } from "./SelectionPreview.types";

/** Renders a non-interactive Fluent MessageBar previewing the active Word selection. */
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
    <MessageBar
      id="selection-preview"
      className={classes.root}
      intent="info"
      icon={<TextEditStyleRegular aria-hidden="true" />}
      layout="multiline"
      politeness="polite"
    >
      <MessageBarBody className={classes.body}>
        <MessageBarTitle>Selección activa · {charCount} caracteres</MessageBarTitle>
        <Body1 className={classes.quote}>“{preview}”</Body1>
        <Caption1Strong className={classes.hint}>Esto se va a analizar</Caption1Strong>
      </MessageBarBody>
    </MessageBar>
  );
}

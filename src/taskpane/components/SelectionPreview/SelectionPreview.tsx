import * as React from "react";
import type { SelectionPreviewProps } from "./SelectionPreview.types";

/** Renders a non-interactive preview of the active Word selection. */
export function SelectionPreview({
  isVisible,
  charCount,
  preview,
}: SelectionPreviewProps): React.JSX.Element | null {
  if (!isVisible) {
    return null;
  }

  return (
    <section
      id="selection-preview"
      className="selection-preview"
      aria-live="polite"
      aria-label="Texto seleccionado para analizar"
    >
      <header className="selection-preview__header">
        <span className="selection-preview__icon" aria-hidden="true">
          ✎
        </span>
        <span className="selection-preview__title">Selección activa</span>
      </header>
      <p className="selection-preview__quote">
        <span className="selection-preview__preview-text">“{preview}”</span>
      </p>
      <footer className="selection-preview__footer">
        <span className="selection-preview__count">{charCount} caracteres</span>
        <span className="selection-preview__hint">Esto se va a analizar</span>
      </footer>
    </section>
  );
}

import * as React from "react";
import type { AnalysisHeroDocumentIllustrationProps } from "./AnalysisHeroDocumentIllustration.types";

/** Renders the shared document silhouette reused by analysis hero states. */
export function AnalysisHeroDocumentIllustration({
  svgClassName,
  documentClassName,
  children,
}: AnalysisHeroDocumentIllustrationProps): React.JSX.Element {
  return (
    <svg
      aria-hidden="true"
      className={svgClassName}
      fill="none"
      role="img"
      viewBox="0 0 240 200"
      xmlns="http://www.w3.org/2000/svg"
    >
      <ellipse cx="120" cy="184" fill="currentColor" opacity="0.08" rx="78" ry="6" />

      <g className={documentClassName}>
        <rect fill="currentColor" height="130" opacity="0.06" rx="8" width="118" x="64" y="38" />
        <rect
          fill="white"
          height="130"
          rx="8"
          stroke="currentColor"
          strokeWidth="1.5"
          width="118"
          x="56"
          y="32"
        />
        {children}
      </g>
    </svg>
  );
}

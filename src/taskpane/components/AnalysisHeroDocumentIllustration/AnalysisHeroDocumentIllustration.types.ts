import type * as React from "react";

/** Props required to render the shared document illustration used by hero states. */
export type AnalysisHeroDocumentIllustrationProps = Readonly<{
  svgClassName: string;
  documentClassName: string;
  children?: React.ReactNode;
}>;

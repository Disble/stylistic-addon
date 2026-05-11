import type * as React from "react";

/** Props required to render the empty-state hero wrapper. */
export type HeroEmptyStateProps = Readonly<{
  children: React.ReactNode;
}>;

/** Griffel class slots consumed by the hero empty state. */
export type HeroEmptyStateClasses = Readonly<{
  root: string;
  illustrationWrapper: string;
  illustration: string;
  sparkle1: string;
  sparkle2: string;
  sparkle3: string;
  title: string;
  subtitle: string;
  actions: string;
}>;

import type * as React from "react";

/** Type-erased props accepted by minimal Fluent test doubles. */
export type MockProps = Record<string, unknown> & {
  children?: React.ReactNode;
  className?: unknown;
  icon?: React.ReactNode;
  label?: React.ReactNode;
  max?: unknown;
  onChange?: ((event: unknown, data?: Record<string, unknown>) => void) | undefined;
  onOptionSelect?: ((event: unknown, data: Record<string, unknown>) => void) | undefined;
  placeholder?: unknown;
  selectedOptions?: unknown;
  size?: unknown;
  validationState?: unknown;
  value?: unknown;
};

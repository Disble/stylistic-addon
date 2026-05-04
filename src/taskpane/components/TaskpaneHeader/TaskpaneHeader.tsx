import * as React from "react";
import type { TaskpaneHeaderProps } from "./TaskpaneHeader.types";

/** Renders the static taskpane header. */
export function TaskpaneHeader({ title, subtitle }: TaskpaneHeaderProps): React.JSX.Element {
  return (
    <header className="stylistic-header">
      <h1>{title}</h1>
      <p>{subtitle}</p>
    </header>
  );
}

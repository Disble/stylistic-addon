import * as React from "react";
import type { DisableTrackChangesSectionProps } from "./DisableTrackChangesSection.types";

/** Renders the Track Changes disable CTA anchor. */
export function DisableTrackChangesSection({
  isVisible,
}: DisableTrackChangesSectionProps): React.JSX.Element {
  return (
    <div
      className="stylistic-section"
      id="disable-track-changes-section"
      style={{ display: isVisible ? "block" : "none" }}
    >
      <button
        id="btn-disable-track-changes"
        type="button"
        className="stylistic-btn stylistic-btn--warning"
      >
        <span id="btn-disable-track-changes-label">Desactivar control de cambios</span>
      </button>
    </div>
  );
}

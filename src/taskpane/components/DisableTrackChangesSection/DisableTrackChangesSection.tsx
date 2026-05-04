import * as React from "react";
import type { DisableTrackChangesSectionProps } from "./DisableTrackChangesSection.types";

/** Renders the Track Changes disable CTA anchor. */
export function DisableTrackChangesSection({
  isLoading,
  isVisible,
  onDisableTrackChanges,
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
        disabled={isLoading}
        onClick={() => {
          void onDisableTrackChanges();
        }}
      >
        <span id="btn-disable-track-changes-label">
          {isLoading ? "Desactivando..." : "Desactivar control de cambios"}
        </span>
      </button>
    </div>
  );
}

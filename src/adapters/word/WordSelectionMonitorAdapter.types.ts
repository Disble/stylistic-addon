import type { SelectionSnapshot } from "../../domain/selection/SelectionSnapshot.types";

/** Configuration for how selection snapshots should be normalized. */
export interface SelectionMonitorOptions {
  previewMaxChars?: number;
}

/** Callback notified when the Word selection snapshot changes. */
export type SelectionListener = (snapshot: SelectionSnapshot) => void;

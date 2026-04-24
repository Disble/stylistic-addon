export type ReplaceTrackedChangeSide = "Added" | "Deleted";

/** Common contract for replace-resolution algorithms selected by the client. */
export interface ReplaceResolutionStrategy {
  readonly actionLabel: "aceptación" | "rechazo";
  readonly semanticOrder: readonly [
    ReplaceTrackedChangeSide,
    ReplaceTrackedChangeSide,
  ];
  priorityFor(trackedChangeType: string): number;
}

/** Replace policy for accept flows: inserted side first, original side second. */
export class AcceptReplaceResolutionStrategy
  implements ReplaceResolutionStrategy
{
  readonly actionLabel = "aceptación" as const;
  readonly semanticOrder = ["Added", "Deleted"] as const;

  priorityFor(trackedChangeType: string): number {
    if (trackedChangeType === "Added") {
      return 0;
    }

    if (trackedChangeType === "Deleted") {
      return 1;
    }

    return 2;
  }
}

/** Replace policy for reject flows: original side first, inserted side second. */
export class RejectReplaceResolutionStrategy
  implements ReplaceResolutionStrategy
{
  readonly actionLabel = "rechazo" as const;
  readonly semanticOrder = ["Deleted", "Added"] as const;

  priorityFor(trackedChangeType: string): number {
    if (trackedChangeType === "Deleted") {
      return 0;
    }

    if (trackedChangeType === "Added") {
      return 1;
    }

    return 2;
  }
}

export type ReplaceTrackedChangeSide = "Added" | "Deleted";

/** Common contract for replace-resolution algorithms selected by the client. */
export interface ReplaceResolutionStrategy {
  readonly actionLabel: "aceptación" | "rechazo";
  readonly semanticOrder: readonly [
    ReplaceTrackedChangeSide,
    ReplaceTrackedChangeSide,
  ];
}

/** Replace policy for accept flows: inserted side first, original side second. */
export class AcceptReplaceResolutionStrategy
  implements ReplaceResolutionStrategy
{
  readonly actionLabel = "aceptación" as const;
  readonly semanticOrder = ["Added", "Deleted"] as const;
}

/** Replace policy for reject flows: original side first, inserted side second. */
export class RejectReplaceResolutionStrategy
  implements ReplaceResolutionStrategy
{
  readonly actionLabel = "rechazo" as const;
  readonly semanticOrder = ["Deleted", "Added"] as const;
}

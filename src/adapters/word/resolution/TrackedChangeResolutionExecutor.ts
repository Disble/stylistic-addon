/** Applies one terminal resolution action to a tracked-change collection. */
export class TrackedChangeResolutionExecutor {
  constructor(
    private readonly suggestionId: string,
    private readonly action: "accept" | "reject",
  ) {}

  /** Applies the requested action to all tracked changes. */
  apply(trackedChanges: Word.TrackedChange[]): void {
    for (const trackedChange of trackedChanges) {
      if (this.action === "accept") {
        trackedChange.accept();
      } else {
        trackedChange.reject();
      }
    }

    console.log(
      `🎯 [TrackedChangeResolutionExecutor] executed ${this.action} on ${trackedChanges.length} tracked changes for suggestionId="${this.suggestionId}"`,
    );
  }
}

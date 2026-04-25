import type {
  ReplaceSuggestionIdentity,
  Suggestion,
  SuggestionObservationStatus,
} from "../../../domain/types";
import {
  getOperationalAnchorLocator,
  isValidOperationalReplaceIdentity,
  parseReplaceIdentityTitle,
} from "../ReplaceIdentityParser";
import type {
  TextLocator,
  WordSearchContainer,
} from "../WordTextLocatorContext";
import { OperationalWrapperGroupResolver } from "./OperationalWrapperGroupResolver";
import type { ReplaceTrackedChangeSide } from "./ReplaceResolutionStrategyContext";
import type {
  ColocatedCommentContext,
  ReplaceObservationContext,
  ReplaceSemanticCandidateMap,
  ReplaceTrackedChangeCandidate,
  ResolutionObservation,
  ResolutionObservationDebugMetadata,
} from "./ResolutionContext";
import type { SuggestionLocator } from "./SuggestionLocator";

type ReplaceTrackedChangeSources = {
  ccTrackedChanges: Word.TrackedChange[];
  ccRangeTrackedChanges: Word.TrackedChange[];
  operationalAnchorTrackedChanges: Word.TrackedChange[];
  commentTrackedChanges: Word.TrackedChange[];
};

type LoadedReplaceObservationSources = {
  sources: ReplaceTrackedChangeSources;
  operationalAnchorFound: boolean;
  baseDebugMetadata: Pick<
    ResolutionObservationDebugMetadata,
    | "ccTrackedChangesCount"
    | "ccRangeTrackedChangesCount"
    | "operationalAnchorTrackedChangesCount"
    | "operationalAnchorFound"
    | "commentTrackedChangesCount"
  >;
};

type LoadedOperationalWrapperMemberSources = LoadedReplaceObservationSources & {
  identity: ReplaceSuggestionIdentity;
};

type SerializedOfficeErrorDiagnostics = {
  message: string;
  name?: string;
  code?: string | number;
  debugInfo?: unknown;
  traceMessages?: unknown;
  stackPreview?: string[];
};

/** Collects and classifies host evidence for one resolution workflow. */
export class SuggestionResolutionObserver {
  private readonly groupResolver = new OperationalWrapperGroupResolver();

  constructor(
    private readonly suggestion: Suggestion,
    private readonly locator: SuggestionLocator,
    private readonly textLocator: TextLocator,
  ) {}

  /** Emits one searchable observe-before diagnostic log entry for this suggestion. */
  private logObserveBefore(
    step: string,
    details?: Record<string, unknown>,
  ): void {
    if (details) {
      console.log(
        `🧪 [SuggestionResolutionObserver] observe-before ${step}`,
        details,
      );
      return;
    }

    console.log(`🧪 [SuggestionResolutionObserver] observe-before ${step}`);
  }

  /** Emits one searchable observe-before diagnostic warning for this suggestion. */
  private warnObserveBefore(
    step: string,
    details?: Record<string, unknown>,
  ): void {
    if (details) {
      console.warn(
        `🧪 [SuggestionResolutionObserver] observe-before ${step}`,
        details,
      );
      return;
    }

    console.warn(`🧪 [SuggestionResolutionObserver] observe-before ${step}`);
  }

  /** Reads one unknown error property defensively so diagnostics never throw while logging. */
  private readUnknownErrorProperty(
    error: unknown,
    propertyName: string,
  ): unknown {
    if (typeof error !== "object" || error === null) {
      return undefined;
    }

    try {
      return (error as Record<string, unknown>)[propertyName];
    } catch {
      return undefined;
    }
  }

  /** Builds one plain diagnostic object from an unknown Office.js-ish error payload. */
  private serializeUnknownError(
    error: unknown,
  ): SerializedOfficeErrorDiagnostics {
    let fallbackMessage = "Unknown error";
    if (error instanceof Error) {
      fallbackMessage = error.message;
    } else if (typeof error === "string") {
      fallbackMessage = error;
    }
    const messageValue = this.readUnknownErrorProperty(error, "message");
    const nameValue = this.readUnknownErrorProperty(error, "name");
    const codeValue = this.readUnknownErrorProperty(error, "code");
    const debugInfo = this.readUnknownErrorProperty(error, "debugInfo");
    const traceMessages = this.readUnknownErrorProperty(error, "traceMessages");
    const stackValue = this.readUnknownErrorProperty(error, "stack");
    const stackPreview =
      typeof stackValue === "string"
        ? stackValue
            .split(/\r?\n/u)
            .map((line) => line.trim())
            .filter((line) => line.length > 0)
            .slice(0, 5)
        : undefined;
    const serialized: SerializedOfficeErrorDiagnostics = {
      message:
        typeof messageValue === "string" && messageValue.length > 0
          ? messageValue
          : fallbackMessage,
    };
    if (typeof nameValue === "string" && nameValue.length > 0)
      serialized.name = nameValue;
    if (typeof codeValue === "string" || typeof codeValue === "number")
      serialized.code = codeValue;
    if (debugInfo !== undefined) serialized.debugInfo = debugInfo;
    if (traceMessages !== undefined) serialized.traceMessages = traceMessages;
    if (stackPreview?.length) serialized.stackPreview = stackPreview;
    return serialized;
  }

  /** Returns one best-effort text snippet only when the range text is already safely available. */
  private getLoadedRangeTextSnippet(range: Word.Range | null): string | null {
    if (!range) {
      return null;
    }

    try {
      const text = (range as { text?: unknown }).text;
      if (typeof text !== "string") {
        return null;
      }
      const normalized = (
        text as string & {
          replaceAll(pattern: RegExp, replacement: string): string;
        }
      )
        .replaceAll(/\s+/gu, " ")
        .trim();
      return normalized.length > 0 ? normalized.slice(0, 80) : null;
    } catch {
      return null;
    }
  }

  /** Builds one tracked-change diagnostic entry with its selected evidence source when known. */
  private describeTrackedChange(
    trackedChange: Word.TrackedChange,
    sources?: ReplaceTrackedChangeSources,
  ): {
    type: string;
    source?: string;
  } {
    return {
      type: trackedChange.type ?? "unknown",
      ...(sources
        ? {
            source: this.identifyTrackedChangeSource(trackedChange, sources),
          }
        : {}),
    };
  }

  /** Builds a compact tracked-change list so one observation pass can be reconstructed later. */
  private describeTrackedChanges(
    trackedChanges: Word.TrackedChange[],
    sources?: ReplaceTrackedChangeSources,
  ): Array<{
    type: string;
    source?: string;
  }> {
    return trackedChanges.map((trackedChange) =>
      this.describeTrackedChange(trackedChange, sources),
    );
  }

  /** Logs every replace evidence source exactly once before the selector starts choosing between them. */
  private logReplaceSourceDiagnostics(
    cc: Word.ContentControl,
    loadedSources: LoadedReplaceObservationSources,
  ): void {
    console.log(
      `🧾 [SuggestionResolutionObserver] suggestionId="${this.suggestion.id}" host-evidence replace-sources cc="${cc.tag}"`,
      {
        hostEvidence: {
          operationalAnchorFound: loadedSources.operationalAnchorFound,
          sources: {
            ccTrackedChanges: this.describeTrackedChanges(
              loadedSources.sources.ccTrackedChanges,
              loadedSources.sources,
            ),
            ccRangeTrackedChanges: this.describeTrackedChanges(
              loadedSources.sources.ccRangeTrackedChanges,
              loadedSources.sources,
            ),
            operationalAnchorTrackedChanges: this.describeTrackedChanges(
              loadedSources.sources.operationalAnchorTrackedChanges,
              loadedSources.sources,
            ),
            commentTrackedChanges: this.describeTrackedChanges(
              loadedSources.sources.commentTrackedChanges,
              loadedSources.sources,
            ),
          },
        },
        heuristicDiagnostics: {
          baseDebugMetadata: loadedSources.baseDebugMetadata,
        },
      },
    );
  }

  /** Resolves one stable source label for a tracked change so post-step logs stay actionable. */
  private identifyTrackedChangeSource(
    trackedChange: Word.TrackedChange | null,
    sources: ReplaceTrackedChangeSources,
  ): string {
    if (!trackedChange) {
      return "none";
    }

    const trackedChangeId = String(
      (trackedChange as { id?: string | number }).id ?? "",
    );
    const belongsToSource = (collection: Word.TrackedChange[]): boolean =>
      collection.some((candidate) => {
        if (candidate === trackedChange) {
          return true;
        }

        const candidateId = String(
          (candidate as { id?: string | number }).id ?? "",
        );
        return trackedChangeId.length > 0 && candidateId === trackedChangeId;
      });

    if (belongsToSource(sources.operationalAnchorTrackedChanges)) {
      return "operationalAnchor";
    }

    if (belongsToSource(sources.ccTrackedChanges)) {
      return "cc";
    }

    if (belongsToSource(sources.ccRangeTrackedChanges)) {
      return "ccRange";
    }

    if (belongsToSource(sources.commentTrackedChanges)) {
      return "comment";
    }

    return "unknown";
  }

  /**
   * Returns `true` only when replace evidence exposes BOTH semantic sides.
   *
   * A native Word replace tracked change is atomic from the user's perspective:
   * one side deletes the original text and the other inserts the replacement.
   * If Word currently exposes only one side, the suggestion is still pending and
   * the workflow must not certify a terminal accept/reject result yet.
   */
  private hasCompleteReplaceTrackedChangePair(
    trackedChanges: Word.TrackedChange[],
  ): boolean {
    let hasAdded = false;
    let hasDeleted = false;

    for (const trackedChange of trackedChanges) {
      if (trackedChange.type === "Added") {
        hasAdded = true;
      }

      if (trackedChange.type === "Deleted") {
        hasDeleted = true;
      }

      if (hasAdded && hasDeleted) {
        return true;
      }
    }

    return false;
  }

  /** Resolves all tracked changes semantically tied to a suggestion CC. */
  private async collectTrackedChangesForContentControl(
    context: Word.RequestContext,
    cc: Word.ContentControl,
  ): Promise<{
    trackedChanges: Word.TrackedChange[];
    ccTrackedChanges: Word.TrackedChange[];
    ccRangeTrackedChanges: Word.TrackedChange[];
    debugMetadata: Pick<
      ResolutionObservationDebugMetadata,
      "ccTrackedChangesCount" | "ccRangeTrackedChangesCount"
    >;
  }> {
    const ccRange = cc.getRange();

    this.logObserveBefore("before collecting CC tracked changes", {
      suggestionId: this.suggestion.id,
      ccTag: cc.tag,
    });
    const ccTrackedChanges = cc.getTrackedChanges();
    ccTrackedChanges.load({ select: "type,id" });

    this.logObserveBefore("before collecting CC range tracked changes", {
      suggestionId: this.suggestion.id,
      ccTag: cc.tag,
    });
    const rangeTrackedChanges = ccRange.getTrackedChanges();
    rangeTrackedChanges.load({ select: "type,id" });

    this.logObserveBefore("before collecting body tracked changes", {
      suggestionId: this.suggestion.id,
      ccTag: cc.tag,
    });
    await context.sync();

    this.logObserveBefore("after collecting CC tracked changes", {
      suggestionId: this.suggestion.id,
      ccTag: cc.tag,
      ccTrackedChangesCount: ccTrackedChanges.items.length,
    });
    this.logObserveBefore("after collecting CC range tracked changes", {
      suggestionId: this.suggestion.id,
      ccTag: cc.tag,
      ccRangeTrackedChangesCount: rangeTrackedChanges.items.length,
    });
    const trackedChangesById = new Map<string, Word.TrackedChange>();
    const trackedChangesWithoutId: Word.TrackedChange[] = [];

    const addTrackedChange = (tc: Word.TrackedChange) => {
      const id = String((tc as { id?: string | number }).id ?? "");
      if (id.length > 0) {
        trackedChangesById.set(id, tc);
      } else if (!trackedChangesWithoutId.includes(tc)) {
        trackedChangesWithoutId.push(tc);
      }
    };

    for (const tc of ccTrackedChanges.items) {
      addTrackedChange(tc);
    }

    for (const tc of rangeTrackedChanges.items) {
      addTrackedChange(tc);
    }

    const trackedChanges = [
      ...Array.from(trackedChangesById.values()),
      ...trackedChangesWithoutId,
    ];

    return {
      trackedChanges,
      ccTrackedChanges: ccTrackedChanges.items,
      ccRangeTrackedChanges: rangeTrackedChanges.items,
      debugMetadata: {
        ccTrackedChangesCount: ccTrackedChanges.items.length,
        ccRangeTrackedChangesCount: rangeTrackedChanges.items.length,
      },
    };
  }

  /** Builds debug metadata for fail-closed observations. */
  private buildEmptyReplaceDebugMetadata(
    cc: Word.ContentControl,
    colocatedComment: ColocatedCommentContext | null,
    observationStatus: SuggestionObservationStatus,
    identityVersion?: string,
  ): ResolutionObservationDebugMetadata {
    return {
      selectedCcTag: cc.tag,
      selectedCcTitleKind: (cc.title ?? "").startsWith("stylistic-meta-v2:")
        ? "operational-wrapper-v1"
        : "invalid-or-missing",
      selectedCommentFound: Boolean(colocatedComment),
      observationStatus,
      ...(identityVersion ? { identityVersion } : {}),
    };
  }

  /** Builds one identity-lost replace observation for structurally invalid metadata. */
  private buildIdentityLostReplaceObservation(
    cc: Word.ContentControl,
    colocatedComment: ColocatedCommentContext | null,
    parsedIdentity: ReplaceSuggestionIdentity | null,
  ): ReplaceObservationContext {
    return {
      identity: parsedIdentity ?? undefined,
      trackedChanges: [],
      observationStatus: "identity-lost",
      debugMetadata: this.buildEmptyReplaceDebugMetadata(
        cc,
        colocatedComment,
        "identity-lost",
        parsedIdentity?.version,
      ),
    };
  }

  /** Loads every replace evidence source once so full and side-specific observation can share it. */
  private async loadReplaceObservationSources(
    context: Word.RequestContext,
    cc: Word.ContentControl,
    colocatedComment: ColocatedCommentContext | null,
    parsedIdentity: ReplaceSuggestionIdentity,
  ): Promise<LoadedReplaceObservationSources> {
    const contentControlObservation =
      await this.collectTrackedChangesForContentControl(context, cc);

    const operationalAnchorRange = await this.resolveOperationalAnchorRange(
      context,
      parsedIdentity,
    );
    const operationalAnchorTrackedChanges = await this.loadRangeTrackedChanges(
      context,
      operationalAnchorRange,
    );

    const commentTrackedChanges = await this.loadCommentTrackedChanges(
      context,
      colocatedComment,
    );

    return {
      sources: {
        ccTrackedChanges: contentControlObservation.ccTrackedChanges,
        ccRangeTrackedChanges: contentControlObservation.ccRangeTrackedChanges,
        operationalAnchorTrackedChanges,
        commentTrackedChanges,
      },
      operationalAnchorFound: Boolean(operationalAnchorRange),
      baseDebugMetadata: {
        ccTrackedChangesCount:
          contentControlObservation.debugMetadata.ccTrackedChangesCount,
        ccRangeTrackedChangesCount:
          contentControlObservation.debugMetadata.ccRangeTrackedChangesCount,
        operationalAnchorTrackedChangesCount:
          operationalAnchorTrackedChanges.length,
        operationalAnchorFound: Boolean(operationalAnchorRange),
        commentTrackedChangesCount: commentTrackedChanges.length,
      },
    };
  }

  /** Loads replace evidence per group member so mixed-decision states can be detected before mutation. */
  private async loadOperationalWrapperGroupMemberSources(
    context: Word.RequestContext,
    group: NonNullable<ReplaceObservationContext["group"]>,
    colocatedComment: ColocatedCommentContext | null,
  ): Promise<LoadedOperationalWrapperMemberSources[]> {
    const memberSources: LoadedOperationalWrapperMemberSources[] = [];

    for (const member of group.members) {
      const loadedSources = await this.loadReplaceObservationSources(
        context,
        member.cc,
        member.cc === group.members[0]?.cc ? colocatedComment : null,
        member.identity,
      );

      memberSources.push({
        ...loadedSources,
        identity: member.identity,
      });
    }

    return memberSources;
  }

  /** Merges member evidence into the legacy source shape consumed by semantic candidate selection. */
  private mergeOperationalWrapperMemberSources(
    memberSources: LoadedOperationalWrapperMemberSources[],
  ): LoadedReplaceObservationSources {
    let merged: LoadedReplaceObservationSources | null = null;

    for (const loadedSources of memberSources) {
      if (!merged) {
        merged = {
          sources: {
            ccTrackedChanges: [...loadedSources.sources.ccTrackedChanges],
            ccRangeTrackedChanges: [
              ...loadedSources.sources.ccRangeTrackedChanges,
            ],
            operationalAnchorTrackedChanges: [
              ...loadedSources.sources.operationalAnchorTrackedChanges,
            ],
            commentTrackedChanges: [
              ...loadedSources.sources.commentTrackedChanges,
            ],
          },
          operationalAnchorFound: loadedSources.operationalAnchorFound,
          baseDebugMetadata: { ...loadedSources.baseDebugMetadata },
        };
        continue;
      }

      merged.sources.ccTrackedChanges.push(
        ...loadedSources.sources.ccTrackedChanges,
      );
      merged.sources.ccRangeTrackedChanges.push(
        ...loadedSources.sources.ccRangeTrackedChanges,
      );
      merged.sources.operationalAnchorTrackedChanges.push(
        ...loadedSources.sources.operationalAnchorTrackedChanges,
      );
      merged.sources.commentTrackedChanges.push(
        ...loadedSources.sources.commentTrackedChanges,
      );
      merged.operationalAnchorFound ||= loadedSources.operationalAnchorFound;
      merged.baseDebugMetadata.ccTrackedChangesCount =
        (merged.baseDebugMetadata.ccTrackedChangesCount ?? 0) +
        (loadedSources.baseDebugMetadata.ccTrackedChangesCount ?? 0);
      merged.baseDebugMetadata.ccRangeTrackedChangesCount =
        (merged.baseDebugMetadata.ccRangeTrackedChangesCount ?? 0) +
        (loadedSources.baseDebugMetadata.ccRangeTrackedChangesCount ?? 0);
      merged.baseDebugMetadata.operationalAnchorTrackedChangesCount =
        (merged.baseDebugMetadata.operationalAnchorTrackedChangesCount ?? 0) +
        (loadedSources.baseDebugMetadata.operationalAnchorTrackedChangesCount ??
          0);
      merged.baseDebugMetadata.commentTrackedChangesCount =
        (merged.baseDebugMetadata.commentTrackedChangesCount ?? 0) +
        (loadedSources.baseDebugMetadata.commentTrackedChangesCount ?? 0);
    }

    return (
      merged ?? {
        sources: {
          ccTrackedChanges: [],
          ccRangeTrackedChanges: [],
          operationalAnchorTrackedChanges: [],
          commentTrackedChanges: [],
        },
        operationalAnchorFound: false,
        baseDebugMetadata: {},
      }
    );
  }

  /**
   * Detects mixed user decisions in a contiguous group before any executor call.
   *
   * A healthy grouped replace member still exposes both `Deleted` and `Added`.
   * When one member exposes only one semantic side while another member exposes
   * the opposite side or a complete pair, Word has already observed incompatible
   * per-member decisions. The workflow must degrade to `mixed-group` instead of
   * trying to auto-complete the group with partial evidence.
   */
  private hasMixedGroupDecision(
    group: NonNullable<ReplaceObservationContext["group"]>,
    memberSources: LoadedOperationalWrapperMemberSources[],
  ): boolean {
    if (group.status !== "contiguous" || memberSources.length <= 1) {
      return false;
    }

    const memberSides = memberSources.map((memberSource) => {
      const semanticCandidates = this.buildReplaceSemanticCandidates(
        memberSource.sources,
      );
      return {
        hasDeleted: semanticCandidates.Deleted.length > 0,
        hasAdded: semanticCandidates.Added.length > 0,
      };
    });

    const hasCompleteMember = memberSides.some(
      (sides) => sides.hasDeleted && sides.hasAdded,
    );
    const hasOnlyDeletedMember = memberSides.some(
      (sides) => sides.hasDeleted && !sides.hasAdded,
    );
    const hasOnlyAddedMember = memberSides.some(
      (sides) => sides.hasAdded && !sides.hasDeleted,
    );

    return (
      (hasOnlyDeletedMember && hasOnlyAddedMember) ||
      (hasCompleteMember && (hasOnlyDeletedMember || hasOnlyAddedMember))
    );
  }

  /** Builds a fail-closed observation for incompatible decisions inside one explicit group. */
  private buildMixedGroupObservation(
    cc: Word.ContentControl,
    colocatedComment: ColocatedCommentContext | null,
    parsedIdentity: ReplaceSuggestionIdentity,
    group: NonNullable<ReplaceObservationContext["group"]>,
  ): ReplaceObservationContext {
    return {
      identity: parsedIdentity,
      trackedChanges: [],
      observationStatus: "mixed-group",
      debugMetadata: {
        ...this.buildEmptyReplaceDebugMetadata(
          cc,
          colocatedComment,
          "mixed-group",
          parsedIdentity.version,
        ),
        wrapperGroupId: group.groupId,
        wrapperGroupSize: group.members.length,
        wrapperGroupStatus: "mixed",
      },
      group: {
        ...group,
        status: "mixed",
      },
    };
  }

  /** Loads tracked changes for one arbitrary range if the range exists. */
  private async loadRangeTrackedChanges(
    context: Word.RequestContext,
    range: Word.Range | null,
  ): Promise<Word.TrackedChange[]> {
    if (!range) {
      return [];
    }

    const trackedChanges = range.getTrackedChanges();
    trackedChanges.load({ select: "type,id" });
    await context.sync();
    return trackedChanges.items;
  }

  /** Loads tracked changes from the colocated Stylistic comment when present. */
  private async loadCommentTrackedChanges(
    context: Word.RequestContext,
    colocatedComment: ColocatedCommentContext | null,
  ): Promise<Word.TrackedChange[]> {
    if (!colocatedComment) {
      return [];
    }

    const commentRangeTrackedChanges =
      colocatedComment.range.getTrackedChanges();
    commentRangeTrackedChanges.load({ select: "type,id" });
    await context.sync();
    return commentRangeTrackedChanges.items;
  }

  /** Builds debug metadata for a fully observed replace pair. */
  private buildReplacePairDebugMetadata(
    cc: Word.ContentControl,
    colocatedComment: ColocatedCommentContext | null,
    parsedIdentity: ReplaceSuggestionIdentity,
    trackedChanges: Word.TrackedChange[],
    baseDebugMetadata: LoadedReplaceObservationSources["baseDebugMetadata"],
  ): ResolutionObservationDebugMetadata {
    const observationStatus = this.hasCompleteReplaceTrackedChangePair(
      trackedChanges,
    )
      ? "confirmed-pending"
      : "unobservable";

    return {
      selectedCcTag: cc.tag,
      selectedCcTitleKind: "operational-wrapper-v1",
      selectedCommentFound: Boolean(colocatedComment),
      observationStatus,
      identityVersion: parsedIdentity.version,
      ...baseDebugMetadata,
    };
  }

  /** Builds debug metadata for a side-specific replace observation. */
  private buildReplaceSemanticSideDebugMetadata(
    cc: Word.ContentControl,
    colocatedComment: ColocatedCommentContext | null,
    parsedIdentity: ReplaceSuggestionIdentity,
    trackedChanges: Word.TrackedChange[],
    baseDebugMetadata: LoadedReplaceObservationSources["baseDebugMetadata"],
  ): ResolutionObservationDebugMetadata {
    const observationStatus =
      trackedChanges.length > 0 ? "confirmed-pending" : "unobservable";

    return {
      selectedCcTag: cc.tag,
      selectedCcTitleKind: "operational-wrapper-v1",
      selectedCommentFound: Boolean(colocatedComment),
      observationStatus,
      identityVersion: parsedIdentity.version,
      ...baseDebugMetadata,
    };
  }

  /** Builds the raw replace-evidence collections without collapsing them to a single winner. */
  private buildReplaceSourceCollections(
    sources: ReplaceTrackedChangeSources,
  ): Array<{
    source: string;
    trackedChanges: Word.TrackedChange[];
  }> {
    return [
      { source: "cc", trackedChanges: sources.ccTrackedChanges },
      { source: "ccRange", trackedChanges: sources.ccRangeTrackedChanges },
      {
        source: "operationalAnchor",
        trackedChanges: sources.operationalAnchorTrackedChanges,
      },
      { source: "comment", trackedChanges: sources.commentTrackedChanges },
    ];
  }

  /** Collects every executable candidate for one replace side without declaring a winner up front. */
  private collectReplaceSemanticCandidates(
    trackedChangeType: ReplaceTrackedChangeSide,
    sources: ReplaceTrackedChangeSources,
  ): ReplaceTrackedChangeCandidate[] {
    const candidatesById = new Map<string, ReplaceTrackedChangeCandidate>();
    const candidatesWithoutId: ReplaceTrackedChangeCandidate[] = [];

    for (const collection of this.buildReplaceSourceCollections(sources)) {
      for (const trackedChange of collection.trackedChanges) {
        this.collectReplaceSemanticCandidate(
          trackedChangeType,
          collection.source,
          trackedChange,
          candidatesById,
          candidatesWithoutId,
        );
      }
    }

    return [...Array.from(candidatesById.values()), ...candidatesWithoutId];
  }

  /** Adds one semantic candidate only when it matches the side and was not already recorded. */
  private collectReplaceSemanticCandidate(
    trackedChangeType: ReplaceTrackedChangeSide,
    source: string,
    trackedChange: Word.TrackedChange,
    candidatesById: Map<string, ReplaceTrackedChangeCandidate>,
    candidatesWithoutId: ReplaceTrackedChangeCandidate[],
  ): void {
    if (trackedChange.type !== trackedChangeType) {
      return;
    }

    const candidate: ReplaceTrackedChangeCandidate = {
      trackedChange,
      source,
    };
    const id = String((trackedChange as { id?: string | number }).id ?? "");

    if (id.length > 0) {
      const sourceScopedId = `${source}:${id}`;
      if (!candidatesById.has(sourceScopedId)) {
        candidatesById.set(sourceScopedId, candidate);
      }
      return;
    }

    if (
      !candidatesWithoutId.some(
        (existingCandidate) =>
          existingCandidate.trackedChange === candidate.trackedChange,
      )
    ) {
      candidatesWithoutId.push(candidate);
    }
  }

  /** Builds the exhaustive candidate lists for both replace sides. */
  private buildReplaceSemanticCandidates(
    sources: ReplaceTrackedChangeSources,
  ): ReplaceSemanticCandidateMap {
    return {
      Deleted: this.collectReplaceSemanticCandidates("Deleted", sources),
      Added: this.collectReplaceSemanticCandidates("Added", sources),
    };
  }

  /** Selects one representative tracked change per replace side for snapshots and result payloads. */
  private selectPrimaryReplaceTrackedChanges(
    semanticCandidates: ReplaceSemanticCandidateMap,
  ): Word.TrackedChange[] {
    return [
      semanticCandidates.Deleted[0]?.trackedChange,
      semanticCandidates.Added[0]?.trackedChange,
    ].filter(
      (trackedChange): trackedChange is Word.TrackedChange =>
        trackedChange !== undefined,
    );
  }

  /** Exposes all candidates for one side during side-specific re-observation. */
  private selectReplaceSemanticSideTrackedChanges(
    trackedChangeType: ReplaceTrackedChangeSide,
    semanticCandidates: ReplaceSemanticCandidateMap,
  ): Word.TrackedChange[] {
    return semanticCandidates[trackedChangeType].map(
      (candidate) => candidate.trackedChange,
    );
  }

  /** Relocates the operational anchor range persisted in operational-wrapper metadata. */
  private async resolveOperationalAnchorRange(
    context: Word.RequestContext,
    identity: ReplaceSuggestionIdentity,
  ): Promise<Word.Range | null> {
    const anchorText =
      getOperationalAnchorLocator(identity, this.suggestion) ?? "";
    if (anchorText.length === 0) {
      return null;
    }

    return this.textLocator.locate({
      context,
      container: context.document.body as unknown as WordSearchContainer,
      searchText: anchorText,
    });
  }

  /** Observes replace suggestion evidence through strict operational-wrapper metadata only. */
  private async observeReplaceSuggestion(
    context: Word.RequestContext,
    cc: Word.ContentControl,
    colocatedComment: ColocatedCommentContext | null,
  ): Promise<ReplaceObservationContext> {
    cc.load("title,tag");
    await context.sync();

    const parsedIdentity = parseReplaceIdentityTitle(
      (cc as { title?: string }).title,
    );
    if (!(cc.title ?? "").startsWith("stylistic-meta-v2:")) {
      return {
        trackedChanges: [],
        observationStatus: "ambiguous-location",
        debugMetadata: this.buildEmptyReplaceDebugMetadata(
          cc,
          colocatedComment,
          "ambiguous-location",
        ),
      };
    }

    if (!isValidOperationalReplaceIdentity(parsedIdentity, this.suggestion)) {
      return this.buildIdentityLostReplaceObservation(
        cc,
        colocatedComment,
        parsedIdentity,
      );
    }

    const group = await this.groupResolver.resolve(context, cc, parsedIdentity);
    if (group.status === "ambiguous") {
      return {
        identity: parsedIdentity,
        trackedChanges: [],
        observationStatus: "ambiguous-location",
        debugMetadata: {
          ...this.buildEmptyReplaceDebugMetadata(
            cc,
            colocatedComment,
            "ambiguous-location",
            parsedIdentity.version,
          ),
          wrapperGroupId: group.groupId,
          wrapperGroupSize: group.members.length,
          wrapperGroupStatus: group.status,
        },
        group,
      };
    }

    const groupMemberSources =
      await this.loadOperationalWrapperGroupMemberSources(
        context,
        group,
        colocatedComment,
      );
    if (this.hasMixedGroupDecision(group, groupMemberSources)) {
      return this.buildMixedGroupObservation(
        cc,
        colocatedComment,
        parsedIdentity,
        group,
      );
    }

    const loadedSources =
      this.mergeOperationalWrapperMemberSources(groupMemberSources);
    this.logReplaceSourceDiagnostics(cc, loadedSources);
    const semanticCandidates = this.buildReplaceSemanticCandidates(
      loadedSources.sources,
    );
    const trackedChanges =
      this.selectPrimaryReplaceTrackedChanges(semanticCandidates);
    const debugMetadata = this.buildReplacePairDebugMetadata(
      cc,
      colocatedComment,
      parsedIdentity,
      trackedChanges,
      loadedSources.baseDebugMetadata,
    );
    console.log(
      `🔬 [SuggestionResolutionObserver] suggestionId="${this.suggestion.id}" replace-selection cc="${cc.tag}"`,
      {
        selectedTrackedChanges: this.describeTrackedChanges(
          trackedChanges,
          loadedSources.sources,
        ),
        debugMetadata,
      },
    );

    return {
      identity: parsedIdentity,
      trackedChanges,
      observationStatus: debugMetadata.observationStatus ?? "unobservable",
      debugMetadata,
      semanticCandidates,
      group,
    };
  }

  /** Observes only one remaining semantic side of a replace after partial progress. */
  private async observeReplaceSuggestionSemanticSide(
    context: Word.RequestContext,
    cc: Word.ContentControl,
    colocatedComment: ColocatedCommentContext | null,
    trackedChangeType: ReplaceTrackedChangeSide,
  ): Promise<ReplaceObservationContext> {
    cc.load("title,tag");
    await context.sync();

    const parsedIdentity = parseReplaceIdentityTitle(
      (cc as { title?: string }).title,
    );
    if (!(cc.title ?? "").startsWith("stylistic-meta-v2:")) {
      return {
        trackedChanges: [],
        observationStatus: "ambiguous-location",
        debugMetadata: this.buildEmptyReplaceDebugMetadata(
          cc,
          colocatedComment,
          "ambiguous-location",
        ),
      };
    }

    if (!isValidOperationalReplaceIdentity(parsedIdentity, this.suggestion)) {
      return this.buildIdentityLostReplaceObservation(
        cc,
        colocatedComment,
        parsedIdentity,
      );
    }

    const group = await this.groupResolver.resolve(context, cc, parsedIdentity);
    if (group.status === "ambiguous") {
      return {
        identity: parsedIdentity,
        trackedChanges: [],
        observationStatus: "ambiguous-location",
        debugMetadata: {
          ...this.buildEmptyReplaceDebugMetadata(
            cc,
            colocatedComment,
            "ambiguous-location",
            parsedIdentity.version,
          ),
          wrapperGroupId: group.groupId,
          wrapperGroupSize: group.members.length,
          wrapperGroupStatus: group.status,
        },
        group,
      };
    }

    const groupMemberSources =
      await this.loadOperationalWrapperGroupMemberSources(
        context,
        group,
        colocatedComment,
      );
    if (this.hasMixedGroupDecision(group, groupMemberSources)) {
      return this.buildMixedGroupObservation(
        cc,
        colocatedComment,
        parsedIdentity,
        group,
      );
    }

    const loadedSources =
      this.mergeOperationalWrapperMemberSources(groupMemberSources);
    const semanticCandidates = this.buildReplaceSemanticCandidates(
      loadedSources.sources,
    );
    const trackedChanges = this.selectReplaceSemanticSideTrackedChanges(
      trackedChangeType,
      semanticCandidates,
    );
    const debugMetadata = this.buildReplaceSemanticSideDebugMetadata(
      cc,
      colocatedComment,
      parsedIdentity,
      trackedChanges,
      loadedSources.baseDebugMetadata,
    );
    console.log(
      `🔬 [SuggestionResolutionObserver] suggestionId="${this.suggestion.id}" replace-semantic-selection cc="${cc.tag}" side=${trackedChangeType}`,
      {
        selectedTrackedChanges: this.describeTrackedChanges(
          trackedChanges,
          loadedSources.sources,
        ),
        debugMetadata,
      },
    );

    return {
      identity: parsedIdentity,
      trackedChanges,
      observationStatus: debugMetadata.observationStatus ?? "unobservable",
      debugMetadata,
      semanticCandidates,
      group,
    };
  }

  /** Observes one resolution candidate using replace evidence only. */
  private async observeResolutionCandidate(
    context: Word.RequestContext,
    candidate: Word.ContentControl,
    colocatedComment: ColocatedCommentContext | null,
  ): Promise<ReplaceObservationContext> {
    const observation = await this.observeReplaceSuggestion(
      context,
      candidate,
      colocatedComment,
    );
    return {
      trackedChanges: observation.trackedChanges,
      observationStatus: observation.observationStatus,
      debugMetadata: observation.debugMetadata,
      semanticCandidates: observation.semanticCandidates,
      group: observation.group,
    };
  }

  /** Chooses the best observed candidate by scanning ranked CCs until evidence is conclusive. */
  async observeResolutionCandidates(
    context: Word.RequestContext,
    candidates: Word.ContentControl[],
    initialCc: Word.ContentControl,
  ): Promise<ResolutionObservation> {
    const observation: ResolutionObservation = {
      selectedCc: initialCc,
      selectedComment: null,
      trackedChanges: [],
      observationStatus: "unobservable",
    };

    for (const [candidateIndex, candidate] of candidates.entries()) {
      this.logObserveBefore("selected CC observation start", {
        suggestionId: this.suggestion.id,
        candidateIndex,
        candidateTag: candidate.tag,
        initialCcTag: initialCc.tag,
      });

      try {
        const colocatedComment =
          await this.locator.findColocatedStylisticComment(context, candidate);
        const candidateObservation = await this.observeResolutionCandidate(
          context,
          candidate,
          colocatedComment,
        );
        console.log(
          `🧪 [SuggestionResolutionObserver] suggestionId="${this.suggestion.id}" candidate-observation tag="${candidate.tag}"`,
          {
            observationStatus: candidateObservation.observationStatus,
            trackedChangesObserved: candidateObservation.trackedChanges.length,
            debugMetadata: candidateObservation.debugMetadata ?? null,
          },
        );

        this.logObserveBefore("selected CC observation end", {
          suggestionId: this.suggestion.id,
          candidateIndex,
          candidateTag: candidate.tag,
          observationStatus: candidateObservation.observationStatus,
          trackedChangesObserved: candidateObservation.trackedChanges.length,
        });

        observation.selectedCc = candidate;
        observation.selectedComment = colocatedComment;
        observation.trackedChanges = candidateObservation.trackedChanges;
        observation.observationStatus = candidateObservation.observationStatus;
        observation.debugMetadata = candidateObservation.debugMetadata;
        observation.semanticCandidates =
          candidateObservation.semanticCandidates;
        observation.group = candidateObservation.group;

        if (
          candidateObservation.observationStatus === "identity-lost" ||
          candidateObservation.observationStatus === "ambiguous-location" ||
          candidateObservation.observationStatus === "mixed-group" ||
          (candidateObservation.observationStatus === "confirmed-pending" &&
            candidateObservation.trackedChanges.length > 0)
        ) {
          break;
        }
      } catch (error) {
        this.warnObserveBefore("selected CC observation failed", {
          suggestionId: this.suggestion.id,
          candidateIndex,
          candidateTag: candidate.tag,
          initialCcTag: initialCc.tag,
          error: this.serializeUnknownError(error),
        });
        if (
          observation.observationStatus === "identity-lost" ||
          observation.trackedChanges.length > 0
        ) {
          continue;
        }
        throw error;
      }
    }

    return observation;
  }

  /** Re-observes only one remaining replace side after one semantic step already succeeded. */
  async observeResolutionCandidatesForSemanticSide(
    context: Word.RequestContext,
    candidates: Word.ContentControl[],
    initialCc: Word.ContentControl,
    trackedChangeType: ReplaceTrackedChangeSide,
  ): Promise<ResolutionObservation> {
    const observation: ResolutionObservation = {
      selectedCc: initialCc,
      selectedComment: null,
      trackedChanges: [],
      observationStatus: "unobservable",
    };

    for (const candidate of candidates) {
      const colocatedComment = await this.locator.findColocatedStylisticComment(
        context,
        candidate,
      );
      const candidateObservation =
        await this.observeReplaceSuggestionSemanticSide(
          context,
          candidate,
          colocatedComment,
          trackedChangeType,
        );

      observation.selectedCc = candidate;
      observation.selectedComment = colocatedComment;
      observation.trackedChanges = candidateObservation.trackedChanges;
      observation.observationStatus = candidateObservation.observationStatus;
      observation.debugMetadata = candidateObservation.debugMetadata;
      observation.semanticCandidates = candidateObservation.semanticCandidates;
      observation.group = candidateObservation.group;

      if (
        candidateObservation.observationStatus === "identity-lost" ||
        candidateObservation.observationStatus === "ambiguous-location" ||
        candidateObservation.observationStatus === "mixed-group" ||
        (candidateObservation.observationStatus === "confirmed-pending" &&
          candidateObservation.trackedChanges.length > 0)
      ) {
        break;
      }
    }

    return observation;
  }
}

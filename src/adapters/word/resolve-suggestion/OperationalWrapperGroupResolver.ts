import type { ReplaceSuggestionIdentity } from "../../../domain/suggestion/Suggestion.types";
import {
  isStructurallyValidOperationalWrapperIdentity,
  parseReplaceIdentityTitle,
} from "../ReplaceIdentityParser";
import type { OperationalWrapperGroup } from "./ResolutionContext";

/** Expands strict operational-wrapper metadata into one explicit contiguous group. */
export class OperationalWrapperGroupResolver {
  /** Resolves the seed wrapper's explicit group without distance heuristics or ranking. */
  async resolve(
    context: Word.RequestContext,
    seedCc: Word.ContentControl,
    seedIdentity: ReplaceSuggestionIdentity
  ): Promise<OperationalWrapperGroup> {
    if (seedIdentity.groupSize === 1) {
      return {
        groupId: seedIdentity.groupId,
        members: [{ cc: seedCc, identity: seedIdentity }],
        status: "single",
      };
    }

    const allContentControls = context.document.contentControls;
    allContentControls.load("items/tag,items/title");
    await context.sync();

    const members = allContentControls.items
      .map((cc) => ({ cc, identity: parseReplaceIdentityTitle(cc.title) }))
      .filter(
        (
          entry
        ): entry is {
          cc: Word.ContentControl;
          identity: ReplaceSuggestionIdentity;
        } =>
          isStructurallyValidOperationalWrapperIdentity(entry.identity) &&
          entry.identity.groupId === seedIdentity.groupId
      )
      .sort((left, right) => left.identity.groupIndex - right.identity.groupIndex);

    if (
      !this.hasCompleteGroup(
        seedIdentity,
        members.map((member) => member.identity)
      )
    ) {
      return {
        groupId: seedIdentity.groupId,
        members,
        status: "ambiguous",
      };
    }

    if (!(await this.hasContiguousWordRanges(context, members))) {
      return {
        groupId: seedIdentity.groupId,
        members,
        status: "ambiguous",
      };
    }

    return {
      groupId: seedIdentity.groupId,
      members,
      status: "contiguous",
    };
  }

  /** Confirms every explicit group slot exists exactly once. */
  private hasCompleteGroup(
    seedIdentity: ReplaceSuggestionIdentity,
    identities: ReplaceSuggestionIdentity[]
  ): boolean {
    if (identities.length !== seedIdentity.groupSize) {
      return false;
    }

    const indexes = new Set(identities.map((identity) => identity.groupIndex));
    for (let index = 0; index < seedIdentity.groupSize; index += 1) {
      if (!indexes.has(index)) {
        return false;
      }
    }

    return true;
  }

  /**
   * Validates actual Word range adjacency for consecutive group members.
   *
   * Metadata proves intended lineage, but it does not prove the wrappers still
   * form one resolvable Word block. The strongest adapter contract available is
   * `Range.compareLocationWith()` between consecutive members in group order;
   * disjoint `Before`/`After` ranges fail closed as ambiguous instead of being
   * treated as a contiguous group by metadata alone.
   */
  private async hasContiguousWordRanges(
    context: Word.RequestContext,
    members: Array<{
      cc: Word.ContentControl;
      identity: ReplaceSuggestionIdentity;
    }>
  ): Promise<boolean> {
    for (let index = 0; index < members.length - 1; index += 1) {
      const currentRange = members[index].cc.getRange();
      const nextRange = members[index + 1].cc.getRange();
      const relation = currentRange.compareLocationWith(nextRange);
      await context.sync();

      if (!this.isContiguousForwardRelation(String(relation.value ?? ""))) {
        return false;
      }
    }

    return true;
  }

  /** Returns true for adjacent or overlapping forward relations between ordered wrappers. */
  private isContiguousForwardRelation(relation: string): boolean {
    return ["AdjacentBefore", "OverlapsBefore", "Equal", "Contains", "Inside"].includes(relation);
  }
}

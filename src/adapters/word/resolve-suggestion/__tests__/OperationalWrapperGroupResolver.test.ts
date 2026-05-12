import { describe, expect, it, vi } from "vitest";
import { makeOperationalWrapperTitle } from "../../WordAdapterActionTestHelper";
import { parseReplaceIdentityTitle } from "../../ReplaceIdentityParser";
import { OperationalWrapperGroupResolver } from "../OperationalWrapperGroupResolver";

/** Builds a minimal Content Control carrying operational-wrapper metadata. */
function makeGroupCc(
  id: string,
  index: number,
  groupSize: number,
  relationWithNext = "AdjacentBefore"
) {
  return {
    tag: `stylistic:track-change:${id}`,
    title: makeOperationalWrapperTitle({
      suggestionId: id,
      insertedTag: `stylistic:track-change:${id}`,
      deletedValue: `original-${id}`,
      anchorValue: `context-${id}`,
      groupId: "group-1",
      groupIndex: index,
      groupSize,
    }),
    getRange: vi.fn(() => ({
      compareLocationWith: vi.fn(() => ({ value: relationWithNext })),
    })),
  } as unknown as Word.ContentControl;
}

/** Builds a minimal Word context exposing all document content controls. */
function makeContext(items: Word.ContentControl[]) {
  return {
    document: {
      contentControls: {
        items,
        load: vi.fn(),
      },
    },
    sync: vi.fn(async () => undefined),
  } as unknown as Word.RequestContext;
}

/** Returns the parsed operational identity or fails the test immediately. */
function requireIdentity(value: ReturnType<typeof parseReplaceIdentityTitle>) {
  if (!value) {
    throw new Error("Expected a valid operational wrapper identity in the test fixture.");
  }

  return value;
}

describe("OperationalWrapperGroupResolver", () => {
  it("resolves an explicit contiguous group for grouped rejectAll/acceptAll execution", async () => {
    const first = makeGroupCc("s-1", 0, 2);
    const second = makeGroupCc("s-2", 1, 2);
    const resolver = new OperationalWrapperGroupResolver();
    const seedIdentity = requireIdentity(parseReplaceIdentityTitle(first.title));

    const group = await resolver.resolve(makeContext([second, first]), first, seedIdentity);

    expect(group.status).toBe("contiguous");
    expect(group.members.map((member) => member.identity.suggestionId)).toEqual(["s-1", "s-2"]);
  });

  it("degrades incomplete explicit groups before any mutation can run", async () => {
    const first = makeGroupCc("s-1", 0, 2);
    const resolver = new OperationalWrapperGroupResolver();
    const seedIdentity = requireIdentity(parseReplaceIdentityTitle(first.title));

    const group = await resolver.resolve(makeContext([first]), first, seedIdentity);

    expect(group.status).toBe("ambiguous");
  });

  it("degrades metadata-complete groups when Word ranges are not adjacent or overlapping", async () => {
    const first = makeGroupCc("s-1", 0, 2, "Before");
    const second = makeGroupCc("s-2", 1, 2);
    const resolver = new OperationalWrapperGroupResolver();
    const seedIdentity = requireIdentity(parseReplaceIdentityTitle(first.title));

    const group = await resolver.resolve(makeContext([first, second]), first, seedIdentity);

    expect(group.status).toBe("ambiguous");
  });
});

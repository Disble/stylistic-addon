/**
 * Helpers for building and identifying Stylistic-generated Word comments.
 *
 * Word comments created through `Range.insertComment()` are authored by the
 * current Office user, not by an arbitrary synthetic author name. Because of
 * that, cleanup and resolution flows must identify Stylistic comments from
 * their persisted content signature, not from `authorName`.
 */

type CommentLike = {
  authorName?: string;
  content?: string;
};

/**
 * Normalizes Word comment text line endings so identification logic is stable
 * across host-specific `\n`, `\r\n`, and `\r` representations.
 */
function normalizeCommentContent(content: string | undefined): string {
  return (content ?? "").replace(/\r\n?/g, "\n");
}

/** Visible content format currently used by Stylistic-generated comments. */
export function buildStylisticCommentContent(
  category: string,
  justification: string,
): string {
  return `[${category}]\n${justification}`;
}

/**
 * Identification for Stylistic comments based on persisted content shape.
 */
export function isStylisticComment(comment: CommentLike): boolean {
  const content = normalizeCommentContent(comment.content);
  return /^\[[^\]\r\n]+\]\n[\s\S]+$/.test(content);
}

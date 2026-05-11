/** Result of deleting resolved suggestion-owned metadata artifacts. */
export type ResolvedTrackChangeMetadataCleanupResult = {
  deletedContentControls: string[];
  failedContentControls: Array<{ tag: string; error: string }>;
};

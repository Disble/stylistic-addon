import type {
  IDocumentPort,
  IUserCorrectionPreferencesPort,
  IUserPreferencesPort,
} from "../domain/ports";
import type { AnalysisProfileId } from "../domain/Profile.types";
import { ReviewSessionMediator } from "../domain/review/ReviewSessionMediator";
import type { UserCorrectionPreferences } from "../domain/user-preferences/UserCorrectionPreferences.types";

/** Collaborators required by taskpane bootstrap and settings handlers. */
export type TaskpaneBootstrapRuntime = {
  documentPort: IDocumentPort;
  reviewSessionMediator: ReviewSessionMediator;
  userCorrectionPreferencesPort: IUserCorrectionPreferencesPort;
  userPreferencesPort: IUserPreferencesPort;
  onSelectionSnapshot: Parameters<IDocumentPort["subscribeSelectionChanges"]>[0];
  setSelectedGenero: (profile: AnalysisProfileId) => void;
  supportedAnalysisProfiles: readonly AnalysisProfileId[];
};

/** Settings save handler contract extracted for reuse by the React entrypoint. */
export type SaveTaskpanePreferencesHandler = (
  correctionInstructions: string | null,
  analysisProfile: AnalysisProfileId
) => Promise<UserCorrectionPreferences>;

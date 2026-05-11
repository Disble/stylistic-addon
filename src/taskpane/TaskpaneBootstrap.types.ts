import type {
  IAuthPort,
  IAuthSessionStoragePort,
  IDocumentPort,
  IUserCorrectionPreferencesPort,
  IUserPreferencesPort,
} from "../domain/ports";
import type { AnalysisProfileId } from "../domain/Profile.types";
import { ReviewSessionMediator } from "../domain/review/ReviewSessionMediator";
import type { UserCorrectionPreferences } from "../domain/user-preferences/UserCorrectionPreferences.types";
import type { AuthSession } from "../domain/auth/AuthSession.types";

/** Collaborators required by taskpane bootstrap, auth, and settings handlers. */
export type TaskpaneBootstrapRuntime = {
  authPort: IAuthPort;
  authSessionStoragePort: IAuthSessionStoragePort;
  documentPort: IDocumentPort;
  officeDialogAuthAdapter: { signIn: () => Promise<AuthSession> };
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

/** Authenticated user summary exposed to the taskpane UI. */
export type AuthenticatedUser = Readonly<{
  id: string;
  name?: string | null;
  email?: string | null;
  image?: string | null;
}>;

/** Persisted Better Auth bearer session used by Mastra API calls. */
export type AuthSession = Readonly<{
  token: string;
  expiresAt?: string | null;
  user: AuthenticatedUser;
}>;

/** Result returned by an OAuth sign-in preparation step. */
export type SocialSignInRequest = Readonly<{
  url: string;
}>;

/** Normalized Better Auth client envelope for adapter responses. */
export type BetterAuthResponse<T> = Readonly<{
  data?: T | null;
  error?: unknown;
}>;

/** Error payload shape returned by Better Auth failures. */
export type BetterAuthErrorPayload = Readonly<{
  message?: string;
  status?: number;
  statusText?: string;
  code?: string;
}>;

/** Session payload shape consumed when mapping Better Auth sessions. */
export type BetterAuthSessionPayload = Readonly<{
  session?: {
    token?: string;
    expiresAt?: string | Date | null;
  } | null;
  user?: {
    id?: string;
    name?: string | null;
    email?: string | null;
    image?: string | null;
  } | null;
}>;

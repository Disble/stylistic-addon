import { createAuthClient } from "better-auth/client";
import type { AuthSession, SocialSignInRequest } from "../../domain/auth/AuthSession.types";
import type { IAuthPort } from "../../domain/ports";
import { BETTER_AUTH_BASE_PATH, MASTRA_BASE_URL } from "../../infrastructure/config";

type BetterAuthResponse<T> = Readonly<{
  data?: T | null;
  error?: unknown;
}>;

type BetterAuthErrorPayload = Readonly<{
  message?: string;
  status?: number;
  statusText?: string;
  code?: string;
}>;

type BetterAuthSessionPayload = Readonly<{
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

/**
 * Better Auth client adapter used by the taskpane and OAuth dialog.
 *
 * The Better Auth client exposes a strongly inferred proxy API. This adapter
 * contains the narrow runtime normalization needed by our hexagonal auth port so
 * the rest of the add-in never depends on Better Auth response shapes directly.
 */
export class BetterAuthAdapter implements IAuthPort {
  private readonly client = createAuthClient({
    baseURL: MASTRA_BASE_URL,
    basePath: BETTER_AUTH_BASE_PATH,
  });

  /** Creates a Google OAuth request without redirecting the taskpane. */
  async createSocialSignInRequest(callbackUrl: string): Promise<SocialSignInRequest> {
    const signIn = this.client.signIn as unknown as {
      social(
        input: Record<string, unknown>
      ): Promise<BetterAuthResponse<{ url?: string }> | { url?: string }>;
    };

    const response = await signIn.social({
      provider: "google",
      callbackURL: callbackUrl,
      disableRedirect: true,
    });

    const url = this.isBetterAuthResponse(response) ? response.data?.url : response.url;
    if (!url) {
      throw new Error(
        `Better Auth no devolvió una URL de inicio de sesión con Google.${this.formatAuthError(
          this.isBetterAuthResponse(response) ? response.error : undefined
        )}`
      );
    }

    return { url };
  }

  /** Resolves the current session through Better Auth, using bearer auth when supplied. */
  async getSession(token?: string): Promise<AuthSession | undefined> {
    const getSession = this.client.getSession as unknown as (
      input?: Record<string, unknown>
    ) => Promise<BetterAuthResponse<BetterAuthSessionPayload> | BetterAuthSessionPayload>;

    const headers = token ? { Authorization: `Bearer ${token}` } : undefined;
    const response = await getSession({
      fetchOptions: {
        credentials: "include",
        headers,
      },
    });

    const payload = this.isBetterAuthResponse(response) ? response.data : response;
    return this.toAuthSession(payload, token);
  }

  /** Signs out the current Better Auth session. */
  async signOut(token?: string): Promise<void> {
    const signOut = this.client.signOut as unknown as (
      input?: Record<string, unknown>
    ) => Promise<unknown>;
    const headers = token ? { Authorization: `Bearer ${token}` } : undefined;
    await signOut({
      fetchOptions: {
        credentials: "include",
        headers,
      },
    });
  }

  private toAuthSession(
    payload: BetterAuthSessionPayload | null | undefined,
    fallbackToken?: string
  ): AuthSession | undefined {
    const token = payload?.session?.token ?? fallbackToken;
    const userId = payload?.user?.id;
    if (!token || !userId) {
      return undefined;
    }

    return {
      token,
      expiresAt: this.normalizeDate(payload?.session?.expiresAt),
      user: {
        id: userId,
        name: payload?.user?.name,
        email: payload?.user?.email,
        image: payload?.user?.image,
      },
    };
  }

  private isBetterAuthResponse<T>(
    value: BetterAuthResponse<T> | T
  ): value is BetterAuthResponse<T> {
    return typeof value === "object" && value !== null && ("data" in value || "error" in value);
  }

  private normalizeDate(value: string | Date | null | undefined): string | null {
    if (!value) {
      return null;
    }
    return value instanceof Date ? value.toISOString() : value;
  }

  private formatAuthError(error: unknown): string {
    if (!error) {
      return " Revisá que el backend tenga GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET y las tablas de Better Auth migradas.";
    }

    if (typeof error === "string") {
      return ` Error: ${error}`;
    }

    if (typeof error !== "object") {
      return "";
    }

    const payload = error as BetterAuthErrorPayload;
    const details = [payload.status, payload.statusText, payload.code, payload.message]
      .filter(Boolean)
      .join(" - ");

    return details ? ` Error: ${details}` : "";
  }
}

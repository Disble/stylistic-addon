import { MastraClient } from "@mastra/client-js";
import { MASTRA_BASE_URL } from "../../infrastructure/config";
import type { MastraAuthTokenProvider } from "./MastraClientFactory.types";

/**
 * Creates Mastra clients with the latest Better Auth bearer token.
 *
 * The factory intentionally creates clients on demand instead of exporting a
 * singleton. Auth state changes after login/logout, and a singleton would freeze
 * stale headers into later workflow calls.
 */
export class MastraClientFactory {
  constructor(private readonly getAuthToken: MastraAuthTokenProvider = () => undefined) {}

  /** Builds a client instance using the current bearer token snapshot. */
  create(): MastraClient {
    const token = this.getAuthToken();
    return new MastraClient({
      baseUrl: MASTRA_BASE_URL,
      headers: token ? { Authorization: `Bearer ${token}` } : undefined,
    });
  }
}

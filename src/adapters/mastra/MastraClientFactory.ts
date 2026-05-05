import { MastraClient } from "@mastra/client-js";
import { MASTRA_BASE_URL } from "../../infrastructure/config";

/** Supplies the current bearer token for Mastra requests. */
export type MastraAuthTokenProvider = () => string | undefined;

/** Creates Mastra clients with the latest Better Auth bearer token. */
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

/* global console */

/**
 * MockFeedbackAdapter — dev/test stub implementing `IFeedbackPort`.
 *
 * Logs the payload to the console and makes no network calls.
 * Used in `taskpane.ts` until the backend feedback workflow is ready.
 * Swap to `FeedbackAdapter` with a one-line change in `taskpane.ts`.
 *
 * @module MockFeedbackAdapter
 */

import { IFeedbackPort } from "../../domain/ports";
import { FeedbackPayload } from "../../domain/types";

export class MockFeedbackAdapter implements IFeedbackPort {
  async sendFeedback(payload: FeedbackPayload): Promise<void> {
    console.log("[MockFeedbackAdapter] sendFeedback:", payload);
  }
}

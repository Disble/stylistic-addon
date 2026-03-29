/* global console */

/**
 * FeedbackAdapter — implements `IFeedbackPort` using `@mastra/client-js`.
 *
 * Sends user feedback to the Mastra feedback workflow via fire-and-forget.
 * All errors are swallowed silently — feedback failures must never surface to the user.
 *
 * Pattern: createRun() + run.start() (same as MastraAdapter — .execute() does not exist
 * in @mastra/client-js v1.7.1).
 *
 * @module FeedbackAdapter
 */

import { MastraClient } from "@mastra/client-js";
import type { IFeedbackPort } from "../../domain/ports";
import type { FeedbackPayload } from "../../domain/types";
import {
  FEEDBACK_WORKFLOW_ID,
  MASTRA_BASE_URL,
} from "../../infrastructure/config";

/** Singleton Mastra client instance, reused across all feedback calls. */
const mastraClient = new MastraClient({ baseUrl: MASTRA_BASE_URL });

export class FeedbackAdapter implements IFeedbackPort {
  /**
   * Sends a feedback payload to the Mastra feedback workflow.
   * Fire-and-forget: never throws, errors swallowed silently.
   */
  async sendFeedback(payload: FeedbackPayload): Promise<void> {
    try {
      console.log("[FeedbackAdapter] Sending feedback", {
        workflowId: FEEDBACK_WORKFLOW_ID,
        payload,
      });

      const workflow = mastraClient.getWorkflow(FEEDBACK_WORKFLOW_ID);
      console.log("[FeedbackAdapter] Workflow reference obtained");

      const run = await workflow.createRun();
      console.log("[FeedbackAdapter] Run created");

      const result = await run.start({ inputData: payload });
      console.log("[FeedbackAdapter] Workflow started successfully", {
        result,
      });
    } catch (error) {
      console.error("[FeedbackAdapter] Failed to send feedback", error);
    }
  }
}

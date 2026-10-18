/* global console */

/**
 * Hub Feedback Adapter — implements `IFeedbackPort` against Stylistic Hub API v1.
 *
 * Posts `{ suggestionId, action, comment? }` to `POST /feedback`; the hub
 * already knows the suggestion by its id and learns from it in the background.
 * Fire-and-forget: every failure is logged and swallowed, because feedback
 * must never surface to the user.
 *
 * @module HubFeedbackAdapter
 */

import type { IFeedbackPort } from "../../domain/ports";
import type { FeedbackPayload } from "../../domain/suggestion/SuggestionResolutionWorkflow.types";
import type { HttpClient } from "../../infrastructure/http/HttpClient";
import { readHubErrorDetail } from "./HubProblemParser.helpers";

/** Sends non-blocking feedback about resolved suggestions to the hub. */
export class HubFeedbackAdapter implements IFeedbackPort {
  constructor(private readonly http: HttpClient) {}

  /** Sends one feedback event. Never throws. */
  async sendFeedback(payload: FeedbackPayload): Promise<void> {
    try {
      await this.http.post("/feedback", payload);
    } catch (error) {
      console.error(`[HubFeedbackAdapter] Feedback not sent: ${readHubErrorDetail(error)}`);
    }
  }
}

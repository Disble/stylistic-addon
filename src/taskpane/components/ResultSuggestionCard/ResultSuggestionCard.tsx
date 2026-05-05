import * as React from "react";
import {
  Badge,
  Body1,
  Body1Strong,
  Body2,
  Button,
  Caption1,
  Caption1Strong,
  Card,
  Textarea,
} from "@fluentui/react-components";
import {
  ArrowRightRegular,
  CheckmarkRegular,
  CommentRegular,
  DismissRegular,
} from "@fluentui/react-icons";
import { getFailedSuggestionCopy } from "./ResultSuggestionCard.helpers";
import type { ResultSuggestionCardProps } from "./ResultSuggestionCard.types";
import { useResultSuggestionCard } from "./useResultSuggestionCard";

/** Renders one suggestion card with Fluent UI v9 components. */
export function ResultSuggestionCard({
  card,
  onAccept,
  onFeedbackCommentChange,
  onNavigate,
  onReject,
  onToggleFeedback,
}: ResultSuggestionCardProps): React.JSX.Element {
  const view = useResultSuggestionCard(card);
  const { classes, isCommentOnly } = view;

  return (
    <li
      className={classes.root}
      data-severity={card.suggestion.severity}
      data-suggestion-id={card.suggestion.id}
    >
      <Card className={classes.card}>
        <span className={classes.severityStripe} aria-hidden="true" />

        <div className={classes.badgeRow}>
          <Badge appearance="filled" color="brand">
            {card.suggestion.category}
          </Badge>
          {!card.isFailed ? (
            <Badge appearance="tint" color={view.severityColor}>
              {view.severityLabel}
            </Badge>
          ) : null}
          {isCommentOnly ? (
            <Badge
              appearance="ghost"
              color="informative"
              data-testid="card-comment-badge"
              icon={<CommentRegular aria-hidden="true" />}
            >
              comentario
            </Badge>
          ) : null}
        </div>

        {card.failure ? (
          <>
            <Caption1Strong className={classes.failureLabel}>
              {getFailedSuggestionCopy(card.failure)}
            </Caption1Strong>
            <Body2 className={classes.justification}>{card.suggestion.justification}</Body2>
            {card.failure.reason !== "not-found" ? (
              <Caption1 className={classes.failureDetail}>{card.failure.message}</Caption1>
            ) : null}
          </>
        ) : (
          <button
            type="button"
            className={classes.clickableArea}
            aria-label="Ir a la sugerencia en el documento"
            onClick={() => {
              void onNavigate(card.suggestion.id);
            }}
          >
            {!isCommentOnly ? (
              <div className={classes.diff} data-testid="card-diff">
                <Body1 className={classes.original}>{card.suggestion.anchor}</Body1>
                <ArrowRightRegular className={classes.diffArrow} aria-hidden="true" />
                <Body1Strong className={classes.suggested}>
                  {card.suggestion.suggestedText ?? ""}
                </Body1Strong>
              </div>
            ) : null}
            <Body2 className={classes.justification}>{card.suggestion.justification}</Body2>
          </button>
        )}

        {card.navigationNote ? (
          <Caption1 className={classes.note}>{card.navigationNote}</Caption1>
        ) : null}

        {card.resolutionNote ? (
          <Caption1 className={classes.note}>{card.resolutionNote}</Caption1>
        ) : null}

        {!card.hideActions ? (
          <div className={classes.footer}>
            <div className={classes.actions}>
              <Button
                appearance="primary"
                aria-label={view.acceptAriaLabel}
                className={classes.primaryAction}
                data-action="accept"
                data-suggestion-id={card.suggestion.id}
                disabled={card.isResolving}
                icon={<CheckmarkRegular aria-hidden="true" />}
                onClick={() => {
                  void onAccept(card.suggestion.id);
                }}
                type="button"
              >
                {view.acceptLabel}
              </Button>
              <Button
                appearance="secondary"
                aria-label={view.rejectAriaLabel}
                data-action="reject"
                data-suggestion-id={card.suggestion.id}
                disabled={card.isResolving}
                icon={<DismissRegular aria-hidden="true" />}
                onClick={() => {
                  void onReject(card.suggestion.id);
                }}
                type="button"
              >
                {view.rejectLabel}
              </Button>
              <Button
                appearance="subtle"
                aria-expanded={card.feedbackOpen}
                aria-label={view.feedbackToggleAriaLabel}
                data-testid="card-feedback-toggle"
                icon={<CommentRegular aria-hidden="true" />}
                onClick={() => {
                  onToggleFeedback(card.suggestion.id);
                }}
                type="button"
              />
            </div>

            {card.feedbackOpen ? (
              <Textarea
                aria-label="Comentario de feedback"
                className={classes.feedbackTextarea}
                data-testid="card-feedback-textarea"
                onChange={(_event, data) => {
                  onFeedbackCommentChange(card.suggestion.id, data.value);
                }}
                placeholder="Comentario opcional..."
                resize="vertical"
                value={card.feedbackComment}
              />
            ) : null}
          </div>
        ) : null}
      </Card>
    </li>
  );
}

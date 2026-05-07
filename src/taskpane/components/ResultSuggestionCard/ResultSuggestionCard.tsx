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
  CheckmarkCircleFilled,
  CheckmarkRegular,
  CommentRegular,
  DismissCircleFilled,
  DismissRegular,
  ErrorCircleFilled,
  SearchInfoFilled,
} from "@fluentui/react-icons";
import { STATUS_ICON_LABEL } from "./ResultSuggestionCard.constants";
import { getFailedSuggestionCopy } from "./ResultSuggestionCard.helpers";
import type { CardVisualState, ResultSuggestionCardProps } from "./ResultSuggestionCard.types";
import { useResultSuggestionCard } from "./ResultSuggestionCard.hooks";

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
  const statusIconLabel = STATUS_ICON_LABEL[view.cardVisualState];
  const statusIconByState: Partial<Record<CardVisualState, React.JSX.Element>> = {
    accepted: <CheckmarkCircleFilled />,
    rejected: <DismissCircleFilled />,
    failed: <ErrorCircleFilled />,
    "not-found": <SearchInfoFilled />,
  };

  const renderStatusIcon = (): React.JSX.Element | null => {
    if (!statusIconLabel) {
      return null;
    }

    return (
      <span
        aria-label={statusIconLabel}
        className={classes.statusIcon}
        data-testid={`card-status-icon-${view.cardVisualState}`}
        role="img"
      >
        {statusIconByState[view.cardVisualState] ?? null}
      </span>
    );
  };

  const renderFailureContent = (): React.JSX.Element => {
    const failureDetail = card.failure?.reason !== "not-found" ? card.failure?.message : null;

    return (
      <>
        <Caption1Strong className={classes.failureLabel}>
          {getFailedSuggestionCopy(card.failure!)}
        </Caption1Strong>
        <Body2 className={classes.justification}>{card.suggestion.justification}</Body2>
        {failureDetail ? <Caption1 className={classes.failureDetail}>{failureDetail}</Caption1> : null}
      </>
    );
  };

  const renderSuggestionContent = (): React.JSX.Element => (
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
          <Body1Strong className={classes.suggested}>{card.suggestion.suggestedText ?? ""}</Body1Strong>
        </div>
      ) : null}
      <Body2 className={classes.justification}>{card.suggestion.justification}</Body2>
    </button>
  );

  const renderNotes = (): React.JSX.Element | null => {
    if (!card.navigationNote && !card.resolutionNote) {
      return null;
    }

    return (
      <>
        {card.navigationNote ? <Caption1 className={classes.note}>{card.navigationNote}</Caption1> : null}
        {card.resolutionNote ? <Caption1 className={classes.note}>{card.resolutionNote}</Caption1> : null}
      </>
    );
  };

  const renderActions = (): React.JSX.Element | null => {
    if (card.hideActions) {
      return null;
    }

    return (
      <div className={classes.footer}>
        <div className={classes.actions}>
          <Button
            appearance="primary"
            aria-label={view.acceptAriaLabel}
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
    );
  };

  return (
    <li
      className={classes.root}
      data-card-state={view.cardVisualState}
      data-category-accent={view.categoryAccent}
      data-severity={card.suggestion.severity}
      data-suggestion-id={card.suggestion.id}
    >
      <Card className={classes.card}>
        <span className={classes.stateStripe} aria-hidden="true" />

        <div className={classes.badgeRow}>
          <span className={classes.categoryPill} data-testid="card-category-pill">
            {card.suggestion.category}
          </span>
          {!card.isFailed ? (
            <span
              aria-label={`Severidad ${view.severityLabel}`}
              className={classes.severityIndicator}
              data-testid="card-severity-indicator"
            >
              <span className={classes.severityDot} aria-hidden="true" />
              <span className={classes.severityLabel}>{view.severityLabel}</span>
            </span>
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
          {renderStatusIcon()}
        </div>

        {card.failure ? renderFailureContent() : renderSuggestionContent()}

        {renderNotes()}

        {renderActions()}
      </Card>
    </li>
  );
}

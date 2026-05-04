import * as React from "react";
import {
  getFailedSuggestionCopy,
  getResultSuggestionCardClassName,
  isCommentOnlyCard,
} from "./ResultSuggestionCard.helpers";
import type { ResultSuggestionCardProps } from "./ResultSuggestionCard.types";

/** Renders one suggestion card with navigation and resolution interactions. */
export function ResultSuggestionCard({
  card,
  onAccept,
  onFeedbackCommentChange,
  onNavigate,
  onReject,
  onToggleFeedback,
}: ResultSuggestionCardProps): React.JSX.Element {
  const isCommentOnly = isCommentOnlyCard(card);

  return (
    <li className={getResultSuggestionCardClassName(card)} data-severity={card.suggestion.severity}>
      <div className="card-meta">
        <span className="result-category">{card.suggestion.category}</span>
        {!card.isFailed ? (
          <>
            <span className={`result-severity result-severity--${card.suggestion.severity}`}>
              {card.suggestion.severity}
            </span>
            {isCommentOnly ? (
              <span className="result-type-badge result-type-badge--comment">comentario</span>
            ) : null}
          </>
        ) : null}
      </div>

      {card.failure ? (
        <>
          <span className="result-failed">{getFailedSuggestionCopy(card.failure)}</span>
          <span className="result-justification">{card.suggestion.justification}</span>
          {card.failure.reason !== "not-found" ? (
            <span className="result-failure-detail">{card.failure.message}</span>
          ) : null}
        </>
      ) : (
        <>
          <div
            className="card-clickable-area"
            onClick={() => {
              void onNavigate(card.suggestion.id);
            }}
          >
            {!isCommentOnly ? (
              <div className="card-diff">
                <span className="result-original">{card.suggestion.anchor}</span>
                <span className="result-suggested">{card.suggestion.suggestedText ?? ""}</span>
              </div>
            ) : null}
            <span className="result-justification">{card.suggestion.justification}</span>
          </div>

          {card.navigationNote ? (
            <span className="result-navigation-note">{card.navigationNote}</span>
          ) : null}

          {card.resolutionNote ? (
            <span className={`result-${card.state}-note`}>{card.resolutionNote}</span>
          ) : null}

          {!card.hideActions ? (
            <>
              <div className="card-footer">
                <span className="result-actions">
                  <button
                    aria-label="Aceptar sugerencia"
                    className={
                      isCommentOnly
                        ? "result-action-btn result-action-btn--text"
                        : "result-action-btn"
                    }
                    data-action="accept"
                    data-suggestion-id={card.suggestion.id}
                    disabled={card.isResolving}
                    onClick={() => {
                      void onAccept(card.suggestion.id);
                    }}
                    type="button"
                  >
                    {isCommentOnly ? "Entendido" : "✓"}
                  </button>
                  <button
                    aria-label="Rechazar sugerencia"
                    className={
                      isCommentOnly
                        ? "result-action-btn result-action-btn--text"
                        : "result-action-btn"
                    }
                    data-action="reject"
                    data-suggestion-id={card.suggestion.id}
                    disabled={card.isResolving}
                    onClick={() => {
                      void onReject(card.suggestion.id);
                    }}
                    type="button"
                  >
                    {isCommentOnly ? "Ignorar" : "✗"}
                  </button>
                  <button
                    aria-label="Dejar feedback"
                    className="feedback-btn"
                    data-action="feedback"
                    onClick={() => {
                      onToggleFeedback(card.suggestion.id);
                    }}
                    type="button"
                  >
                    💬
                  </button>
                </span>
              </div>

              <div
                className={`feedback-accordion${card.feedbackOpen ? " feedback-accordion--open" : ""}`}
              >
                <textarea
                  className="feedback-textarea"
                  onChange={(event) => {
                    onFeedbackCommentChange(card.suggestion.id, event.target.value);
                  }}
                  placeholder="Comentario opcional..."
                  value={card.feedbackComment}
                />
              </div>
            </>
          ) : null}
        </>
      )}
    </li>
  );
}

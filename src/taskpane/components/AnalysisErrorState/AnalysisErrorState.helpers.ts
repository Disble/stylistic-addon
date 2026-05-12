import type {
  AnalysisErrorRetryPresentation,
  AnalysisErrorStateProps,
} from "./AnalysisErrorState.types";

/** Returns the retry copy associated with one error retry mode. */
export function resolveRetryPresentation(
  retryKind: AnalysisErrorStateProps["error"]["retryKind"]
): AnalysisErrorRetryPresentation {
  if (retryKind === "retry-query") {
    return {
      title: "No pudimos recuperar el resultado",
      guidance: "Reintentá la consulta del mismo run sin reenviar el texto al backend.",
      actionLabel: "Reintentar consulta",
    };
  }

  return {
    title: "El análisis no pudo completarse",
    guidance: "Reintentá el análisis completo para volver a enviar la petición al backend.",
    actionLabel: "Reintentar análisis",
  };
}

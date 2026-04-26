import type { WorkflowOutput } from "../../domain/mastra/MastraWorkflow.types";

/**
 * Deterministic development output returned when the Mastra poll bypass is
 * enabled from `config.ts`.
 *
 * Keep this payload structurally identical to the backend success contract so
 * the rest of the pipeline behaves exactly as if Mastra had returned it.
 */
export const MOCK_MASTRA_POLL_OUTPUT: WorkflowOutput = {
  suggestions: [
    {
      type: "track-change",
      anchor: "ni Shu",
      context:
        "Xia no tenía idea de lo que estaba pasando por la mente de Mei ni Shu.",
      category: "gramática",
      severity: "high",
      justification: "Omisión de preposición en estructuras paralelas.",
      suggestedText: "ni de Shu",
    },
    {
      type: "track-change",
      anchor: "eran meras suposiciones",
      context: "Cualquier cosa que pudiera decir eran meras suposiciones.",
      category: "gramática",
      severity: "high",
      justification: "Concordancia verbal con sujeto singular.",
      suggestedText: "era una mera suposición",
    },
    {
      type: "track-change",
      anchor: "se quedaron en silencio",
      context: "Todas se quedaron en silencio.",
      category: "estilo",
      severity: "medium",
      justification:
        "Eco léxico por repetición de 'silencio' en párrafos cercanos.",
      suggestedText: "guardaron silencio",
    },
    {
      type: "comment-only",
      anchor: "desde allí",
      context:
        "desperté en unas instalaciones de WEPO lejos de las garras de Jack y desde allí no volví a saber de él.",
      category: "estilo",
      severity: "low",
      justification:
        "Evaluar si se refiere a la ubicación (correcto) o al momento temporal (debería ser 'desde entonces').",
    },
  ],
};

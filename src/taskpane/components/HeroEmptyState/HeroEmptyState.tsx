import * as React from "react";
import type { HeroEmptyStateProps } from "./HeroEmptyState.types";
import { useHeroEmptyState } from "./HeroEmptyState.hooks";

/**
 * First-impression surface shown when the taskpane is idle (no analysis
 * running, no progress, no rendered results). Centers a brand illustration
 * with a short headline and exposes the primary CTA(s) below it via
 * `children` so the analyze flow stays visually associated with the artwork.
 *
 * The inline SVG paints a stylized document plus a magic wand and three
 * sparkles. It uses `currentColor` so the wrapper themes it via Fluent's
 * brand foreground token. Sparkle classes are wired in from the parent
 * Griffel hook so animation timings live next to the rest of the styles.
 */
export function HeroEmptyState({ children }: HeroEmptyStateProps): React.JSX.Element {
  const classes = useHeroEmptyState();

  return (
    <section
      aria-labelledby="hero-empty-state-title"
      className={classes.root}
      data-testid="hero-empty-state"
    >
      <div className={classes.illustrationWrapper}>
        <svg
          aria-hidden="true"
          className={classes.illustration}
          fill="none"
          role="img"
          viewBox="0 0 240 200"
          xmlns="http://www.w3.org/2000/svg"
        >
          <ellipse cx="120" cy="184" fill="currentColor" opacity="0.08" rx="78" ry="6" />

          <rect
            fill="currentColor"
            height="130"
            opacity="0.06"
            rx="8"
            width="118"
            x="64"
            y="38"
          />
          <rect
            fill="white"
            height="130"
            rx="8"
            stroke="currentColor"
            strokeWidth="1.5"
            width="118"
            x="56"
            y="32"
          />

          <rect fill="currentColor" height="5" opacity="0.30" rx="2.5" width="68" x="68" y="54" />
          <rect fill="currentColor" height="5" opacity="0.18" rx="2.5" width="92" x="68" y="68" />
          <rect fill="currentColor" height="5" opacity="0.18" rx="2.5" width="56" x="68" y="82" />

          <rect fill="currentColor" height="5" opacity="0.18" rx="2.5" width="92" x="68" y="106" />
          <rect fill="currentColor" height="5" opacity="0.18" rx="2.5" width="78" x="68" y="120" />
          <rect fill="currentColor" height="5" opacity="0.18" rx="2.5" width="44" x="68" y="134" />

          <line
            stroke="currentColor"
            strokeLinecap="round"
            strokeWidth="4"
            x1="190"
            x2="148"
            y1="22"
            y2="64"
          />
          <path
            d="M198 12 L201.5 19 L208.5 22 L201.5 25 L198 32 L194.5 25 L187.5 22 L194.5 19 Z"
            fill="currentColor"
          />

          <g className={classes.sparkle1}>
            <path
              d="M218 56 L219.2 60.5 L223.5 61.7 L219.2 62.9 L218 67.4 L216.8 62.9 L212.5 61.7 L216.8 60.5 Z"
              fill="currentColor"
            />
          </g>
          <g className={classes.sparkle2}>
            <path
              d="M28 96 L29 99.5 L32.5 100.5 L29 101.5 L28 105 L27 101.5 L23.5 100.5 L27 99.5 Z"
              fill="currentColor"
            />
          </g>
          <g className={classes.sparkle3}>
            <circle cx="206" cy="108" fill="currentColor" r="2.5" />
          </g>
        </svg>
        <div>
          <h2 className={classes.title} id="hero-empty-state-title">
            Listo para revisar tu texto
          </h2>
          <p className={classes.subtitle}>
            Analizá el documento completo o seleccioná un fragmento para empezar.
          </p>
        </div>
      </div>
      <div className={classes.actions}>{children}</div>
    </section>
  );
}

import type { AuthSectionClasses } from "./AuthSection.types";
import { useAuthSectionStyles } from "./AuthSection.styles";

/** Returns Griffel classes for the authentication section. */
export function useAuthSection(): AuthSectionClasses {
  const styles = useAuthSectionStyles();
  return {
    root: styles.root,
    content: styles.content,
    title: styles.title,
    description: styles.description,
    button: styles.button,
  };
}

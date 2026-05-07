import type { TaskpaneStatusType } from "../../TaskpaneShellStore.types";
import type { StatusBarIntent } from "./StatusBar.types";

/** Maps shell status slugs to the corresponding Fluent MessageBar intent. */
export const STATUS_BAR_INTENT: Record<TaskpaneStatusType, StatusBarIntent> = {
  success: "success",
  error: "error",
};

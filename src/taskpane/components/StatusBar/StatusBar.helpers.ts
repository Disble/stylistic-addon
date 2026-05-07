import type { TaskpaneStatusType } from "../../TaskpaneShellStore.types";
import { STATUS_BAR_INTENT } from "./StatusBar.constants";
import type { StatusBarIntent } from "./StatusBar.types";

/** Maps the taskpane status type to the matching Fluent MessageBar intent. */
export function getStatusBarIntent(type: TaskpaneStatusType): StatusBarIntent {
  return STATUS_BAR_INTENT[type] ?? "info";
}

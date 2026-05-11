/**
 * Maps unknown runtime failures to taskpane-friendly user messages.
 */
export function toTaskpaneUserMessage(error: unknown): string {
  if (!(error instanceof Error)) {
    if (typeof error === "string") {
      return error;
    }

    try {
      return JSON.stringify(error);
    } catch {
      return "Error no serializable";
    }
  }

  const officeError = error as Error & { code?: string };
  switch (officeError.code) {
    case "AccessDenied":
      return "El documento está protegido o es de solo lectura.";
    case "InvalidArgument":
      return "Argumento inválido al comunicarse con Word.";
    case "ItemNotFound":
      return "No se encontró el elemento solicitado en el documento.";
    default:
      return officeError.message;
  }
}

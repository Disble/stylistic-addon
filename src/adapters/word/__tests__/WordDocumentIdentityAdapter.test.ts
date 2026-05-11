import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { DOCUMENT_UUID_SETTINGS_KEY } from "../../../infrastructure/config";
import { installOfficeDocumentSettings } from "../WordAdapterTestHelper";
import { WordDocumentIdentityAdapter } from "../WordDocumentIdentityAdapter";

describe("WordDocumentIdentityAdapter", () => {
  let adapter: WordDocumentIdentityAdapter;

  beforeEach(() => {
    adapter = new WordDocumentIdentityAdapter();
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("returns the persisted UUID when the document already has one", async () => {
    const existingUuid = "11111111-1111-4111-8111-111111111111";
    const office = installOfficeDocumentSettings({ existingValue: existingUuid });

    await expect(adapter.getDocumentUuid()).resolves.toBe(existingUuid);

    expect(office.get).toHaveBeenCalledWith(DOCUMENT_UUID_SETTINGS_KEY);
    expect(office.set).not.toHaveBeenCalled();
    expect(office.saveAsync).not.toHaveBeenCalled();
  });

  it("creates and persists a UUID when the document does not have one", async () => {
    installOfficeDocumentSettings();
    const randomUuidSpy = vi
      .spyOn(globalThis.crypto, "randomUUID")
      .mockReturnValue("22222222-2222-4222-8222-222222222222");

    await expect(adapter.getDocumentUuid()).resolves.toBe("22222222-2222-4222-8222-222222222222");

    randomUuidSpy.mockRestore();
  });

  it("replaces invalid persisted values with a new UUID", async () => {
    const office = installOfficeDocumentSettings({ existingValue: "not-a-uuid" });
    const randomUuidSpy = vi
      .spyOn(globalThis.crypto, "randomUUID")
      .mockReturnValue("33333333-3333-4333-8333-333333333333");

    await expect(adapter.getDocumentUuid()).resolves.toBe("33333333-3333-4333-8333-333333333333");

    expect(office.set).toHaveBeenCalledWith(
      DOCUMENT_UUID_SETTINGS_KEY,
      "33333333-3333-4333-8333-333333333333"
    );

    randomUuidSpy.mockRestore();
  });

  it("fails closed when Office document settings are unavailable", async () => {
    const officeGlobal = globalThis as unknown as { Office?: unknown };
    delete officeGlobal.Office;

    await expect(adapter.getDocumentUuid()).rejects.toThrow(
      "Tu versión de Office no soporta la identidad persistida de documento para Stylistic."
    );
  });

  it("surfaces persistence errors when settings.saveAsync fails", async () => {
    installOfficeDocumentSettings({ saveErrorMessage: "save failed" });
    const randomUuidSpy = vi
      .spyOn(globalThis.crypto, "randomUUID")
      .mockReturnValue("44444444-4444-4444-8444-444444444444");

    await expect(adapter.getDocumentUuid()).rejects.toThrow("save failed");

    randomUuidSpy.mockRestore();
  });
});

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { ANALYSIS_PROFILE_STORAGE_KEY } from "../../../infrastructure/config";
import { OfficeUserPreferencesAdapter } from "../OfficeUserPreferencesAdapter";

type StorageMock = {
  getItem: ReturnType<typeof vi.fn>;
  setItem: ReturnType<typeof vi.fn>;
  removeItem: ReturnType<typeof vi.fn>;
};

function installOfficeRuntime(storage: StorageMock | undefined): void {
  (globalThis as unknown as { OfficeRuntime?: { storage?: StorageMock } }).OfficeRuntime = storage
    ? { storage }
    : undefined;
}

function makeStorageMock(): StorageMock {
  return {
    getItem: vi.fn(),
    setItem: vi.fn(),
    removeItem: vi.fn(),
  };
}

describe("OfficeUserPreferencesAdapter.getAnalysisProfile", () => {
  let storage: StorageMock;

  beforeEach(() => {
    storage = makeStorageMock();
    installOfficeRuntime(storage);
  });

  afterEach(() => {
    installOfficeRuntime(undefined);
  });

  it("returns undefined when storage has no value", async () => {
    storage.getItem.mockResolvedValueOnce(null);
    const adapter = new OfficeUserPreferencesAdapter();

    await expect(adapter.getAnalysisProfile()).resolves.toBeUndefined();
    expect(storage.getItem).toHaveBeenCalledWith(ANALYSIS_PROFILE_STORAGE_KEY);
  });

  it("returns the stored profile when present", async () => {
    storage.getItem.mockResolvedValueOnce("ensayo-academico");
    const adapter = new OfficeUserPreferencesAdapter();

    await expect(adapter.getAnalysisProfile()).resolves.toBe("ensayo-academico");
  });

  it("returns undefined when the stored value is an empty string", async () => {
    storage.getItem.mockResolvedValueOnce("");
    const adapter = new OfficeUserPreferencesAdapter();

    await expect(adapter.getAnalysisProfile()).resolves.toBeUndefined();
  });

  it("returns undefined when the stored value is not a string", async () => {
    storage.getItem.mockResolvedValueOnce(42 as unknown as string);
    const adapter = new OfficeUserPreferencesAdapter();

    await expect(adapter.getAnalysisProfile()).resolves.toBeUndefined();
  });
});

describe("OfficeUserPreferencesAdapter.setAnalysisProfile", () => {
  let storage: StorageMock;

  beforeEach(() => {
    storage = makeStorageMock();
    installOfficeRuntime(storage);
  });

  afterEach(() => {
    installOfficeRuntime(undefined);
  });

  it("persists the value under the analysis-profile storage key", async () => {
    storage.setItem.mockResolvedValueOnce(undefined);
    const adapter = new OfficeUserPreferencesAdapter();

    await adapter.setAnalysisProfile("periodismo-cultural");

    expect(storage.setItem).toHaveBeenCalledWith(
      ANALYSIS_PROFILE_STORAGE_KEY,
      "periodismo-cultural"
    );
  });
});

describe("OfficeUserPreferencesAdapter — host without OfficeRuntime", () => {
  beforeEach(() => {
    installOfficeRuntime(undefined);
  });

  afterEach(() => {
    installOfficeRuntime(undefined);
  });

  it("throws when reading from a host without OfficeRuntime.storage", async () => {
    const adapter = new OfficeUserPreferencesAdapter();
    await expect(adapter.getAnalysisProfile()).rejects.toThrow();
  });

  it("throws when writing to a host without OfficeRuntime.storage", async () => {
    const adapter = new OfficeUserPreferencesAdapter();
    await expect(adapter.setAnalysisProfile("general")).rejects.toThrow();
  });
});

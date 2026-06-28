/**
 * PivotOps — Storage Abstraction
 * Uses OfficeRuntime.storage (immune to tracking prevention) with localStorage fallback.
 * OfficeRuntime.storage is async and works reliably in Office Add-in WebViews.
 */

function hasOfficeStorage(): boolean {
  try {
    return typeof OfficeRuntime !== "undefined" && !!OfficeRuntime.storage;
  } catch {
    return false;
  }
}

export async function storageGetItem(key: string): Promise<string | null> {
  if (hasOfficeStorage()) {
    try {
      const value = await OfficeRuntime.storage.getItem(key);
      return value ?? null;
    } catch {
      // Fall through to localStorage
    }
  }
  try {
    return localStorage.getItem(key);
  } catch {
    return null;
  }
}

export async function storageSetItem(key: string, value: string): Promise<void> {
  if (hasOfficeStorage()) {
    try {
      await OfficeRuntime.storage.setItem(key, value);
      return;
    } catch {
      // Fall through to localStorage
    }
  }
  try {
    localStorage.setItem(key, value);
  } catch {
    // Storage completely unavailable
  }
}

export async function storageRemoveItem(key: string): Promise<void> {
  if (hasOfficeStorage()) {
    try {
      await OfficeRuntime.storage.removeItem(key);
      return;
    } catch {
      // Fall through to localStorage
    }
  }
  try {
    localStorage.removeItem(key);
  } catch {
    // Storage completely unavailable
  }
}

const getStorage = (): Storage | null => {
  if (typeof window === "undefined") {
    return null;
  }

  try {
    return window.localStorage;
  } catch (error) {
    console.warn("LocalStorage is not available:", error);
    return null;
  }
};

const tryWithStorage = <T>(fn: (storage: Storage) => T, fallback: T): T => {
  const storage = getStorage();
  if (!storage) {
    return fallback;
  }

  try {
    return fn(storage);
  } catch (error) {
    console.warn("LocalStorage operation failed:", error);
    return fallback;
  }
};

export const safeLocalStorage = {
  isAvailable: () => getStorage() !== null,
  getItem: (key: string) => tryWithStorage(storage => storage.getItem(key), null),
  setItem: (key: string, value: string) => {
    void tryWithStorage(storage => {
      storage.setItem(key, value);
      return true;
    }, true);
  },
  removeItem: (key: string) => {
    void tryWithStorage(storage => {
      storage.removeItem(key);
      return true;
    }, true);
  },
  getJSON: <T>(key: string, fallback: T): T => {
    const raw = safeLocalStorage.getItem(key);
    if (!raw) {
      return fallback;
    }

    try {
      return JSON.parse(raw) as T;
    } catch (error) {
      console.warn(`Failed to parse localStorage key "${key}":`, error);
      return fallback;
    }
  },
  setJSON: (key: string, value: unknown) => {
    try {
      const raw = JSON.stringify(value);
      safeLocalStorage.setItem(key, raw);
    } catch (error) {
      console.warn(`Failed to serialise localStorage key "${key}":`, error);
    }
  },
  keys: (): string[] =>
    tryWithStorage(storage => {
      const result: string[] = [];
      for (let index = 0; index < storage.length; index += 1) {
        const key = storage.key(index);
        if (key) {
          result.push(key);
        }
      }
      return result;
    }, [] as string[]),
  dispatchStorageEvent: (key: string, newValue: string | null) => {
    if (typeof window === "undefined" || !safeLocalStorage.isAvailable()) {
      return;
    }

    try {
      const storage = window.localStorage;
      const event = new StorageEvent("storage", {
        key,
        newValue,
        storageArea: storage,
      });
      window.dispatchEvent(event);
    } catch (error) {
      console.warn("Failed to dispatch synthetic storage event:", error);
    }
  },
};

export const getSafeStorage = (): Storage | null => getStorage();

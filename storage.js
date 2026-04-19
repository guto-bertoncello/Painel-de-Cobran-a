const DB_NAME = "site-cobranca-local";
const DB_VERSION = 1;
const STORE_NAME = "imports";

export async function saveStoredImport(record) {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(STORE_NAME, "readwrite");
    transaction.oncomplete = () => resolve(record);
    transaction.onerror = () => reject(transaction.error);
    transaction.objectStore(STORE_NAME).put(record);
  });
}

export async function getStoredImport(slot) {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(STORE_NAME, "readonly");
    const request = transaction.objectStore(STORE_NAME).get(slot);
    request.onsuccess = () => resolve(request.result ?? null);
    request.onerror = () => reject(request.error);
  });
}

export async function deleteStoredImport(slot) {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(STORE_NAME, "readwrite");
    transaction.oncomplete = () => resolve(true);
    transaction.onerror = () => reject(transaction.error);
    transaction.objectStore(STORE_NAME).delete(slot);
  });
}

function openDatabase() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onupgradeneeded = () => {
      const database = request.result;
      if (!database.objectStoreNames.contains(STORE_NAME)) {
        database.createObjectStore(STORE_NAME, { keyPath: "slot" });
      }
    };

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

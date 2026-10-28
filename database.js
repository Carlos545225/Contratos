/**
 * Persistencia local en IndexedDB. Los registros y archivos permanecen en
 * este navegador y no se envían a ningún servidor.
 */
const DB = (() => {
  const STORES = {
    ENTITIES: 'entities',
    CONTRACTS: 'contracts',
    EXCEL_FILES: 'excel_files',
    DOCUMENTS: 'documents',
    QUOTES: 'quotes',
    AUDIT_LOG: 'audit_log'
  };
  const STORE_NAMES = Object.values(STORES);
  const DATABASE_NAME = 'SIGEC_IPS_DB';
  let databasePromise;

  function createMissingStores(database) {
    for (const storeName of STORE_NAMES) {
      if (!database.objectStoreNames.contains(storeName)) {
        database.createObjectStore(storeName, { keyPath: 'id' });
      }
    }
  }

  function finishDatabaseOpen(database, resolve) {
    database.onversionchange = () => database.close();
    resolve(database);
  }

  function openDatabase() {
    if (!databasePromise) {
      databasePromise = new Promise((resolve, reject) => {
        const request = indexedDB.open(DATABASE_NAME);

        request.onupgradeneeded = () => createMissingStores(request.result);

        request.onsuccess = () => {
          const database = request.result;
          const missingStores = STORE_NAMES.filter(name => !database.objectStoreNames.contains(name));
          if (missingStores.length === 0) {
            finishDatabaseOpen(database, resolve);
            return;
          }

          const nextVersion = database.version + 1;
          database.close();
          const upgradeRequest = indexedDB.open(DATABASE_NAME, nextVersion);
          upgradeRequest.onupgradeneeded = () => createMissingStores(upgradeRequest.result);
          upgradeRequest.onsuccess = () => {
            const upgradedDatabase = upgradeRequest.result;
            const stillMissing = STORE_NAMES.filter(name => !upgradedDatabase.objectStoreNames.contains(name));
            if (stillMissing.length > 0) {
              upgradedDatabase.close();
              reject(new Error(`No se pudieron crear los almacenes locales faltantes: ${stillMissing.join(', ')}.`));
              return;
            }
            finishDatabaseOpen(upgradedDatabase, resolve);
          };
          upgradeRequest.onerror = () => reject(upgradeRequest.error || new Error('No se pudo actualizar el almacenamiento local.'));
          upgradeRequest.onblocked = () => reject(new Error('La actualización del almacenamiento local está bloqueada por otra pestaña. Cierre las otras pestañas de SIGEC-IPS e inténtelo de nuevo.'));
        };

        request.onerror = () => {
          databasePromise = null;
          reject(request.error || new Error('No se pudo abrir el almacenamiento local.'));
        };
        request.onblocked = () => {
          databasePromise = null;
          reject(new Error('El almacenamiento local está bloqueado por otra pestaña. Cierre las otras pestañas de SIGEC-IPS e inténtelo de nuevo.'));
        };
      });
    }
    return databasePromise;
  }

  function transactionError(transaction, actionError) {
    const error = actionError || transaction.error;
    if (error?.name === 'QuotaExceededError') {
      return new Error('El navegador se quedó sin espacio para guardar estos datos. Libere espacio o elimine archivos y vuelva a intentarlo.');
    }
    return error || new Error('Falló la operación de almacenamiento local.');
  }

  async function runTransaction(storeNames, mode, action) {
    const database = await openDatabase();
    return new Promise((resolve, reject) => {
      let transaction;
      try {
        transaction = database.transaction(storeNames, mode);
      } catch (error) {
        reject(error);
        return;
      }

      let result;
      let actionError;
      try {
        result = action(transaction);
      } catch (error) {
        actionError = error;
        transaction.abort();
      }

      transaction.oncomplete = () => resolve(result);
      transaction.onerror = () => reject(transactionError(transaction, actionError));
      transaction.onabort = () => reject(transactionError(transaction, actionError));
    });
  }

  function validateStore(storeName) {
    if (!STORE_NAMES.includes(storeName)) throw new Error(`Almacén no permitido: ${storeName}`);
  }

  async function get(storeName, key) {
    validateStore(storeName);
    return runTransaction(storeName, 'readonly', transaction => new Promise((resolve, reject) => {
      const request = transaction.objectStore(storeName).get(key);
      request.onsuccess = () => resolve(request.result ?? null);
      request.onerror = () => reject(request.error);
    }));
  }

  async function getAll(storeName) {
    validateStore(storeName);
    return runTransaction(storeName, 'readonly', transaction => new Promise((resolve, reject) => {
      const request = transaction.objectStore(storeName).getAll();
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    }));
  }

  async function getAllMetadata(storeName) {
    const records = await getAll(storeName);
    return records.map(record => {
      const metadata = { ...record };
      delete metadata.rawBuffer;
      if (metadata.attachment) {
        metadata.attachment = { ...metadata.attachment };
        delete metadata.attachment.dataBase64;
      }
      return metadata;
    });
  }

  async function getByIndex(storeName, indexName, value) {
    const records = await getAllMetadata(storeName);
    return records.filter(record => record[indexName] === value);
  }

  async function put(storeName, record) {
    validateStore(storeName);
    if (!record || typeof record.id !== 'string' || !record.id) {
      throw new Error('El registro requiere un identificador válido.');
    }
    return runTransaction(storeName, 'readwrite', transaction => {
      transaction.objectStore(storeName).put(record);
      return record.id;
    });
  }

  async function remove(storeName, key) {
    validateStore(storeName);
    return runTransaction(storeName, 'readwrite', transaction => {
      transaction.objectStore(storeName).delete(key);
      return true;
    });
  }

  async function clearAllData() {
    return runTransaction(STORE_NAMES, 'readwrite', transaction => {
      for (const storeName of STORE_NAMES) {
        transaction.objectStore(storeName).clear();
      }
      return true;
    });
  }

  return {
    STORES,
    get,
    getAll,
    getAllMetadata,
    getByIndex,
    put,
    remove,
    clearAllData
  };
})();

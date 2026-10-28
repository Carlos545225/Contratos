/**
 * BACKUP MODULE
 * Serialización binaria y estructural completa:
 * Entidades, Contratos, Archivos Excel (en Base64), Cotizaciones, Auditoría y Adjuntos.
 */

const BackupModule = (() => {
  let pendingRestoreData = null;
  let lastBackupAt = null;

  /**
   * Convierte un ArrayBuffer a cadena Base64
   */
  function arrayBufferToBase64(buffer) {
    if (!buffer) return null;
    let binary = '';
    const bytes = new Uint8Array(buffer);
    const len = bytes.byteLength;
    for (let i = 0; i < len; i++) {
      binary += String.fromCharCode(bytes[i]);
    }
    return window.btoa(binary);
  }

  /**
   * Convierte una cadena Base64 a ArrayBuffer
   */
  function base64ToArrayBuffer(base64) {
    if (!base64) return null;
    const binaryString = window.atob(base64);
    const len = binaryString.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes.buffer;
  }

  async function restoreBackupPayload(payload) {
    const { entities, contracts, quotes, excelFiles, documents, auditLog } = payload;

    if (Array.isArray(entities)) {
      for (const entity of entities) {
        await DB.put(DB.STORES.ENTITIES, entity);
      }
    }

    if (Array.isArray(contracts)) {
      for (const contract of contracts) {
        await DB.put(DB.STORES.CONTRACTS, contract);
      }
    }

    if (Array.isArray(quotes)) {
      for (const quote of quotes) {
        await DB.put(DB.STORES.QUOTES, quote);
      }
    }

    if (Array.isArray(excelFiles)) {
      for (const file of excelFiles) {
        await DB.put(DB.STORES.EXCEL_FILES, {
          ...file,
          rawBuffer: file.rawBuffer ? base64ToArrayBuffer(file.rawBuffer) : null
        });
      }
    }

    if (Array.isArray(documents)) {
      for (const file of documents) {
        await DB.put(DB.STORES.DOCUMENTS, {
          ...file,
          rawBuffer: file.rawBuffer ? base64ToArrayBuffer(file.rawBuffer) : null
        });
      }
    }

    if (Array.isArray(auditLog)) {
      for (const log of auditLog) {
        await DB.put(DB.STORES.AUDIT_LOG, log);
      }
    }
  }

  /**
   * Genera y descarga el archivo JSON de copia de seguridad con todos los datos y binarios
   */
  async function exportFullBackup(downloadName) {
    try {
      App.showToast('Generando copia de respaldo de la IPS...', 'info');

      const entities = await DB.getAll(DB.STORES.ENTITIES);
      const contracts = await DB.getAll(DB.STORES.CONTRACTS);
      const quotes = await DB.getAll(DB.STORES.QUOTES);
      const excelFilesRaw = await DB.getAll(DB.STORES.EXCEL_FILES);
      const documentsRaw = await DB.getAll(DB.STORES.DOCUMENTS);
      const auditLog = await DB.getAll(DB.STORES.AUDIT_LOG);

      // Serializar binarios a Base64 para garantizar portabilidad en JSON
      const excelFilesExport = excelFilesRaw.map(fileItem => {
        return {
          ...fileItem,
          rawBuffer: fileItem.rawBuffer ? arrayBufferToBase64(fileItem.rawBuffer) : null
        };
      });
      const documentsExport = documentsRaw.map(fileItem => ({
        ...fileItem,
        rawBuffer: fileItem.rawBuffer ? arrayBufferToBase64(fileItem.rawBuffer) : null
      }));

      const backupPackage = {
        appSignature: 'SIGEC_IPS_BACKUP_V1',
        createdAt: new Date().toISOString(),
        institution: 'IPS Colombia',
        version: 1,
        stats: {
          entitiesCount: entities.length,
          contractsCount: contracts.length,
          quotesCount: quotes.length,
          excelFilesCount: excelFilesExport.length,
          documentsCount: documentsExport.length
        },
        payload: {
          entities,
          contracts,
          quotes,
          excelFiles: excelFilesExport,
          documents: documentsExport,
          auditLog
        }
      };

      const jsonStr = JSON.stringify(backupPackage, null, 2);
      const blob = new Blob([jsonStr], { type: 'application/json' });
      const url = URL.createObjectURL(blob);

      const downloadAnchor = document.createElement('a');
      const timeStamp = new Date().toISOString().replace(/[:.]/g, '-');
      downloadAnchor.href = url;
      downloadAnchor.download = downloadName || `SIGEC_IPS_BACKUP_${timeStamp}.json`;
      document.body.appendChild(downloadAnchor);
      downloadAnchor.click();
      document.body.removeChild(downloadAnchor);
      URL.revokeObjectURL(url);

      const dateNowFormatted = new Date().toLocaleString();
      lastBackupAt = dateNowFormatted;
      updateLastBackupLabel();

      App.showToast('Copia de respaldo generada exitosamente.', 'success');
    } catch (err) {
      console.error(err);
      App.showToast('Error generando copia de respaldo: ' + err.message, 'error');
    }
  }

  async function exportRepositoryBackup() {
    await exportFullBackup('respaldo.json');
  }

  async function loadRepositoryBackup() {
    if (!/^https?:$/.test(window.location.protocol)) return false;

    const response = await fetch(new URL('data/respaldo.json', document.baseURI), {
      cache: 'no-store'
    });
    if (!response.ok) {
      throw new Error(`No se pudo leer data/respaldo.json (HTTP ${response.status}).`);
    }

    const backup = await response.json();
    if (
      !backup.appSignature?.startsWith('SIGEC_IPS') ||
      !backup.payload ||
      typeof backup.createdAt !== 'string'
    ) {
      throw new Error('data/respaldo.json no tiene un formato de respaldo válido.');
    }

    const lastImportedAt = localStorage.getItem('sigecIpsRepositoryBackupCreatedAt');
    if (lastImportedAt === backup.createdAt) return false;

    await restoreBackupPayload(backup.payload);
    localStorage.setItem('sigecIpsRepositoryBackupCreatedAt', backup.createdAt);
    App.showToast('Copia del repositorio cargada. Para publicar cambios, reemplaza data/respaldo.json con una nueva exportación.', 'success');
    return true;
  }

  function updateLastBackupLabel() {
    const lbl = document.getElementById('last-backup-label');
    if (lbl) {
      const val = lastBackupAt || 'No descargado en esta sesión';
      lbl.textContent = `Último respaldo: ${val}`;
    }
  }

  /**
   * Valida la estructura del archivo JSON cargado por el usuario
   */
  function handleFileSelected(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const parsed = JSON.parse(e.target.result);

        if (!parsed.appSignature || !parsed.appSignature.startsWith('SIGEC_IPS') || !parsed.payload) {
          throw new Error('El archivo seleccionado no corresponde a un respaldo válido de SIGEC-IPS.');
        }

        pendingRestoreData = parsed;

        // Mostrar estadísticas del paquete detectado
        const statsList = document.getElementById('backup-stats-list');
        statsList.innerHTML = `
          <li><strong>Fecha de Emisión:</strong> ${new Date(parsed.createdAt).toLocaleString()}</li>
          <li><strong>Entidades:</strong> ${parsed.stats?.entitiesCount ?? parsed.payload.entities?.length ?? 0}</li>
          <li><strong>Contratos:</strong> ${parsed.stats?.contractsCount ?? parsed.payload.contracts?.length ?? 0}</li>
          <li><strong>Cotizaciones:</strong> ${parsed.stats?.quotesCount ?? parsed.payload.quotes?.length ?? 0}</li>
          <li><strong>Archivos Tarifarios (Excel):</strong> ${parsed.stats?.excelFilesCount ?? parsed.payload.excelFiles?.length ?? 0}</li>
          <li><strong>Documentos:</strong> ${parsed.stats?.documentsCount ?? parsed.payload.documents?.length ?? 0}</li>
        `;

        document.getElementById('backup-preview-card').classList.remove('hidden');
      } catch (err) {
        pendingRestoreData = null;
        document.getElementById('backup-preview-card').classList.add('hidden');
        App.showToast('Archivo de respaldo no válido: ' + err.message, 'error');
      }
    };

    reader.readAsText(file);
  }

  /**
   * Ejecuta la restauración en base a la opción elegida (combinación o sustitución total)
   */
  async function executeRestore() {
    if (!pendingRestoreData) {
      App.showToast('No se ha cargado ningún respaldo válido.', 'error');
      return;
    }

    const mode = document.querySelector('input[name="restore-mode"]:checked').value;

    App.showConfirmModal(
      '¿Desea proceder con la restauración de la base de datos?',
      async () => {
        try {
          App.showToast('Restaurando base de datos...', 'info');

          if (mode === 'replace') {
            await DB.clearAllData();
          }

          await restoreBackupPayload(pendingRestoreData.payload);

          App.showToast('Restauración completada con éxito.', 'success');
          document.getElementById('backup-preview-card').classList.add('hidden');
          document.getElementById('backup-file-input').value = '';
          pendingRestoreData = null;

          App.refreshDashboard();
          App.navigate('dashboard');
        } catch (err) {
          console.error(err);
          App.showToast('Fallo durante la restauración: ' + err.message, 'error');
        }
      }
    );
  }

  return {
    exportFullBackup,
    exportRepositoryBackup,
    loadRepositoryBackup,
    handleFileSelected,
    executeRestore,
    updateLastBackupLabel
  };
})();

/**
 * APPLICATION CONTROLLER (Vanilla JS)
 * Orquestación central de UI, paginación, filtros dinámicos, CRUD y vistas duales.
 */

const App = (() => {
  // Estado global de la aplicación
  let currentContractId = null;
  let activeExcelPage = 1;
  const EXCEL_PAGE_SIZE = 15;
  const consultationMode = new URLSearchParams(window.location.search).get('modo') === 'consulta';
  const isReadOnly = consultationMode;
  let filteredExcelRows = [];
  let currentExcelDataset = null;
  let currentExcelFile = null;
  let currentExcelSheets = [];
  let quoteAttachmentData = null; // Almacenamiento temporal para adjunto de cotización
  let confirmCallback = null;
  let crossContracts = [];
  let crossSourceSheets = [];
  let crossUploadSheets = [];
  let crossUploadBuffer = null;
  let crossUploadFileName = '';
  let crossResultRows = [];
  let crossResultHeaders = [];
  let crossColumnMappings = [];
  let crossColumnMappingId = 0;

  /**
   * Inicialización al cargar el DOM
   */
  async function init() {
    document.body.classList.toggle('read-only-mode', consultationMode);
    if (!consultationMode) BackupModule.updateLastBackupLabel();
    try {
      await BackupModule.loadRepositoryBackup();
    } catch (error) {
      console.error(error);
      showToast(`No se pudo cargar la copia del repositorio: ${error.message}`, 'error');
    }
    await refreshDashboard();
  }

  // ==================== NAVEGACIÓN Y TOASTS ====================

  function denyReadOnlyWrite() {
    if (!isReadOnly) return false;
    showToast('Esta página es de solo lectura. Abra index.html para editar.', 'error');
    return true;
  }

  function navigate(viewName) {
    if (isReadOnly && (viewName === 'cross-information' || viewName === 'backup')) {
      return;
    }
    document.querySelectorAll('.app-view').forEach(v => v.classList.remove('active'));
    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));

    const targetView = document.getElementById(`view-${viewName}`);
    const targetNav = document.querySelector(`.nav-btn[data-view="${viewName}"]`);

    if (targetView) targetView.classList.add('active');
    if (targetNav) targetNav.classList.add('active');

    if (viewName === 'dashboard') refreshDashboard();
    if (viewName === 'entities') renderEntitiesTable();
    if (viewName === 'contracts') renderContractsTable();
    if (viewName === 'cross-information') loadCrossContracts();
    if (viewName === 'backup') BackupModule.updateLastBackupLabel();
  }

  function switchContractTab(tabName) {
    const selectedTab = document.querySelector(`[data-contract-tab="${tabName}"]`);
    if (!selectedTab) return;

    document.querySelectorAll('[data-contract-tab]').forEach(tab => {
      const isSelected = tab === selectedTab;
      tab.classList.toggle('active', isSelected);
      tab.setAttribute('aria-selected', String(isSelected));
      tab.setAttribute('tabindex', isSelected ? '0' : '-1');
    });

    document.querySelectorAll('[data-contract-panel]').forEach(panel => {
      panel.hidden = panel.dataset.contractPanel !== tabName;
    });
  }

  function handleContractTabKeydown(event) {
    const tabs = Array.from(document.querySelectorAll('[data-contract-tab]'));
    const currentIndex = tabs.indexOf(event.currentTarget);
    let nextIndex;

    if (event.key === 'ArrowRight') {
      nextIndex = (currentIndex + 1) % tabs.length;
    } else if (event.key === 'ArrowLeft') {
      nextIndex = (currentIndex - 1 + tabs.length) % tabs.length;
    } else if (event.key === 'Home') {
      nextIndex = 0;
    } else if (event.key === 'End') {
      nextIndex = tabs.length - 1;
    } else {
      return;
    }

    event.preventDefault();
    const nextTab = tabs[nextIndex];
    switchContractTab(nextTab.dataset.contractTab);
    nextTab.focus();
  }

  function showToast(message, type = 'info') {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = message;

    container.appendChild(toast);
    setTimeout(() => {
      toast.style.opacity = '0';
      setTimeout(() => toast.remove(), 300);
    }, 3500);
  }

  function showConfirmModal(message, onConfirm) {
    document.getElementById('confirm-modal-message').textContent = message;
    confirmCallback = onConfirm;
    document.getElementById('confirm-modal').classList.remove('hidden');
    document.getElementById('confirm-modal-ok-btn').onclick = () => {
      if (confirmCallback) confirmCallback();
      closeConfirmModal();
    };
  }

  function closeConfirmModal() {
    document.getElementById('confirm-modal').classList.add('hidden');
    confirmCallback = null;
  }

  // ==================== CRUCE DE INFORMACIÓN ====================

  async function loadCrossContracts() {
    try {
      const [contracts, entities, excelFiles] = await Promise.all([
        DB.getAll(DB.STORES.CONTRACTS),
        DB.getAll(DB.STORES.ENTITIES),
        DB.getAllMetadata(DB.STORES.EXCEL_FILES)
      ]);
      const selectedContractId = document.getElementById('cross-contract-select').value;
      crossContracts = contracts.reduce((available, contract) => {
        const files = excelFiles.filter(file => file.contractId === contract.id);
        const file = files.find(item => item.version === contract.currentVersion) || files[files.length - 1];
        if (file) {
          const entity = entities.find(item => item.id === contract.entityId);
          available.push({
            contract,
            file,
            label: `${contract.number} - ${entity ? entity.name : 'Entidad no encontrada'}`
          });
        }
        return available;
      }, []);

      const select = document.getElementById('cross-contract-select');
      select.innerHTML = '<option value="">Seleccione un contrato...</option>';
      crossContracts.forEach(item => {
        const option = document.createElement('option');
        option.value = item.contract.id;
        option.textContent = item.label;
        select.appendChild(option);
      });

      if (crossContracts.some(item => item.contract.id === selectedContractId)) {
        select.value = selectedContractId;
      } else {
        select.value = '';
        crossSourceSheets = [];
        populateCrossSheetSelect('cross-contract-sheet', [], 'Seleccione una hoja...');
        populateCrossColumnSelect('cross-contract-key', []);
        populateCrossColumnSelect('cross-contract-price', []);
      }

      if (crossContracts.length === 0) {
        clearCrossResult('No hay contratos con un tarifario Excel vigente para comparar.');
        return;
      }
      if (select.value) await handleCrossContractChange(select.value);
      else updateCrossAvailability();
    } catch (error) {
      console.error(error);
      showToast('Error cargando los tarifarios disponibles para el cruce: ' + error.message, 'error');
    }
  }

  async function handleCrossContractChange(contractId) {
    clearCrossResult();
    const contractEntry = crossContracts.find(item => item.contract.id === contractId);
    if (!contractEntry) {
      crossSourceSheets = [];
      populateCrossSheetSelect('cross-contract-sheet', [], 'Seleccione una hoja...');
      populateCrossColumnSelect('cross-contract-key', []);
      populateCrossColumnSelect('cross-contract-price', []);
      updateCrossAvailability();
      return;
    }

    try {
      const file = await DB.get(DB.STORES.EXCEL_FILES, contractEntry.file.id);
      crossSourceSheets = Array.isArray(file.sheets) && file.sheets.length > 0
        ? file.sheets
        : [{
            sheetName: file.sheetName || 'Hoja 1',
            headers: file.headers || [],
            rows: file.rows || [],
            displayRows: file.displayRows || []
          }];
      if (crossSourceSheets.some(sheet => !Array.isArray(sheet.headers)) && file.rawBuffer) {
        crossSourceSheets = ExcelParser.getFormattedSheets(file.rawBuffer, file.headerRow || 1);
      }
      populateCrossSheetSelect(
        'cross-contract-sheet',
        crossSourceSheets.map(sheet => sheet.sheetName),
        'Seleccione una hoja...'
      );
      if (crossSourceSheets.length > 0) {
        document.getElementById('cross-contract-sheet').value = crossSourceSheets[0].sheetName;
        handleCrossSourceSheetChange();
      } else {
        populateCrossColumnSelect('cross-contract-key', []);
        populateCrossColumnSelect('cross-contract-price', []);
        updateCrossAvailability();
      }
    } catch (error) {
      console.error(error);
      showToast('No se pudo preparar el tarifario contractual: ' + error.message, 'error');
    }
  }

  function populateCrossSheetSelect(selectId, sheetNames, placeholder) {
    const select = document.getElementById(selectId);
    select.innerHTML = '';
    const initialOption = document.createElement('option');
    initialOption.value = '';
    initialOption.textContent = placeholder;
    select.appendChild(initialOption);
    sheetNames.forEach(sheetName => {
      const option = document.createElement('option');
      option.value = sheetName;
      option.textContent = sheetName;
      select.appendChild(option);
    });
    select.disabled = sheetNames.length === 0;
  }

  function populateCrossColumnSelect(selectId, headers) {
    const select = document.getElementById(selectId);
    const previousValue = select.value;
    select.innerHTML = '';
    const initialOption = document.createElement('option');
    initialOption.value = '';
    initialOption.textContent = headers.length ? 'Seleccione una columna...' : 'Cargue una tabla primero...';
    select.appendChild(initialOption);
    headers.forEach(header => {
      const option = document.createElement('option');
      option.value = header;
      option.textContent = header;
      select.appendChild(option);
    });
    select.disabled = headers.length === 0;
    if (headers.includes(previousValue)) select.value = previousValue;
  }

  function getCrossSourceDataset() {
    const sheetName = document.getElementById('cross-contract-sheet').value;
    return crossSourceSheets.find(sheet => sheet.sheetName === sheetName) || null;
  }

  function getCrossUploadDataset() {
    const sheetName = document.getElementById('cross-upload-sheet').value;
    return crossUploadSheets.find(sheet => sheet.sheetName === sheetName) || null;
  }

  function updateCrossColumnOptions() {
    const source = getCrossSourceDataset();
    const upload = getCrossUploadDataset();
    populateCrossColumnSelect('cross-contract-key', source ? source.headers : []);
    populateCrossColumnSelect('cross-contract-price', source ? source.headers : []);
    populateCrossColumnSelect('cross-upload-key', upload ? upload.headers : []);
    populateCrossColumnSelect('cross-upload-price', upload ? upload.headers : []);
    renderCrossColumnMappings();
  }

  function addCrossColumnMapping() {
    crossColumnMappings.push({
      id: ++crossColumnMappingId,
      sourceColumn: '',
      uploadColumn: '',
      useForMatch: false
    });
    renderCrossColumnMappings();
    clearCrossResult();
  }

  function renderCrossColumnMappings() {
    const container = document.getElementById('cross-column-mappings');
    if (!container) return;
    const source = getCrossSourceDataset();
    const upload = getCrossUploadDataset();
    const sourceHeaders = source ? source.headers : [];
    const uploadHeaders = upload ? upload.headers : [];
    container.innerHTML = '';

    crossColumnMappings.forEach(mapping => {
      const row = document.createElement('div');
      row.className = 'cross-column-mapping';

      const sourceField = document.createElement('div');
      sourceField.className = 'form-group';
      const sourceLabel = document.createElement('label');
      sourceLabel.textContent = 'Columna contractual';
      const sourceSelect = createCrossMappingSelect(sourceHeaders, mapping.sourceColumn, 'Seleccione una columna...');
      sourceSelect.addEventListener('change', () => {
        mapping.sourceColumn = sourceSelect.value;
        handleCrossColumnChange();
      });
      sourceField.append(sourceLabel, sourceSelect);

      const uploadField = document.createElement('div');
      uploadField.className = 'form-group';
      const uploadLabel = document.createElement('label');
      uploadLabel.textContent = 'Columna del archivo nuevo';
      const uploadSelect = createCrossMappingSelect(uploadHeaders, mapping.uploadColumn, 'Seleccione una columna...');
      uploadSelect.addEventListener('change', () => {
        mapping.uploadColumn = uploadSelect.value;
        handleCrossColumnChange();
      });
      uploadField.append(uploadLabel, uploadSelect);

      const matchLabel = document.createElement('label');
      matchLabel.className = 'cross-match-option';
      const matchCheckbox = document.createElement('input');
      matchCheckbox.type = 'checkbox';
      matchCheckbox.checked = mapping.useForMatch;
      matchCheckbox.addEventListener('change', () => {
        mapping.useForMatch = matchCheckbox.checked;
        handleCrossColumnChange();
      });
      const matchText = document.createElement('span');
      matchText.textContent = 'Usar también para buscar coincidencias';
      matchLabel.append(matchCheckbox, matchText);

      const removeButton = document.createElement('button');
      removeButton.type = 'button';
      removeButton.className = 'btn btn-sm btn-danger cross-remove-column';
      removeButton.textContent = 'Quitar';
      removeButton.setAttribute('aria-label', 'Quitar par de columnas');
      removeButton.addEventListener('click', () => {
        crossColumnMappings = crossColumnMappings.filter(item => item.id !== mapping.id);
        renderCrossColumnMappings();
        handleCrossColumnChange();
      });

      row.append(sourceField, uploadField, matchLabel, removeButton);
      container.appendChild(row);
    });
  }

  function createCrossMappingSelect(headers, selectedValue, placeholder) {
    const select = document.createElement('select');
    select.className = 'cross-mapping-select';
    const initialOption = document.createElement('option');
    initialOption.value = '';
    initialOption.textContent = headers.length ? placeholder : 'Seleccione las tablas primero...';
    select.appendChild(initialOption);
    headers.forEach(header => {
      const option = document.createElement('option');
      option.value = header;
      option.textContent = header;
      select.appendChild(option);
    });
    select.value = headers.includes(selectedValue) ? selectedValue : '';
    select.disabled = headers.length === 0;
    return select;
  }

  function handleCrossSourceSheetChange() {
    clearCrossResult();
    updateCrossColumnOptions();
    updateCrossAvailability();
  }

  function handleCrossUploadSheetChange() {
    clearCrossResult();
    updateCrossColumnOptions();
    updateCrossAvailability();
  }

  async function handleCrossFileSelection(event) {
    const file = event.target.files[0];
    if (!file) return;

    crossUploadFileName = file.name;
    crossUploadSheets = [];
    crossUploadBuffer = null;
    document.getElementById('cross-file-name').textContent = `Leyendo ${file.name}...`;
    clearCrossResult();
    populateCrossSheetSelect('cross-upload-sheet', [], 'Cargue un archivo...');
    updateCrossColumnOptions();
    updateCrossAvailability();

    const reader = new FileReader();
    reader.onload = () => {
      try {
        crossUploadBuffer = reader.result;
        parseCrossUploadFile();
        document.getElementById('cross-file-name').textContent =
          `${crossUploadFileName} (${crossUploadSheets.length} hoja(s))`;
      } catch (error) {
        console.error(error);
        crossUploadSheets = [];
        crossUploadBuffer = null;
        populateCrossSheetSelect('cross-upload-sheet', [], 'No se pudo leer el archivo.');
        updateCrossColumnOptions();
        updateCrossAvailability();
        document.getElementById('cross-file-name').textContent = 'No fue posible leer el archivo.';
        showToast('Error al leer el archivo para comparar: ' + error.message, 'error');
      }
    };
    reader.onerror = () => {
      crossUploadSheets = [];
      crossUploadBuffer = null;
      document.getElementById('cross-file-name').textContent = 'No fue posible leer el archivo.';
      updateCrossAvailability();
      showToast('No se pudo leer el archivo seleccionado.', 'error');
    };
    reader.readAsArrayBuffer(file);
  }

  function parseCrossUploadFile() {
    if (!crossUploadBuffer) return;
    const previousSheetName = document.getElementById('cross-upload-sheet').value;
    const headerRow = Number(document.getElementById('cross-header-row').value) || 1;
    crossUploadSheets = ExcelParser.parseArrayBuffer(crossUploadBuffer, headerRow);
    populateCrossSheetSelect(
      'cross-upload-sheet',
      crossUploadSheets.map(sheet => sheet.sheetName),
      'Seleccione una hoja...'
    );
    const sheetSelect = document.getElementById('cross-upload-sheet');
    if (crossUploadSheets.some(sheet => sheet.sheetName === previousSheetName)) {
      sheetSelect.value = previousSheetName;
    } else if (crossUploadSheets.length > 0) {
      sheetSelect.value = crossUploadSheets[0].sheetName;
    }
    updateCrossColumnOptions();
    updateCrossAvailability();
  }

  function handleCrossHeaderRowChange() {
    if (!crossUploadBuffer) return;
    clearCrossResult();
    try {
      parseCrossUploadFile();
    } catch (error) {
      console.error(error);
      crossUploadSheets = [];
      populateCrossSheetSelect('cross-upload-sheet', [], 'No se pudo leer la fila de encabezados.');
      updateCrossColumnOptions();
      updateCrossAvailability();
      showToast('No se pudo interpretar la fila de encabezados: ' + error.message, 'error');
    }
  }

  function handleCrossColumnChange() {
    clearCrossResult();
    updateCrossAvailability();
  }

  function resetInformationCross() {
    crossSourceSheets = [];
    crossUploadSheets = [];
    crossUploadBuffer = null;
    crossUploadFileName = '';
    crossColumnMappings = [];

    document.getElementById('cross-contract-select').value = '';
    document.getElementById('cross-header-row').value = '1';
    document.getElementById('cross-file').value = '';
    document.getElementById('cross-file-name').textContent = 'Admite .xlsx, .xls y .csv.';

    populateCrossSheetSelect('cross-contract-sheet', [], 'Seleccione una hoja...');
    populateCrossSheetSelect('cross-upload-sheet', [], 'Cargue un archivo...');
    populateCrossColumnSelect('cross-contract-key', []);
    populateCrossColumnSelect('cross-contract-price', []);
    populateCrossColumnSelect('cross-upload-key', []);
    populateCrossColumnSelect('cross-upload-price', []);

    const mappings = document.getElementById('cross-column-mappings');
    if (mappings) mappings.innerHTML = '';
    clearCrossResult();
    updateCrossAvailability();
    showToast('Cruce limpiado. Puede iniciar uno nuevo.', 'success');
  }

  function updateCrossAvailability() {
    const ready = Boolean(
      getCrossSourceDataset()
      && getCrossUploadDataset()
      && document.getElementById('cross-contract-key').value
      && document.getElementById('cross-contract-price').value
      && document.getElementById('cross-upload-key').value
      && document.getElementById('cross-upload-price').value
    );
    document.getElementById('btn-run-cross').disabled = !ready;
  }

  function clearCrossResult(message = 'Configure los archivos y columnas para generar el resultado.') {
    crossResultRows = [];
    crossResultHeaders = [];
    document.getElementById('cross-result-summary').textContent = message;
    document.getElementById('cross-result-head').innerHTML = '<tr><th>Resultado</th></tr>';
    document.getElementById('cross-result-body').innerHTML =
      '<tr><td class="text-center">Aún no se ha realizado ningún cruce.</td></tr>';
    document.getElementById('btn-export-cross').disabled = true;
    document.getElementById('cross-result-stats').classList.add('hidden');
    document.getElementById('cross-result-search').value = '';
    document.getElementById('cross-result-search').disabled = true;
    document.getElementById('cross-result-filter').value = '';
    document.getElementById('cross-result-filter').disabled = true;
    document.getElementById('cross-result-count').textContent = '';
  }

  function normalizeCrossKey(value) {
    return String(value === null || value === undefined ? '' : value)
      .normalize('NFKC')
      .trim()
      .replace(/\s+/g, ' ')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toLocaleLowerCase();
  }

  function parseCrossMoney(value) {
    if (typeof value === 'number') return Number.isFinite(value) ? value : null;
    if (value === null || value === undefined || String(value).trim() === '') return null;

    const original = String(value).trim();
    const isNegative = original.startsWith('(') && original.endsWith(')');
    let text = original.replace(/[^\d,.\-]/g, '');
    if (!/\d/.test(text)) return null;
    if (isNegative) text = text.replace(/-/g, '');
    const commaIndex = text.lastIndexOf(',');
    const dotIndex = text.lastIndexOf('.');

    if (commaIndex >= 0 && dotIndex >= 0) {
      const decimalSeparator = commaIndex > dotIndex ? ',' : '.';
      const groupingSeparator = decimalSeparator === ',' ? '.' : ',';
      text = text.split(groupingSeparator).join('');
      if (decimalSeparator === ',') text = text.replace(',', '.');
    } else if (commaIndex >= 0 || dotIndex >= 0) {
      const separator = commaIndex >= 0 ? ',' : '.';
      const parts = text.split(separator);
      const hasRepeatedSeparators = parts.length > 2;
      const lastPart = parts[parts.length - 1];
      const isThousandsFormat = hasRepeatedSeparators || lastPart.length === 3;
      if (isThousandsFormat) text = parts.join('');
      else text = `${parts.slice(0, -1).join('')}.${lastPart}`;
    }

    const amount = Number(text);
    if (!Number.isFinite(amount)) return null;
    return isNegative ? -Math.abs(amount) : amount;
  }

  function normalizeCrossMoneyForComparison(value) {
    const amount = parseCrossMoney(value);
    return amount === null ? null : Math.round(amount);
  }

  function formatCrossMoney(value) {
    if (value === null || value === undefined || !Number.isFinite(value)) return '—';
    return `$ ${value.toLocaleString('es-CO', { maximumFractionDigits: 0 })}`;
  }

  function runInformationCross() {
    const source = getCrossSourceDataset();
    const upload = getCrossUploadDataset();
    const sourceKey = document.getElementById('cross-contract-key').value;
    const sourcePrice = document.getElementById('cross-contract-price').value;
    const uploadKey = document.getElementById('cross-upload-key').value;
    const uploadPrice = document.getElementById('cross-upload-price').value;

    if (!source || !upload || !sourceKey || !sourcePrice || !uploadKey || !uploadPrice) {
      showToast('Seleccione las dos tablas y todas las columnas requeridas.', 'error');
      return;
    }
    if (crossColumnMappings.some(mapping => Boolean(mapping.sourceColumn) !== Boolean(mapping.uploadColumn))) {
      showToast('Seleccione una columna en ambos archivos para cada columna adicional.', 'error');
      return;
    }

    const columnMappings = crossColumnMappings.filter(mapping => (
      mapping.sourceColumn && mapping.uploadColumn
    ));
    const matchMappings = columnMappings.filter(mapping => mapping.useForMatch);
    const buildCrossIndex = (rows, keyColumn, useUploadColumns) => {
      const index = new Map();
      rows.forEach(row => {
        const keyValues = [
          normalizeCrossKey(row[keyColumn]),
          ...matchMappings.map(mapping => normalizeCrossKey(
            row[useUploadColumns ? mapping.uploadColumn : mapping.sourceColumn]
          ))
        ];
        if (!keyValues[0]) return;
        const key = JSON.stringify(keyValues);
        if (!index.has(key)) index.set(key, []);
        index.get(key).push(row);
      });
      return index;
    };
    const sourceByKey = buildCrossIndex(source.rows, sourceKey, false);
    const uploadByKey = buildCrossIndex(upload.rows, uploadKey, true);
    const keyLabel = matchMappings.length
      ? `${sourceKey} / ${uploadKey} + ${matchMappings.map(mapping => `${mapping.sourceColumn} / ${mapping.uploadColumn}`).join(' + ')}`
      : `${sourceKey} / ${uploadKey}`;
    const beforeLabel = `Precio contractual (${sourcePrice})`;
    const afterLabel = `Precio nuevo (${uploadPrice})`;
    const resultColumnHeaders = [];
    columnMappings.forEach((mapping, index) => {
      resultColumnHeaders.push(
        `Dato ${index + 1} contractual (${mapping.sourceColumn})`,
        `Dato ${index + 1} nuevo (${mapping.uploadColumn})`
      );
    });
    const result = [];
    const matchedSourceRows = new Set();
    const matchedUploadRows = new Set();
    const buildResultRow = (sourceRow, uploadRow, status) => {
      const sourceKeyValue = sourceRow ? sourceRow[sourceKey] : null;
      const uploadKeyValue = uploadRow ? uploadRow[uploadKey] : null;
      const before = sourceRow ? parseCrossMoney(sourceRow[sourcePrice]) : null;
      const after = uploadRow ? parseCrossMoney(uploadRow[uploadPrice]) : null;
      const comparableBefore = sourceRow ? normalizeCrossMoneyForComparison(sourceRow[sourcePrice]) : null;
      const comparableAfter = uploadRow ? normalizeCrossMoneyForComparison(uploadRow[uploadPrice]) : null;
      const settlesByUvr = uploadRow && comparableAfter === 0;
      const valuesAreEqual = status === 'Igual';
      const difference = sourceRow && uploadRow && comparableBefore !== null && comparableAfter !== null
        ? comparableAfter - comparableBefore
        : null;
      const percentage = difference === null || comparableBefore === 0
        ? (comparableBefore === 0 && comparableAfter === 0 ? 0 : null)
        : (difference / comparableBefore) * 100;
      const resultRow = {
        [keyLabel]: sourceRow ? sourceKeyValue : uploadKeyValue,
        [beforeLabel]: comparableBefore,
        [afterLabel]: !uploadRow ? null : settlesByUvr ? 'Se liquida por UVR' : comparableAfter,
        'Diferencia ($)': difference,
        'Diferencia (%)': percentage === null ? null : Math.round(percentage),
        'Comparación': status,
        display: [
          String((sourceRow ? sourceKeyValue : uploadKeyValue) ?? ''),
          sourceRow ? formatCrossMoney(before) : '—',
          !uploadRow ? '—' : settlesByUvr ? 'Se liquida por UVR' : formatCrossMoney(after),
          formatCrossMoney(difference),
          percentage === null ? 'No aplica' : `${Math.round(percentage).toLocaleString('es-CO')} %`,
          status
        ]
      };
      columnMappings.forEach((mapping, mappingIndex) => {
        const sourceValue = sourceRow ? sourceRow[mapping.sourceColumn] : '';
        const uploadValue = uploadRow ? uploadRow[mapping.uploadColumn] : '';
        resultRow[resultColumnHeaders[mappingIndex * 2]] = sourceValue;
        resultRow[resultColumnHeaders[mappingIndex * 2 + 1]] = uploadValue;
        resultRow.display.push(
          String(sourceValue ?? ''),
          String(uploadValue ?? '')
        );
      });
      return resultRow;
    };

    sourceByKey.forEach((sourceRows, key) => {
      const uploadRows = uploadByKey.get(key) || [];
      const availableUploadIndexes = new Set(uploadRows.map((row, index) => index));
      const pairs = [];
      const unmatchedSourceRows = [];

      sourceRows.forEach(sourceRow => {
        const sourceValue = normalizeCrossMoneyForComparison(sourceRow[sourcePrice]);
        const equalUploadIndex = sourceValue === null
          ? undefined
          : Array.from(availableUploadIndexes).find(index =>
            normalizeCrossMoneyForComparison(uploadRows[index][uploadPrice]) === sourceValue
          );

        if (equalUploadIndex === undefined) {
          unmatchedSourceRows.push(sourceRow);
          return;
        }

        pairs.push({ sourceRow, uploadRow: uploadRows[equalUploadIndex] });
        availableUploadIndexes.delete(equalUploadIndex);
      });

      const remainingUploadRows = Array.from(availableUploadIndexes, index => uploadRows[index]);
      const remainingPairCount = Math.min(unmatchedSourceRows.length, remainingUploadRows.length);
      for (let index = 0; index < remainingPairCount; index++) {
        pairs.push({
          sourceRow: unmatchedSourceRows[index],
          uploadRow: remainingUploadRows[index]
        });
      }

      pairs.forEach(({ sourceRow, uploadRow }) => {
        matchedSourceRows.add(sourceRow);
        matchedUploadRows.add(uploadRow);
        const comparableBefore = normalizeCrossMoneyForComparison(sourceRow[sourcePrice]);
        const comparableAfter = normalizeCrossMoneyForComparison(uploadRow[uploadPrice]);
        const settlesByUvr = comparableAfter === 0;
        const valuesAreEqual = !settlesByUvr
          && comparableBefore !== null
          && comparableAfter !== null
          && comparableBefore === comparableAfter;
        result.push(buildResultRow(sourceRow, uploadRow, valuesAreEqual ? 'Igual' : 'Diferente'));
      });

      unmatchedSourceRows.slice(remainingPairCount).forEach(sourceRow => {
        matchedSourceRows.add(sourceRow);
        result.push(buildResultRow(sourceRow, null, 'No se encuentra en archivo nuevo'));
      });
    });

    source.rows.forEach(sourceRow => {
      if (matchedSourceRows.has(sourceRow)) return;
      const status = normalizeCrossKey(sourceRow[sourceKey])
        ? 'No se encuentra en archivo nuevo'
        : 'Sin clave para cruzar';
      result.push(buildResultRow(sourceRow, null, status));
    });

    upload.rows.forEach(uploadRow => {
      if (matchedUploadRows.has(uploadRow)) return;
      const status = normalizeCrossKey(uploadRow[uploadKey])
        ? 'No se encuentra en contrato'
        : 'Sin clave para cruzar';
      result.push(buildResultRow(null, uploadRow, status));
    });

    crossResultHeaders = [keyLabel, beforeLabel, afterLabel, 'Diferencia ($)', 'Diferencia (%)', 'Comparación', ...resultColumnHeaders];
    crossResultRows = result;
    const counts = result.reduce((totals, row) => {
      if (row['Comparación'] === 'Igual') totals.equal++;
      else if (row['Comparación'] === 'Diferente') totals.different++;
      else if (row['Comparación'] === 'No se encuentra en archivo nuevo') totals.sourceOnly++;
      else if (row['Comparación'] === 'No se encuentra en contrato') totals.uploadOnly++;
      else totals.missingKey++;
      return totals;
    }, { equal: 0, different: 0, sourceOnly: 0, uploadOnly: 0, missingKey: 0 });
    document.getElementById('cross-result-summary').textContent =
      `${result.length} fila(s) comparada(s). Las coincidencias se emparejan por clave y, cuando hay duplicados, se prioriza el mismo precio.${counts.missingKey ? ` ${counts.missingKey} fila(s) sin clave para cruzar.` : ''}`;
    document.getElementById('cross-stat-matched').textContent = counts.equal + counts.different;
    document.getElementById('cross-stat-equal').textContent = counts.equal;
    document.getElementById('cross-stat-different').textContent = counts.different;
    document.getElementById('cross-stat-source-only').textContent = counts.sourceOnly;
    document.getElementById('cross-stat-upload-only').textContent = counts.uploadOnly;
    document.getElementById('cross-result-stats').classList.remove('hidden');
    document.getElementById('cross-result-search').disabled = false;
    document.getElementById('cross-result-filter').disabled = false;
    renderInformationCrossResult();
    document.getElementById('btn-export-cross').disabled = result.length === 0;
    if (counts.equal + counts.different === 0) {
      showToast('No se encontraron coincidencias. Se muestran las filas sin pareja o sin clave para cruzar.', 'info');
    }
  }

  function filterInformationCrossResult() {
    renderInformationCrossResult();
  }

  function renderInformationCrossResult() {
    const tableHead = document.getElementById('cross-result-head');
    const tableBody = document.getElementById('cross-result-body');
    tableHead.innerHTML = '';
    const headerRow = document.createElement('tr');
    crossResultHeaders.forEach(header => {
      const th = document.createElement('th');
      th.textContent = header;
      headerRow.appendChild(th);
    });
    tableHead.appendChild(headerRow);
    tableBody.innerHTML = '';

    if (crossResultRows.length === 0) {
      tableBody.innerHTML = `<tr><td colspan="${crossResultHeaders.length}" class="text-center">No se encontraron coincidencias para las columnas seleccionadas.</td></tr>`;
      return;
    }
    const term = document.getElementById('cross-result-search').value.trim().toLocaleLowerCase();
    const status = document.getElementById('cross-result-filter').value;
    const filteredRows = crossResultRows.filter(result => {
      const matchesStatus = !status || result['Comparación'] === status;
      const matchesTerm = !term || result.display.join(' ').toLocaleLowerCase().includes(term);
      return matchesStatus && matchesTerm;
    });
    document.getElementById('cross-result-count').textContent =
      `${filteredRows.length} de ${crossResultRows.length} fila(s)`;
    if (filteredRows.length === 0) {
      tableBody.innerHTML = `<tr><td colspan="${crossResultHeaders.length}" class="text-center">No hay filas que coincidan con los filtros.</td></tr>`;
      return;
    }
    filteredRows.forEach(result => {
      const tr = document.createElement('tr');
      if (result['Comparación'] === 'Igual') tr.classList.add('cross-row-equal');
      else if (result['Comparación'] === 'Diferente') tr.classList.add('cross-row-different');
      else tr.classList.add('cross-row-unmatched');
      result.display.forEach(value => {
        const td = document.createElement('td');
        td.textContent = value;
        tr.appendChild(td);
      });
      tableBody.appendChild(tr);
    });
  }

  function exportInformationCross() {
    const exportRows = crossResultRows.map(result => {
      const row = {};
      crossResultHeaders.forEach(header => {
        row[header] = result[header];
      });
      return row;
    });
    ExcelParser.exportJsonToExcel(exportRows, 'Resultado_Cruce_Informacion.xlsx', crossResultHeaders);
  }

  // ==================== DASHBOARD ====================

  async function refreshDashboard() {
    try {
      const [entities, contracts, excels, quotes] = await Promise.all([
        DB.getAll(DB.STORES.ENTITIES),
        DB.getAll(DB.STORES.CONTRACTS),
        DB.getAllMetadata(DB.STORES.EXCEL_FILES),
        DB.getAllMetadata(DB.STORES.QUOTES)
      ]);

      document.getElementById('kpi-entities').textContent = entities.filter(e => e.status === 'Activo').length;
      document.getElementById('kpi-contracts').textContent = contracts.length;
      document.getElementById('kpi-excels').textContent = excels.length;
      document.getElementById('kpi-quotes').textContent = quotes.length;

      // Tabla de Contratos Recientes
      const recentContractsTbody = document.getElementById('dashboard-recent-contracts');
      recentContractsTbody.innerHTML = '';
      const recentList = contracts.slice(-5).reverse();

      if (recentList.length === 0) {
        recentContractsTbody.innerHTML = '<tr><td colspan="5" class="text-center">No hay contratos registrados.</td></tr>';
      } else {
        for (const c of recentList) {
          const ent = entities.find(e => e.id === c.entityId);
          const tr = document.createElement('tr');
          tr.innerHTML = `
            <td><strong>${c.number}</strong></td>
            <td>${ent ? ent.name : 'Entidad desconocida'}</td>
            <td>${c.startDate} / ${c.endDate}</td>
            <td><span class="badge ${c.status === 'Activo' ? 'badge-success' : 'badge-secondary'}">${c.status}</span></td>
            <td><button class="btn btn-sm btn-outline" onclick="App.viewContractDetail('${c.id}')">Ver Detalle</button></td>
          `;
          recentContractsTbody.appendChild(tr);
        }
      }

      // Tabla de Cotizaciones Recientes
      const recentQuotesTbody = document.getElementById('dashboard-recent-quotes');
      recentQuotesTbody.innerHTML = '';
      const recentQuotes = quotes.slice(-5).reverse();

      if (recentQuotes.length === 0) {
        recentQuotesTbody.innerHTML = '<tr><td colspan="4" class="text-center">No hay cotizaciones registradas.</td></tr>';
      } else {
        recentQuotes.forEach(q => {
          const tr = document.createElement('tr');
          const codeCell = document.createElement('td');
          const code = document.createElement('code');
          code.textContent = q.code || '';
          codeCell.appendChild(code);

          const descriptionCell = document.createElement('td');
          descriptionCell.textContent = q.description || '';

          const valueCell = document.createElement('td');
          const value = document.createElement('strong');
          value.textContent = `$ ${Number(q.totalValue || 0).toLocaleString()}`;
          valueCell.appendChild(value);

          const dateCell = document.createElement('td');
          dateCell.textContent = q.date || '';

          tr.append(codeCell, descriptionCell, valueCell, dateCell);
          recentQuotesTbody.appendChild(tr);
        });
      }
    } catch (err) {
      console.error(err);
      showToast('Error cargando métricas: ' + err.message, 'error');
    }
  }

  // ==================== BÚSQUEDA GLOBAL ====================

  async function handleGlobalSearch(event) {
    const term = event.target.value.toLowerCase().trim();
    const flyout = document.getElementById('global-search-results');

    if (!term || term.length < 2) {
      flyout.classList.add('hidden');
      flyout.innerHTML = '';
      return;
    }

    const [entities, contracts, quotes] = await Promise.all([
      DB.getAll(DB.STORES.ENTITIES),
      DB.getAll(DB.STORES.CONTRACTS),
      DB.getAllMetadata(DB.STORES.QUOTES)
    ]);

    const matches = [];

    entities.forEach(e => {
      if (e.name.toLowerCase().includes(term) || (e.nit && e.nit.includes(term))) {
        matches.push({ type: 'Entidad', text: `${e.name} (NIT: ${e.nit || 'S/N'})`, action: () => navigate('entities') });
      }
    });

    contracts.forEach(c => {
      if (c.number.toLowerCase().includes(term) || c.description.toLowerCase().includes(term)) {
        matches.push({ type: 'Contrato', text: `Contrato ${c.number}: ${c.description}`, action: () => viewContractDetail(c.id) });
      }
    });

    quotes.forEach(q => {
      if (q.code.toLowerCase().includes(term) || q.description.toLowerCase().includes(term)) {
        matches.push({ type: 'Cotización', text: `Cotización [${q.code}]: ${q.description}`, action: () => viewContractDetail(q.contractId) });
      }
    });

    flyout.innerHTML = '';
    if (matches.length === 0) {
      flyout.innerHTML = '<div class="search-item text-muted">Sin coincidencias halladas.</div>';
    } else {
      matches.slice(0, 8).forEach(m => {
        const item = document.createElement('div');
        item.className = 'search-item';
        const badge = document.createElement('span');
        badge.className = 'badge badge-secondary';
        badge.style.marginRight = '8px';
        badge.textContent = m.type;
        item.append(badge, document.createTextNode(` ${m.text}`));
        item.onclick = () => {
          m.action();
          flyout.classList.add('hidden');
          event.target.value = '';
        };
        flyout.appendChild(item);
      });
    }

    flyout.classList.remove('hidden');
  }

  // ==================== GESTIÓN DE ENTIDADES ====================

  async function renderEntitiesTable() {
    const entities = await DB.getAll(DB.STORES.ENTITIES);
    const contracts = await DB.getAll(DB.STORES.CONTRACTS);
    const tbody = document.getElementById('entities-table-body');
    tbody.innerHTML = '';

    if (entities.length === 0) {
      tbody.innerHTML = `<tr><td colspan="${isReadOnly ? 6 : 7}" class="text-center">No se han registrado entidades aún.</td></tr>`;
      return;
    }

    entities.forEach(ent => {
      const relatedContractsCount = contracts.filter(c => c.entityId === ent.id).length;
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><strong>${ent.name}</strong></td>
        <td>${ent.nit || '-'}</td>
        <td>${ent.code || '-'}</td>
        <td><span class="badge badge-secondary">${relatedContractsCount} contrato(s)</span></td>
        <td><span class="badge ${ent.status === 'Activo' ? 'badge-success' : 'badge-danger'}">${ent.status}</span></td>
        <td>${ent.notes || '-'}</td>
        ${isReadOnly ? '' : `<td data-write-only>
          <button class="btn btn-sm btn-outline" onclick="App.openEntityModal('${ent.id}')">Editar</button>
          <button class="btn btn-sm ${ent.status === 'Activo' ? 'btn-danger' : 'btn-secondary'}" onclick="App.toggleEntityStatus('${ent.id}')">
            ${ent.status === 'Activo' ? 'Desactivar' : 'Activar'}
          </button>
        </td>`}
      `;
      tbody.appendChild(tr);
    });
  }

  function filterEntitiesTable() {
    const term = document.getElementById('entity-filter-input').value.toLowerCase();
    const rows = document.querySelectorAll('#entities-table-body tr');
    rows.forEach(r => {
      r.style.display = r.textContent.toLowerCase().includes(term) ? '' : 'none';
    });
  }

  function openEntityModal(entityId = null) {
    if (denyReadOnlyWrite()) return;
    const form = document.getElementById('form-entity');
    form.reset();
    document.getElementById('entity-id').value = '';

    if (entityId) {
      document.getElementById('modal-entity-title').textContent = 'Editar Entidad';
      DB.get(DB.STORES.ENTITIES, entityId).then(ent => {
        if (ent) {
          document.getElementById('entity-id').value = ent.id;
          document.getElementById('entity-name').value = ent.name;
          document.getElementById('entity-nit').value = ent.nit || '';
          document.getElementById('entity-code').value = ent.code || '';
          document.getElementById('entity-status').value = ent.status || 'Activo';
          document.getElementById('entity-notes').value = ent.notes || '';
        }
      });
    } else {
      document.getElementById('modal-entity-title').textContent = 'Registrar Nueva Entidad';
    }
    document.getElementById('modal-entity').classList.remove('hidden');
  }

  function closeEntityModal() {
    document.getElementById('modal-entity').classList.add('hidden');
  }

  async function saveEntity(event) {
    if (denyReadOnlyWrite()) {
      event.preventDefault();
      return;
    }
    event.preventDefault();
    const id = document.getElementById('entity-id').value || `ent_${crypto.randomUUID()}`;
    const entity = {
      id: id,
      name: document.getElementById('entity-name').value.trim(),
      nit: document.getElementById('entity-nit').value.trim(),
      code: document.getElementById('entity-code').value.trim(),
      status: document.getElementById('entity-status').value,
      notes: document.getElementById('entity-notes').value.trim(),
      updatedAt: new Date().toISOString()
    };

    try {
      await DB.put(DB.STORES.ENTITIES, entity);
      showToast('Entidad guardada en este navegador.', 'success');
      closeEntityModal();
      renderEntitiesTable();
      refreshDashboard();
    } catch (e) {
      showToast('Error al registrar la entidad: ' + e, 'error');
    }
  }

  async function toggleEntityStatus(entityId) {
    if (denyReadOnlyWrite()) return;
    try {
      const ent = await DB.get(DB.STORES.ENTITIES, entityId);
      if (!ent) return;
      const newStatus = ent.status === 'Activo' ? 'Inactivo' : 'Activo';

      showConfirmModal(`¿Desea cambiar el estado de la entidad "${ent.name}" a ${newStatus}?`, async () => {
        try {
          ent.status = newStatus;
          await DB.put(DB.STORES.ENTITIES, ent);
          showToast('Estado de la entidad actualizado.', 'success');
          await renderEntitiesTable();
        } catch (error) {
          showToast(`No se pudo actualizar el estado de la entidad: ${error.message}`, 'error');
        }
      });
    } catch (error) {
      showToast(`No se pudo consultar la entidad: ${error.message}`, 'error');
    }
  }

  // ==================== GESTIÓN DE CONTRATOS ====================

  async function openNewContractModal() {
    if (denyReadOnlyWrite()) return;
    const form = document.getElementById('form-contract');
    form.reset();
    ExcelParser.resetPayload();

    // Poblar desplegable de entidades activas
    const entities = await DB.getAll(DB.STORES.ENTITIES);
    const select = document.getElementById('contract-entity-id');
    select.innerHTML = '<option value="">Seleccione una entidad...</option>';

    entities.forEach(ent => {
      const opt = document.createElement('option');
      opt.value = ent.id;
      opt.textContent = `${ent.name} ${ent.nit ? `(${ent.nit})` : ''}`;
      select.appendChild(opt);
    });

    document.getElementById('modal-contract').classList.remove('hidden');
  }

  function closeContractModal() {
    document.getElementById('modal-contract').classList.add('hidden');
    ExcelParser.resetPayload();
  }

  async function saveContract(event) {
    if (denyReadOnlyWrite()) {
      event.preventDefault();
      return;
    }
    event.preventDefault();
    const entityId = document.getElementById('contract-entity-id').value;
    const number = document.getElementById('contract-number').value.trim();
    const description = document.getElementById('contract-description').value.trim();
    const startDate = document.getElementById('contract-start-date').value;
    const endDate = document.getElementById('contract-end-date').value;
    const status = document.getElementById('contract-status').value;
    const notes = document.getElementById('contract-notes').value.trim();

    if (!entityId || !number || !startDate || !endDate) {
      showToast('Complete los campos obligatorios del contrato.', 'error');
      return;
    }

    const contractId = `cntr_${crypto.randomUUID()}`;
    const parsedExcel = ExcelParser.getParsedPayload();

    const newContract = {
      id: contractId,
      entityId,
      number,
      description,
      startDate,
      endDate,
      status,
      notes,
      currentVersion: parsedExcel ? 1 : 0,
      createdAt: new Date().toISOString()
    };

    try {
      const currentContracts = await DB.getAll(DB.STORES.CONTRACTS);
      if (currentContracts.some(contract =>
        contract.entityId === entityId && contract.number.toLowerCase() === number.toLowerCase()
      )) {
        showToast('Ya existe un contrato con ese número para la entidad seleccionada.', 'error');
        return;
      }
      // Guardar contrato
      await DB.put(DB.STORES.CONTRACTS, newContract);

      // Si se cargó un Excel contractual, guardarlo en el store de archivos
      if (parsedExcel) {
        const excelRecord = {
          id: `file_${contractId}_v1`,
          contractId: contractId,
          version: 1,
          fileName: parsedExcel.fileName,
          sheetName: parsedExcel.sheetName,
          headerRow: parsedExcel.headerRow,
          headers: parsedExcel.headers,
          rows: parsedExcel.rows,
          sheets: parsedExcel.sheets,
          rawBuffer: parsedExcel.rawBuffer,
          uploadedAt: new Date().toISOString()
        };
        await DB.put(DB.STORES.EXCEL_FILES, excelRecord);
      }

      showToast('Contrato creado exitosamente con sus tablas asociadas.', 'success');
      closeContractModal();
      navigate('contracts');
    } catch (e) {
      showToast('Fallo al registrar contrato: ' + e, 'error');
    }
  }

  async function renderContractsTable() {
    const contracts = await DB.getAll(DB.STORES.CONTRACTS);
    const entities = await DB.getAll(DB.STORES.ENTITIES);
    const tbody = document.getElementById('contracts-table-body');
    tbody.innerHTML = '';

    if (contracts.length === 0) {
      tbody.innerHTML = '<tr><td colspan="7" class="text-center">No existen contratos registrados.</td></tr>';
      return;
    }

    contracts.forEach(cntr => {
      const ent = entities.find(e => e.id === cntr.entityId);
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><strong>${cntr.number}</strong></td>
        <td>${ent ? ent.name : 'Entidad no encontrada'}</td>
        <td>${cntr.description}</td>
        <td>${cntr.startDate} al ${cntr.endDate}</td>
        <td><span class="badge badge-secondary">v${cntr.currentVersion || 0}</span></td>
        <td><span class="badge ${cntr.status === 'Activo' ? 'badge-success' : 'badge-warning'}">${cntr.status}</span></td>
        <td>
          <button class="btn btn-sm btn-primary" onclick="App.viewContractDetail('${cntr.id}')">Ver Tablas</button>
        </td>
      `;
      tbody.appendChild(tr);
    });
  }

  function filterContractsTable() {
    const term = document.getElementById('contract-filter-input').value.toLowerCase();
    const status = document.getElementById('contract-status-filter').value;
    const rows = document.querySelectorAll('#contracts-table-body tr');

    rows.forEach(r => {
      const text = r.textContent.toLowerCase();
      const matchesTerm = text.includes(term);
      const matchesStatus = !status || text.includes(status.toLowerCase());
      r.style.display = matchesTerm && matchesStatus ? '' : 'none';
    });
  }

  // ==================== VISTA INDIVIDUAL DEL CONTRATO ====================

  async function viewContractDetail(contractId) {
    currentContractId = contractId;
    const contract = await DB.get(DB.STORES.CONTRACTS, contractId);
    if (!contract) {
      showToast('Contrato no encontrado.', 'error');
      return;
    }

    const entity = await DB.get(DB.STORES.ENTITIES, contract.entityId);

    // Actualizar encabezados de la vista
    document.getElementById('detail-contract-number').textContent = contract.number;
    document.getElementById('detail-contract-title').textContent = `Contrato: ${contract.number}`;
    document.getElementById('detail-contract-subtitle').textContent = `Entidad: ${entity ? entity.name : 'N/A'} | Objeto: ${contract.description}`;
    switchContractTab('excel');

    // Barra de Metadatos
    const metaBar = document.getElementById('detail-contract-meta-bar');
    metaBar.innerHTML = `
      <div class="meta-item"><span>Estado:</span><span><span class="badge ${contract.status === 'Activo' ? 'badge-success' : 'badge-secondary'}">${contract.status}</span></span></div>
      <div class="meta-item"><span>Inicio Vigencia:</span><span>${contract.startDate}</span></div>
      <div class="meta-item"><span>Fin Vigencia:</span><span>${contract.endDate}</span></div>
      <div class="meta-item"><span>Versión Tarifario:</span><span>v${contract.currentVersion || 0}</span></div>
      <div class="meta-item"><span>Observaciones:</span><span>${contract.notes || 'Ninguna'}</span></div>
    `;

    // Cargar los datos de ambas pestañas
    await loadContractExcelData(contractId, contract.currentVersion);
    await loadContractQuotes(contractId);
    await renderContractDocuments(contractId);

    navigate('contract-detail');
  }

  async function renderContractDocuments(contractId) {
    const tbody = document.getElementById('contract-documents-body');
    const documents = await DB.getByIndex(DB.STORES.DOCUMENTS, 'contractId', contractId);
    tbody.replaceChildren();

    if (documents.length === 0) {
      tbody.innerHTML = '<tr><td colspan="5" class="text-center">No hay documentos cargados para este contrato.</td></tr>';
      return;
    }

    documents.sort((a, b) => new Date(b.uploadedAt) - new Date(a.uploadedAt)).forEach(file => {
      const row = document.createElement('tr');
      const nameCell = document.createElement('td');
      const typeCell = document.createElement('td');
      const sizeCell = document.createElement('td');
      const dateCell = document.createElement('td');
      const actionCell = document.createElement('td');
      nameCell.textContent = file.fileName;
      typeCell.textContent = file.mimeType || 'Desconocido';
      sizeCell.textContent = `${(file.size / 1024).toFixed(1)} KB`;
      dateCell.textContent = new Date(file.uploadedAt).toLocaleString();

      const downloadButton = document.createElement('button');
      downloadButton.type = 'button';
      downloadButton.className = 'btn btn-sm btn-outline';
      downloadButton.textContent = 'Descargar';
      downloadButton.addEventListener('click', () => downloadContractDocument(file.id));
      actionCell.appendChild(downloadButton);

      if (!isReadOnly) {
        const deleteButton = document.createElement('button');
        deleteButton.type = 'button';
        deleteButton.className = 'btn btn-sm btn-danger';
        deleteButton.textContent = 'Eliminar';
        deleteButton.addEventListener('click', () => deleteContractDocument(file.id));
        actionCell.appendChild(deleteButton);
      }

      row.append(nameCell, typeCell, sizeCell, dateCell, actionCell);
      tbody.appendChild(row);
    });
  }

  async function uploadContractDocument(event) {
    if (denyReadOnlyWrite()) {
      event.target.value = '';
      return;
    }
    const file = event.target.files[0];
    if (!file) return;
    if (!currentContractId) {
      showToast('Seleccione primero un contrato.', 'error');
      event.target.value = '';
      return;
    }
    try {
      const record = {
        id: `doc_${crypto.randomUUID()}`,
        contractId: currentContractId,
        fileName: file.name,
        mimeType: file.type || 'application/octet-stream',
        size: file.size,
        rawBuffer: await file.arrayBuffer(),
        uploadedAt: new Date().toISOString()
      };
      await DB.put(DB.STORES.DOCUMENTS, record);
      showToast(`Documento "${file.name}" guardado localmente en este navegador.`, 'success');
      await renderContractDocuments(currentContractId);
    } catch (error) {
      showToast(`No se guardó el documento: ${error.message}`, 'error');
    } finally {
      event.target.value = '';
    }
  }

  async function downloadContractDocument(documentId) {
    try {
      const file = await DB.get(DB.STORES.DOCUMENTS, documentId);
      if (!file?.rawBuffer) throw new Error('El archivo no está disponible en este navegador.');
      const blob = new Blob([file.rawBuffer], { type: file.mimeType || 'application/octet-stream' });
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement('a');
      anchor.href = url;
      anchor.download = file.fileName;
      anchor.click();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (error) {
      showToast(`No se pudo descargar el documento: ${error.message}`, 'error');
    }
  }

  function deleteContractDocument(documentId) {
    if (denyReadOnlyWrite()) return;
    showConfirmModal('¿Desea eliminar este documento de este navegador?', async () => {
      try {
        await DB.remove(DB.STORES.DOCUMENTS, documentId);
        showToast('Documento eliminado del almacenamiento local.', 'success');
        await renderContractDocuments(currentContractId);
      } catch (error) {
        showToast(`No se pudo eliminar el documento: ${error.message}`, 'error');
      }
    });
  }

  // ==================== TABLA 1: DATOS DINÁMICOS EXCEL ====================

  async function loadContractExcelData(contractId, versionNumber) {
    const excelFiles = await DB.getByIndex(DB.STORES.EXCEL_FILES, 'contractId', contractId);
    const activeMetadata = excelFiles.find(f => f.version === versionNumber) || excelFiles[excelFiles.length - 1];
    const activeFile = activeMetadata ? await DB.get(DB.STORES.EXCEL_FILES, activeMetadata.id) : null;

    const fileInfo = document.getElementById('contract-excel-file-info');
    const sheetTabs = document.getElementById('excel-sheet-tabs');
    if (!activeFile) {
      currentExcelFile = null;
      currentExcelSheets = [];
      currentExcelDataset = null;
      filteredExcelRows = [];
      fileInfo.textContent = 'Sin tarifario Excel cargado.';
      sheetTabs.innerHTML = '';
      sheetTabs.classList.add('hidden');
      document.getElementById('excel-table-head').innerHTML = '<tr><th>Sin Columnas</th></tr>';
      document.getElementById('excel-table-body').innerHTML = '<tr><td class="text-center">No hay registros importados en este contrato. Cargue un archivo con el botón superior.</td></tr>';
      activeExcelPage = 1;
      updatePaginationControls(0);
      return;
    }

    currentExcelFile = activeFile;
    currentExcelSheets = Array.isArray(activeFile.sheets) && activeFile.sheets.length > 0
      ? activeFile.sheets
      : [{
          sheetName: activeFile.sheetName || 'Hoja 1',
          headerRow: activeFile.headerRow,
          headers: activeFile.headers || [],
          rows: activeFile.rows || []
        }];

    if (currentExcelSheets.some(sheet => !Array.isArray(sheet.displayRows)) && activeFile.rawBuffer) {
      try {
        currentExcelSheets = ExcelParser.getFormattedSheets(activeFile.rawBuffer, activeFile.headerRow || 1);
      } catch (error) {
        showToast(`No se pudo recuperar el formato de las celdas Excel: ${error.message}`, 'error');
      }
    }

    sheetTabs.innerHTML = '';
    sheetTabs.classList.toggle('hidden', currentExcelSheets.length < 2);
    currentExcelSheets.forEach((sheet, index) => {
      const tab = document.createElement('button');
      tab.type = 'button';
      tab.className = 'excel-sheet-tab';
      tab.id = `excel-sheet-tab-${index}`;
      tab.role = 'tab';
      tab.setAttribute('aria-selected', String(index === 0));
      tab.setAttribute('aria-controls', 'excel-table-container');
      tab.setAttribute('tabindex', index === 0 ? '0' : '-1');
      tab.title = sheet.sheetName;
      tab.textContent = sheet.sheetName;
      tab.onclick = () => selectContractExcelSheet(index);
      tab.onkeydown = handleExcelSheetKeydown;
      sheetTabs.appendChild(tab);
    });

    selectContractExcelSheet(0);
  }

  function selectContractExcelSheet(sheetIndex) {
    const sheet = currentExcelSheets[sheetIndex];
    if (!sheet || !currentExcelFile) return;

    document.querySelectorAll('.excel-sheet-tab').forEach((tab, index) => {
      const isSelected = index === sheetIndex;
      tab.classList.toggle('active', isSelected);
      tab.setAttribute('aria-selected', String(isSelected));
      tab.setAttribute('tabindex', isSelected ? '0' : '-1');
    });

    currentExcelDataset = { ...currentExcelFile, ...sheet };
    filteredExcelRows = (sheet.rows || []).map((row, index) => ({
      data: row,
      display: sheet.displayRows?.[index] || row
    }));
    activeExcelPage = 1;
    document.getElementById('excel-data-search').value = '';
    document.getElementById('contract-excel-file-info').textContent =
      `Archivo: ${currentExcelFile.fileName} (Hoja: ${sheet.sheetName} | Versión ${currentExcelFile.version})`;
    renderCurrentExcelSheet();
  }

  function handleExcelSheetKeydown(event) {
    const tabs = Array.from(document.querySelectorAll('.excel-sheet-tab'));
    const currentIndex = tabs.indexOf(event.currentTarget);
    let nextIndex;

    if (event.key === 'ArrowRight') {
      nextIndex = (currentIndex + 1) % tabs.length;
    } else if (event.key === 'ArrowLeft') {
      nextIndex = (currentIndex - 1 + tabs.length) % tabs.length;
    } else if (event.key === 'Home') {
      nextIndex = 0;
    } else if (event.key === 'End') {
      nextIndex = tabs.length - 1;
    } else {
      return;
    }

    event.preventDefault();
    tabs[nextIndex].focus();
    selectContractExcelSheet(nextIndex);
  }

  function renderCurrentExcelSheet() {
    const tableHead = document.getElementById('excel-table-head');
    const tableBody = document.getElementById('excel-table-body');
    tableHead.innerHTML = '';
    const trH = document.createElement('tr');
    const headers = currentExcelDataset.headers || [];
    headers.forEach(headerText => {
      const th = document.createElement('th');
      th.textContent = headerText;
      th.style.cursor = 'pointer';
      th.title = 'Clic para ordenar por esta columna';
      th.onclick = () => sortExcelDataByColumn(headerText);
      trH.appendChild(th);
    });
    tableHead.appendChild(trH);

    if (headers.length === 0 || currentExcelDataset.rows.length === 0) {
      tableBody.innerHTML = `<tr><td colspan="${Math.max(headers.length, 1)}" class="text-center">Esta hoja no contiene registros para mostrar.</td></tr>`;
      updatePaginationControls(0);
      return;
    }

    renderCurrentExcelPage();
  }

  function renderCurrentExcelPage() {
    const tableBody = document.getElementById('excel-table-body');
    tableBody.innerHTML = '';

    if (filteredExcelRows.length === 0) {
      tableBody.innerHTML = '<tr><td colspan="100%" class="text-center">No se encontraron filas que coincidan con la búsqueda.</td></tr>';
      updatePaginationControls(0);
      return;
    }

    const startIdx = (activeExcelPage - 1) * EXCEL_PAGE_SIZE;
    const endIdx = startIdx + EXCEL_PAGE_SIZE;
    const pageRows = filteredExcelRows.slice(startIdx, endIdx);

    const headers = currentExcelDataset.headers;

    pageRows.forEach(({ data, display }) => {
      const tr = document.createElement('tr');
      headers.forEach(h => {
        const td = document.createElement('td');
        td.textContent = display[h] !== undefined ? display[h] : (data[h] !== undefined ? data[h] : '');
        tr.appendChild(td);
      });
      tableBody.appendChild(tr);
    });

    updatePaginationControls(filteredExcelRows.length);
  }

  function updatePaginationControls(totalRows) {
    const totalPages = Math.ceil(totalRows / EXCEL_PAGE_SIZE) || 1;
    document.getElementById('excel-pagination-info').textContent = `Total: ${totalRows} filas | Página ${activeExcelPage} de ${totalPages}`;
    document.getElementById('excel-current-page-label').textContent = `Página ${activeExcelPage} / ${totalPages}`;
    document.getElementById('btn-excel-prev').disabled = activeExcelPage <= 1;
    document.getElementById('btn-excel-next').disabled = activeExcelPage >= totalPages;
  }

  function changeExcelPage(offset) {
    activeExcelPage += offset;
    renderCurrentExcelPage();
  }

  function filterExcelDataTable() {
    const term = document.getElementById('excel-data-search').value.toLowerCase().trim();
    if (!currentExcelDataset) return;

    if (!term) {
      filteredExcelRows = currentExcelDataset.rows.map((row, index) => ({
        data: row,
        display: currentExcelDataset.displayRows?.[index] || row
      }));
    } else {
      filteredExcelRows = currentExcelDataset.rows.reduce((matches, row, index) => {
        const isMatch = Object.values(row).some(val => String(val).toLowerCase().includes(term))
          || Object.values(currentExcelDataset.displayRows?.[index] || {}).some(val => String(val).toLowerCase().includes(term));
        if (isMatch) {
          matches.push({
            data: row,
            display: currentExcelDataset.displayRows?.[index] || row
          });
        }
        return matches;
      }, []);
    }
    activeExcelPage = 1;
    renderCurrentExcelPage();
  }

  let sortAscending = true;
  function sortExcelDataByColumn(columnName) {
    filteredExcelRows.sort((a, b) => {
      const valA = a.data[columnName] !== undefined ? a.data[columnName] : '';
      const valB = b.data[columnName] !== undefined ? b.data[columnName] : '';
      if (!isNaN(valA) && !isNaN(valB) && valA !== '' && valB !== '') {
        return sortAscending ? valA - valB : valB - valA;
      }
      return sortAscending ? String(valA).localeCompare(String(valB)) : String(valB).localeCompare(String(valA));
    });
    sortAscending = !sortAscending;
    renderCurrentExcelPage();
  }

  function exportCurrentContractExcel() {
    if (!currentExcelDataset || !currentExcelDataset.rows || currentExcelDataset.rows.length === 0) {
      showToast('No hay datos en el tarifario actual para exportar.', 'error');
      return;
    }
    ExcelParser.exportJsonToExcel(
      currentExcelDataset.rows,
      `Tarifario_${currentContractId}_v${currentExcelDataset.version}.xlsx`,
      currentExcelDataset.headers,
      currentExcelDataset.numberFormats
    );
  }

  async function downloadOriginalContractExcel() {
    if (!currentExcelDataset || !currentExcelDataset.rawBuffer) {
      showToast('No existe el archivo binario original en este navegador.', 'error');
      return;
    }

    const blob = new Blob([currentExcelDataset.rawBuffer], { type: 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = currentExcelDataset.fileName || `Contrato_${currentContractId}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // ==================== ACTUALIZACIÓN DE VERSIÓN DE EXCEL ====================

  function openUpdateExcelModal() {
    if (denyReadOnlyWrite()) return;
    if (!currentContractId) return;
    document.getElementById('form-update-excel').reset();
    ExcelParser.resetUpdatePayload();
    document.getElementById('modal-update-excel').classList.remove('hidden');
  }

  function closeUpdateExcelModal() {
    document.getElementById('modal-update-excel').classList.add('hidden');
    ExcelParser.resetUpdatePayload();
  }

  async function submitUpdateExcel(event) {
    if (denyReadOnlyWrite()) {
      event.preventDefault();
      return;
    }
    event.preventDefault();
    const payload = ExcelParser.getUpdateParsedPayload();
    if (!payload) {
      showToast('Seleccione un archivo Excel válido.', 'error');
      return;
    }

    const contract = await DB.get(DB.STORES.CONTRACTS, currentContractId);
    const newVersion = (contract.currentVersion || 0) + 1;

    showConfirmModal(`¿Desea establecer este archivo como la versión vigente v${newVersion}? Las cotizaciones existentes se mantendrán intactas.`, async () => {
      try {
        const excelRecord = {
          id: `file_${currentContractId}_v${newVersion}`,
          contractId: currentContractId,
          version: newVersion,
          fileName: payload.fileName,
          sheetName: payload.sheetName,
          headerRow: payload.headerRow,
          headers: payload.headers,
          rows: payload.rows,
          sheets: payload.sheets,
          rawBuffer: payload.rawBuffer,
          uploadedAt: new Date().toISOString()
        };

        await DB.put(DB.STORES.EXCEL_FILES, excelRecord);

        // Actualizar contrato a la nueva versión
        contract.currentVersion = newVersion;
        await DB.put(DB.STORES.CONTRACTS, contract);

        showToast(`Versión v${newVersion} establecida como vigente.`, 'success');
        closeUpdateExcelModal();
        viewContractDetail(currentContractId);
      } catch (e) {
        showToast('Error al actualizar el archivo: ' + e, 'error');
      }
    });
  }

  async function openVersionHistoryModal() {
    const files = await DB.getByIndex(DB.STORES.EXCEL_FILES, 'contractId', currentContractId);
    const tbody = document.getElementById('history-versions-tbody');
    tbody.innerHTML = '';

    if (files.length === 0) {
      tbody.innerHTML = '<tr><td colspan="5" class="text-center">Sin historial de archivos.</td></tr>';
    } else {
      files.sort((a, b) => b.version - a.version).forEach(f => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td><strong>v${f.version}</strong></td>
          <td>${f.fileName}</td>
          <td>${new Date(f.uploadedAt).toLocaleString()}</td>
          <td>${Array.isArray(f.sheets) && f.sheets.length > 0
            ? `${f.sheets.length} hoja(s) · ${f.sheets.reduce((total, sheet) => total + sheet.rows.length, 0)} filas`
            : `${f.rows ? f.rows.length : 0} filas`}</td>
          <td>
            <button class="btn btn-sm btn-outline" onclick="App.restoreSpecificVersion(${f.version})">Visualizar</button>
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    document.getElementById('modal-history').classList.remove('hidden');
  }

  function closeVersionHistoryModal() {
    document.getElementById('modal-history').classList.add('hidden');
  }

  async function restoreSpecificVersion(verNum) {
    closeVersionHistoryModal();
    await loadContractExcelData(currentContractId, verNum);
    showToast(`Mostrando datos de la versión v${verNum}.`, 'info');
  }

  // ==================== TABLA 2: COTIZACIONES ====================

  async function loadContractQuotes(contractId) {
    const quotes = await DB.getByIndex(DB.STORES.QUOTES, 'contractId', contractId);
    const tbody = document.getElementById('quotes-table-body');
    tbody.innerHTML = '';

    if (quotes.length === 0) {
      tbody.innerHTML = `<tr><td colspan="${isReadOnly ? 8 : 9}" class="text-center">No hay cotizaciones para este contrato.${isReadOnly ? '' : ' Utilice "+ Añadir Cotización".'}</td></tr>`;
      return;
    }

    quotes.forEach(q => {
      const tr = document.createElement('tr');
      const hasAttachment = Boolean(q.attachment);
      const codeCell = document.createElement('td');
      const code = document.createElement('strong');
      code.textContent = q.code || '';
      codeCell.appendChild(code);

      const descriptionCell = document.createElement('td');
      descriptionCell.textContent = q.description || '';

      const quantityCell = document.createElement('td');
      quantityCell.textContent = q.quantity ?? '';

      const unitValueCell = document.createElement('td');
      unitValueCell.textContent = `$ ${Number(q.unitValue).toLocaleString()}`;

      const totalValueCell = document.createElement('td');
      const totalValue = document.createElement('strong');
      totalValue.textContent = `$ ${Number(q.totalValue).toLocaleString()}`;
      totalValueCell.appendChild(totalValue);

      const tariffTypeCell = document.createElement('td');
      const tariffType = document.createElement('span');
      tariffType.className = 'badge badge-secondary';
      tariffType.textContent = q.tariffType || '';
      tariffTypeCell.appendChild(tariffType);

      const statusCell = document.createElement('td');
      const status = document.createElement('span');
      status.className = `badge ${getQuoteBadgeClass(q.status)}`;
      status.textContent = q.status || '';
      statusCell.appendChild(status);

      const attachmentCell = document.createElement('td');
      if (hasAttachment) {
        const downloadButton = document.createElement('button');
        downloadButton.type = 'button';
        downloadButton.className = 'btn btn-sm btn-outline';
        downloadButton.textContent = '📎 Ver';
        downloadButton.addEventListener('click', () => downloadQuoteAttachment(q.id));
        attachmentCell.appendChild(downloadButton);
      } else {
        attachmentCell.textContent = '-';
      }

      tr.append(codeCell, descriptionCell, quantityCell, unitValueCell, totalValueCell, tariffTypeCell, statusCell, attachmentCell);
      if (!isReadOnly) {
        const actionsCell = document.createElement('td');
        actionsCell.dataset.writeOnly = '';
        const editButton = document.createElement('button');
        editButton.type = 'button';
        editButton.className = 'btn btn-sm btn-outline';
        editButton.textContent = 'Editar';
        editButton.addEventListener('click', () => openNewQuoteModal(q.id));
        const deleteButton = document.createElement('button');
        deleteButton.type = 'button';
        deleteButton.className = 'btn btn-sm btn-danger';
        deleteButton.textContent = 'Eliminar';
        deleteButton.addEventListener('click', () => deleteQuote(q.id));
        actionsCell.append(editButton, deleteButton);
        tr.appendChild(actionsCell);
      }
      tbody.appendChild(tr);
    });
  }

  function getQuoteBadgeClass(status) {
    switch (status) {
      case 'Aprobada': return 'badge-success';
      case 'Pendiente': return 'badge-warning';
      case 'Rechazada': return 'badge-danger';
      default: return 'badge-secondary';
    }
  }

  function openNewQuoteModal(quoteId = null) {
    if (denyReadOnlyWrite()) return;
    const form = document.getElementById('form-quote');
    form.reset();
    document.getElementById('quote-id').value = '';
    document.getElementById('quote-custom-fields-container').innerHTML = '';
    quoteAttachmentData = null;
    document.getElementById('quote-attachment-indicator').classList.add('hidden');

    if (quoteId) {
      document.getElementById('modal-quote-title').textContent = 'Editar Cotización';
      DB.get(DB.STORES.QUOTES, quoteId).then(q => {
        if (q) {
          document.getElementById('quote-id').value = q.id;
          document.getElementById('quote-code').value = q.code;
          document.getElementById('quote-tariff-type').value = q.tariffType;
          document.getElementById('quote-desc').value = q.description;
          document.getElementById('quote-quantity').value = q.quantity;
          document.getElementById('quote-unit-value').value = q.unitValue;
          document.getElementById('quote-total-value').value = q.totalValue;
          document.getElementById('quote-provider').value = q.provider || '';
          document.getElementById('quote-date').value = q.date;
          document.getElementById('quote-status').value = q.status;
          document.getElementById('quote-observations').value = q.observations || '';

          if (q.customFields && Array.isArray(q.customFields)) {
            q.customFields.forEach(cf => addCustomFieldRow(cf.name, cf.value));
          }

          if (q.attachment) {
            quoteAttachmentData = q.attachment;
            const ind = document.getElementById('quote-attachment-indicator');
            ind.textContent = `Archivo adjunto actual: ${q.attachment.name}`;
            ind.classList.remove('hidden');
          }
        }
      });
    } else {
      document.getElementById('modal-quote-title').textContent = 'Registrar Cotización';
      document.getElementById('quote-date').value = new Date().toISOString().split('T')[0];
    }

    document.getElementById('modal-quote').classList.remove('hidden');
  }

  function closeQuoteModal() {
    document.getElementById('modal-quote').classList.add('hidden');
  }

  function calculateQuoteTotal() {
    const qty = parseFloat(document.getElementById('quote-quantity').value) || 0;
    const unitVal = parseFloat(document.getElementById('quote-unit-value').value) || 0;
    document.getElementById('quote-total-value').value = qty * unitVal;
  }

  function addCustomFieldRow(name = '', value = '') {
    if (denyReadOnlyWrite()) return;
    const container = document.getElementById('quote-custom-fields-container');
    const row = document.createElement('div');
    row.className = 'custom-field-item';
    row.innerHTML = `
      <input type="text" placeholder="Nombre Campo (Ej: Tiempo de Entrega)" value="${name}" class="cf-name">
      <input type="text" placeholder="Valor" value="${value}" class="cf-val">
      <button type="button" class="btn btn-sm btn-danger" onclick="this.parentElement.remove()">✕</button>
    `;
    container.appendChild(row);
  }

  function handleQuoteAttachment(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      quoteAttachmentData = {
        name: file.name,
        type: file.type,
        size: file.size,
        dataBase64: e.target.result // Data URL
      };
      const ind = document.getElementById('quote-attachment-indicator');
      ind.textContent = `Archivo adjunto cargado: ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
      ind.classList.remove('hidden');
    };
    reader.readAsDataURL(file);
  }

  async function saveQuote(event) {
    if (denyReadOnlyWrite()) {
      event.preventDefault();
      return;
    }
    event.preventDefault();
    const id = document.getElementById('quote-id').value || `quote_${crypto.randomUUID()}`;
    const code = document.getElementById('quote-code').value.trim();
    const tariffType = document.getElementById('quote-tariff-type').value.trim();
    const description = document.getElementById('quote-desc').value.trim();
    const quantity = parseFloat(document.getElementById('quote-quantity').value) || 0;
    const unitValue = parseFloat(document.getElementById('quote-unit-value').value) || 0;
    const totalValue = quantity * unitValue;
    const provider = document.getElementById('quote-provider').value.trim();
    const date = document.getElementById('quote-date').value;
    const status = document.getElementById('quote-status').value;
    const observations = document.getElementById('quote-observations').value.trim();

    // Recopilar campos personalizados
    const customFields = [];
    document.querySelectorAll('.custom-field-item').forEach(el => {
      const name = el.querySelector('.cf-name').value.trim();
      const val = el.querySelector('.cf-val').value.trim();
      if (name) customFields.push({ name, value: val });
    });

    const quoteRecord = {
      id,
      contractId: currentContractId,
      code,
      tariffType,
      description,
      quantity,
      unitValue,
      totalValue,
      provider,
      date,
      status,
      observations,
      customFields,
      attachment: quoteAttachmentData,
      updatedAt: new Date().toISOString()
    };

    try {
      await DB.put(DB.STORES.QUOTES, quoteRecord);
      showToast('Cotización guardada exitosamente.', 'success');
      closeQuoteModal();
      loadContractQuotes(currentContractId);
    } catch (e) {
      showToast('Error guardando cotización: ' + e, 'error');
    }
  }

  function deleteQuote(quoteId) {
    if (denyReadOnlyWrite()) return;
    showConfirmModal('¿Está seguro de eliminar esta cotización de forma permanente?', async () => {
      try {
        await DB.remove(DB.STORES.QUOTES, quoteId);
        showToast('Cotización eliminada.', 'success');
        loadContractQuotes(currentContractId);
      } catch (e) {
        showToast('Error al eliminar cotización: ' + e, 'error');
      }
    });
  }

  async function downloadQuoteAttachment(quoteId) {
    const q = await DB.get(DB.STORES.QUOTES, quoteId);
    if (!q || !q.attachment) {
      showToast('La cotización no posee archivos adjuntos.', 'error');
      return;
    }
    const a = document.createElement('a');
    a.href = q.attachment.dataBase64;
    a.download = q.attachment.name || 'Adjunto_Cotizacion';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }

  function filterQuotesTable() {
    const term = document.getElementById('quote-search-input').value.toLowerCase();
    const status = document.getElementById('quote-status-filter').value;
    const rows = document.querySelectorAll('#quotes-table-body tr');

    rows.forEach(r => {
      const text = r.textContent.toLowerCase();
      const matchesTerm = text.includes(term);
      const matchesStatus = !status || text.includes(status.toLowerCase());
      r.style.display = matchesTerm && matchesStatus ? '' : 'none';
    });
  }

  async function exportQuotesToExcel() {
    const quotes = await DB.getByIndex(DB.STORES.QUOTES, 'contractId', currentContractId);
    if (quotes.length === 0) {
      showToast('No existen cotizaciones para exportar.', 'error');
      return;
    }
    const flatQuotes = quotes.map(q => ({
      Codigo: q.code,
      Descripcion: q.description,
      Cantidad: q.quantity,
      ValorUnitario: q.unitValue,
      ValorTotal: q.totalValue,
      ManualTarifario: q.tariffType,
      Proveedor: q.provider || '',
      Fecha: q.date,
      Estado: q.status,
      Observaciones: q.observations || ''
    }));

    ExcelParser.exportJsonToExcel(flatQuotes, `Cotizaciones_Contrato_${currentContractId}.xlsx`);
  }

  // Inicializar al cargar
  window.addEventListener('DOMContentLoaded', init);

  return {
    navigate,
    switchContractTab,
    handleContractTabKeydown,
    showToast,
    showConfirmModal,
    closeConfirmModal,
    handleCrossContractChange,
    handleCrossSourceSheetChange,
    handleCrossUploadSheetChange,
    handleCrossFileSelection,
    handleCrossHeaderRowChange,
    handleCrossColumnChange,
    addCrossColumnMapping,
    resetInformationCross,
    updateCrossAvailability,
    runInformationCross,
    filterInformationCrossResult,
    exportInformationCross,
    refreshDashboard,
    handleGlobalSearch,
    renderEntitiesTable,
    filterEntitiesTable,
    openEntityModal,
    closeEntityModal,
    saveEntity,
    toggleEntityStatus,
    openNewContractModal,
    closeContractModal,
    saveContract,
    renderContractsTable,
    filterContractsTable,
    viewContractDetail,
    uploadContractDocument,
    downloadContractDocument,
    deleteContractDocument,
    changeExcelPage,
    filterExcelDataTable,
    exportCurrentContractExcel,
    downloadOriginalContractExcel,
    openUpdateExcelModal,
    closeUpdateExcelModal,
    submitUpdateExcel,
    openVersionHistoryModal,
    closeVersionHistoryModal,
    restoreSpecificVersion,
    openNewQuoteModal,
    closeQuoteModal,
    calculateQuoteTotal,
    addCustomFieldRow,
    handleQuoteAttachment,
    saveQuote,
    deleteQuote,
    downloadQuoteAttachment,
    filterQuotesTable,
    exportQuotesToExcel
  };
})();

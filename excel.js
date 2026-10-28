/**
 * EXCEL MODULE (SheetJS Integration)
 * Manejo dinámico de archivos Excel, detección agnóstica de columnas,
 * cambio de hojas, selección de fila de encabezados y conversión a binario/base64.
 */

const ExcelParser = (() => {
  // Estado volátil de carga para el asistente del modal de nuevo contrato
  let currentFileArrayBuffer = null;
  let currentFileName = '';
  let currentWorkbook = null;
  let parsedPayload = null;

  // Estado volátil para actualización de versión
  let updateFileArrayBuffer = null;
  let updateFileName = '';
  let updateWorkbook = null;
  let updateParsedPayload = null;

  /**
   * Manejador de evento al seleccionar un archivo en el modal de nuevo contrato
   */
  async function handleFileSelection(event) {
    const file = event.target.files[0];
    if (!file) return;

    currentFileName = file.name;
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        currentFileArrayBuffer = e.target.result;
        currentWorkbook = XLSX.read(new Uint8Array(currentFileArrayBuffer), { type: 'array', cellNF: true });

        const sheetSelect = document.getElementById('excel-select-sheet');
        sheetSelect.innerHTML = '';
        currentWorkbook.SheetNames.forEach((sheetName) => {
          const opt = document.createElement('option');
          opt.value = sheetName;
          opt.textContent = sheetName;
          sheetSelect.appendChild(opt);
        });

        document.getElementById('excel-sheet-selector-box').classList.remove('hidden');
        renderPreview();
      } catch (err) {
        App.showToast('Error al leer el archivo Excel: ' + err.message, 'error');
      }
    };

    reader.readAsArrayBuffer(file);
  }

  function handleSheetChange(sheetName) {
    renderPreview();
  }

  function handleHeaderRowChange(rowNumber) {
    renderPreview();
  }

  function parseWorkbookSheets(workbook, headerRowIndex, emptyHeaderPrefix) {
    return workbook.SheetNames.map(sheetName => {
      const worksheet = workbook.Sheets[sheetName];
      const rawMatrix = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
      const headerRow = rawMatrix[headerRowIndex] || [];
      const range = worksheet['!ref'] ? XLSX.utils.decode_range(worksheet['!ref']) : { s: { r: 0, c: 0 } };
      const headers = headerRow.map((header, index) => (
        header ? String(header).trim() : `${emptyHeaderPrefix}${index + 1}`
      ));
      const rows = [];
      const displayRows = [];
      const numberFormats = [];

      for (let rowIndex = headerRowIndex + 1; rowIndex < rawMatrix.length; rowIndex++) {
        const row = rawMatrix[rowIndex];
        if (row.length === 0 || row.every(value => value === '' || value === null || value === undefined)) continue;

        const rowObject = {};
        const displayRow = {};
        const numberFormatRow = {};
        headers.forEach((header, index) => {
          const value = row[index] !== undefined ? row[index] : '';
          const address = XLSX.utils.encode_cell({
            r: range.s.r + rowIndex,
            c: range.s.c + index
          });
          const cell = worksheet[address];

          rowObject[header] = value;
          displayRow[header] = cell
            ? (cell.w !== undefined ? cell.w : XLSX.utils.format_cell(cell))
            : value;
          if (cell && typeof cell.v === 'number' && cell.z) {
            numberFormatRow[header] = cell.z;
          }
        });
        rows.push(rowObject);
        displayRows.push(displayRow);
        numberFormats.push(numberFormatRow);
      }

      return {
        sheetName,
        headerRow: headerRowIndex + 1,
        headers,
        rows,
        displayRows,
        numberFormats
      };
    });
  }

  /**
   * Genera la previsualización tabular del archivo seleccionado
   */
  function renderPreview() {
    if (!currentWorkbook) return;

    const sheetName = document.getElementById('excel-select-sheet').value || currentWorkbook.SheetNames[0];
    const headerRowIndex = Math.max(0, (parseInt(document.getElementById('excel-header-row').value, 10) || 1) - 1);
    const sheets = parseWorkbookSheets(currentWorkbook, headerRowIndex, 'COLUMNA_');
    const selectedSheet = sheets.find(sheet => sheet.sheetName === sheetName);
    if (!selectedSheet) return;

    parsedPayload = {
      fileName: currentFileName,
      ...selectedSheet,
      sheets,
      rawBuffer: currentFileArrayBuffer
    };

    // Renderizar tabla previa (primeras 5 filas)
    const statsBox = document.getElementById('excel-preview-stats');
    statsBox.innerHTML = `<strong>Hojas del archivo:</strong> ${sheets.length} | <strong>Columnas:</strong> ${selectedSheet.headers.length} | <strong>Registros:</strong> ${selectedSheet.rows.length}`;

    const table = document.getElementById('excel-preview-table');
    table.innerHTML = '';

    const thead = document.createElement('thead');
    const trH = document.createElement('tr');
    selectedSheet.headers.forEach(h => {
      const th = document.createElement('th');
      th.textContent = h;
      trH.appendChild(th);
    });
    thead.appendChild(trH);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    selectedSheet.rows.slice(0, 5).forEach((row, index) => {
      const tr = document.createElement('tr');
      selectedSheet.headers.forEach(h => {
        const td = document.createElement('td');
        td.textContent = selectedSheet.displayRows[index]?.[h] ?? row[h];
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);

    document.getElementById('excel-preview-box').classList.remove('hidden');
  }

  function getParsedPayload() {
    return parsedPayload;
  }

  function getFormattedSheets(rawBuffer, headerRow = 1) {
    const workbook = XLSX.read(new Uint8Array(rawBuffer), { type: 'array', cellNF: true });
    const headerRowIndex = Math.max(0, (Number(headerRow) || 1) - 1);
    return parseWorkbookSheets(workbook, headerRowIndex, 'COLUMNA_');
  }

  function parseArrayBuffer(rawBuffer, headerRow = 1) {
    if (!rawBuffer) throw new Error('El archivo Excel no contiene datos.');
    const workbook = XLSX.read(new Uint8Array(rawBuffer), { type: 'array', cellNF: true });
    const headerRowIndex = Math.max(0, (Number(headerRow) || 1) - 1);
    return parseWorkbookSheets(workbook, headerRowIndex, 'COLUMNA_');
  }

  function resetPayload() {
    currentFileArrayBuffer = null;
    currentFileName = '';
    currentWorkbook = null;
    parsedPayload = null;
    const previewBox = document.getElementById('excel-preview-box');
    if (previewBox) previewBox.classList.add('hidden');
    const selectorBox = document.getElementById('excel-sheet-selector-box');
    if (selectorBox) selectorBox.classList.add('hidden');
  }

  // ==================== ASISTENTE DE ACTUALIZACIÓN ====================

  async function handleUpdateFileSelection(event) {
    const file = event.target.files[0];
    if (!file) return;

    updateFileName = file.name;
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        updateFileArrayBuffer = e.target.result;
        updateWorkbook = XLSX.read(new Uint8Array(updateFileArrayBuffer), { type: 'array', cellNF: true });

        const sheetSelect = document.getElementById('update-excel-select-sheet');
        sheetSelect.innerHTML = '';
        updateWorkbook.SheetNames.forEach((sheetName) => {
          const opt = document.createElement('option');
          opt.value = sheetName;
          opt.textContent = sheetName;
          sheetSelect.appendChild(opt);
        });

        document.getElementById('update-excel-sheet-box').classList.remove('hidden');
        renderUpdatePreview();
      } catch (err) {
        App.showToast('Error al leer el archivo Excel: ' + err.message, 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function handleUpdateSheetChange() { renderUpdatePreview(); }
  function handleUpdateHeaderRowChange() { renderUpdatePreview(); }

  function renderUpdatePreview() {
    if (!updateWorkbook) return;

    const sheetName = document.getElementById('update-excel-select-sheet').value || updateWorkbook.SheetNames[0];
    const headerRowIndex = Math.max(0, (parseInt(document.getElementById('update-excel-header-row').value, 10) || 1) - 1);
    const sheets = parseWorkbookSheets(updateWorkbook, headerRowIndex, 'COL_');
    const selectedSheet = sheets.find(sheet => sheet.sheetName === sheetName);
    if (!selectedSheet) return;

    updateParsedPayload = {
      fileName: updateFileName,
      ...selectedSheet,
      sheets,
      rawBuffer: updateFileArrayBuffer
    };

    const statsBox = document.getElementById('update-excel-preview-stats');
    statsBox.innerHTML = `<strong>Hojas del archivo:</strong> ${sheets.length} | <strong>Columnas:</strong> ${selectedSheet.headers.length} | <strong>Registros:</strong> ${selectedSheet.rows.length}`;

    const table = document.getElementById('update-excel-preview-table');
    table.innerHTML = '';
    const thead = document.createElement('thead');
    const trH = document.createElement('tr');
    selectedSheet.headers.forEach(h => {
      const th = document.createElement('th');
      th.textContent = h;
      trH.appendChild(th);
    });
    thead.appendChild(trH);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    selectedSheet.rows.slice(0, 5).forEach((row, index) => {
      const tr = document.createElement('tr');
      selectedSheet.headers.forEach(h => {
        const td = document.createElement('td');
        td.textContent = selectedSheet.displayRows[index]?.[h] ?? row[h];
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);

    document.getElementById('update-excel-preview-box').classList.remove('hidden');
  }

  function getUpdateParsedPayload() {
    return updateParsedPayload;
  }

  function resetUpdatePayload() {
    updateFileArrayBuffer = null;
    updateFileName = '';
    updateWorkbook = null;
    updateParsedPayload = null;
    const pBox = document.getElementById('update-excel-preview-box');
    if (pBox) pBox.classList.add('hidden');
    const sBox = document.getElementById('update-excel-sheet-box');
    if (sBox) sBox.classList.add('hidden');
  }

  /**
   * Exporta cualquier arreglo de objetos JSON a un archivo .xlsx y dispara la descarga en navegador
   */
  function exportJsonToExcel(dataArray, defaultFilename = 'Exportacion.xlsx', headers, numberFormats) {
    if (!dataArray || dataArray.length === 0) {
      App.showToast('No existen registros para exportar.', 'error');
      return;
    }

    const columnHeaders = headers || Array.from(new Set(dataArray.flatMap(row => Object.keys(row))));
    const worksheet = XLSX.utils.json_to_sheet(dataArray, { header: columnHeaders });
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    const headerStyle = {
      fill: { patternType: 'solid', fgColor: { rgb: '0F2D52' } },
      font: { name: 'Aptos', sz: 11, bold: true, color: { rgb: 'FFFFFF' } },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      border: {
        top: { style: 'thin', color: { rgb: '0F2D52' } },
        bottom: { style: 'medium', color: { rgb: '0284C7' } },
        left: { style: 'thin', color: { rgb: '0F2D52' } },
        right: { style: 'thin', color: { rgb: '0F2D52' } }
      }
    };
    const border = {
      top: { style: 'thin', color: { rgb: 'D8E1EB' } },
      bottom: { style: 'thin', color: { rgb: 'D8E1EB' } },
      left: { style: 'thin', color: { rgb: 'D8E1EB' } },
      right: { style: 'thin', color: { rgb: 'D8E1EB' } }
    };

    for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex++) {
      const headerCell = worksheet[XLSX.utils.encode_cell({ r: 0, c: columnIndex })];
      if (headerCell) headerCell.s = headerStyle;

      const header = columnHeaders[columnIndex] || '';
      const normalizedHeader = String(header).toLocaleLowerCase();
      const isPercentage = normalizedHeader.includes('%') || normalizedHeader.includes('porcentaje');
      const isCurrency = /precio|valor|diferencia|total|unitario|tarifa/.test(normalizedHeader);
      const columnFormat = isPercentage ? '0"%"' : (isCurrency ? '#,##0;[Red]-#,##0' : null);
      const sampleValues = dataArray.slice(0, 100).map(row => String(row[header] ?? ''));
      const contentWidth = Math.max(
        String(header).length,
        ...sampleValues.map(value => value.length)
      );
      worksheet['!cols'] = worksheet['!cols'] || [];
      worksheet['!cols'][columnIndex] = {
        wch: Math.min(Math.max(contentWidth + 3, 12), 42)
      };

      for (let rowIndex = 1; rowIndex <= range.e.r; rowIndex++) {
        const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
        const cell = worksheet[cellAddress];
        if (!cell) continue;
        cell.s = {
          fill: {
            patternType: 'solid',
            fgColor: { rgb: rowIndex % 2 === 0 ? 'F1F6FB' : 'FFFFFF' }
          },
          font: { name: 'Aptos', sz: 10, color: { rgb: '1E293B' } },
          alignment: {
            vertical: 'center',
            horizontal: typeof cell.v === 'number' ? 'right' : 'left'
          },
          border
        };
        if (columnFormat && typeof cell.v === 'number') cell.z = columnFormat;
      }
    }

    worksheet['!rows'] = [{ hpt: 30 }];
    worksheet['!autofilter'] = { ref: XLSX.utils.encode_range(range) };
    dataArray.forEach((row, rowIndex) => {
      const statusEntry = Object.entries(row).find(([header]) =>
        /comparaci[oó]n|estado/i.test(header)
      );
      const statusCellValue = statusEntry?.[1];
      if (!statusCellValue) return;
      const status = String(statusCellValue).toLocaleLowerCase();
      const statusStyle = ['igual', 'aprobada', 'activo'].includes(status)
        ? { fill: 'DCFCE7', font: '15803D' }
        : ['diferente', 'rechazada', 'finalizado', 'suspendido'].includes(status)
          ? { fill: 'FEE2E2', font: 'B91C1C' }
          : ['pendiente', 'en liquidación'].includes(status)
            ? { fill: 'FEF3C7', font: 'B45309' }
            : { fill: 'E2E8F0', font: '475569' };
      columnHeaders.forEach((header, columnIndex) => {
        if (header !== statusEntry[0]) return;
        const cell = worksheet[XLSX.utils.encode_cell({ r: rowIndex + 1, c: columnIndex })];
        if (cell) {
          cell.s = {
            ...cell.s,
            fill: {
              patternType: 'solid',
              fgColor: { rgb: statusStyle.fill }
            },
            font: { name: 'Aptos', sz: 10, bold: true, color: { rgb: statusStyle.font } },
            alignment: { horizontal: 'center', vertical: 'center' }
          };
        }
      });
    });

    if (headers && numberFormats) {
      numberFormats.forEach((rowFormats, rowIndex) => {
        headers.forEach((header, columnIndex) => {
          const format = rowFormats[header];
          const cellAddress = XLSX.utils.encode_cell({ r: rowIndex + 1, c: columnIndex });
          const cell = worksheet[cellAddress];
          if (format && cell) cell.z = format;
        });
      });
    }
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Datos');
    workbook.Props = {
      Title: defaultFilename.replace(/\.xlsx$/i, ''),
      Subject: 'Exportación de datos SIGEC-IPS',
      Creator: 'SIGEC-IPS'
    };
    XLSX.writeFile(workbook, defaultFilename);
  }

  return {
    handleFileSelection,
    handleSheetChange,
    handleHeaderRowChange,
    getParsedPayload,
    getFormattedSheets,
    parseArrayBuffer,
    resetPayload,
    handleUpdateFileSelection,
    handleUpdateSheetChange,
    handleUpdateHeaderRowChange,
    getUpdateParsedPayload,
    resetUpdatePayload,
    exportJsonToExcel
  };
})();

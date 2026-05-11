(function () {
  "use strict";

  const REQUIRED_HEADERS = [
    "Área",
    "Zona",
    "Centro",
    "CIAS",
    "PROFESIONAL",
    "Tipo visita",
    "Accesibilidad",
    "Categoría",
    "Código de centro",
    "Fecha Primer Hueco Libre",
    "Fecha Primer Hueco ID",
    "Fecha Corte"
  ];

  const IGNORED_HEADERS = new Set([
    "CIAS ID",
    "UID"
  ]);

  const DATE_COLUMNS = new Set([
    "Fecha Primer Hueco Libre",
    "Fecha Primer Hueco ID",
    "Fecha Corte"
  ]);

  const FILTERS = [
    { key: "area", column: "Área", elementId: "filterArea" },
    { key: "zona", column: "Zona", elementId: "filterZona" },
    { key: "centro", column: "Centro", elementId: "filterCentro" },
    { key: "profesional", column: "PROFESIONAL", elementId: "filterProfesional" },
    { key: "categoria", column: "Categoría", elementId: "filterCategoria" },
    { key: "tipoVisita", column: "Tipo visita", elementId: "filterTipoVisita" }
  ];

  const PALETTE = {
    blue: "#2876b9",
    teal: "#0e9d90",
    green: "#169b72",
    amber: "#ecb343",
    red: "#d55663",
    slate: "#5f7e9b",
    sky: "#4ca0df"
  };

  const state = {
    activeTab: "upload",
    workbookName: "",
    sheetName: "",
    loadedAt: "",
    rows: [],
    filters: {
      area: "",
      zona: "",
      centro: "",
      profesional: "",
      categoria: "",
      tipoVisita: ""
    },
    tableSearch: "",
    tableSort: {
      column: "Accesibilidad",
      direction: "desc"
    },
    currentPage: 1,
    pageSize: 20
  };

  const refs = {
    statusIndicator: document.getElementById("statusIndicator"),
    statusLabel: document.getElementById("statusLabel"),
    statusMessage: document.getElementById("statusMessage"),
    statusMeta: document.getElementById("statusMeta"),
    uploadFeedback: document.getElementById("uploadFeedback"),
    expectedColumnsList: document.getElementById("expectedColumnsList"),
    selectFileButton: document.getElementById("selectFileButton"),
    fileInput: document.getElementById("fileInput"),
    filtersPanel: document.getElementById("filtersPanel"),
    filtersSummary: document.getElementById("filtersSummary"),
    resetFiltersButton: document.getElementById("resetFiltersButton"),
    exportCsvButton: document.getElementById("exportCsvButton"),
    loadedFileName: document.getElementById("loadedFileName"),
    loadedSheetName: document.getElementById("loadedSheetName"),
    loadedRowsCount: document.getElementById("loadedRowsCount"),
    loadedValidCount: document.getElementById("loadedValidCount"),
    loadedAt: document.getElementById("loadedAt"),
    loadedState: document.getElementById("loadedState"),
    tabButtons: Array.from(document.querySelectorAll(".tab-button")),
    tabPanels: Array.from(document.querySelectorAll(".tab-panel")),
    tableSearchInput: document.getElementById("tableSearchInput"),
    pageSizeSelect: document.getElementById("pageSizeSelect"),
    tableVisibleCount: document.getElementById("tableVisibleCount"),
    tableValidVisibleCount: document.getElementById("tableValidVisibleCount"),
    tableEmptyState: document.getElementById("tableEmptyState"),
    tableWrapper: document.getElementById("tableWrapper"),
    dataTableHead: document.getElementById("dataTableHead"),
    dataTableBody: document.getElementById("dataTableBody"),
    paginationInfo: document.getElementById("paginationInfo"),
    prevPageButton: document.getElementById("prevPageButton"),
    nextPageButton: document.getElementById("nextPageButton"),
    summaryScope: document.getElementById("summaryScope"),
    kpiPct0to2: document.getElementById("kpiPct0to2"),
    kpiPct0to3: document.getElementById("kpiPct0to3"),
    kpiPct0to6: document.getElementById("kpiPct0to6"),
    kpiPct7Plus: document.getElementById("kpiPct7Plus"),
    kpiTotalValid: document.getElementById("kpiTotalValid"),
    kpiMean: document.getElementById("kpiMean"),
    kpiMedian: document.getElementById("kpiMedian"),
    kpiMax: document.getElementById("kpiMax"),
    summaryBandsChart: document.getElementById("summaryBandsChart"),
    summaryDonutChart: document.getElementById("summaryDonutChart"),
    centersChart: document.getElementById("centersChart"),
    zonesChart: document.getElementById("zonesChart"),
    categoryChart: document.getElementById("categoryChart"),
    visitTypeChart: document.getElementById("visitTypeChart"),
    histogramChart: document.getElementById("histogramChart"),
    centerBandsChart: document.getElementById("centerBandsChart"),
    printReportButton: document.getElementById("printReportButton"),
    reportScope: document.getElementById("reportScope"),
    reportArea: document.getElementById("reportArea"),
    reportZones: document.getElementById("reportZones"),
    reportZonesDetail: document.getElementById("reportZonesDetail"),
    reportCutoffDate: document.getElementById("reportCutoffDate"),
    reportSummaryBody: document.getElementById("reportSummaryBody"),
    reportEmptyState: document.getElementById("reportEmptyState"),
    reportContent: document.getElementById("reportContent"),
    reportKpiPct0to2: document.getElementById("reportKpiPct0to2"),
    reportKpiPct0to3: document.getElementById("reportKpiPct0to3"),
    reportKpiPct0to6: document.getElementById("reportKpiPct0to6"),
    reportKpiPct7Plus: document.getElementById("reportKpiPct7Plus"),
    reportKpiTotalValid: document.getElementById("reportKpiTotalValid"),
    reportKpiMean: document.getElementById("reportKpiMean"),
    reportKpiMedian: document.getElementById("reportKpiMedian"),
    reportKpiMax: document.getElementById("reportKpiMax"),
    reportBandsChart: document.getElementById("reportBandsChart"),
    reportCentersChart: document.getElementById("reportCentersChart")
  };

  FILTERS.forEach((filter) => {
    refs[filter.key] = document.getElementById(filter.elementId);
  });

  initialize();

  function initialize() {
    renderExpectedColumns();
    renderTableHeader();
    bindEvents();
    renderAll();
  }

  function bindEvents() {
    refs.selectFileButton.addEventListener("click", () => refs.fileInput.click());
    refs.fileInput.addEventListener("change", handleFileSelection);
    refs.resetFiltersButton.addEventListener("click", resetFilters);
    refs.exportCsvButton.addEventListener("click", exportFilteredTable);
    refs.tableSearchInput.addEventListener("input", handleTableSearch);
    refs.pageSizeSelect.addEventListener("change", handlePageSizeChange);
    refs.prevPageButton.addEventListener("click", () => changePage(-1));
    refs.nextPageButton.addEventListener("click", () => changePage(1));
    refs.dataTableHead.addEventListener("click", handleSortClick);

    refs.tabButtons.forEach((button) => {
      button.addEventListener("click", () => setActiveTab(button.dataset.tabTarget || "upload"));
    });

    FILTERS.forEach((filter) => {
      refs[filter.key].addEventListener("change", () => {
        state.filters[filter.key] = refs[filter.key].value;
        state.currentPage = 1;
        renderAll();
      });
    });

    document.addEventListener("click", async (event) => {
      const exportButton = event.target.closest("[data-chart-export]");
      if (!exportButton) {
        return;
      }

      const chartId = exportButton.getAttribute("data-chart-export");
      const container = document.getElementById(chartId);
      if (!container) {
        return;
      }

      try {
        await window.DemorasCharts.exportContainerAsPng(container, buildExportFileName(chartId, "png"));
      } catch (error) {
        setStatus("error", "Error al exportar", error.message, "Revise si el grafico tiene datos visibles.");
      }
    });
  }

  function renderExpectedColumns() {
    refs.expectedColumnsList.innerHTML = REQUIRED_HEADERS.map((header) => "<span>" + escapeHtml(header) + "</span>").join("");
  }

  function renderTableHeader() {
    const cells = REQUIRED_HEADERS.map((header) => {
      const isActive = state.tableSort.column === header;
      const icon = isActive ? (state.tableSort.direction === "asc" ? "▲" : "▼") : "↕";
      return (
        "<th>" +
        '<button class="sort-button' +
        (isActive ? " is-active" : "") +
        '" type="button" data-sort-column="' +
        escapeHtml(header) +
        '">' +
        "<span>" +
        escapeHtml(header) +
        '</span><span class="sort-button__icon">' +
        icon +
        "</span></button></th>"
      );
    }).join("");

    refs.dataTableHead.innerHTML = "<tr>" + cells + "</tr>";
  }

  async function handleFileSelection(event) {
    const [file] = event.target.files || [];
    event.target.value = "";

    if (!file) {
      return;
    }

    if (!/\.xlsx$/i.test(file.name)) {
      clearData();
      setFeedback("error", "El archivo seleccionado no es un .xlsx valido.");
      setStatus("error", "Archivo no valido", "Seleccione un libro de Excel con extension .xlsx.", "Columnas obligatorias: " + REQUIRED_HEADERS.join(", "));
      updateLoadProfile("Error", file.name, "-", 0, 0);
      renderAll();
      return;
    }

    setFeedback("info", "Leyendo la primera hoja del archivo y validando columnas...");
    setStatus("info", "Procesando archivo", "Se esta leyendo el libro Excel seleccionado.", file.name);
    updateLoadProfile("Procesando", file.name, "-", 0, 0);

    try {
      const workbook = await window.XLSXLite.readWorkbook(file);
      const workbookData = readSheetFromRawRows(workbook.rawRows || []);

      state.workbookName = file.name;
      state.sheetName = workbook.sheetName;
      state.loadedAt = formatDateTime(new Date());
      state.rows = normalizeRows(workbookData.rows);
      state.currentPage = 1;
      state.tableSearch = "";
      state.tableSort = {
        column: "Accesibilidad",
        direction: "desc"
      };
      refs.tableSearchInput.value = "";
      refs.pageSizeSelect.value = String(state.pageSize);
      resetFilterValues();
      renderTableHeader();

      const validCount = state.rows.filter((row) => row.__hasValidAcc).length;
      updateLoadProfile("Correcto", file.name, workbook.sheetName, state.rows.length, validCount);
      setFeedback("success", "Carga correcta. Ya puede consultar la tabla, el resumen ejecutivo y los graficos.");
      setStatus(
        "success",
        "Archivo cargado",
        "Se han leido " + formatInteger(state.rows.length) + " filas y " + formatInteger(validCount) + " agendas validas.",
        file.name + " · Hoja: " + workbook.sheetName
      );
      setActiveTab("summary");
      renderAll();
    } catch (error) {
      if (error && error.code === "INVALID_SHEET_STRUCTURE") {
        clearData();
        updateLoadProfile("Error estructural", file.name, "-", error.rowCount || 0, 0);
        setFeedback("error", error.message);
        setStatus(
          "error",
          "Columnas no validas",
          "El archivo no contiene todas las cabeceras obligatorias.",
          "Detectadas: " + formatDetectedHeaders(error.detectedHeaders)
        );
        renderAll();
        setActiveTab("upload");
        return;
      }

      clearData();
      updateLoadProfile("Error", file.name, "-", 0, 0);
      setFeedback("error", error.message || "No se ha podido leer el archivo Excel.");
      setStatus(
        "error",
        "Error de lectura",
        error.message || "No se ha podido leer el archivo Excel.",
        "Asegurese de usar un libro .xlsx valido con la estructura requerida."
      );
      renderAll();
      setActiveTab("upload");
    }
  }

  function normalizeRows(rows) {
    return rows.map((rawRow, index) => {
      const normalizedRow = {};
      const dateSortValues = {};

      REQUIRED_HEADERS.forEach((header) => {
        const rawValue = rawRow[header];
        if (DATE_COLUMNS.has(header)) {
          const dateInfo = normalizeDateValue(rawValue);
          normalizedRow[header] = dateInfo.display;
          dateSortValues[header] = dateInfo.sortValue;
        } else {
          normalizedRow[header] = normalizeText(rawValue);
        }
      });

      const accessibilityValue = parseNumericValue(rawRow["Accesibilidad"]);
      return Object.assign(normalizedRow, {
        __rowId: index + 1,
        __acc: accessibilityValue,
        __hasValidAcc: Number.isFinite(accessibilityValue),
        __dateSort: dateSortValues
      });
    });
  }

  function readSheetFromRawRows(rawRows) {
    const raw = Array.isArray(rawRows) ? rawRows.filter((row) => !isEmptyRow(row)) : [];
    if (!raw.length) {
      throw new Error("La hoja está vacía.");
    }

    let headers = (raw[0] || []).map((value) => normalizeHeader(value));
    headers = headers.map((header) => (IGNORED_HEADERS.has(header) ? null : header));
    headers = dedupeHeaders(headers);

    const detectedHeaders = headers.filter(Boolean);
    const missingHeaders = REQUIRED_HEADERS.filter((column) => !detectedHeaders.includes(column));
    const rows = raw
      .slice(1)
      .filter((row) => !isEmptyRow(row))
      .map((row) => buildRowObject(row, headers));

    if (missingHeaders.length > 0) {
      throw createSheetStructureError(missingHeaders, detectedHeaders, rows.length);
    }

    return {
      headers: detectedHeaders,
      rows: rows,
      rowCount: rows.length
    };
  }

  function buildRowObject(row, processedHeaders) {
    const rowObject = {};

    processedHeaders.forEach((header, index) => {
      if (header) {
        rowObject[header] = row[index] == null ? "" : row[index];
      }
    });

    return rowObject;
  }

  function normalizeHeader(value) {
    return String(value == null ? "" : value)
      .trim()
      .replace(/\s+/g, " ");
  }

  function dedupeHeaders(headers) {
    const seenHeaders = Object.create(null);

    return headers.map((header) => {
      const normalizedHeader = normalizeHeader(header);
      if (!normalizedHeader) {
        return null;
      }
      if (!seenHeaders[normalizedHeader]) {
        seenHeaders[normalizedHeader] = 1;
        return normalizedHeader;
      }
      return normalizedHeader + "." + seenHeaders[normalizedHeader]++;
    });
  }

  function isEmptyRow(row) {
    return !Array.isArray(row) || row.every((value) => normalizeText(value) === "");
  }

  function createSheetStructureError(missingHeaders, detectedHeaders, rowCount) {
    const error = new Error(
      "Faltan columnas obligatorias: " +
        missingHeaders.join(", ") +
        ". Cabeceras detectadas: " +
        formatDetectedHeaders(detectedHeaders) +
        "."
    );
    error.code = "INVALID_SHEET_STRUCTURE";
    error.missingHeaders = missingHeaders.slice();
    error.detectedHeaders = detectedHeaders.slice();
    error.rowCount = rowCount || 0;
    return error;
  }

  function normalizeDateValue(value) {
    const textValue = value == null ? "" : String(value).trim();
    if (!textValue) {
      return { display: "", sortValue: null };
    }

    if (/^-?\d+(?:[.,]\d+)?$/.test(textValue)) {
      const serial = parseFloat(textValue.replace(",", "."));
      if (Number.isFinite(serial) && serial > 0) {
        const date = excelSerialToDate(serial);
        if (date) {
          return {
            display: formatDate(date),
            sortValue: date.getTime()
          };
        }
      }
    }

    const parsed = parseDateString(textValue);
    if (parsed) {
      return {
        display: formatDate(parsed),
        sortValue: parsed.getTime()
      };
    }

    return {
      display: textValue,
      sortValue: null
    };
  }

  function excelSerialToDate(serial) {
    const wholeDays = Math.floor(serial);
    const adjustedDays = wholeDays > 59 ? wholeDays - 1 : wholeDays;
    const milliseconds = Date.UTC(1899, 11, 31) + adjustedDays * 86400000;
    const date = new Date(milliseconds);
    return Number.isNaN(date.getTime()) ? null : date;
  }

  function parseDateString(value) {
    const isoMatch = value.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (isoMatch) {
      return new Date(Date.UTC(Number(isoMatch[1]), Number(isoMatch[2]) - 1, Number(isoMatch[3])));
    }

    const slashMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (slashMatch) {
      const day = Number(slashMatch[1]);
      const month = Number(slashMatch[2]) - 1;
      const year = Number(slashMatch[3].length === 2 ? "20" + slashMatch[3] : slashMatch[3]);
      return new Date(Date.UTC(year, month, day));
    }

    return null;
  }

  function setActiveTab(tabName) {
    state.activeTab = tabName;
    refs.tabButtons.forEach((button) => {
      button.classList.toggle("is-active", button.dataset.tabTarget === tabName);
    });
    refs.tabPanels.forEach((panel) => {
      panel.classList.toggle("is-active", panel.dataset.tabPanel === tabName);
    });
  }

  function renderAll() {
    syncFilterSelects();
    renderFilterSummary();
    renderTableHeader();
    renderTable();
    renderSummary();
    renderCharts();
    renderReport();
    toggleControls();
  }

  function syncFilterSelects() {
    let safetyCounter = 0;
    let changed = true;

    while (changed && safetyCounter < 5) {
      changed = false;
      safetyCounter += 1;

      FILTERS.forEach((filter) => {
        const options = uniqueSorted(
          getRowsForFilterOptions(filter.key).map((row) => normalizeText(row[filter.column])).filter(Boolean)
        );

        if (state.filters[filter.key] && !options.includes(state.filters[filter.key])) {
          state.filters[filter.key] = "";
          changed = true;
        }

        renderSelectOptions(refs[filter.key], options, state.filters[filter.key], !state.rows.length);
      });
    }
  }

  function renderSelectOptions(selectElement, values, currentValue, disabled) {
    const optionsMarkup = ['<option value="">Todos</option>']
      .concat(
        values.map((value) => {
          const selected = value === currentValue ? ' selected' : "";
          return '<option value="' + escapeHtml(value) + '"' + selected + ">" + escapeHtml(value) + "</option>";
        })
      )
      .join("");

    selectElement.innerHTML = optionsMarkup;
    selectElement.disabled = Boolean(disabled);
  }

  function getRowsForFilterOptions(filterKey) {
    return state.rows.filter((row) => {
      return FILTERS.every((filter) => {
        if (filter.key === filterKey) {
          return true;
        }
        const selectedValue = state.filters[filter.key];
        return !selectedValue || row[filter.column] === selectedValue;
      });
    });
  }

  function renderFilterSummary() {
    if (!state.rows.length) {
      refs.filtersPanel.classList.add("is-disabled");
      refs.filtersSummary.textContent = "Sin datos cargados.";
      return;
    }

    refs.filtersPanel.classList.remove("is-disabled");
    const filteredRows = getFilteredRows();
    const validRows = filteredRows.filter((row) => row.__hasValidAcc);
    const parts = FILTERS.map((filter) => {
      const value = state.filters[filter.key] || "Todos";
      return filter.column + ": " + value;
    });

    parts.push("Filas filtradas: " + formatInteger(filteredRows.length));
    parts.push("Agendas validas: " + formatInteger(validRows.length));
    refs.filtersSummary.textContent = parts.join(" | ");
  }

  function toggleControls() {
    const hasRows = state.rows.length > 0;
    const tableRows = getTableWorkingRows();
    refs.tableSearchInput.disabled = !hasRows;
    refs.pageSizeSelect.disabled = !hasRows;
    refs.resetFiltersButton.disabled = !hasRows;
    refs.exportCsvButton.disabled = !tableRows.length;
    refs.printReportButton.disabled = !hasRows;
    refs.prevPageButton.disabled = !hasRows || state.currentPage <= 1;
    refs.nextPageButton.disabled = !hasRows || state.currentPage >= getTotalPages(tableRows.length);
  }

  function renderTable() {
    if (!state.rows.length) {
      refs.tableEmptyState.classList.remove("is-hidden");
      refs.tableWrapper.classList.add("is-hidden");
      refs.tableEmptyState.textContent = "Cargue un Excel para visualizar la tabla.";
      refs.tableVisibleCount.textContent = "0 filas visibles";
      refs.tableValidVisibleCount.textContent = "0 agendas validas para analisis";
      refs.dataTableBody.innerHTML = "";
      refs.paginationInfo.textContent = "Pagina 0 de 0";
      return;
    }

    const tableRows = getTableWorkingRows();
    const validVisibleRows = tableRows.filter((row) => row.__hasValidAcc).length;
    refs.tableVisibleCount.textContent = formatInteger(tableRows.length) + " filas visibles";
    refs.tableValidVisibleCount.textContent = formatInteger(validVisibleRows) + " agendas validas para analisis";

    if (!tableRows.length) {
      refs.tableEmptyState.classList.remove("is-hidden");
      refs.tableWrapper.classList.add("is-hidden");
      refs.tableEmptyState.textContent = "No hay filas que coincidan con los filtros y la busqueda actual.";
      refs.dataTableBody.innerHTML = "";
      refs.paginationInfo.textContent = "Pagina 0 de 0";
      return;
    }

    refs.tableEmptyState.classList.add("is-hidden");
    refs.tableWrapper.classList.remove("is-hidden");
    const totalPages = getTotalPages(tableRows.length);
    state.currentPage = clamp(state.currentPage, 1, totalPages);
    const pageStart = (state.currentPage - 1) * state.pageSize;
    const pageRows = tableRows.slice(pageStart, pageStart + state.pageSize);

    refs.dataTableBody.innerHTML = pageRows
      .map((row) => {
        const cells = REQUIRED_HEADERS.map((header) => "<td>" + renderCellValue(header, row) + "</td>").join("");
        return "<tr>" + cells + "</tr>";
      })
      .join("");

    refs.paginationInfo.textContent = "Pagina " + state.currentPage + " de " + totalPages;
  }

  function renderSummary() {
    const filteredRows = getFilteredRows();
    const validRows = filteredRows.filter((row) => row.__hasValidAcc);
    const metrics = calculateExecutiveMetrics(validRows);

    refs.summaryScope.textContent = validRows.length
      ? formatInteger(validRows.length) + " agendas validas tras filtros"
      : "Sin datos analiticos";

    if (!metrics) {
      setKpiValue(refs.kpiPct0to2, "--");
      setKpiValue(refs.kpiPct0to3, "--");
      setKpiValue(refs.kpiPct0to6, "--");
      setKpiValue(refs.kpiPct7Plus, "--");
      setKpiValue(refs.kpiTotalValid, "--");
      setKpiValue(refs.kpiMean, "--");
      setKpiValue(refs.kpiMedian, "--");
      setKpiValue(refs.kpiMax, "--");
      window.DemorasCharts.setEmpty(refs.summaryBandsChart, "No hay agendas validas con los filtros actuales.");
      window.DemorasCharts.setEmpty(refs.summaryDonutChart, "No hay agendas validas con los filtros actuales.");
      return;
    }

    setKpiValue(refs.kpiPct0to2, formatPercent(metrics.pct0to2));
    setKpiValue(refs.kpiPct0to3, formatPercent(metrics.pct0to3));
    setKpiValue(refs.kpiPct0to6, formatPercent(metrics.pct0to6));
    setKpiValue(refs.kpiPct7Plus, formatPercent(metrics.pct7Plus));
    setKpiValue(refs.kpiTotalValid, formatInteger(metrics.totalValid));
    setKpiValue(refs.kpiMean, formatDayMetric(metrics.mean));
    setKpiValue(refs.kpiMedian, formatDayMetric(metrics.median));
    setKpiValue(refs.kpiMax, formatDayMetric(metrics.max));

    renderSummaryCharts(validRows);
  }

  function renderSummaryCharts(validRows) {
    const delayBands = buildExecutiveDelayBands(validRows);
    const total = validRows.length;

    window.DemorasCharts.renderVerticalBarChart(refs.summaryBandsChart, {
      items: delayBands.map((band) => ({
        label: band.label,
        value: band.count,
        color: band.color,
        share: band.share
      })),
      valueFormatter: (value) => formatInteger(value),
      annotationFormatter: (item) => formatPercent(item.share),
      titleFormatter: (item) => formatInteger(item.value) + " agendas"
    });

    window.DemorasCharts.renderDonutChart(refs.summaryDonutChart, {
      segments: [
        { label: "0-2 días", value: delayBands[0].count, color: delayBands[0].color },
        { label: "3 días", value: delayBands[1].count, color: delayBands[1].color },
        { label: "4-6 días", value: delayBands[2].count, color: delayBands[2].color },
        { label: "7 o más días", value: delayBands[3].count, color: PALETTE.red }
      ],
      centerLabel: "Agendas",
      centerValue: formatInteger(total)
    });
  }

  function renderReport() {
    const filteredRows = getFilteredRows();
    const validRows = filteredRows.filter((row) => row.__hasValidAcc);
    const metrics = calculateExecutiveMetrics(validRows);
    const latestCutoff = getLatestDateSortValue(filteredRows, "Fecha Corte");
    const zoneMeta = getReportZoneMeta(filteredRows);

    refs.reportArea.textContent = state.filters.area || "Todas las áreas";
    refs.reportZones.textContent = zoneMeta.value;
    refs.reportZonesDetail.textContent = zoneMeta.detail;
    refs.reportCutoffDate.textContent = latestCutoff == null ? "Fecha de corte no disponible" : formatLongDate(latestCutoff);
    refs.reportScope.textContent = metrics
      ? formatInteger(metrics.totalValid) + " agendas válidas analizadas"
      : "Sin información analítica";

    if (!state.rows.length) {
      refs.reportSummaryBody.innerHTML = "<p>Cargue un archivo Excel para generar el informe.</p>";
      refs.reportEmptyState.classList.remove("is-hidden");
      refs.reportEmptyState.textContent = "Cargue un archivo Excel para comenzar la elaboración del informe.";
      refs.reportContent.classList.add("is-hidden");
      setReportKpisEmpty();
      window.DemorasCharts.setEmpty(refs.reportBandsChart, "Cargue un Excel para generar el gráfico del informe.");
      window.DemorasCharts.setEmpty(refs.reportCentersChart, "Cargue un Excel para generar el gráfico del informe.");
      return;
    }

    refs.reportSummaryBody.innerHTML = buildReportSummaryMarkup(filteredRows, metrics, latestCutoff);

    if (!metrics) {
      refs.reportEmptyState.classList.remove("is-hidden");
      refs.reportEmptyState.textContent = filteredRows.length
        ? "No hay agendas válidas con Accesibilidad numérica en el subconjunto filtrado."
        : "No hay información disponible para generar el informe con los filtros actuales.";
      refs.reportContent.classList.add("is-hidden");
      setReportKpisEmpty();
      window.DemorasCharts.setEmpty(refs.reportBandsChart, "No hay agendas válidas con los filtros actuales.");
      window.DemorasCharts.setEmpty(refs.reportCentersChart, "No hay agendas válidas con los filtros actuales.");
      return;
    }

    refs.reportEmptyState.classList.add("is-hidden");
    refs.reportContent.classList.remove("is-hidden");
    setKpiValue(refs.reportKpiPct0to2, formatPercent(metrics.pct0to2));
    setKpiValue(refs.reportKpiPct0to3, formatPercent(metrics.pct0to3));
    setKpiValue(refs.reportKpiPct0to6, formatPercent(metrics.pct0to6));
    setKpiValue(refs.reportKpiPct7Plus, formatPercent(metrics.pct7Plus));
    setKpiValue(refs.reportKpiTotalValid, formatInteger(metrics.totalValid));
    setKpiValue(refs.reportKpiMean, formatDayMetric(metrics.mean));
    setKpiValue(refs.reportKpiMedian, formatDayMetric(metrics.median));
    setKpiValue(refs.reportKpiMax, formatDayMetric(metrics.max));

    renderReportCharts(validRows);
  }

  function renderReportCharts(validRows) {
    const delayBands = buildExecutiveDelayBands(validRows);
    const topCenters = aggregateByField(validRows, "Centro")
      .sort((a, b) => b.mean - a.mean || b.count - a.count)
      .slice(0, 6)
      .map((item) => ({
        label: item.label,
        value: item.mean,
        color: PALETTE.blue,
        note: "n=" + formatInteger(item.count)
      }));

    window.DemorasCharts.renderVerticalBarChart(refs.reportBandsChart, {
      items: delayBands.map((band) => ({
        label: band.label,
        value: band.count,
        color: band.color,
        share: band.share
      })),
      valueFormatter: formatInteger,
      annotationFormatter: (item) => formatPercent(item.share),
      titleFormatter: (item) => formatInteger(item.value) + " agendas"
    });

    window.DemorasCharts.renderHorizontalBarChart(refs.reportCentersChart, {
      items: topCenters,
      valueFormatter: formatDayMetric,
      emptyMessage: "No hay centros con agendas válidas para representar en el informe."
    });
  }

  function setReportKpisEmpty() {
    setKpiValue(refs.reportKpiPct0to2, "--");
    setKpiValue(refs.reportKpiPct0to3, "--");
    setKpiValue(refs.reportKpiPct0to6, "--");
    setKpiValue(refs.reportKpiPct7Plus, "--");
    setKpiValue(refs.reportKpiTotalValid, "--");
    setKpiValue(refs.reportKpiMean, "--");
    setKpiValue(refs.reportKpiMedian, "--");
    setKpiValue(refs.reportKpiMax, "--");
  }

  function buildReportSummaryMarkup(filteredRows, metrics, latestCutoff) {
    if (!filteredRows.length) {
      return "<p>No hay información disponible para generar un informe con la selección actual.</p>";
    }

    if (!metrics) {
      return "<p>El subconjunto filtrado no dispone de agendas con Accesibilidad numérica válida, por lo que no es posible elaborar indicadores ejecutivos en este momento.</p>";
    }

    const cutoffText = latestCutoff == null
      ? "No se dispone de una fecha de corte válida en el subconjunto filtrado."
      : "La fecha de corte más reciente disponible es " + formatLongDate(latestCutoff) + ".";
    const filteredBaseText = filteredRows.length !== metrics.totalValid
      ? " sobre un total de " + formatInteger(filteredRows.length) + " registros filtrados"
      : "";
    const assessment = buildExecutiveAssessment(metrics);

    const paragraphs = [
      "El análisis realizado sobre las agendas filtradas incorpora " +
        formatInteger(metrics.totalValid) +
        " agendas válidas para el cálculo ejecutivo" +
        filteredBaseText +
        ". " +
        cutoffText,
      "En términos de accesibilidad, el " +
        formatPercent(metrics.pct0to2) +
        " se sitúa entre 0 y 2 días, el " +
        formatPercent(metrics.pct0to3) +
        " entre 0 y 3 días y el " +
        formatPercent(metrics.pct0to6) +
        " entre 0 y 6 días. Las agendas con demora de 7 o más días representan el " +
        formatPercent(metrics.pct7Plus) +
        " del total analizado, con una demora media de " +
        formatDaysText(metrics.mean) +
        ", una mediana de " +
        formatDaysText(metrics.median) +
        " y una demora máxima de " +
        formatDaysText(metrics.max) +
        ".",
      assessment
    ];

    return paragraphs.map((paragraph) => "<p>" + escapeHtml(paragraph) + "</p>").join("");
  }

  function buildExecutiveAssessment(metrics) {
    if (metrics.pct7Plus >= 35 || metrics.mean >= 7) {
      return "En conjunto, la distribución observada refleja una presión relevante en la accesibilidad, con un peso significativo de agendas en demora prolongada.";
    }

    if (metrics.pct0to2 >= 60 && metrics.pct7Plus <= 15) {
      return "En conjunto, la situación muestra un comportamiento favorable, con predominio de agendas en tramos de demora corta y una presencia contenida de demoras prolongadas.";
    }

    return "En conjunto, la situación presenta un comportamiento intermedio, con margen de mejora en la reducción de las agendas que concentran mayores demoras.";
  }

  function getLatestDateSortValue(rows, column) {
    const timestamps = rows
      .map((row) => row.__dateSort && row.__dateSort[column])
      .filter((value) => Number.isFinite(value));

    if (!timestamps.length) {
      return null;
    }

    return Math.max.apply(null, timestamps);
  }

  function getReportZoneMeta(filteredRows) {
    if (state.filters.zona) {
      return {
        value: state.filters.zona,
        detail: ""
      };
    }

    const filteredZones = uniqueSorted(filteredRows.map((row) => normalizeText(row["Zona"])).filter(Boolean));
    const allZones = uniqueSorted(state.rows.map((row) => normalizeText(row["Zona"])).filter(Boolean));

    if (!filteredZones.length) {
      return {
        value: "Todas las zonas",
        detail: ""
      };
    }

    if (filteredZones.length === allZones.length) {
      return {
        value: "Todas las zonas",
        detail: ""
      };
    }

    return {
      value: "Todas las zonas",
      detail: "Zonas incluidas en el subconjunto filtrado: " + summarizeTextList(filteredZones, 4) + "."
    };
  }

  function summarizeTextList(values, visibleItems) {
    if (values.length <= visibleItems) {
      return values.join(", ");
    }

    const visible = values.slice(0, visibleItems).join(", ");
    return visible + " y " + formatInteger(values.length - visibleItems) + " más";
  }

  function renderCharts() {
    if (!state.rows.length) {
      renderChartsEmpty("Cargue un Excel para comenzar el analisis grafico.");
      return;
    }

    const validRows = getFilteredRows().filter((row) => row.__hasValidAcc);
    if (!validRows.length) {
      renderChartsEmpty("No hay agendas validas con los filtros actuales.");
      return;
    }

    const centers = aggregateByField(validRows, "Centro")
      .sort((a, b) => b.mean - a.mean || b.count - a.count)
      .slice(0, 10)
      .map((item) => ({
        label: item.label,
        value: item.mean,
        color: PALETTE.blue,
        note: "n=" + formatInteger(item.count)
      }));

    const zones = aggregateByField(validRows, "Zona")
      .sort((a, b) => b.mean - a.mean || b.count - a.count)
      .slice(0, 10)
      .map((item) => ({
        label: item.label,
        value: item.mean,
        color: PALETTE.teal,
        note: "n=" + formatInteger(item.count)
      }));

    const categories = aggregateByField(validRows, "Categoría")
      .sort((a, b) => b.mean - a.mean || b.count - a.count)
      .slice(0, 8)
      .map((item) => ({
        label: item.label,
        value: item.mean,
        color: PALETTE.green,
        note: "n=" + formatInteger(item.count)
      }));

    const visitTypes = aggregateByField(validRows, "Tipo visita")
      .sort((a, b) => b.mean - a.mean || b.count - a.count)
      .slice(0, 8)
      .map((item) => ({
        label: item.label,
        value: item.mean,
        color: PALETTE.amber,
        note: "n=" + formatInteger(item.count)
      }));

    const histogram = buildHistogramBins(validRows).map((item) => ({
      label: item.label,
      value: item.count,
      color: item.color
    }));

    const centerBands = aggregateCenterBands(validRows)
      .sort((a, b) => b.total - a.total || b.mean - a.mean)
      .slice(0, 8)
      .map((item) => ({
        label: item.label,
        total: item.total,
        segments: [
          { label: "0-2 días", value: item.band0to2, color: PALETTE.green },
          { label: "3 días", value: item.band3, color: PALETTE.teal },
          { label: "4-6 días", value: item.band4to6, color: PALETTE.amber },
          { label: "7 o más días", value: item.band7Plus, color: PALETTE.red }
        ]
      }));

    window.DemorasCharts.renderHorizontalBarChart(refs.centersChart, {
      items: centers,
      valueFormatter: formatDayMetric
    });

    window.DemorasCharts.renderHorizontalBarChart(refs.zonesChart, {
      items: zones,
      valueFormatter: formatDayMetric
    });

    window.DemorasCharts.renderHorizontalBarChart(refs.categoryChart, {
      items: categories,
      valueFormatter: formatDayMetric
    });

    window.DemorasCharts.renderHorizontalBarChart(refs.visitTypeChart, {
      items: visitTypes,
      valueFormatter: formatDayMetric
    });

    window.DemorasCharts.renderVerticalBarChart(refs.histogramChart, {
      items: histogram,
      valueFormatter: formatInteger,
      titleFormatter: (item) => formatInteger(item.value) + " agendas"
    });

    window.DemorasCharts.renderStackedBarChart(refs.centerBandsChart, {
      items: centerBands,
      legend: [
        { label: "0-2 días", color: PALETTE.green },
        { label: "3 días", color: PALETTE.teal },
        { label: "4-6 días", color: PALETTE.amber },
        { label: "7 o más días", color: PALETTE.red }
      ]
    });
  }

  function renderChartsEmpty(message) {
    window.DemorasCharts.setEmpty(refs.centersChart, message);
    window.DemorasCharts.setEmpty(refs.zonesChart, message);
    window.DemorasCharts.setEmpty(refs.categoryChart, message);
    window.DemorasCharts.setEmpty(refs.visitTypeChart, message);
    window.DemorasCharts.setEmpty(refs.histogramChart, message);
    window.DemorasCharts.setEmpty(refs.centerBandsChart, message);
  }

  function aggregateByField(rows, column) {
    const grouped = new Map();

    rows.forEach((row) => {
      const label = normalizeText(row[column]) || "Sin dato";
      if (!grouped.has(label)) {
        grouped.set(label, {
          label,
          values: []
        });
      }
      grouped.get(label).values.push(row.__acc);
    });

    return Array.from(grouped.values()).map((group) => ({
      label: group.label,
      count: group.values.length,
      mean: average(group.values),
      median: median(group.values),
      max: Math.max(...group.values)
    }));
  }

  function aggregateCenterBands(rows) {
    const grouped = new Map();

    rows.forEach((row) => {
      const label = normalizeText(row["Centro"]) || "Sin dato";
      if (!grouped.has(label)) {
        grouped.set(label, {
          label: label,
          total: 0,
          sum: 0,
          band0to2: 0,
          band3: 0,
          band4to6: 0,
          band7Plus: 0
        });
      }

      const group = grouped.get(label);
      group.total += 1;
      group.sum += row.__acc;

      if (row.__acc <= 2) {
        group.band0to2 += 1;
      } else if (row.__acc === 3) {
        group.band3 += 1;
      } else if (row.__acc <= 6) {
        group.band4to6 += 1;
      } else {
        group.band7Plus += 1;
      }
    });

    return Array.from(grouped.values()).map((group) => ({
      label: group.label,
      total: group.total,
      mean: group.sum / group.total,
      band0to2: group.band0to2,
      band3: group.band3,
      band4to6: group.band4to6,
      band7Plus: group.band7Plus
    }));
  }

  function buildHistogramBins(rows) {
    const definitions = [
      { label: "0", min: 0, max: 0, color: PALETTE.green },
      { label: "1", min: 1, max: 1, color: PALETTE.green },
      { label: "2", min: 2, max: 2, color: PALETTE.teal },
      { label: "3", min: 3, max: 3, color: PALETTE.teal },
      { label: "4", min: 4, max: 4, color: PALETTE.blue },
      { label: "5-6", min: 5, max: 6, color: PALETTE.amber },
      { label: "7-9", min: 7, max: 9, color: "#f08d4b" },
      { label: "10-14", min: 10, max: 14, color: PALETTE.red },
      { label: "15+", min: 15, max: Number.POSITIVE_INFINITY, color: "#9f4c69" }
    ];

    definitions.forEach((definition) => {
      definition.count = 0;
    });

    rows.forEach((row) => {
      const rounded = Math.max(0, Math.round(row.__acc));
      const bin = definitions.find((definition) => rounded >= definition.min && rounded <= definition.max);
      if (bin) {
        bin.count += 1;
      }
    });

    return definitions;
  }

  function buildExecutiveDelayBands(rows) {
    const total = rows.length || 1;
    const bands = [
      { label: "0-2 días", count: 0, color: PALETTE.green },
      { label: "3 días", count: 0, color: PALETTE.teal },
      { label: "4-6 días", count: 0, color: PALETTE.amber },
      { label: "7 o más días", count: 0, color: PALETTE.red }
    ];

    rows.forEach((row) => {
      if (row.__acc <= 2) {
        bands[0].count += 1;
      } else if (row.__acc === 3) {
        bands[1].count += 1;
      } else if (row.__acc <= 6) {
        bands[2].count += 1;
      } else {
        bands[3].count += 1;
      }
    });

    bands.forEach((band) => {
      band.share = (band.count / total) * 100;
    });

    return bands;
  }

  function calculateExecutiveMetrics(validRows) {
    if (!validRows.length) {
      return null;
    }

    const values = validRows.map((row) => row.__acc).slice().sort((a, b) => a - b);
    const totalValid = values.length;
    const count0to2 = values.filter((value) => value >= 0 && value <= 2).length;
    const count0to3 = values.filter((value) => value >= 0 && value <= 3).length;
    const count0to6 = values.filter((value) => value >= 0 && value <= 6).length;
    const count7Plus = values.filter((value) => value >= 7).length;

    return {
      totalValid: totalValid,
      pct0to2: (count0to2 / totalValid) * 100,
      pct0to3: (count0to3 / totalValid) * 100,
      pct0to6: (count0to6 / totalValid) * 100,
      pct7Plus: (count7Plus / totalValid) * 100,
      mean: average(values),
      median: median(values),
      max: Math.max(...values)
    };
  }

  function getFilteredRows() {
    return state.rows.filter((row) => {
      return FILTERS.every((filter) => {
        const selectedValue = state.filters[filter.key];
        return !selectedValue || row[filter.column] === selectedValue;
      });
    });
  }

  function getTableWorkingRows() {
    const filteredRows = getFilteredRows();
    const searchTerm = state.tableSearch.trim().toLowerCase();
    const searchedRows = searchTerm
      ? filteredRows.filter((row) =>
          REQUIRED_HEADERS.some((header) => normalizeText(row[header]).toLowerCase().includes(searchTerm))
        )
      : filteredRows;

    return searchedRows.slice().sort(compareRows);
  }

  function compareRows(rowA, rowB) {
    const column = state.tableSort.column;
    const direction = state.tableSort.direction === "asc" ? 1 : -1;
    const valueA = getComparableValue(rowA, column);
    const valueB = getComparableValue(rowB, column);
    const isBlankA = valueA == null || valueA === "";
    const isBlankB = valueB == null || valueB === "";

    if (isBlankA && isBlankB) {
      return 0;
    }
    if (isBlankA) {
      return 1;
    }
    if (isBlankB) {
      return -1;
    }

    if (typeof valueA === "number" && typeof valueB === "number") {
      return (valueA - valueB) * direction;
    }

    return String(valueA).localeCompare(String(valueB), "es", {
      sensitivity: "base",
      numeric: true
    }) * direction;
  }

  function getComparableValue(row, column) {
    if (column === "Accesibilidad") {
      return row.__hasValidAcc ? row.__acc : null;
    }
    if (DATE_COLUMNS.has(column)) {
      return row.__dateSort[column] == null ? null : row.__dateSort[column];
    }
    return normalizeText(row[column]);
  }

  function handleSortClick(event) {
    const button = event.target.closest("[data-sort-column]");
    if (!button) {
      return;
    }

    const column = button.getAttribute("data-sort-column");
    if (!column) {
      return;
    }

    if (state.tableSort.column === column) {
      state.tableSort.direction = state.tableSort.direction === "asc" ? "desc" : "asc";
    } else {
      state.tableSort.column = column;
      state.tableSort.direction = column === "Accesibilidad" ? "desc" : "asc";
    }

    state.currentPage = 1;
    renderAll();
  }

  function handleTableSearch(event) {
    state.tableSearch = event.target.value || "";
    state.currentPage = 1;
    renderAll();
  }

  function handlePageSizeChange(event) {
    state.pageSize = Number(event.target.value) || 20;
    state.currentPage = 1;
    renderAll();
  }

  function changePage(delta) {
    const totalPages = getTotalPages(getTableWorkingRows().length);
    state.currentPage = clamp(state.currentPage + delta, 1, totalPages);
    renderAll();
  }

  function resetFilters() {
    resetFilterValues();
    state.tableSearch = "";
    state.currentPage = 1;
    refs.tableSearchInput.value = "";
    renderAll();
  }

  function handlePrintReport() {
    setActiveTab("report");
    window.print();
  }

  function resetFilterValues() {
    FILTERS.forEach((filter) => {
      state.filters[filter.key] = "";
      if (refs[filter.key]) {
        refs[filter.key].value = "";
      }
    });
  }

  function exportFilteredTable() {
    const rows = getTableWorkingRows();
    if (!rows.length) {
      setStatus("info", "Sin filas para exportar", "No hay datos visibles en tabla para generar el CSV.", "Aplique otros filtros o revise la busqueda.");
      return;
    }

    const headerLine = REQUIRED_HEADERS.join(";");
    const dataLines = rows.map((row) =>
      REQUIRED_HEADERS.map((header) => escapeCsvValue(row[header])).join(";")
    );
    const csvContent = "\ufeff" + [headerLine].concat(dataLines).join("\r\n");
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = buildExportFileName("tabla_filtrada", "csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }

  function renderCellValue(header, row) {
    const value = row[header];
    if (!normalizeText(value)) {
      return '<span class="cell-chip">-</span>';
    }

    if (header === "Accesibilidad" && row.__hasValidAcc) {
      let className = "cell-chip";
      if (row.__acc <= 2) {
        className += " cell-chip--success";
      } else if (row.__acc <= 6) {
        className += " cell-chip--warning";
      } else {
        className += " cell-chip--danger";
      }

      return '<span class="' + className + '">' + escapeHtml(formatDayMetric(row.__acc)) + "</span>";
    }

    if (header === "Categoría" || header === "Tipo visita" || header === "Área" || header === "Zona") {
      return '<span class="cell-chip">' + escapeHtml(value) + "</span>";
    }

    return escapeHtml(value);
  }

  function setFeedback(type, message) {
    refs.uploadFeedback.className = "feedback-banner";
    if (type === "success") {
      refs.uploadFeedback.classList.add("feedback-banner--success");
    } else if (type === "error") {
      refs.uploadFeedback.classList.add("feedback-banner--error");
    } else {
      refs.uploadFeedback.classList.add("feedback-banner--info");
    }
    refs.uploadFeedback.textContent = message;
  }

  function setStatus(type, label, message, meta) {
    refs.statusIndicator.className = "status-indicator";
    if (type === "success") {
      refs.statusIndicator.classList.add("is-success");
    } else if (type === "error") {
      refs.statusIndicator.classList.add("is-error");
    } else {
      refs.statusIndicator.classList.add("is-info");
    }

    refs.statusLabel.textContent = label;
    refs.statusMessage.textContent = message;
    refs.statusMeta.textContent = meta;
  }

  function updateLoadProfile(status, fileName, sheetName, rowCount, validCount) {
    refs.loadedState.textContent = status;
    refs.loadedFileName.textContent = fileName || "Sin cargar";
    refs.loadedSheetName.textContent = sheetName || "-";
    refs.loadedRowsCount.textContent = formatInteger(rowCount || 0);
    refs.loadedValidCount.textContent = formatInteger(validCount || 0);
    refs.loadedAt.textContent = state.loadedAt || "-";
  }

  function clearData() {
    state.workbookName = "";
    state.sheetName = "";
    state.loadedAt = "";
    state.rows = [];
    state.tableSearch = "";
    state.currentPage = 1;
    state.tableSort = {
      column: "Accesibilidad",
      direction: "desc"
    };
    resetFilterValues();
    refs.tableSearchInput.value = "";
  }

  function setKpiValue(element, value) {
    if (!element) {
      return;
    }
    element.textContent = value;
  }

  function parseNumericValue(value) {
    if (typeof value === "number" && Number.isFinite(value)) {
      return value >= 0 ? value : NaN;
    }

    const textValue = normalizeText(value);
    if (!textValue) {
      return NaN;
    }

    const normalized = textValue.replace(/\s+/g, "").replace(",", ".");
    const parsed = Number(normalized);
    return Number.isFinite(parsed) && parsed >= 0 ? parsed : NaN;
  }

  function normalizeText(value) {
    return value == null ? "" : String(value).trim();
  }

  function formatDetectedHeaders(headers) {
    const visibleHeaders = (headers || []).map((header) => normalizeHeader(header)).filter(Boolean);
    return visibleHeaders.length ? visibleHeaders.join(" | ") : "No se detectaron cabeceras en la primera fila";
  }

  function uniqueSorted(values) {
    return Array.from(new Set(values)).sort((left, right) =>
      String(left).localeCompare(String(right), "es", {
        sensitivity: "base",
        numeric: true
      })
    );
  }

  function average(values) {
    if (!values.length) {
      return 0;
    }
    return values.reduce((sum, value) => sum + value, 0) / values.length;
  }

  function median(values) {
    if (!values.length) {
      return 0;
    }

    const sorted = values.slice().sort((a, b) => a - b);
    const middle = Math.floor(sorted.length / 2);
    if (sorted.length % 2 !== 0) {
      return sorted[middle];
    }
    return (sorted[middle - 1] + sorted[middle]) / 2;
  }

  function formatPercent(value) {
    return formatNumber(value, 1) + "%";
  }

  function formatDayMetric(value) {
    const decimals = Math.abs(value - Math.round(value)) < 0.01 ? 0 : 1;
    return formatNumber(value, decimals) + " d";
  }

  function formatDaysText(value) {
    const decimals = Math.abs(value - Math.round(value)) < 0.01 ? 0 : 1;
    return formatNumber(value, decimals) + " días";
  }

  function formatInteger(value) {
    return formatNumber(value, 0);
  }

  function formatNumber(value, decimals) {
    return new Intl.NumberFormat("es-ES", {
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals
    }).format(value);
  }

  function formatDate(date) {
    return new Intl.DateTimeFormat("es-ES", {
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
      timeZone: "UTC"
    }).format(date);
  }

  function formatLongDate(timestamp) {
    return new Intl.DateTimeFormat("es-ES", {
      day: "numeric",
      month: "long",
      year: "numeric",
      timeZone: "UTC"
    }).format(new Date(timestamp));
  }

  function formatDateTime(date) {
    return new Intl.DateTimeFormat("es-ES", {
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit"
    }).format(date);
  }

  function getTotalPages(totalRows) {
    return Math.max(1, Math.ceil(totalRows / state.pageSize));
  }

  function clamp(value, min, max) {
    return Math.min(Math.max(value, min), max);
  }

  function escapeHtml(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function escapeCsvValue(value) {
    const text = normalizeText(value);
    if (!text) {
      return "";
    }
    return '"' + text.replace(/"/g, '""') + '"';
  }

  function buildExportFileName(baseName, extension) {
    const stamp = new Date()
      .toISOString()
      .replace(/[:T]/g, "-")
      .replace(/\..+/, "");
    return baseName + "_" + stamp + "." + extension;
  }
})();

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

  const ZBS_COORDS = {
    "ABENOJAR": { lat: 38.879, lng: -4.357 },
    "AGUDO": { lat: 38.981, lng: -4.872 },
    "ALBALADEJO": { lat: 38.618, lng: -2.805 },
    "ALCAZAR DE SAN JUAN": { lat: 39.391, lng: -3.210 },
    "ALCAZAR I": { lat: 39.391, lng: -3.210 },
    "ALCAZAR II": { lat: 39.391, lng: -3.210 },
    "ALCAZAR DE SAN JUAN I": { lat: 39.391, lng: -3.210 },
    "ALCAZAR DE SAN JUAN II": { lat: 39.391, lng: -3.210 },
    "ALCOBA DE LOS MONTES": { lat: 39.258, lng: -4.477 },
    "ALCOLEA DE CALATRAVA": { lat: 38.986, lng: -4.114 },
    "ALDEA DEL REY": { lat: 38.738, lng: -3.839 },
    "ALMADEN": { lat: 38.775, lng: -4.831 },
    "ALMAGRO": { lat: 38.889, lng: -3.711 },
    "ALMEDINA": { lat: 38.624, lng: -2.954 },
    "ALMODOVAR DEL CAMPO": { lat: 38.709, lng: -4.179 },
    "ALMURADIEL": { lat: 38.513, lng: -3.497 },
    "ARGAMASILLA DE ALBA": { lat: 39.129, lng: -3.092 },
    "ARGAMASILLA DE CALATRAVA": { lat: 38.729, lng: -4.078 },
    "ARENAS DE SAN JUAN": { lat: 39.218, lng: -3.502 },
    "BOLAÑOS": { lat: 38.906, lng: -3.664 },
    "BOLAÑOS DE CALATRAVA": { lat: 38.906, lng: -3.664 },
    "CALZADA DE CALATRAVA": { lat: 38.704, lng: -3.776 },
    "CAMPO DE CRIPTANA": { lat: 39.404, lng: -3.124 },
    "CARRION DE CALATRAVA": { lat: 39.018, lng: -3.816 },
    "CHILLON": { lat: 38.794, lng: -4.866 },
    "CIUDAD REAL": { lat: 38.986, lng: -3.929 },
    "CIUDAD REAL 1": { lat: 38.986, lng: -3.929 },
    "CIUDAD REAL 2": { lat: 38.986, lng: -3.929 },
    "CIUDAD REAL 3": { lat: 38.986, lng: -3.929 },
    "CIUDAD REAL 4": { lat: 38.986, lng: -3.929 },
    "CIUDAD REAL I": { lat: 38.986, lng: -3.929 },
    "CIUDAD REAL II": { lat: 38.986, lng: -3.929 },
    "CIUDAD REAL III": { lat: 38.986, lng: -3.929 },
    "CIUDAD REAL IV": { lat: 38.986, lng: -3.929 },
    "CORRAL DE CALATRAVA": { lat: 38.858, lng: -4.080 },
    "DAIMIEL": { lat: 39.070, lng: -3.615 },
    "DAIMIEL 2": { lat: 39.070, lng: -3.615 },
    "DAIMIEL II": { lat: 39.070, lng: -3.615 },
    "HERENCIA": { lat: 39.366, lng: -3.356 },
    "HORCAJO DE LOS MONTES": { lat: 39.325, lng: -4.648 },
    "LA SOLANA": { lat: 38.945, lng: -3.238 },
    "MALAGON": { lat: 39.166, lng: -3.854 },
    "MANZANARES": { lat: 38.999, lng: -3.371 },
    "MANZANARES 1": { lat: 38.999, lng: -3.371 },
    "MANZANARES 2": { lat: 38.999, lng: -3.371 },
    "MANZANARES I": { lat: 38.999, lng: -3.371 },
    "MANZANARES II": { lat: 38.999, lng: -3.371 },
    "MEMBRILLA": { lat: 38.972, lng: -3.343 },
    "MIGUELTURRA": { lat: 38.964, lng: -3.891 },
    "MORAL DE CALATRAVA": { lat: 38.829, lng: -3.579 },
    "PEDRO MUÑOZ": { lat: 39.402, lng: -2.946 },
    "PIEDRABUENA": { lat: 39.035, lng: -4.175 },
    "PORZUNA": { lat: 39.146, lng: -4.154 },
    "POZUELO DE CALATRAVA": { lat: 38.912, lng: -3.837 },
    "PUERTOLLANO": { lat: 38.687, lng: -4.112 },
    "PUERTOLLANO I": { lat: 38.687, lng: -4.112 },
    "PUERTOLLANO II": { lat: 38.687, lng: -4.112 },
    "PUERTOLLANO III": { lat: 38.687, lng: -4.112 },
    "PUERTOLLANO IV": { lat: 38.687, lng: -4.112 },
    "RETUERTA DEL BULLAQUE": { lat: 39.463, lng: -4.409 },
    "SANTA CRUZ DE MUDELA": { lat: 38.642, lng: -3.467 },
    "SOCUELLAMOS": { lat: 39.285, lng: -2.793 },
    "TOMELLOSO": { lat: 39.157, lng: -3.021 },
    "TOMELLOSO 1": { lat: 39.157, lng: -3.021 },
    "TOMELLOSO 2": { lat: 39.157, lng: -3.021 },
    "TOMELLOSO I": { lat: 39.157, lng: -3.021 },
    "TOMELLOSO II": { lat: 39.157, lng: -3.021 },
    "TORRALBA DE CALATRAVA": { lat: 39.018, lng: -3.750 },
    "VALDEPEÑAS": { lat: 38.762, lng: -3.384 },
    "VALDEPEÑAS I": { lat: 38.762, lng: -3.384 },
    "VALDEPEÑAS II": { lat: 38.762, lng: -3.384 },
    "VILLAHERMOSA": { lat: 38.750, lng: -2.871 },
    "VILLANUEVA DE LA FUENTE": { lat: 38.694, lng: -2.697 },
    "VILLANUEVA DE LOS INFANTES": { lat: 38.734, lng: -3.013 },
    "VILLARRUBIA DE LOS OJOS": { lat: 39.219, lng: -3.608 },
    "VISO DEL MARQUES": { lat: 38.522, lng: -3.563 }
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
    filtersCompactTicking: false,
    currentPage: 1,
    pageSize: 20,
    zoneMap: null,
    zoneMapLayer: null,
    zoneMapTileLayer: null
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
    zoneGeoMap: document.getElementById("zoneGeoMap"),
    zoneMapEmptyState: document.getElementById("zoneMapEmptyState"),
    zoneMapMissing: document.getElementById("zoneMapMissing"),
    zoneMapMissingList: document.getElementById("zoneMapMissingList"),
    printReportButton: document.getElementById("printReportButton"),
    reportPrintNotice: document.getElementById("reportPrintNotice"),
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
    reportCategoryEmptyState: document.getElementById("reportCategoryEmptyState"),
    reportCategoryTableWrapper: document.getElementById("reportCategoryTableWrapper"),
    reportCategoryTableBody: document.getElementById("reportCategoryTableBody"),
    reportCategoryVisitEmptyState: document.getElementById("reportCategoryVisitEmptyState"),
    reportCategoryVisitTableWrapper: document.getElementById("reportCategoryVisitTableWrapper"),
    reportCategoryVisitTableBody: document.getElementById("reportCategoryVisitTableBody"),
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
    refs.printReportButton.addEventListener("click", handlePrintReport);
    refs.tableSearchInput.addEventListener("input", handleTableSearch);
    refs.pageSizeSelect.addEventListener("change", handlePageSizeChange);
    refs.prevPageButton.addEventListener("click", () => changePage(-1));
    refs.nextPageButton.addEventListener("click", () => changePage(1));
    refs.dataTableHead.addEventListener("click", handleSortClick);
    window.addEventListener("scroll", handleFiltersCompactScroll, { passive: true });
    window.addEventListener("resize", handleFiltersCompactScroll, { passive: true });

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

    updateFiltersCompactMode();
  }

  function handleFiltersCompactScroll() {
    if (!refs.filtersPanel || state.filtersCompactTicking) {
      return;
    }

    state.filtersCompactTicking = true;
    window.requestAnimationFrame(updateFiltersCompactMode);
  }

  function updateFiltersCompactMode() {
    state.filtersCompactTicking = false;
    if (!refs.filtersPanel) {
      return;
    }

    const canCompact = window.matchMedia("(min-width: 921px)").matches;
    if (!canCompact) {
      refs.filtersPanel.classList.remove("filters-panel--compact");
      return;
    }

    const scrollY = window.scrollY || document.documentElement.scrollTop || 0;
    if (scrollY > 140) {
      refs.filtersPanel.classList.add("filters-panel--compact");
    } else if (scrollY < 80) {
      refs.filtersPanel.classList.remove("filters-panel--compact");
    }
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
    if (tabName === "charts") {
      scheduleMapResize();
    }
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
      setReportDetailTablesEmpty("Cargue un Excel para generar el detalle analítico.");
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
      setReportDetailTablesEmpty("No hay agendas válidas con los filtros actuales.");
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

    renderReportDetailTables(validRows);
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

  function renderReportDetailTables(validRows) {
    const categoryRows = buildCategorySummaries(validRows);
    const categoryVisitRows = buildCategoryVisitSummaries(validRows);

    renderReportTable(
      refs.reportCategoryEmptyState,
      refs.reportCategoryTableWrapper,
      refs.reportCategoryTableBody,
      categoryRows,
      (item) =>
        "<tr>" +
        "<th scope=\"row\">" + escapeHtml(item.category) + "</th>" +
        renderMetricCells(item) +
        "</tr>"
    );

    renderReportTable(
      refs.reportCategoryVisitEmptyState,
      refs.reportCategoryVisitTableWrapper,
      refs.reportCategoryVisitTableBody,
      categoryVisitRows,
      (item) =>
        "<tr>" +
        "<th scope=\"row\">" + escapeHtml(item.category) + "</th>" +
        "<td>" + escapeHtml(item.visitType) + "</td>" +
        renderMetricCells(item) +
        "</tr>"
    );
  }

  function renderReportTable(emptyElement, wrapperElement, bodyElement, rows, rowRenderer) {
    if (!emptyElement || !wrapperElement || !bodyElement) {
      return;
    }

    if (!rows.length) {
      emptyElement.classList.remove("is-hidden");
      wrapperElement.classList.add("is-hidden");
      bodyElement.innerHTML = "";
      return;
    }

    emptyElement.classList.add("is-hidden");
    wrapperElement.classList.remove("is-hidden");
    bodyElement.innerHTML = rows.map(rowRenderer).join("");
  }

  function renderMetricCells(item) {
    return (
      '<td class="report-table__numeric">' + formatInteger(item.totalValid) + "</td>" +
      '<td class="report-table__numeric">' + formatPercent(item.pct0to2) + "</td>" +
      '<td class="report-table__numeric">' + formatPercent(item.pct0to3) + "</td>" +
      '<td class="report-table__numeric">' + formatPercent(item.pct0to6) + "</td>" +
      '<td class="report-table__numeric report-table__signal">' + formatPercent(item.pct7Plus) + "</td>" +
      '<td class="report-table__numeric">' + formatDayMetric(item.mean) + "</td>" +
      '<td class="report-table__numeric">' + formatDayMetric(item.median) + "</td>" +
      '<td class="report-table__numeric">' + formatDayMetric(item.max) + "</td>"
    );
  }

  function setReportDetailTablesEmpty(message) {
    [
      [refs.reportCategoryEmptyState, refs.reportCategoryTableWrapper, refs.reportCategoryTableBody],
      [refs.reportCategoryVisitEmptyState, refs.reportCategoryVisitTableWrapper, refs.reportCategoryVisitTableBody]
    ].forEach(([emptyElement, wrapperElement, bodyElement]) => {
      if (!emptyElement || !wrapperElement || !bodyElement) {
        return;
      }
      emptyElement.textContent = message;
      emptyElement.classList.remove("is-hidden");
      wrapperElement.classList.add("is-hidden");
      bodyElement.innerHTML = "";
    });
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
    const assessment = [buildExecutiveAssessment(metrics), buildReportDetailAssessment(filteredRows)]
      .filter(Boolean)
      .join(" ");

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

  function buildReportDetailAssessment(filteredRows) {
    const validRows = filteredRows.filter((row) => row.__hasValidAcc);
    const categoryRows = buildCategorySummaries(validRows);
    const categoryVisitRows = buildCategoryVisitSummaries(validRows);
    const sentences = [];

    if (categoryRows.length > 1 && categoryRows[0].totalValid >= 3) {
      sentences.push(
        "Por categoría, " +
          categoryRows[0].category +
          " concentra el comportamiento menos favorable del subconjunto, con una demora media de " +
          formatDaysText(categoryRows[0].mean) +
          " y un " +
          formatPercent(categoryRows[0].pct7Plus) +
          " de agendas con 7 o más días."
      );
    }

    if (categoryVisitRows.length > 1 && categoryVisitRows[0].totalValid >= 3) {
      sentences.push(
        "En el cruce categoría y tipo de visita destaca " +
          categoryVisitRows[0].category +
          " - " +
          categoryVisitRows[0].visitType +
          ", con " +
          formatInteger(categoryVisitRows[0].totalValid) +
          " agendas válidas y un " +
          formatPercent(categoryVisitRows[0].pct7Plus) +
          " en 7 o más días."
      );
    }

    return sentences.join(" ");
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

    renderGeographicZoneMap(validRows);

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
    setMapEmpty(message);
  }

  function renderGeographicZoneMap(validRows) {
    if (!refs.zoneGeoMap) {
      return;
    }

    if (!window.L) {
      setMapEmpty("No se ha podido cargar Leaflet. Revise la conexión o use una copia local de la librería.");
      return;
    }

    const zoneRows = aggregateByZoneForMap(validRows);
    const zonesWithCoords = zoneRows.filter((zone) => zone.coordinates);
    const missingZones = zoneRows.filter((zone) => !zone.coordinates).map((zone) => zone.label);

    renderMissingZoneDiagnostic(missingZones);

    if (!zonesWithCoords.length) {
      setMapEmpty("No hay zonas con coordenadas configuradas para el subconjunto filtrado.");
      return;
    }

    ensureZoneMap();
    clearMapEmpty();
    state.zoneMapLayer.clearLayers();

    const maxTotal = Math.max(...zonesWithCoords.map((zone) => zone.totalValid));
    const minTotal = Math.min(...zonesWithCoords.map((zone) => zone.totalValid));
    const coordinateGroups = groupZonesByCoordinate(zonesWithCoords);
    const bounds = [];

    coordinateGroups.forEach((group) => {
      group.forEach((zone, index) => {
        const point = applyCoordinateOffset(zone.coordinates, index, group.length);
        const marker = window.L.circleMarker([point.lat, point.lng], {
          radius: scaleMapRadius(zone.totalValid, minTotal, maxTotal),
          color: getMapColor(zone.mean),
          weight: 2,
          opacity: 0.95,
          fillColor: getMapColor(zone.mean),
          fillOpacity: 0.68
        });

        marker.bindPopup(buildZoneMapPopup(zone), {
          maxWidth: 320,
          className: "zone-map-popup"
        });
        marker.bindTooltip(zone.label + " · " + formatDayMetric(zone.mean), {
          sticky: true,
          direction: "top"
        });
        marker.addTo(state.zoneMapLayer);
        bounds.push([point.lat, point.lng]);
      });
    });

    if (bounds.length === 1) {
      state.zoneMap.setView(bounds[0], 11);
    } else {
      state.zoneMap.fitBounds(bounds, {
        padding: [28, 28],
        maxZoom: 11
      });
    }
    scheduleMapResize();
  }

  function ensureZoneMap() {
    if (state.zoneMap || !refs.zoneGeoMap || !window.L) {
      return;
    }

    state.zoneMap = window.L.map(refs.zoneGeoMap, {
      zoomControl: true,
      scrollWheelZoom: true
    }).setView([38.99, -3.93], 9);

    state.zoneMapTileLayer = window.L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
      attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>',
      maxZoom: 18
    }).addTo(state.zoneMap);

    state.zoneMapLayer = window.L.layerGroup().addTo(state.zoneMap);
  }

  function aggregateByZoneForMap(rows) {
    const grouped = new Map();

    rows.forEach((row) => {
      const label = normalizeText(row["Zona"]) || "Sin zona";
      const key = normalizeZoneName(label);
      if (!grouped.has(key)) {
        grouped.set(key, {
          label: label,
          rows: []
        });
      }
      grouped.get(key).rows.push(row);
    });

    return Array.from(grouped.values()).map((group) => {
      const metrics = calculateExecutiveMetrics(group.rows);
      return Object.assign({}, group, metrics, {
        coordinates: getZoneCoordinates(group.label)
      });
    }).sort((left, right) => right.mean - left.mean || right.totalValid - left.totalValid);
  }

  function getZoneCoordinates(zoneName) {
    const normalizedZoneName = normalizeZoneName(zoneName);
    if (ZBS_COORDS[normalizedZoneName]) {
      return ZBS_COORDS[normalizedZoneName];
    }

    const matchingKey = Object.keys(ZBS_COORDS).find((key) => normalizeZoneName(key) === normalizedZoneName);
    return matchingKey ? ZBS_COORDS[matchingKey] : null;
  }

  function normalizeZoneName(value) {
    return normalizeText(value)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[ºª]/g, "")
      .replace(/[-_/.,;:()]+/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .toUpperCase();
  }

  function groupZonesByCoordinate(zones) {
    const grouped = new Map();
    zones.forEach((zone) => {
      const key = zone.coordinates.lat.toFixed(4) + "," + zone.coordinates.lng.toFixed(4);
      if (!grouped.has(key)) {
        grouped.set(key, []);
      }
      grouped.get(key).push(zone);
    });
    return Array.from(grouped.values());
  }

  function applyCoordinateOffset(coordinates, index, total) {
    if (total <= 1) {
      return coordinates;
    }

    const angle = (Math.PI * 2 * index) / total;
    const distance = 0.010 + Math.min(total, 6) * 0.0015;
    const latOffset = Math.sin(angle) * distance;
    const lngOffset = (Math.cos(angle) * distance) / Math.max(0.45, Math.cos((coordinates.lat * Math.PI) / 180));
    return {
      lat: coordinates.lat + latOffset,
      lng: coordinates.lng + lngOffset
    };
  }

  function scaleMapRadius(value, minValue, maxValue) {
    const minRadius = 8;
    const maxRadius = 34;
    if (maxValue === minValue) {
      return Math.round((minRadius + maxRadius) / 2);
    }
    return minRadius + ((value - minValue) / (maxValue - minValue)) * (maxRadius - minRadius);
  }

  function getMapColor(mean) {
    if (mean <= 2) {
      return PALETTE.green;
    }
    if (mean <= 3) {
      return PALETTE.amber;
    }
    if (mean <= 6) {
      return "#e9823a";
    }
    return PALETTE.red;
  }

  function buildZoneMapPopup(zone) {
    return (
      '<div class="map-popup">' +
      "<h4>" + escapeHtml(zone.label) + "</h4>" +
      '<div class="map-popup__grid">' +
      "<span>Agendas válidas</span><span>" + formatInteger(zone.totalValid) + "</span>" +
      "<span>% 0-2 días</span><span>" + formatPercent(zone.pct0to2) + "</span>" +
      "<span>% 0-3 días</span><span>" + formatPercent(zone.pct0to3) + "</span>" +
      "<span>% 0-6 días</span><span>" + formatPercent(zone.pct0to6) + "</span>" +
      "<span>% 7 o más días</span><span>" + formatPercent(zone.pct7Plus) + "</span>" +
      "<span>Demora media</span><span>" + formatDayMetric(zone.mean) + "</span>" +
      "<span>Mediana</span><span>" + formatDayMetric(zone.median) + "</span>" +
      "<span>Máxima</span><span>" + formatDayMetric(zone.max) + "</span>" +
      "</div>" +
      "</div>"
    );
  }

  function renderMissingZoneDiagnostic(missingZones) {
    if (!refs.zoneMapMissing || !refs.zoneMapMissingList) {
      return;
    }

    if (!missingZones.length) {
      refs.zoneMapMissing.classList.add("is-hidden");
      refs.zoneMapMissingList.textContent = "";
      return;
    }

    refs.zoneMapMissing.classList.remove("is-hidden");
    refs.zoneMapMissingList.textContent = summarizeTextList(uniqueSorted(missingZones), 8);
    console.info("Zonas sin coordenadas configuradas para el mapa:", missingZones);
  }

  function setMapEmpty(message) {
    if (refs.zoneMapEmptyState) {
      refs.zoneMapEmptyState.textContent = message;
      refs.zoneMapEmptyState.classList.remove("is-hidden");
    }
    if (state.zoneMapLayer) {
      state.zoneMapLayer.clearLayers();
    }
  }

  function clearMapEmpty() {
    if (refs.zoneMapEmptyState) {
      refs.zoneMapEmptyState.classList.add("is-hidden");
    }
  }

  function scheduleMapResize() {
    if (!state.zoneMap) {
      return;
    }
    window.setTimeout(() => {
      state.zoneMap.invalidateSize();
    }, 80);
  }

  function buildCategorySummaries(validRows) {
    return summarizeGroupedRows(validRows, [
      {
        key: "category",
        accessor: getRecordCategory,
        fallback: "Sin categoría"
      }
    ]).sort(compareByDelayPressure);
  }

  function buildCategoryVisitSummaries(validRows) {
    return summarizeGroupedRows(validRows, [
      {
        key: "category",
        accessor: getRecordCategory,
        fallback: "Sin categoría"
      },
      {
        key: "visitType",
        accessor: getRecordVisitType,
        fallback: "Sin tipo de visita"
      }
    ]).sort((left, right) => {
      const pressureOrder = compareByDelayPressure(left, right);
      return pressureOrder;
    });
  }

  function buildWorstProfessionalsByCategory(validRows) {
    const grouped = new Map();

    validRows.forEach((row) => {
      const category = getRecordCategory(row);
      const professional = normalizeText(row["PROFESIONAL"]) || "Sin profesional informado";
      const center = normalizeText(row["Centro"]) || "Sin centro";
      const key = category + "\u001f" + professional;

      if (!grouped.has(key)) {
        grouped.set(key, {
          category,
          professional,
          centers: new Set(),
          rows: []
        });
      }

      const group = grouped.get(key);
      group.centers.add(center);
      group.rows.push(row);
    });

    const professionalRows = Array.from(grouped.values()).map((group) => {
      const metrics = calculateExecutiveMetrics(group.rows);
      return Object.assign({}, group, metrics, {
        center: summarizeTextList(uniqueSorted(Array.from(group.centers)), 2)
      });
    }).sort(compareProfessionalDelayPressure);

    const byCategory = new Map();
    professionalRows.forEach((row) => {
      if (!byCategory.has(row.category)) {
        byCategory.set(row.category, []);
      }
      byCategory.get(row.category).push(row);
    });

    return Array.from(byCategory.entries()).map(([category, professionals]) => ({
      category,
      professionals: professionals.slice(0, 5)
    })).sort((left, right) => {
      const worstLeft = left.professionals[0];
      const worstRight = right.professionals[0];
      if (!worstLeft || !worstRight) {
        return !worstLeft ? 1 : -1;
      }
      return compareProfessionalDelayPressure(worstLeft, worstRight) ||
        left.category.localeCompare(right.category, "es", { sensitivity: "base", numeric: true });
    });
  }

  function buildProfessionalsWithDelayOver7(validRows) {
    const grouped = new Map();

    validRows.forEach((row) => {
      const category = getRecordCategory(row);
      const professional = normalizeText(row["PROFESIONAL"]);
      const cias = normalizeText(row["CIAS"]);

      if (!professional && !cias) {
        return;
      }

      const key = [category, professional || "Sin profesional", cias || "Sin CIAS"].join("\u001f");
      if (!grouped.has(key)) {
        grouped.set(key, {
          category,
          professional: professional || "CIAS " + cias,
          centers: new Set(),
          zones: new Set(),
          rows: []
        });
      }

      const group = grouped.get(key);
      group.centers.add(normalizeText(row["Centro"]) || "Sin centro");
      group.zones.add(normalizeText(row["Zona"]) || "Sin zona");
      group.rows.push(row);
    });

    return Array.from(grouped.values()).map((group) => {
      const metrics = calculateExecutiveMetrics(group.rows);
      return {
        category: group.category,
        professional: group.professional,
        center: summarizeTextList(uniqueSorted(Array.from(group.centers)), 2),
        zone: summarizeTextList(uniqueSorted(Array.from(group.zones)), 2),
        totalValid: metrics.totalValid,
        mean: metrics.mean,
        median: metrics.median,
        max: metrics.max,
        pct7Plus: metrics.pct7Plus
      };
    }).filter((row) => row.totalValid > 0 && row.mean >= 7)
      .sort((left, right) => (
        right.mean - left.mean ||
        right.max - left.max ||
        right.pct7Plus - left.pct7Plus ||
        right.totalValid - left.totalValid ||
        left.professional.localeCompare(right.professional, "es", { sensitivity: "base", numeric: true })
      ));
  }

  function compareProfessionalDelayPressure(left, right) {
    return (
      right.mean - left.mean ||
      right.pct7Plus - left.pct7Plus ||
      right.max - left.max ||
      right.totalValid - left.totalValid ||
      left.professional.localeCompare(right.professional, "es", {
        sensitivity: "base",
        numeric: true
      })
    );
  }

  function getRecordCategory(row) {
    return (
      normalizeText(row["Categoría"]) ||
      normalizeText(row["Categoria"]) ||
      normalizeText(row.categoria) ||
      normalizeText(row.category) ||
      "Sin categoría"
    );
  }

  function getRecordVisitType(row) {
    return (
      normalizeText(row["Tipo visita"]) ||
      normalizeText(row["Tipo Visita"]) ||
      normalizeText(row.tipoVisita) ||
      normalizeText(row.visitType) ||
      "Sin tipo de visita"
    );
  }

  function summarizeGroupedRows(validRows, descriptors) {
    const grouped = new Map();

    validRows.forEach((row) => {
      const labels = descriptors.map((descriptor) => {
        const value = descriptor.accessor ? descriptor.accessor(row) : row[descriptor.column];
        return normalizeText(value) || descriptor.fallback;
      });
      const groupKey = labels.join("\u001f");

      if (!grouped.has(groupKey)) {
        const group = { rows: [] };
        descriptors.forEach((descriptor, index) => {
          group[descriptor.key] = labels[index];
        });
        grouped.set(groupKey, group);
      }

      grouped.get(groupKey).rows.push(row);
    });

    return Array.from(grouped.values()).map((group) => {
      const metrics = calculateExecutiveMetrics(group.rows);
      return Object.assign({}, group, metrics);
    });
  }

  function compareByDelayPressure(left, right) {
    return (
      right.pct7Plus - left.pct7Plus ||
      right.mean - left.mean ||
      right.totalValid - left.totalValid ||
      String(left.category || left.label || "").localeCompare(String(right.category || right.label || ""), "es", {
        sensitivity: "base",
        numeric: true
      }) ||
      String(left.visitType || "").localeCompare(String(right.visitType || ""), "es", {
        sensitivity: "base",
        numeric: true
      })
    );
  }

  function aggregateByField(rows, column) {
    const grouped = new Map();

    rows.forEach((row) => {
      const label = normalizeText(row[column]) || "Sin dato";
      if (!grouped.has(label)) {
        grouped.set(label, {
          label,
          values: [],
          rows: []
        });
      }
      grouped.get(label).values.push(row.__acc);
      grouped.get(label).rows.push(row);
    });

    return Array.from(grouped.values()).map((group) => {
      const metrics = calculateExecutiveMetrics(group.rows);
      return {
        label: group.label,
        count: metrics.totalValid,
        totalValid: metrics.totalValid,
        pct0to2: metrics.pct0to2,
        pct0to3: metrics.pct0to3,
        pct0to6: metrics.pct0to6,
        pct7Plus: metrics.pct7Plus,
        mean: metrics.mean,
        median: metrics.median,
        max: metrics.max
      };
    });
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

  async function handlePrintReport(event) {
    if (event) {
      event.preventDefault();
      event.stopPropagation();
    }

    if (!state.rows.length || refs.printReportButton.disabled) {
      return;
    }

    hidePrintNotice();
    setActiveTab("report");

    if (!window.jspdf || typeof window.jspdf.jsPDF !== "function") {
      showPrintNotice("No se ha podido cargar el generador local de PDF.");
      return;
    }

    const originalText = refs.printReportButton.textContent;
    refs.printReportButton.disabled = true;
    refs.printReportButton.textContent = "Generando PDF...";

    try {
      const logoSources = await resolvePdfLogoSources();
      const filename = buildReportPdfFileName();
      const doc = await generateNativeReportPdf(logoSources);

      if (window.AndroidPdf && typeof window.AndroidPdf.saveReportPdf === "function") {
        const dataUri = doc.output("datauristring");
        const base64 = String(dataUri || "").split(",")[1] || "";
        if (!base64) {
          throw new Error("No se ha podido generar el contenido PDF.");
        }
        window.AndroidPdf.saveReportPdf(base64, filename);
        showPrintNotice("PDF generado localmente. Revise la carpeta Descargas del dispositivo.");
      } else {
        const pdfBlob = doc.output("blob");
        downloadBlob(pdfBlob, filename);
      }
    } catch (error) {
      console.error("Error al generar el PDF del informe:", error);
      showPrintNotice("No se ha podido generar el PDF local. Revise el archivo cargado e intentelo de nuevo.");
    } finally {
      refs.printReportButton.textContent = originalText;
      refs.printReportButton.disabled = !state.rows.length;
    }
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    link.rel = "noopener";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  async function resolvePdfLogoSources() {
    const configuredLogos = Object.assign({
      jccm: "assets/img/sescam.jpg",
      gai: "assets/img/gai-cr.jpg"
    }, window.DEMORAS_PDF_LOGOS || {});

    const [jccm, gai] = await Promise.all([
      configuredLogos.jccm ? loadLocalImageIfAvailable(configuredLogos.jccm) : Promise.resolve(""),
      configuredLogos.gai ? loadLocalImageIfAvailable(configuredLogos.gai) : Promise.resolve("")
    ]);

    return { jccm, gai };
  }

  function loadLocalImageIfAvailable(src) {
    return new Promise((resolve) => {
      const image = new Image();
      const finish = (value) => {
        image.onload = null;
        image.onerror = null;
        resolve(value);
      };

      image.onload = () => {
        try {
          const canvas = document.createElement("canvas");
          canvas.width = image.naturalWidth || image.width;
          canvas.height = image.naturalHeight || image.height;
          const context = canvas.getContext("2d");
          context.drawImage(image, 0, 0);
          finish({
            dataUrl: canvas.toDataURL("image/png"),
            width: canvas.width,
            height: canvas.height
          });
        } catch (error) {
          finish("");
        }
      };
      image.onerror = () => finish("");
      image.src = src;
    });
  }

  async function generateNativeReportPdf(logoSources) {
    const jsPDF = window.jspdf && window.jspdf.jsPDF;
    attachAutoTablePlugin(jsPDF);

    const doc = new jsPDF({
      orientation: "portrait",
      unit: "mm",
      format: "a4",
      compress: true
    });

    if (typeof doc.autoTable !== "function") {
      throw new Error("No se ha cargado jsPDF-AutoTable.");
    }

    const filteredRows = getFilteredRows();
    const validRows = filteredRows.filter((row) => row.__hasValidAcc);
    const metrics = calculateExecutiveMetrics(validRows);
    const latestCutoff = getLatestDateSortValue(filteredRows, "Fecha Corte");
    const zoneMeta = getReportZoneMeta(filteredRows);
    const context = {
      filteredRows,
      validRows,
      metrics,
      latestCutoff,
      zoneMeta,
      logoSources: logoSources || {}
    };

    doc.setProperties({
      title: "Informe sobre accesibilidad",
      subject: "Programa de analisis de las demoras en Atencion Primaria",
      author: "Gerencia de Atencion Integrada de Ciudad Real",
      creator: "Programa de analisis de las demoras en Atencion Primaria"
    });

    drawNativeCoverPage(doc, context);

    if (metrics) {
      drawNativeDetailPage(doc, context);
      drawNativeProfessionalsPage(doc, context);
      drawNativeRankingsPage(doc, context);
      drawNativeProfessionalsOver7Page(doc, context);
      drawNativeMethodologyPage(doc);
    }

    addNativePdfFooters(doc);
    return doc;
  }

  function drawNativeCoverPage(doc, context) {
    const metrics = context.metrics;
    drawNativeHeader(doc, context.logoSources);
    drawNativeTitleBlock(doc, context);

    if (!state.rows.length) {
      drawNativeEmptyState(doc, 86, "Cargue un archivo Excel para generar el informe PDF.");
      return;
    }

    if (!metrics) {
      drawNativeContextCards(doc, context, 82);
      drawNativeEmptyState(doc, 126, "No hay agendas validas con Accesibilidad numerica en el subconjunto filtrado.");
      return;
    }

    drawNativeContextCards(doc, context, 82);
    const summaryY = drawNativeSummary(doc, context, 132);
    drawNativeKpiGrid(doc, metrics, summaryY + 8);
  }

  function drawNativeHeader(doc, logoSources) {
    const margin = 12;
    const pageWidth = 210;
    const logoWidth = 56;
    const logoHeight = 16;

    doc.setDrawColor(220, 232, 242);
    doc.setFillColor(255, 255, 255);
    doc.roundedRect(margin, 10, logoWidth, logoHeight, 3, 3, "FD");
    doc.roundedRect(pageWidth - margin - logoWidth, 10, logoWidth, logoHeight, 3, 3, "FD");

    if (isPdfLogoReady(logoSources && logoSources.jccm)) {
      drawPdfImageContain(doc, logoSources.jccm, margin + 2, 12, logoWidth - 4, logoHeight - 4);
    } else {
      doc.setFont("helvetica", "bold");
      doc.setFontSize(9);
      doc.setTextColor(31, 59, 90);
      doc.text("Castilla-La Mancha", margin + logoWidth / 2, 20.8, { align: "center" });
    }

    if (isPdfLogoReady(logoSources && logoSources.gai)) {
      drawPdfImageContain(doc, logoSources.gai, pageWidth - margin - logoWidth + 2, 12, logoWidth - 4, logoHeight - 4);
    } else {
      doc.setFont("helvetica", "bold");
      doc.setFontSize(9);
      doc.setTextColor(31, 59, 90);
      doc.text("GAI Ciudad Real", pageWidth - margin - logoWidth / 2, 20.8, { align: "center" });
    }

    doc.setFont("helvetica", "bold");
    doc.setFontSize(11);
    doc.setTextColor(18, 58, 90);
    doc.text("Servicio de Salud de Castilla-La Mancha", pageWidth / 2, 17, { align: "center" });
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8);
    doc.setTextColor(82, 113, 141);
    doc.text("Gerencia de Atencion Integrada de Ciudad Real", pageWidth / 2, 23, { align: "center" });

    doc.setDrawColor(216, 228, 239);
    doc.line(margin, 34, pageWidth - margin, 34);
  }

  function isPdfLogoReady(logo) {
    return Boolean(logo && /^data:image\//.test(logo.dataUrl || ""));
  }

  function drawPdfImageContain(doc, logo, x, y, maxWidth, maxHeight) {
    const sourceWidth = Math.max(1, Number(logo.width) || maxWidth);
    const sourceHeight = Math.max(1, Number(logo.height) || maxHeight);
    const scale = Math.min(maxWidth / sourceWidth, maxHeight / sourceHeight);
    const width = sourceWidth * scale;
    const height = sourceHeight * scale;
    const offsetX = x + (maxWidth - width) / 2;
    const offsetY = y + (maxHeight - height) / 2;
    doc.addImage(logo.dataUrl, "PNG", offsetX, offsetY, width, height, undefined, "FAST");
  }

  function drawNativeTitleBlock(doc, context) {
    const margin = 12;
    const width = 186;
    const y = 42;
    const cutoff = context.latestCutoff == null ? "Fecha de corte no disponible" : formatLongDate(context.latestCutoff);

    doc.setFillColor(18, 58, 90);
    doc.setDrawColor(18, 58, 90);
    doc.roundedRect(margin, y, width, 30, 4, 4, "FD");
    doc.setFont("helvetica", "bold");
    doc.setFontSize(21);
    doc.setTextColor(255, 255, 255);
    doc.text("Informe sobre accesibilidad", margin + 8, y + 12);
    doc.setFontSize(10);
    doc.text("Programa de analisis de las demoras en Atencion Primaria", margin + 8, y + 20);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8.5);
    doc.setTextColor(224, 237, 247);
    doc.text("Subdireccion de Enfermeria de Atencion Primaria · Gerencia de Atencion Integrada de Ciudad Real", margin + 8, y + 26);
    doc.setFillColor(245, 250, 253);
    doc.roundedRect(margin + width - 58, y + 7, 50, 10, 4, 4, "F");
    doc.setFont("helvetica", "bold");
    doc.setFontSize(7.5);
    doc.setTextColor(18, 58, 90);
    doc.text(cutoff, margin + width - 33, y + 13.5, { align: "center" });
  }

  function drawNativeContextCards(doc, context, startY) {
    const metrics = context.metrics;
    const cards = [
      ["Area", state.filters.area || "Todas las areas"],
      ["Zona", context.zoneMeta.value],
      ["Fecha de corte", context.latestCutoff == null ? "No disponible" : formatLongDate(context.latestCutoff)],
      ["Agendas validas", metrics ? formatInteger(metrics.totalValid) : "--"]
    ];

    if (state.filters.centro) {
      cards.splice(2, 0, ["Centro", state.filters.centro]);
    }
    if (state.filters.categoria) {
      cards.splice(3, 0, ["Categoria", state.filters.categoria]);
    }
    if (state.filters.tipoVisita) {
      cards.splice(4, 0, ["Tipo visita", state.filters.tipoVisita]);
    }

    const margin = 12;
    const gap = 4;
    const columns = 4;
    const cardWidth = (186 - gap * (columns - 1)) / columns;
    const cardHeight = 20;

    cards.slice(0, 8).forEach((card, index) => {
      const col = index % columns;
      const row = Math.floor(index / columns);
      const x = margin + col * (cardWidth + gap);
      const y = startY + row * (cardHeight + gap);
      doc.setFillColor(247, 251, 254);
      doc.setDrawColor(224, 235, 243);
      doc.roundedRect(x, y, cardWidth, cardHeight, 3, 3, "FD");
      doc.setFont("helvetica", "bold");
      doc.setFontSize(6.7);
      doc.setTextColor(106, 130, 152);
      doc.text(String(card[0]).toUpperCase(), x + 3, y + 6);
      doc.setFontSize(8.2);
      doc.setTextColor(27, 51, 74);
      doc.text(doc.splitTextToSize(String(card[1]), cardWidth - 6).slice(0, 2), x + 3, y + 12);
    });
  }

  function drawNativeSummary(doc, context, startY) {
    const summaryHtml = buildReportSummaryMarkup(context.filteredRows, context.metrics, context.latestCutoff);
    const summaryText = htmlToPlainText(summaryHtml);
    const lines = doc.splitTextToSize(summaryText, 176).slice(0, 12);
    const boxHeight = Math.max(34, lines.length * 4.1 + 14);

    doc.setFillColor(255, 255, 255);
    doc.setDrawColor(223, 234, 242);
    doc.roundedRect(12, startY, 186, boxHeight, 4, 4, "FD");
    doc.setFont("helvetica", "bold");
    doc.setFontSize(11);
    doc.setTextColor(18, 58, 90);
    doc.text("Resumen ejecutivo", 18, startY + 8);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8.5);
    doc.setTextColor(52, 77, 102);
    doc.text(lines, 18, startY + 15, { lineHeightFactor: 1.25 });
    return startY + boxHeight;
  }

  function drawNativeKpiGrid(doc, metrics, startY) {
    const kpis = [
      ["% 0-2 dias", formatPercent(metrics.pct0to2), "Acumulativo DGAP"],
      ["% 0-3 dias", formatPercent(metrics.pct0to3), "Acumulativo DGAP"],
      ["% 0-6 dias", formatPercent(metrics.pct0to6), "Acumulativo DGAP"],
      ["% >=7 dias", formatPercent(metrics.pct7Plus), "Demora prolongada"],
      ["Agendas", formatInteger(metrics.totalValid), "Base analizada"],
      ["Media", formatDayMetric(metrics.mean), "Dias"],
      ["Mediana", formatDayMetric(metrics.median), "Valor central"],
      ["Max.", formatDayMetric(metrics.max), "Mayor demora"]
    ];
    const margin = 12;
    const gap = 4;
    const cardWidth = (186 - gap * 3) / 4;
    const cardHeight = 28;

    doc.setFont("helvetica", "bold");
    doc.setFontSize(12);
    doc.setTextColor(18, 58, 90);
    doc.text("Indicadores clave", margin, startY + 4);

    kpis.forEach((kpi, index) => {
      const col = index % 4;
      const row = Math.floor(index / 4);
      const x = margin + col * (cardWidth + gap);
      const y = startY + 10 + row * (cardHeight + gap);
      const isAlert = index === 3;
      doc.setFillColor(isAlert ? 255 : 247, isAlert ? 243 : 251, isAlert ? 244 : 254);
      doc.setDrawColor(isAlert ? 230 : 223, isAlert ? 190 : 234, isAlert ? 196 : 242);
      doc.roundedRect(x, y, cardWidth, cardHeight, 3, 3, "FD");
      doc.setFont("helvetica", "bold");
      doc.setFontSize(6.8);
      doc.setTextColor(106, 130, 152);
      doc.text(kpi[0].toUpperCase(), x + 3, y + 6);
      doc.setFontSize(15);
      doc.setTextColor(isAlert ? 173 : 18, isAlert ? 72 : 58, isAlert ? 82 : 90);
      doc.text(kpi[1], x + 3, y + 16);
      doc.setFont("helvetica", "normal");
      doc.setFontSize(6.8);
      doc.setTextColor(112, 136, 157);
      doc.text(kpi[2], x + 3, y + 23);
    });
  }

  function drawNativeDetailPage(doc, context) {
    doc.addPage();
    drawNativeSectionTitle(doc, "Detalle de accesibilidad", "Tablas calculadas sobre agendas validas del subconjunto filtrado");
    const categoryRows = buildCategorySummaries(context.validRows);
    const categoryVisitRows = buildCategoryVisitSummaries(context.validRows);
    drawNativeTable(doc, "Detalle por categoria", 31, ["Categoria", "Agendas", "% 0-2", "% 0-3", "% 0-6", "% >=7 d", "Media", "Mediana", "Max."], categoryRows.map((row) => [
      row.category,
      formatInteger(row.totalValid),
      formatPercent(row.pct0to2),
      formatPercent(row.pct0to3),
      formatPercent(row.pct0to6),
      formatPercent(row.pct7Plus),
      formatReportDelay(row.mean),
      formatReportDelay(row.median),
      formatReportDelay(row.max)
    ]), {
      0: { cellWidth: 42 },
      1: { cellWidth: 18, halign: "right" },
      2: { cellWidth: 18, halign: "right" },
      3: { cellWidth: 18, halign: "right" },
      4: { cellWidth: 18, halign: "right" },
      5: { cellWidth: 18, halign: "right" },
      6: { cellWidth: 18, halign: "right" },
      7: { cellWidth: 18, halign: "right" },
      8: { cellWidth: 18, halign: "right" }
    });

    const nextY = Math.max(86, (doc.lastAutoTable && doc.lastAutoTable.finalY ? doc.lastAutoTable.finalY + 12 : 88));
    drawNativeTable(doc, "Detalle por categoria y tipo de visita", nextY, ["Categoria", "Tipo visita", "Agendas", "% 0-2", "% 0-3", "% 0-6", "% >=7 d", "Media", "Mediana", "Max."], categoryVisitRows.map((row) => [
      row.category,
      row.visitType,
      formatInteger(row.totalValid),
      formatPercent(row.pct0to2),
      formatPercent(row.pct0to3),
      formatPercent(row.pct0to6),
      formatPercent(row.pct7Plus),
      formatReportDelay(row.mean),
      formatReportDelay(row.median),
      formatReportDelay(row.max)
    ]), {
      0: { cellWidth: 28 },
      1: { cellWidth: 38 },
      2: { cellWidth: 15, halign: "right" },
      3: { cellWidth: 15, halign: "right" },
      4: { cellWidth: 15, halign: "right" },
      5: { cellWidth: 15, halign: "right" },
      6: { cellWidth: 15, halign: "right" },
      7: { cellWidth: 15, halign: "right" },
      8: { cellWidth: 15, halign: "right" },
      9: { cellWidth: 15, halign: "right" }
    });
  }

  function drawNativeProfessionalsPage(doc, context) {
    doc.addPage();
    drawNativeSectionTitle(doc, "Profesionales con mayor demora por categoria", "Top 5 por categoria profesional");
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8.3);
    doc.setTextColor(57, 85, 108);
    doc.text(doc.splitTextToSize("Se muestran, para cada categoria profesional, los cinco profesionales con mayor demora media dentro del subconjunto filtrado. El ranking se ordena por demora media descendente y, en caso de empate, por % >=7 dias, demora maxima y agendas validas.", 186), 12, 31);

    let y = 48;
    buildWorstProfessionalsByCategory(context.validRows).forEach((block) => {
      if (y > 238) {
        doc.addPage();
        y = 20;
      }
      doc.setFillColor(234, 246, 244);
      doc.setDrawColor(205, 229, 225);
      doc.roundedRect(12, y - 5, 186, 9, 2.5, 2.5, "FD");
      doc.setFont("helvetica", "bold");
      doc.setFontSize(10);
      doc.setTextColor(11, 109, 100);
      doc.text(block.category, 16, y + 1);
      drawNativeTable(doc, "", y + 8, ["Profesional", "Centro", "Agendas", "% 0-2", "% 0-3", "% 0-6", "% >=7 d", "Media", "Mediana", "Max."], block.professionals.map((row) => [
        row.professional,
        row.center,
        formatInteger(row.totalValid),
        formatPercent(row.pct0to2),
        formatPercent(row.pct0to3),
        formatPercent(row.pct0to6),
        formatPercent(row.pct7Plus),
        formatReportDelay(row.mean),
        formatReportDelay(row.median),
        formatReportDelay(row.max)
      ]), {
        0: { cellWidth: 43 },
        1: { cellWidth: 32 },
        2: { cellWidth: 14, halign: "right" },
        3: { cellWidth: 15, halign: "right" },
        4: { cellWidth: 15, halign: "right" },
        5: { cellWidth: 15, halign: "right" },
        6: { cellWidth: 15, halign: "right" },
        7: { cellWidth: 15, halign: "right" },
        8: { cellWidth: 15, halign: "right" },
        9: { cellWidth: 12, halign: "right" }
      }, 7.1);
      y = (doc.lastAutoTable && doc.lastAutoTable.finalY ? doc.lastAutoTable.finalY : y + 38) + 12;
    });
  }

  function drawNativeRankingsPage(doc, context) {
    doc.addPage();
    drawNativeSectionTitle(doc, "Rankings ejecutivos", "Visualizacion estable sin graficos canvas");
    const centers = aggregateByField(context.validRows, "Centro")
      .sort((left, right) => right.mean - left.mean || right.totalValid - left.totalValid)
      .slice(0, 10);
    const zones = aggregateByField(context.validRows, "Zona")
      .sort((left, right) => right.mean - left.mean || right.totalValid - left.totalValid)
      .slice(0, 10);
    drawNativeRankingBars(doc, "Centros con mayor demora media", centers, 32);
    drawNativeRankingBars(doc, "Zonas con mayor demora media", zones, 152);
  }

  function drawNativeProfessionalsOver7Page(doc, context) {
    doc.addPage();
    drawNativeSectionTitle(doc, "Profesionales con demora media de 7 o mas dias", "Tabla operativa calculada sobre el subconjunto filtrado");
    const rows = buildProfessionalsWithDelayOver7(context.validRows);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8.3);
    doc.setTextColor(57, 85, 108);
    doc.text(doc.splitTextToSize("Se incluyen exclusivamente profesionales con demora media igual o superior a 7 dias. El orden prioriza mayor demora media, mayor demora maxima, mayor porcentaje de agendas con 7 o mas dias y mayor numero de agendas validas.", 186), 12, 31);

    if (!rows.length) {
      drawNativeEmptyState(doc, 54, "No se identifican profesionales con demora media igual o superior a 7 dias en el subconjunto filtrado.");
      return;
    }

    drawNativeTable(doc, "", 54, ["Categoria", "Profesional", "Centro", "Zona", "Agendas", "Media", "Mediana", "Max.", "% >=7 d"], rows.map((row) => [
      row.category,
      row.professional,
      row.center,
      row.zone,
      formatInteger(row.totalValid),
      formatReportDelay(row.mean),
      formatReportDelay(row.median),
      formatReportDelay(row.max),
      formatPercent(row.pct7Plus)
    ]), {
      0: { cellWidth: 24 },
      1: { cellWidth: 42 },
      2: { cellWidth: 32 },
      3: { cellWidth: 25 },
      4: { cellWidth: 14, halign: "right" },
      5: { cellWidth: 14, halign: "right" },
      6: { cellWidth: 14, halign: "right" },
      7: { cellWidth: 11, halign: "right" },
      8: { cellWidth: 10, halign: "right" }
    }, 6.8);
  }

  function drawNativeMethodologyPage(doc) {
    doc.addPage();
    drawNativeSectionTitle(doc, "Metodologia y privacidad", "Criterios de calculo");
    const paragraphs = [
      "Los indicadores se calculan sobre agendas con valor numerico y no negativo en la columna Accesibilidad.",
      "Los porcentajes 0-2 dias, 0-3 dias y 0-6 dias son indicadores acumulativos DGAP. Una agenda con Accesibilidad 2 cuenta en 0-2, 0-3 y 0-6.",
      "% 7 o mas dias recoge las agendas con Accesibilidad igual o superior a 7 dias.",
      "Las tablas por categoria, tipo de visita y profesional calculan porcentajes sobre el total valido de cada grupo, no sobre el total global.",
      "El ranking de profesionales se ordena por demora media descendente; en empate, por % >=7 dias, demora maxima y agendas validas.",
      "El fichero Excel se procesa localmente en el dispositivo. No se envian datos a servidores externos. El PDF se genera localmente en el navegador o en la WebView de Android."
    ];
    let y = 36;
    paragraphs.forEach((paragraph, index) => {
      doc.setFillColor(index === paragraphs.length - 1 ? 243 : 255, index === paragraphs.length - 1 ? 251 : 255, index === paragraphs.length - 1 ? 249 : 255);
      doc.setDrawColor(223, 234, 242);
      const lines = doc.splitTextToSize(paragraph, 174);
      const height = Math.max(16, lines.length * 4.2 + 8);
      doc.roundedRect(12, y, 186, height, 3, 3, "FD");
      doc.setFont("helvetica", index === paragraphs.length - 1 ? "bold" : "normal");
      doc.setFontSize(9);
      doc.setTextColor(52, 77, 102);
      doc.text(lines, 18, y + 7, { lineHeightFactor: 1.25 });
      y += height + 7;
    });
  }

  function drawNativeSectionTitle(doc, title, subtitle) {
    doc.setFillColor(255, 255, 255);
    doc.setDrawColor(216, 228, 239);
    doc.line(12, 24, 198, 24);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(17);
    doc.setTextColor(18, 58, 90);
    doc.text(title, 12, 17);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8);
    doc.setTextColor(82, 113, 141);
    doc.text(subtitle, 12, 22);
  }

  function drawNativeTable(doc, title, startY, head, body, columnStyles, fontSize) {
    if (title) {
      doc.setFont("helvetica", "bold");
      doc.setFontSize(10.5);
      doc.setTextColor(18, 58, 90);
      doc.text(title, 12, startY - 4);
    }

    doc.autoTable({
      startY,
      head: [head],
      body,
      margin: { left: 12, right: 12, bottom: 18, top: 18 },
      tableWidth: 186,
      theme: "grid",
      styles: {
        font: "helvetica",
        fontSize: fontSize || 7.5,
        cellPadding: 1.4,
        overflow: "linebreak",
        valign: "middle",
        lineColor: [226, 235, 242],
        lineWidth: 0.15,
        textColor: [42, 64, 86]
      },
      headStyles: {
        fillColor: [237, 245, 251],
        textColor: [27, 51, 74],
        fontStyle: "bold",
        halign: "left"
      },
      alternateRowStyles: {
        fillColor: [248, 251, 253]
      },
      columnStyles,
      didParseCell: (data) => {
        if (data.section === "body" && data.column.index >= Math.max(1, head.length - 8)) {
          data.cell.styles.halign = "right";
        }
        if (data.section === "body" && head[data.column.index] === "% >=7 d") {
          data.cell.styles.textColor = [173, 72, 82];
          data.cell.styles.fontStyle = "bold";
        }
        if (data.section === "body" && head[data.column.index] === "Media") {
          const numericValue = Number(String(data.cell.raw || "").replace(",", ".").replace(/[^\d.-]/g, ""));
          data.cell.styles.fontStyle = "bold";
          data.cell.styles.textColor = numericValue >= 10 ? [173, 72, 82] : [18, 58, 90];
          data.cell.styles.fillColor = numericValue >= 10 ? [255, 243, 244] : [239, 247, 252];
        }
      }
    });
  }

  function drawNativeRankingBars(doc, title, rows, startY) {
    doc.setFont("helvetica", "bold");
    doc.setFontSize(11);
    doc.setTextColor(18, 58, 90);
    doc.text(title, 12, startY);
    const maxValue = Math.max(1, ...rows.map((row) => row.mean));
    let y = startY + 10;
    rows.forEach((row) => {
      const barWidth = (row.mean / maxValue) * 82;
      doc.setFont("helvetica", "normal");
      doc.setFontSize(8);
      doc.setTextColor(36, 61, 85);
      doc.text(doc.splitTextToSize(row.label, 62).slice(0, 1), 12, y);
      doc.setFillColor(232, 240, 246);
      doc.roundedRect(78, y - 4, 82, 5, 2, 2, "F");
      doc.setFillColor(row.mean >= 7 ? 213 : row.mean >= 4 ? 236 : 14, row.mean >= 7 ? 86 : row.mean >= 4 ? 179 : 157, row.mean >= 7 ? 99 : row.mean >= 4 ? 67 : 144);
      doc.roundedRect(78, y - 4, Math.max(2, barWidth), 5, 2, 2, "F");
      doc.setFont("helvetica", "bold");
      doc.setTextColor(27, 51, 74);
      doc.text(formatReportDelay(row.mean), 165, y, { align: "right" });
      doc.setFont("helvetica", "normal");
      doc.setTextColor(112, 136, 157);
      doc.text("n=" + formatInteger(row.totalValid), 197, y, { align: "right" });
      y += 10.5;
    });
  }

  function drawNativeEmptyState(doc, y, message) {
    doc.setFillColor(247, 251, 254);
    doc.setDrawColor(223, 234, 242);
    doc.roundedRect(12, y, 186, 34, 4, 4, "FD");
    doc.setFont("helvetica", "bold");
    doc.setFontSize(11);
    doc.setTextColor(82, 113, 141);
    doc.text(message, 105, y + 19, { align: "center" });
  }

  function addNativePdfFooters(doc) {
    const totalPages = doc.internal.getNumberOfPages();
    for (let page = 1; page <= totalPages; page += 1) {
      doc.setPage(page);
      doc.setDrawColor(223, 234, 242);
      doc.line(12, 282, 198, 282);
      doc.setFont("helvetica", "normal");
      doc.setFontSize(7);
      doc.setTextColor(112, 136, 157);
      doc.text("Programa de analisis de las demoras en Atencion Primaria", 12, 287);
      doc.text("Gerencia de Atencion Integrada de Ciudad Real", 105, 287, { align: "center" });
      doc.setFont("helvetica", "bold");
      doc.setTextColor(27, 51, 74);
      doc.text("Pagina " + page + " de " + totalPages, 198, 287, { align: "right" });
    }
  }

  function attachAutoTablePlugin(jsPDF) {
    if (!jsPDF || jsPDF.API.autoTable) {
      return;
    }

    const pluginObject = window.jspdfAutoTable || window.jsPDFAutoTable;
    if (pluginObject && typeof pluginObject.applyPlugin === "function") {
      pluginObject.applyPlugin(jsPDF);
      return;
    }

    if (typeof window.applyPlugin === "function") {
      window.applyPlugin(jsPDF);
      return;
    }

    if (typeof window.autoTable === "function") {
      jsPDF.API.autoTable = function (options) {
        return window.autoTable(this, options);
      };
    }
  }

  function htmlToPlainText(html) {
    const container = document.createElement("div");
    container.innerHTML = html;
    return normalizeText(container.textContent).replace(/\s+/g, " ");
  }

  function buildReportPdfFileName() {
    const filteredRows = getFilteredRows();
    const latestCutoff = getLatestDateSortValue(filteredRows, "Fecha Corte");
    const datePart = latestCutoff == null ? formatFileDate(new Date()) : formatFileDate(new Date(latestCutoff));
    const scope = state.filters.centro || state.filters.zona || state.filters.area || "AP";
    return "Informe_accesibilidad_" + sanitizeFileNamePart(scope) + "_" + datePart + ".pdf";
  }

  function formatFileDate(date) {
    const year = date.getUTCFullYear();
    const month = String(date.getUTCMonth() + 1).padStart(2, "0");
    const day = String(date.getUTCDate()).padStart(2, "0");
    return year + "-" + month + "-" + day;
  }

  function sanitizeFileNamePart(value) {
    return normalizeText(value)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-zA-Z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "")
      .slice(0, 48) || "AP";
  }

  function showPrintNotice(message) {
    if (!refs.reportPrintNotice) {
      window.alert(message);
      return;
    }

    refs.reportPrintNotice.textContent = message;
    refs.reportPrintNotice.classList.remove("is-hidden");
  }

  function hidePrintNotice() {
    if (!refs.reportPrintNotice) {
      return;
    }

    refs.reportPrintNotice.textContent = "";
    refs.reportPrintNotice.classList.add("is-hidden");
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

  function formatReportDelay(value) {
    return formatNumber(value, 1) + " d";
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



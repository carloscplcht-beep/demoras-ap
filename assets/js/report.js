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

  const IGNORED_HEADERS = new Set(["CIAS ID", "UID"]);
  const DATE_COLUMNS = new Set(["Fecha Primer Hueco Libre", "Fecha Primer Hueco ID", "Fecha Corte"]);

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
    red: "#d55663"
  };

  const state = {
    rows: [],
    hasAttemptedLoad: false,
    readToken: 0,
    renderTimer: 0
  };

  const refs = {
    fileInput: document.getElementById("fileInput"),
    resetFiltersButton: document.getElementById("resetFiltersButton"),
    printReportButton: document.getElementById("printReportButton"),
    reportScope: document.getElementById("reportScope"),
    reportArea: document.getElementById("reportArea"),
    reportZones: document.getElementById("reportZones"),
    reportZonesDetail: document.getElementById("reportZonesDetail"),
    reportCutoffDate: document.getElementById("reportCutoffDate"),
    reportSummaryBody: document.getElementById("reportSummaryBody"),
    reportEmptyState: document.getElementById("reportEmptyState"),
    reportContent: document.getElementById("reportContent"),
    reportKpiPct48h: document.getElementById("reportKpiPct48h"),
    reportKpiPct5d: document.getElementById("reportKpiPct5d"),
    reportKpiPct7d: document.getElementById("reportKpiPct7d"),
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

  if (!refs.fileInput || !refs.printReportButton || !refs.reportSummaryBody) {
    return;
  }

  initialize();

  function initialize() {
    bindEvents();
    renderReport();
  }

  function bindEvents() {
    // Capturamos el archivo antes de que otros listeners limpien el input.
    refs.fileInput.addEventListener("change", handleFileSelection, true);
    refs.printReportButton.addEventListener("click", handlePrintReport);

    FILTERS.forEach((filter) => {
      if (refs[filter.key]) {
        refs[filter.key].addEventListener("change", queueRender);
      }
    });

    if (refs.resetFiltersButton) {
      refs.resetFiltersButton.addEventListener("click", queueRender);
    }
  }

  async function handleFileSelection(event) {
    const file = event.target.files && event.target.files[0];
    const token = ++state.readToken;

    if (!file || !/\.xlsx$/i.test(file.name) || !window.XLSXLite) {
      state.rows = [];
      state.hasAttemptedLoad = Boolean(file);
      queueRender();
      return;
    }

    try {
      const workbook = await window.XLSXLite.readWorkbook(file);
      if (token !== state.readToken) {
        return;
      }

      const workbookData = readSheetFromRawRows(workbook.rawRows || []);
      state.rows = normalizeRows(workbookData.rows);
      state.hasAttemptedLoad = true;
      queueRender();
    } catch (error) {
      if (token !== state.readToken) {
        return;
      }
      state.rows = [];
      state.hasAttemptedLoad = true;
      queueRender();
    }
  }

  function queueRender() {
    window.clearTimeout(state.renderTimer);
    state.renderTimer = window.setTimeout(renderReport, 0);
  }

  function renderReport() {
    const filteredRows = getFilteredRows();
    const validRows = filteredRows.filter((row) => row.__hasValidAcc);
    const metrics = calculateExecutiveMetrics(validRows);
    const latestCutoff = getLatestDateSortValue(filteredRows, "Fecha Corte");
    const zoneMeta = getReportZoneMeta(filteredRows);

    updatePrintButton(metrics);
    refs.reportArea.textContent = refs.area && refs.area.value ? refs.area.value : "Todas las áreas";
    refs.reportZones.textContent = zoneMeta.value;
    refs.reportZonesDetail.textContent = zoneMeta.detail;
    refs.reportCutoffDate.textContent = latestCutoff == null ? "Fecha de corte no disponible" : formatLongDate(latestCutoff);
    refs.reportScope.textContent = metrics
      ? formatInteger(metrics.totalValid) + " agendas válidas analizadas"
      : "Sin datos analíticos";

    if (!state.hasAttemptedLoad) {
      refs.reportSummaryBody.innerHTML = "<p>Cargue un archivo Excel para generar el informe.</p>";
      refs.reportEmptyState.classList.remove("is-hidden");
      refs.reportEmptyState.textContent = "Cargue un archivo Excel para comenzar la elaboración del informe.";
      refs.reportContent.classList.add("is-hidden");
      setReportKpisEmpty();
      setReportChartsEmpty("Cargue un Excel para generar el gráfico del informe.");
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
      setReportChartsEmpty("No hay agendas válidas con los filtros actuales.");
      return;
    }

    refs.reportEmptyState.classList.add("is-hidden");
    refs.reportContent.classList.remove("is-hidden");
    setKpiValue(refs.reportKpiPct48h, formatPercent(metrics.pctUnder48h));
    setKpiValue(refs.reportKpiPct5d, formatPercent(metrics.pctUnder5d));
    setKpiValue(refs.reportKpiPct7d, formatPercent(metrics.pct7OrMore));
    setKpiValue(refs.reportKpiTotalValid, formatInteger(metrics.totalValid));
    setKpiValue(refs.reportKpiMean, formatDayMetric(metrics.mean));
    setKpiValue(refs.reportKpiMedian, formatDayMetric(metrics.median));
    setKpiValue(refs.reportKpiMax, formatDayMetric(metrics.max));
    renderReportCharts(validRows);
  }

  function renderReportCharts(validRows) {
    if (!window.DemorasCharts) {
      return;
    }

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
      annotationFormatter: function (item) {
        return formatPercent(item.share);
      },
      titleFormatter: function (item) {
        return formatInteger(item.value) + " agendas";
      }
    });

    window.DemorasCharts.renderHorizontalBarChart(refs.reportCentersChart, {
      items: topCenters,
      valueFormatter: formatDayMetric,
      emptyMessage: "No hay centros con agendas válidas para representar en el informe."
    });
  }

  function setReportChartsEmpty(message) {
    if (!window.DemorasCharts) {
      return;
    }
    window.DemorasCharts.setEmpty(refs.reportBandsChart, message);
    window.DemorasCharts.setEmpty(refs.reportCentersChart, message);
  }

  function setReportKpisEmpty() {
    setKpiValue(refs.reportKpiPct48h, "--");
    setKpiValue(refs.reportKpiPct5d, "--");
    setKpiValue(refs.reportKpiPct7d, "--");
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

    const paragraphs = [
      "El análisis realizado sobre las agendas filtradas incorpora " +
        formatInteger(metrics.totalValid) +
        " agendas válidas para el cálculo ejecutivo" +
        filteredBaseText +
        ". " +
        cutoffText,
      "En términos de accesibilidad, el " +
        formatPercent(metrics.pctUnder48h) +
        " de las agendas presenta una demora inferior a 48 horas y el " +
        formatPercent(metrics.pctUnder5d) +
        " se sitúa por debajo de 5 días. Por su parte, el " +
        formatPercent(metrics.pct7OrMore) +
        " alcanza o supera los 7 días, con una demora media de " +
        formatDaysText(metrics.mean) +
        ", una mediana de " +
        formatDaysText(metrics.median) +
        " y una demora máxima de " +
        formatDaysText(metrics.max) +
        ".",
      buildExecutiveAssessment(metrics)
    ];

    return paragraphs.map(function (paragraph) {
      return "<p>" + escapeHtml(paragraph) + "</p>";
    }).join("");
  }

  function buildExecutiveAssessment(metrics) {
    if (metrics.pct7OrMore >= 35 || metrics.mean >= 7) {
      return "En conjunto, la distribución observada refleja una presión relevante en la accesibilidad, con un peso significativo de agendas en demora prolongada.";
    }

    if (metrics.pctUnder48h >= 60 && metrics.pct7OrMore <= 15) {
      return "En conjunto, la situación muestra un comportamiento favorable, con predominio de agendas en tramos de demora corta y una presencia contenida de demoras prolongadas.";
    }

    return "En conjunto, la situación presenta un comportamiento intermedio, con margen de mejora en la reducción de las agendas que concentran mayores demoras.";
  }

  function getReportZoneMeta(filteredRows) {
    if (refs.zona && refs.zona.value) {
      return {
        value: refs.zona.value,
        detail: ""
      };
    }

    const filteredZones = uniqueSorted(filteredRows.map(function (row) {
      return normalizeText(row["Zona"]);
    }).filter(Boolean));
    const allZones = uniqueSorted(state.rows.map(function (row) {
      return normalizeText(row["Zona"]);
    }).filter(Boolean));

    if (!filteredZones.length || filteredZones.length === allZones.length) {
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

    return values.slice(0, visibleItems).join(", ") + " y " + formatInteger(values.length - visibleItems) + " más";
  }

  function readSheetFromRawRows(rawRows) {
    const raw = Array.isArray(rawRows) ? rawRows.filter(function (row) { return !isEmptyRow(row); }) : [];
    if (!raw.length) {
      throw new Error("La hoja está vacía.");
    }

    let headers = (raw[0] || []).map(normalizeHeader);
    headers = headers.map(function (header) {
      return IGNORED_HEADERS.has(header) ? null : header;
    });
    headers = dedupeHeaders(headers);

    const detectedHeaders = headers.filter(Boolean);
    const missingHeaders = REQUIRED_HEADERS.filter(function (column) {
      return !detectedHeaders.includes(column);
    });

    if (missingHeaders.length) {
      throw new Error("Estructura no válida.");
    }

    return {
      rows: raw
        .slice(1)
        .filter(function (row) { return !isEmptyRow(row); })
        .map(function (row) { return buildRowObject(row, headers); })
    };
  }

  function normalizeRows(rows) {
    return rows.map(function (rawRow, index) {
      const normalizedRow = {};
      const dateSortValues = {};

      REQUIRED_HEADERS.forEach(function (header) {
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

  function buildRowObject(row, processedHeaders) {
    const rowObject = {};
    processedHeaders.forEach(function (header, index) {
      if (header) {
        rowObject[header] = row[index] == null ? "" : row[index];
      }
    });
    return rowObject;
  }

  function dedupeHeaders(headers) {
    const seenHeaders = Object.create(null);
    return headers.map(function (header) {
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

  function normalizeHeader(value) {
    return String(value == null ? "" : value).trim().replace(/\s+/g, " ");
  }

  function isEmptyRow(row) {
    return !Array.isArray(row) || row.every(function (value) {
      return normalizeText(value) === "";
    });
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

  function parseNumericValue(value) {
    if (typeof value === "number" && Number.isFinite(value)) {
      return value;
    }

    const textValue = normalizeText(value);
    if (!textValue) {
      return NaN;
    }

    const normalized = textValue.replace(/\s+/g, "").replace(",", ".");
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : NaN;
  }

  function getFilteredRows() {
    return state.rows.filter(function (row) {
      return FILTERS.every(function (filter) {
        const element = refs[filter.key];
        const selectedValue = element ? element.value : "";
        return !selectedValue || row[filter.column] === selectedValue;
      });
    });
  }

  function calculateExecutiveMetrics(validRows) {
    if (!validRows.length) {
      return null;
    }

    const values = validRows.map(function (row) {
      return row.__acc;
    }).slice().sort(function (left, right) {
      return left - right;
    });

    const totalValid = values.length;
    const countUnder48h = values.filter(function (value) { return value < 2; }).length;
    const countUnder5d = values.filter(function (value) { return value < 5; }).length;
    const count7OrMore = values.filter(function (value) { return value >= 7; }).length;

    return {
      totalValid: totalValid,
      pctUnder48h: (countUnder48h / totalValid) * 100,
      pctUnder5d: (countUnder5d / totalValid) * 100,
      pct7OrMore: (count7OrMore / totalValid) * 100,
      mean: average(values),
      median: median(values),
      max: Math.max.apply(null, values)
    };
  }

  function buildExecutiveDelayBands(rows) {
    const total = rows.length || 1;
    const bands = [
      { label: "<48h", count: 0, color: PALETTE.green },
      { label: "2-4 días", count: 0, color: PALETTE.teal },
      { label: "5-6 días", count: 0, color: PALETTE.amber },
      { label: "7-13 días", count: 0, color: PALETTE.red },
      { label: "14+ días", count: 0, color: "#9f4c69" }
    ];

    rows.forEach(function (row) {
      if (row.__acc < 2) {
        bands[0].count += 1;
      } else if (row.__acc < 5) {
        bands[1].count += 1;
      } else if (row.__acc < 7) {
        bands[2].count += 1;
      } else if (row.__acc < 14) {
        bands[3].count += 1;
      } else {
        bands[4].count += 1;
      }
    });

    bands.forEach(function (band) {
      band.share = (band.count / total) * 100;
    });

    return bands;
  }

  function aggregateByField(rows, column) {
    const grouped = new Map();

    rows.forEach(function (row) {
      const label = normalizeText(row[column]) || "Sin dato";
      if (!grouped.has(label)) {
        grouped.set(label, {
          label: label,
          values: []
        });
      }
      grouped.get(label).values.push(row.__acc);
    });

    return Array.from(grouped.values()).map(function (group) {
      return {
        label: group.label,
        count: group.values.length,
        mean: average(group.values)
      };
    });
  }

  function getLatestDateSortValue(rows, column) {
    const timestamps = rows
      .map(function (row) { return row.__dateSort && row.__dateSort[column]; })
      .filter(function (value) { return Number.isFinite(value); });

    return timestamps.length ? Math.max.apply(null, timestamps) : null;
  }

  function average(values) {
    return values.length
      ? values.reduce(function (sum, value) { return sum + value; }, 0) / values.length
      : 0;
  }

  function median(values) {
    if (!values.length) {
      return 0;
    }

    const sorted = values.slice().sort(function (left, right) { return left - right; });
    const middle = Math.floor(sorted.length / 2);
    return sorted.length % 2 !== 0 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
  }

  function handlePrintReport() {
    if (refs.printReportButton.disabled) {
      return;
    }

    const reportTabButton = document.querySelector('[data-tab-target="report"]');
    if (reportTabButton) {
      reportTabButton.click();
    }
    window.print();
  }

  function updatePrintButton(metrics) {
    const canPrint = Boolean(metrics);
    refs.printReportButton.disabled = !canPrint;
    refs.printReportButton.title = canPrint
      ? ""
      : "Cargue un Excel con agendas válidas para imprimir el informe.";
  }

  function setKpiValue(element, value) {
    if (element) {
      element.textContent = value;
    }
  }

  function normalizeText(value) {
    return value == null ? "" : String(value).trim();
  }

  function uniqueSorted(values) {
    return Array.from(new Set(values)).sort(function (left, right) {
      return String(left).localeCompare(String(right), "es", {
        sensitivity: "base",
        numeric: true
      });
    });
  }

  function formatPercent(value) {
    return formatNumber(value, 1) + "%";
  }

  function formatInteger(value) {
    return formatNumber(value, 0);
  }

  function formatDayMetric(value) {
    const decimals = Math.abs(value - Math.round(value)) < 0.01 ? 0 : 1;
    return formatNumber(value, decimals) + " d";
  }

  function formatDaysText(value) {
    const decimals = Math.abs(value - Math.round(value)) < 0.01 ? 0 : 1;
    return formatNumber(value, decimals) + " días";
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

  function escapeHtml(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }
})();

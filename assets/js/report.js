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

  ensureDgapInterface();
  ensureReportDetailUi();
  ensureReportDetailStyles();

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
    summaryKpiPct0to2: document.getElementById("kpiPct0to2"),
    summaryKpiPct0to3: document.getElementById("kpiPct0to3"),
    summaryKpiPct0to6: document.getElementById("kpiPct0to6"),
    summaryKpiPct7Plus: document.getElementById("kpiPct7Plus"),
    summaryKpiTotalValid: document.getElementById("kpiTotalValid"),
    summaryKpiMean: document.getElementById("kpiMean"),
    summaryKpiMedian: document.getElementById("kpiMedian"),
    summaryKpiMax: document.getElementById("kpiMax"),
    summaryBandsChart: document.getElementById("summaryBandsChart"),
    summaryDonutChart: document.getElementById("summaryDonutChart"),
    centersChart: document.getElementById("centersChart"),
    zonesChart: document.getElementById("zonesChart"),
    categoryChart: document.getElementById("categoryChart"),
    visitTypeChart: document.getElementById("visitTypeChart"),
    histogramChart: document.getElementById("histogramChart"),
    centerBandsChart: document.getElementById("centerBandsChart"),
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

  if (!refs.fileInput || !refs.printReportButton || !refs.reportSummaryBody) {
    return;
  }

  initialize();

  function initialize() {
    bindEvents();
    renderReport();
  }

  function ensureDgapInterface() {
    patchKpiCard("kpiPct0to2", "% 0-2 días", "Indicador DGAP acumulativo", "Accesibilidad entre 0 y 2 días.");
    insertKpiCardAfter("kpiPct0to2", "kpiPct0to3", "% 0-3 días", "Indicador DGAP acumulativo", "Accesibilidad entre 0 y 3 días.");
    patchKpiCard("kpiPct0to6", "% 0-6 días", "Indicador DGAP acumulativo", "Accesibilidad entre 0 y 6 días.");
    patchKpiCard("kpiPct7Plus", "% 7 o más días", "Demora prolongada", "Accesibilidad igual o superior a 7 días.");

    patchKpiCard("reportKpiPct0to2", "% 0-2 días", "Indicador DGAP acumulativo", "Accesibilidad entre 0 y 2 días.");
    insertKpiCardAfter("reportKpiPct0to2", "reportKpiPct0to3", "% 0-3 días", "Indicador DGAP acumulativo", "Accesibilidad entre 0 y 3 días.");
    patchKpiCard("reportKpiPct0to6", "% 0-6 días", "Indicador DGAP acumulativo", "Accesibilidad entre 0 y 6 días.");
    patchKpiCard("reportKpiPct7Plus", "% 7 o más días", "Demora prolongada", "Accesibilidad igual o superior a 7 días.");

    document.querySelectorAll(".method-list").forEach(function (list) {
      list.innerHTML =
        "<li><strong>% 0-2 días:</strong> agendas con Accesibilidad entre 0 y 2 días.</li>" +
        "<li><strong>% 0-3 días:</strong> agendas con Accesibilidad entre 0 y 3 días.</li>" +
        "<li><strong>% 0-6 días:</strong> agendas con Accesibilidad entre 0 y 6 días.</li>" +
        "<li><strong>% 7 o más días:</strong> agendas con Accesibilidad igual o superior a 7 días.</li>" +
        "<li>Los indicadores 0-2, 0-3 y 0-6 días son acumulativos. Los gráficos de distribución usan tramos excluyentes: 0-2, 3, 4-6 y 7 o más días.</li>";
    });
  }

  function insertKpiCardAfter(anchorValueId, newValueId, label, kicker, detail) {
    if (document.getElementById(newValueId)) {
      patchKpiCard(newValueId, label, kicker, detail);
      return;
    }

    const anchorValue = document.getElementById(anchorValueId);
    const anchorCard = anchorValue ? anchorValue.closest(".kpi-card") : null;
    if (!anchorCard || !anchorCard.parentNode) {
      return;
    }

    const clone = anchorCard.cloneNode(true);
    const clonedValue = clone.querySelector(".kpi-card__value");
    if (clonedValue) {
      clonedValue.id = newValueId;
      clonedValue.textContent = "--";
    }
    anchorCard.parentNode.insertBefore(clone, anchorCard.nextSibling);
    patchKpiCard(newValueId, label, kicker, detail);
  }

  function patchKpiCard(valueId, label, kicker, detail) {
    const value = document.getElementById(valueId);
    const card = value ? value.closest(".kpi-card") : null;
    if (!card) {
      return;
    }

    const kickerElement = card.querySelector(".kpi-card__kicker");
    const labelElement = card.querySelector(".kpi-card__label");
    const detailElement = card.querySelector(".kpi-card__detail");

    if (kickerElement) {
      kickerElement.textContent = kicker;
    }
    if (labelElement) {
      labelElement.textContent = label;
    }
    if (detailElement) {
      detailElement.textContent = detail;
    }
  }

  function ensureReportDetailUi() {
    if (document.getElementById("reportCategoryTableBody")) {
      return;
    }

    const chartsSection = document.querySelector(".report-charts");
    if (!chartsSection) {
      return;
    }

    const detailSection = document.createElement("section");
    detailSection.className = "report-detail-tables";
    detailSection.setAttribute("aria-labelledby", "reportDetailTablesTitle");
    detailSection.innerHTML =
      '<div class="report-section-head">' +
        "<div>" +
          '<p class="section-kicker">Análisis desagregado</p>' +
          '<h3 id="reportDetailTablesTitle">Detalle de accesibilidad</h3>' +
        "</div>" +
      "</div>" +
      '<article class="report-table-card" aria-labelledby="reportCategoryTitle">' +
        '<div class="report-table-card__head">' +
          "<div>" +
            '<p class="section-kicker">Categoría profesional</p>' +
            '<h4 id="reportCategoryTitle">Detalle por categoría</h4>' +
          "</div>" +
        "</div>" +
        '<div class="report-table-empty is-hidden" id="reportCategoryEmptyState">No hay agendas válidas para mostrar el detalle por categoría.</div>' +
        '<div class="report-table-wrapper" id="reportCategoryTableWrapper">' +
          '<table class="report-table" aria-describedby="reportCategoryTitle">' +
            "<thead><tr>" +
              '<th scope="col">Categoría</th>' +
              '<th scope="col" class="report-table__numeric">Agendas válidas</th>' +
              '<th scope="col" class="report-table__numeric">% 0-2 días</th>' +
              '<th scope="col" class="report-table__numeric">% 0-3 días</th>' +
              '<th scope="col" class="report-table__numeric">% 0-6 días</th>' +
              '<th scope="col" class="report-table__numeric">% 7 o más días</th>' +
              '<th scope="col" class="report-table__numeric">Media</th>' +
              '<th scope="col" class="report-table__numeric">Mediana</th>' +
              '<th scope="col" class="report-table__numeric">Máxima</th>' +
            "</tr></thead>" +
            '<tbody id="reportCategoryTableBody"></tbody>' +
          "</table>" +
        "</div>" +
      "</article>" +
      '<article class="report-table-card" aria-labelledby="reportCategoryVisitTitle">' +
        '<div class="report-table-card__head">' +
          "<div>" +
            '<p class="section-kicker">Categoría y agenda</p>' +
            '<h4 id="reportCategoryVisitTitle">Detalle por categoría y tipo de visita</h4>' +
          "</div>" +
        "</div>" +
        '<div class="report-table-empty is-hidden" id="reportCategoryVisitEmptyState">No hay agendas válidas para mostrar el detalle por categoría y tipo de visita.</div>' +
        '<div class="report-table-wrapper" id="reportCategoryVisitTableWrapper">' +
          '<table class="report-table report-table--compact" aria-describedby="reportCategoryVisitTitle">' +
            "<thead><tr>" +
              '<th scope="col">Categoría</th>' +
              '<th scope="col">Tipo visita</th>' +
              '<th scope="col" class="report-table__numeric">Agendas válidas</th>' +
              '<th scope="col" class="report-table__numeric">% 0-2 días</th>' +
              '<th scope="col" class="report-table__numeric">% 0-3 días</th>' +
              '<th scope="col" class="report-table__numeric">% 0-6 días</th>' +
              '<th scope="col" class="report-table__numeric">% 7 o más días</th>' +
              '<th scope="col" class="report-table__numeric">Media</th>' +
              '<th scope="col" class="report-table__numeric">Mediana</th>' +
              '<th scope="col" class="report-table__numeric">Máxima</th>' +
            "</tr></thead>" +
            '<tbody id="reportCategoryVisitTableBody"></tbody>' +
          "</table>" +
        "</div>" +
      "</article>";

    chartsSection.parentNode.insertBefore(detailSection, chartsSection);
  }

  function ensureReportDetailStyles() {
    if (document.getElementById("reportDetailRuntimeStyles")) {
      return;
    }

    const style = document.createElement("style");
    style.id = "reportDetailRuntimeStyles";
    style.textContent =
      ".report-detail-tables{display:grid;gap:18px}" +
      ".report-table-card{padding:22px 24px;border-radius:24px;background:rgba(247,251,255,.82);border:1px solid rgba(27,55,88,.08);display:grid;gap:14px}" +
      ".report-table-card__head{display:flex;justify-content:space-between;gap:16px;align-items:start}" +
      ".report-table-card__head h4{margin:8px 0 0;font-family:var(--font-display);font-size:1.12rem;color:var(--ink-900)}" +
      ".report-table-wrapper{overflow:auto;max-height:430px;border-radius:18px;border:1px solid rgba(27,55,88,.09);background:rgba(255,255,255,.86)}" +
      ".report-table{width:100%;min-width:920px;border-collapse:separate;border-spacing:0;font-size:.9rem}" +
      ".report-table--compact{min-width:1080px;font-size:.86rem}" +
      ".report-table th,.report-table td{padding:12px 14px;border-bottom:1px solid rgba(27,55,88,.08);color:var(--ink-700);vertical-align:top}" +
      ".report-table thead th{position:sticky;top:0;z-index:1;background:linear-gradient(180deg,#f4f9ff,#edf5fb);color:var(--ink-900);font-size:.76rem;font-weight:800;text-transform:uppercase;letter-spacing:.055em;text-align:left;white-space:nowrap}" +
      ".report-table tbody th{font-weight:800;color:var(--ink-900)}" +
      ".report-table tbody tr:last-child th,.report-table tbody tr:last-child td{border-bottom:0}" +
      ".report-table tbody tr:nth-child(even){background:rgba(40,118,185,.035)}" +
      ".report-table__numeric{text-align:right!important;white-space:nowrap;font-variant-numeric:tabular-nums}" +
      ".report-table__signal{color:#ad4852!important;font-weight:800}" +
      ".report-table-empty{padding:22px;border-radius:18px;border:1px dashed rgba(40,118,185,.18);background:rgba(255,255,255,.74);color:var(--ink-500);text-align:center}" +
      "@media print{.report-table-card{padding:12px 0;border-left:0;border-right:0;border-radius:0}.report-table-wrapper{max-height:none;overflow:visible;border-radius:0}.report-table,.report-table--compact{min-width:0;font-size:.72rem}.report-table thead{display:table-header-group}.report-table thead th{position:static;background:#eef4f9!important}.report-table th,.report-table td{padding:7px 6px;color:#24384d!important}.report-table tr{break-inside:avoid}}";
    document.head.appendChild(style);
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

    renderDgapDashboard(filteredRows, validRows, metrics);
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
      setReportDetailTablesEmpty("Cargue un Excel para generar el detalle analítico.");
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
      setReportDetailTablesEmpty("No hay agendas válidas con los filtros actuales.");
      setReportChartsEmpty("No hay agendas válidas con los filtros actuales.");
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

  function renderDgapDashboard(filteredRows, validRows, metrics) {
    if (!state.hasAttemptedLoad || !metrics) {
      setSummaryKpisEmpty();
      setDashboardChartsEmpty(!state.hasAttemptedLoad
        ? "Cargue un Excel para generar los gráficos."
        : "No hay agendas válidas con los filtros actuales.");
      return;
    }

    setKpiValue(refs.summaryKpiPct0to2, formatPercent(metrics.pct0to2));
    setKpiValue(refs.summaryKpiPct0to3, formatPercent(metrics.pct0to3));
    setKpiValue(refs.summaryKpiPct0to6, formatPercent(metrics.pct0to6));
    setKpiValue(refs.summaryKpiPct7Plus, formatPercent(metrics.pct7Plus));
    setKpiValue(refs.summaryKpiTotalValid, formatInteger(metrics.totalValid));
    setKpiValue(refs.summaryKpiMean, formatDayMetric(metrics.mean));
    setKpiValue(refs.summaryKpiMedian, formatDayMetric(metrics.median));
    setKpiValue(refs.summaryKpiMax, formatDayMetric(metrics.max));

    renderDgapCharts(validRows);
  }

  function setSummaryKpisEmpty() {
    [
      refs.summaryKpiPct0to2,
      refs.summaryKpiPct0to3,
      refs.summaryKpiPct0to6,
      refs.summaryKpiPct7Plus,
      refs.summaryKpiTotalValid,
      refs.summaryKpiMean,
      refs.summaryKpiMedian,
      refs.summaryKpiMax
    ].forEach(function (element) {
      setKpiValue(element, "--");
    });
  }

  function renderDgapCharts(validRows) {
    if (!window.DemorasCharts) {
      return;
    }

    const exclusiveBands = buildExclusiveDelayBands(validRows);
    const bandItems = exclusiveBands.map(function (band) {
      return {
        label: band.label,
        value: band.count,
        color: band.color,
        share: band.share
      };
    });

    window.DemorasCharts.renderVerticalBarChart(refs.summaryBandsChart, {
      items: bandItems,
      valueFormatter: formatInteger,
      annotationFormatter: function (item) { return formatPercent(item.share); },
      titleFormatter: function (item) { return formatInteger(item.value) + " agendas"; }
    });

    window.DemorasCharts.renderDonutChart(refs.summaryDonutChart, {
      items: bandItems,
      valueFormatter: formatInteger,
      annotationFormatter: function (item) { return formatPercent(item.share); },
      titleFormatter: function (item) { return item.label + ": " + formatInteger(item.value) + " agendas"; },
      centerLabel: formatInteger(validRows.length),
      centerSubLabel: "agendas válidas"
    });

    renderRankingChart(refs.centersChart, validRows, "Centro", PALETTE.blue, 10, "No hay centros con agendas válidas.");
    renderRankingChart(refs.zonesChart, validRows, "Zona", PALETTE.teal, 10, "No hay zonas con agendas válidas.");
    renderRankingChart(refs.categoryChart, validRows, "Categoría", PALETTE.green, 8, "No hay categorías con agendas válidas.");
    renderRankingChart(refs.visitTypeChart, validRows, "Tipo visita", PALETTE.amber, 8, "No hay tipos de visita con agendas válidas.");
    renderHistogramChart(validRows);
    renderCenterBandsChart(validRows);
  }

  function renderRankingChart(container, rows, column, color, limit, emptyMessage) {
    if (!window.DemorasCharts || !container) {
      return;
    }

    const items = aggregateByField(rows, column)
      .sort(function (left, right) {
        return right.mean - left.mean || right.count - left.count || left.label.localeCompare(right.label, "es", { sensitivity: "base", numeric: true });
      })
      .slice(0, limit)
      .map(function (item) {
        return {
          label: item.label,
          value: item.mean,
          color: color,
          note: "n=" + formatInteger(item.count)
        };
      });

    window.DemorasCharts.renderHorizontalBarChart(container, {
      items: items,
      valueFormatter: formatDayMetric,
      emptyMessage: emptyMessage
    });
  }

  function renderHistogramChart(rows) {
    if (!window.DemorasCharts || !refs.histogramChart) {
      return;
    }

    const histogramBands = [
      { label: "0-2 días", min: 0, max: 2, count: 0, color: PALETTE.green },
      { label: "3 días", min: 3, max: 3, count: 0, color: PALETTE.teal },
      { label: "4-6 días", min: 4, max: 6, count: 0, color: PALETTE.amber },
      { label: "7-9 días", min: 7, max: 9, count: 0, color: PALETTE.red },
      { label: "10-14 días", min: 10, max: 14, count: 0, color: "#b45568" },
      { label: "15+ días", min: 15, max: Infinity, count: 0, color: "#7d4a66" }
    ];

    rows.forEach(function (row) {
      const band = histogramBands.find(function (item) {
        return row.__acc >= item.min && row.__acc <= item.max;
      });
      if (band) {
        band.count += 1;
      }
    });

    window.DemorasCharts.renderVerticalBarChart(refs.histogramChart, {
      items: histogramBands.map(function (band) {
        return {
          label: band.label,
          value: band.count,
          color: band.color
        };
      }),
      valueFormatter: formatInteger,
      emptyMessage: "No hay agendas válidas para construir el histograma."
    });
  }

  function renderCenterBandsChart(rows) {
    if (!window.DemorasCharts || !refs.centerBandsChart) {
      return;
    }

    const centers = aggregateCenterBands(rows).slice(0, 8);
    window.DemorasCharts.renderStackedBarChart(refs.centerBandsChart, {
      items: centers,
      segments: [
        { key: "band0to2", label: "0-2 días", color: PALETTE.green },
        { key: "band3", label: "3 días", color: PALETTE.teal },
        { key: "band4to6", label: "4-6 días", color: PALETTE.amber },
        { key: "band7Plus", label: "7 o más días", color: PALETTE.red }
      ],
      valueFormatter: formatInteger,
      emptyMessage: "No hay centros con agendas válidas para representar tramos."
    });
  }

  function setDashboardChartsEmpty(message) {
    if (!window.DemorasCharts) {
      return;
    }

    [
      refs.summaryBandsChart,
      refs.summaryDonutChart,
      refs.centersChart,
      refs.zonesChart,
      refs.categoryChart,
      refs.visitTypeChart,
      refs.histogramChart,
      refs.centerBandsChart
    ].forEach(function (container) {
      if (container) {
        window.DemorasCharts.setEmpty(container, message);
      }
    });
  }

  function renderReportDetailTables(validRows) {
    const categoryRows = buildCategorySummaries(validRows);
    const categoryVisitRows = buildCategoryVisitSummaries(validRows);

    renderReportTable({
      rows: categoryRows,
      body: refs.reportCategoryTableBody,
      wrapper: refs.reportCategoryTableWrapper,
      emptyState: refs.reportCategoryEmptyState,
      emptyMessage: "No hay agendas válidas para mostrar el detalle por categoría.",
      rowRenderer: renderCategoryDetailRow
    });

    renderReportTable({
      rows: categoryVisitRows,
      body: refs.reportCategoryVisitTableBody,
      wrapper: refs.reportCategoryVisitTableWrapper,
      emptyState: refs.reportCategoryVisitEmptyState,
      emptyMessage: "No hay agendas válidas para mostrar el detalle por categoría y tipo de visita.",
      rowRenderer: renderCategoryVisitDetailRow
    });
  }

  function renderReportTable(options) {
    if (!options.body || !options.wrapper || !options.emptyState) {
      return;
    }

    if (!options.rows.length) {
      options.body.innerHTML = "";
      options.wrapper.classList.add("is-hidden");
      options.emptyState.classList.remove("is-hidden");
      options.emptyState.textContent = options.emptyMessage;
      return;
    }

    options.emptyState.classList.add("is-hidden");
    options.wrapper.classList.remove("is-hidden");
    options.body.innerHTML = options.rows.map(options.rowRenderer).join("");
  }

  function setReportDetailTablesEmpty(message) {
    renderReportTable({
      rows: [],
      body: refs.reportCategoryTableBody,
      wrapper: refs.reportCategoryTableWrapper,
      emptyState: refs.reportCategoryEmptyState,
      emptyMessage: message,
      rowRenderer: renderCategoryDetailRow
    });

    renderReportTable({
      rows: [],
      body: refs.reportCategoryVisitTableBody,
      wrapper: refs.reportCategoryVisitTableWrapper,
      emptyState: refs.reportCategoryVisitEmptyState,
      emptyMessage: message,
      rowRenderer: renderCategoryVisitDetailRow
    });
  }

  function renderCategoryDetailRow(row) {
    return (
      "<tr>" +
      '<th scope="row">' + escapeHtml(row.category) + "</th>" +
      renderNumericCell(formatInteger(row.totalValid)) +
      renderNumericCell(formatPercent(row.pct0to2)) +
      renderNumericCell(formatPercent(row.pct0to3)) +
      renderNumericCell(formatPercent(row.pct0to6)) +
      renderNumericCell(formatPercent(row.pct7Plus), row.pct7Plus >= 20 ? "report-table__signal" : "") +
      renderNumericCell(formatReportDelay(row.mean)) +
      renderNumericCell(formatReportDelay(row.median)) +
      renderNumericCell(formatReportDelay(row.max)) +
      "</tr>"
    );
  }

  function renderCategoryVisitDetailRow(row) {
    return (
      "<tr>" +
      '<th scope="row">' + escapeHtml(row.category) + "</th>" +
      "<td>" + escapeHtml(row.visitType) + "</td>" +
      renderNumericCell(formatInteger(row.totalValid)) +
      renderNumericCell(formatPercent(row.pct0to2)) +
      renderNumericCell(formatPercent(row.pct0to3)) +
      renderNumericCell(formatPercent(row.pct0to6)) +
      renderNumericCell(formatPercent(row.pct7Plus), row.pct7Plus >= 20 ? "report-table__signal" : "") +
      renderNumericCell(formatReportDelay(row.mean)) +
      renderNumericCell(formatReportDelay(row.median)) +
      renderNumericCell(formatReportDelay(row.max)) +
      "</tr>"
    );
  }

  function renderNumericCell(value, extraClass) {
    const className = ["report-table__numeric", extraClass].filter(Boolean).join(" ");
    return '<td class="' + className + '">' + escapeHtml(value) + "</td>";
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

    const detailAssessment = buildReportDetailAssessment(filteredRows);
    const closingAssessment = [buildExecutiveAssessment(metrics), detailAssessment].filter(Boolean).join(" ");
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
      closingAssessment
    ];

    return paragraphs.map(function (paragraph) {
      return "<p>" + escapeHtml(paragraph) + "</p>";
    }).join("");
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
    const validRows = filteredRows.filter(function (row) {
      return row.__hasValidAcc;
    });
    const categoryRows = buildCategorySummaries(validRows);
    const categoryVisitRows = buildCategoryVisitSummaries(validRows);
    const sentences = [];

    if (categoryRows.length > 1) {
      sentences.push(
        "El análisis por categoría muestra que " +
          categoryRows[0].category +
          " presenta el comportamiento menos favorable dentro del subconjunto filtrado, con una demora media de " +
          formatDaysText(categoryRows[0].mean) +
          " y un " +
          formatPercent(categoryRows[0].pct7Plus) +
          " de agendas con 7 o más días."
      );
    }

    if (categoryVisitRows.length > 1) {
      sentences.push(
        "En el cruce por categoría y tipo de visita, la combinación " +
          categoryVisitRows[0].category +
          " - " +
          categoryVisitRows[0].visitType +
          " concentra los valores más elevados según el criterio de demora prolongada y demora media."
      );
    }

    return sentences.join(" ");
  }

  function buildCategorySummaries(validRows) {
    return summarizeGroupedRows(validRows, [
      {
        key: "category",
        column: "Categoría",
        fallback: "Sin categoría"
      }
    ]).sort(compareByDelayPressure);
  }

  function buildCategoryVisitSummaries(validRows) {
    return summarizeGroupedRows(validRows, [
      {
        key: "category",
        column: "Categoría",
        fallback: "Sin categoría"
      },
      {
        key: "visitType",
        column: "Tipo visita",
        fallback: "Sin tipo de visita"
      }
    ]).sort(function (left, right) {
      const categoryOrder = left.category.localeCompare(right.category, "es", {
        sensitivity: "base",
        numeric: true
      });

      if (categoryOrder !== 0) {
        return categoryOrder;
      }

      return compareByDelayPressure(left, right);
    });
  }

  function summarizeGroupedRows(validRows, descriptors) {
    const grouped = new Map();

    validRows.forEach(function (row) {
      const labels = descriptors.map(function (descriptor) {
        return normalizeText(row[descriptor.column]) || descriptor.fallback;
      });
      const groupKey = labels.join("\u001f");

      if (!grouped.has(groupKey)) {
        const group = { rows: [] };
        descriptors.forEach(function (descriptor, index) {
          group[descriptor.key] = labels[index];
        });
        grouped.set(groupKey, group);
      }

      grouped.get(groupKey).rows.push(row);
    });

    return Array.from(grouped.values()).map(function (group) {
      const metrics = calculateExecutiveMetrics(group.rows);
      return Object.assign({}, group, metrics);
    });
  }

  function compareByDelayPressure(left, right) {
    return (
      right.pct7Plus - left.pct7Plus ||
      right.mean - left.mean ||
      right.totalValid - left.totalValid ||
      String(left.category || "").localeCompare(String(right.category || ""), "es", {
        sensitivity: "base",
        numeric: true
      }) ||
      String(left.visitType || "").localeCompare(String(right.visitType || ""), "es", {
        sensitivity: "base",
        numeric: true
      })
    );
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
    const count0to2 = values.filter(function (value) { return value >= 0 && value <= 2; }).length;
    const count0to3 = values.filter(function (value) { return value >= 0 && value <= 3; }).length;
    const count0to6 = values.filter(function (value) { return value >= 0 && value <= 6; }).length;
    const count7Plus = values.filter(function (value) { return value >= 7; }).length;

    return {
      totalValid: totalValid,
      pct0to2: (count0to2 / totalValid) * 100,
      pct0to3: (count0to3 / totalValid) * 100,
      pct0to6: (count0to6 / totalValid) * 100,
      pct7Plus: (count7Plus / totalValid) * 100,
      mean: average(values),
      median: median(values),
      max: Math.max.apply(null, values)
    };
  }

  function buildExecutiveDelayBands(rows) {
    return buildExclusiveDelayBands(rows);
  }

  function buildExclusiveDelayBands(rows) {
    const total = rows.length || 1;
    const bands = [
      { label: "0-2 días", count: 0, color: PALETTE.green },
      { label: "3 días", count: 0, color: PALETTE.teal },
      { label: "4-6 días", count: 0, color: PALETTE.amber },
      { label: "7 o más días", count: 0, color: PALETTE.red }
    ];

    rows.forEach(function (row) {
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

  function aggregateCenterBands(rows) {
    const grouped = new Map();

    rows.forEach(function (row) {
      const label = normalizeText(row["Centro"]) || "Sin dato";
      if (!grouped.has(label)) {
        grouped.set(label, {
          label: label,
          total: 0,
          band0to2: 0,
          band3: 0,
          band4to6: 0,
          band7Plus: 0
        });
      }

      const group = grouped.get(label);
      group.total += 1;
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

    return Array.from(grouped.values()).sort(function (left, right) {
      return right.band7Plus - left.band7Plus || right.total - left.total || left.label.localeCompare(right.label, "es", {
        sensitivity: "base",
        numeric: true
      });
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

  function formatReportDelay(value) {
    return formatNumber(value, 1) + " d";
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

(function () {
  "use strict";

  function cssVariable(name, fallback) {
    const value = getComputedStyle(document.documentElement).getPropertyValue(name).trim();
    return value || fallback;
  }

  function getTheme() {
    return {
      surface: cssVariable("--surface-strong", "#ffffff"),
      surfaceSoft: cssVariable("--surface-soft", "#f7fbff"),
      ink: cssVariable("--ink-900", "#17324d"),
      muted: cssVariable("--ink-500", "#6f879e"),
      line: cssVariable("--grid", "#d8e3ee"),
      accent: cssVariable("--accent", "#2876b9"),
      teal: cssVariable("--accent-alt", "#0e9d90"),
      green: cssVariable("--good", "#169b72"),
      amber: cssVariable("--warn", "#ecb343"),
      red: cssVariable("--bad", "#d55663"),
      accentSoft: cssVariable("--accent-soft", "#ddecfa")
    };
  }

  function escapeXml(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  function formatNumber(value, decimals) {
    return new Intl.NumberFormat("es-ES", {
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals
    }).format(value);
  }

  function niceMax(value) {
    if (!Number.isFinite(value) || value <= 0) {
      return 1;
    }

    const exponent = Math.floor(Math.log10(value));
    const fraction = value / Math.pow(10, exponent);
    let niceFraction = 1;

    if (fraction <= 1) {
      niceFraction = 1;
    } else if (fraction <= 2) {
      niceFraction = 2;
    } else if (fraction <= 5) {
      niceFraction = 5;
    } else {
      niceFraction = 10;
    }

    return niceFraction * Math.pow(10, exponent);
  }

  function wrapLabel(label, maxCharsPerLine, maxLines) {
    const normalized = String(label || "").trim();
    if (!normalized) {
      return [""];
    }

    const words = normalized.split(/\s+/);
    const lines = [];
    let currentLine = "";

    for (const word of words) {
      const candidate = currentLine ? currentLine + " " + word : word;
      if (candidate.length <= maxCharsPerLine || !currentLine) {
        currentLine = candidate;
      } else {
        lines.push(currentLine);
        currentLine = word;
      }
    }

    if (currentLine) {
      lines.push(currentLine);
    }

    if (lines.length <= maxLines) {
      return lines;
    }

    const truncated = lines.slice(0, maxLines);
    truncated[maxLines - 1] = truncated[maxLines - 1].slice(0, Math.max(0, maxCharsPerLine - 1)).trim() + "...";
    return truncated;
  }

  function setEmpty(container, message) {
    container.innerHTML = '<div class="chart-empty">' + escapeXml(message) + "</div>";
    container.dataset.svgMarkup = "";
  }

  function mountSvg(container, width, height, content) {
    const markup =
      '<svg class="chart-svg" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ' +
      width +
      " " +
      height +
      '" role="img" aria-hidden="true">' +
      content +
      "</svg>";

    container.innerHTML = markup;
    container.dataset.svgMarkup = markup;
  }

  function renderTextLines(x, y, lines, attrs, lineHeight) {
    const attributes = Object.entries(attrs || {})
      .map(([key, value]) => key + '="' + escapeXml(value) + '"')
      .join(" ");

    return (
      '<text ' +
      'x="' +
      x +
      '" y="' +
      y +
      '" ' +
      attributes +
      ">" +
      lines
        .map((line, index) => {
          const dy = index === 0 ? 0 : lineHeight;
          return '<tspan x="' + x + '" dy="' + dy + '">' + escapeXml(line) + "</tspan>";
        })
        .join("") +
      "</text>"
    );
  }

  function renderVerticalBarChart(container, options) {
    const items = (options.items || []).filter((item) => Number.isFinite(item.value));
    if (!items.length) {
      setEmpty(container, options.emptyMessage || "No hay datos suficientes para este grafico.");
      return;
    }

    const theme = getTheme();
    const width = Math.max(820, items.length * 86 + 110);
    const height = 360;
    const padding = { top: 28, right: 28, bottom: 94, left: 62 };
    const chartWidth = width - padding.left - padding.right;
    const chartHeight = height - padding.top - padding.bottom;
    const maxValue = niceMax(Math.max(...items.map((item) => item.value)));
    const stepCount = 4;
    const slotWidth = chartWidth / items.length;
    const barWidth = Math.max(24, Math.min(64, slotWidth * 0.58));

    let grid = '<rect x="0" y="0" width="' + width + '" height="' + height + '" rx="26" fill="' + theme.surfaceSoft + '"/>';
    for (let step = 0; step <= stepCount; step += 1) {
      const value = (maxValue / stepCount) * step;
      const y = padding.top + chartHeight - (value / maxValue) * chartHeight;
      grid +=
        '<line x1="' +
        padding.left +
        '" y1="' +
        y +
        '" x2="' +
        (width - padding.right) +
        '" y2="' +
        y +
        '" stroke="' +
        theme.line +
        '" stroke-width="1"/>' +
        '<text x="' +
        (padding.left - 10) +
        '" y="' +
        (y + 4) +
        '" fill="' +
        theme.muted +
        '" font-size="12" text-anchor="end">' +
        escapeXml(options.valueFormatter ? options.valueFormatter(value) : formatNumber(value, 0)) +
        "</text>";
    }

    const bars = items
      .map((item, index) => {
        const value = item.value;
        const barHeight = (value / maxValue) * chartHeight;
        const x = padding.left + index * slotWidth + (slotWidth - barWidth) / 2;
        const y = padding.top + chartHeight - barHeight;
        const lines = wrapLabel(item.label, 12, 2);
        const annotation = options.annotationFormatter ? options.annotationFormatter(item) : null;
        const annotationY = Math.max(padding.top + 12, y - 12);
        const barColor = item.color || options.barColor || theme.accent;

        return (
          '<g>' +
          '<rect x="' +
          x +
          '" y="' +
          y +
          '" width="' +
          barWidth +
          '" height="' +
          Math.max(barHeight, 3) +
          '" rx="14" fill="' +
          barColor +
          '" opacity="0.95">' +
          "<title>" +
          escapeXml(item.label + ": " + (options.titleFormatter ? options.titleFormatter(item) : item.value)) +
          "</title></rect>" +
          (annotation
            ? '<text x="' +
              (x + barWidth / 2) +
              '" y="' +
              annotationY +
              '" text-anchor="middle" fill="' +
              theme.ink +
              '" font-size="12" font-weight="700">' +
              escapeXml(annotation) +
              "</text>"
            : "") +
          renderTextLines(
            x + barWidth / 2,
            height - 48,
            lines,
            {
              fill: theme.ink,
              "font-size": 12,
              "font-weight": 600,
              "text-anchor": "middle"
            },
            15
          ) +
          "</g>"
        );
      })
      .join("");

    mountSvg(container, width, height, grid + bars);
  }

  function renderHorizontalBarChart(container, options) {
    const items = (options.items || []).filter((item) => Number.isFinite(item.value));
    if (!items.length) {
      setEmpty(container, options.emptyMessage || "No hay datos suficientes para este grafico.");
      return;
    }

    const theme = getTheme();
    const rowHeight = 56;
    const width = 920;
    const height = Math.max(320, 96 + items.length * rowHeight);
    const padding = { top: 34, right: 76, bottom: 28, left: 280 };
    const chartWidth = width - padding.left - padding.right;
    const maxValue = niceMax(Math.max(...items.map((item) => item.value)));
    const stepCount = 4;

    let grid = '<rect x="0" y="0" width="' + width + '" height="' + height + '" rx="26" fill="' + theme.surfaceSoft + '"/>';
    for (let step = 0; step <= stepCount; step += 1) {
      const value = (maxValue / stepCount) * step;
      const x = padding.left + (value / maxValue) * chartWidth;
      grid +=
        '<line x1="' +
        x +
        '" y1="' +
        padding.top +
        '" x2="' +
        x +
        '" y2="' +
        (height - padding.bottom) +
        '" stroke="' +
        theme.line +
        '" stroke-width="1"/>' +
        '<text x="' +
        x +
        '" y="' +
        (height - 8) +
        '" fill="' +
        theme.muted +
        '" font-size="12" text-anchor="middle">' +
        escapeXml(options.valueFormatter ? options.valueFormatter(value) : formatNumber(value, 0)) +
        "</text>";
    }

    const bars = items
      .map((item, index) => {
        const y = padding.top + index * rowHeight + 2;
        const barWidth = (item.value / maxValue) * chartWidth;
        const labelLines = wrapLabel(item.label, 28, 2);
        const noteText = item.note ? '<text x="' + (padding.left - 16) + '" y="' + (y + 31) + '" text-anchor="end" fill="' + theme.muted + '" font-size="11">' + escapeXml(item.note) + "</text>" : "";
        const color = item.color || options.barColor || theme.accent;

        return (
          '<g>' +
          renderTextLines(
            padding.left - 16,
            y + 16,
            labelLines,
            {
              fill: theme.ink,
              "font-size": 13,
              "font-weight": 600,
              "text-anchor": "end"
            },
            14
          ) +
          noteText +
          '<rect x="' +
          padding.left +
          '" y="' +
          (y + 6) +
          '" width="' +
          chartWidth +
          '" height="16" rx="9" fill="' +
          theme.accentSoft +
          '" opacity="0.55"/>' +
          '<rect x="' +
          padding.left +
          '" y="' +
          (y + 6) +
          '" width="' +
          Math.max(barWidth, 3) +
          '" height="16" rx="9" fill="' +
          color +
          '">' +
          "<title>" +
          escapeXml(item.label + ": " + (options.titleFormatter ? options.titleFormatter(item) : item.value)) +
          "</title></rect>" +
          '<text x="' +
          (padding.left + Math.max(barWidth + 10, 12)) +
          '" y="' +
          (y + 19) +
          '" fill="' +
          theme.ink +
          '" font-size="12" font-weight="700">' +
          escapeXml(options.valueFormatter ? options.valueFormatter(item.value) : formatNumber(item.value, 1)) +
          "</text>" +
          "</g>"
        );
      })
      .join("");

    mountSvg(container, width, height, grid + bars);
  }

  function describeArc(cx, cy, radius, startAngle, endAngle) {
    const start = polarToCartesian(cx, cy, radius, endAngle);
    const end = polarToCartesian(cx, cy, radius, startAngle);
    const largeArcFlag = endAngle - startAngle <= 180 ? "0" : "1";

    return [
      "M",
      start.x,
      start.y,
      "A",
      radius,
      radius,
      0,
      largeArcFlag,
      0,
      end.x,
      end.y
    ].join(" ");
  }

  function polarToCartesian(cx, cy, radius, angleInDegrees) {
    const angleInRadians = ((angleInDegrees - 90) * Math.PI) / 180;
    return {
      x: cx + radius * Math.cos(angleInRadians),
      y: cy + radius * Math.sin(angleInRadians)
    };
  }

  function renderDonutChart(container, options) {
    const segments = (options.segments || []).filter((segment) => Number.isFinite(segment.value) && segment.value > 0);
    if (!segments.length) {
      setEmpty(container, options.emptyMessage || "No hay datos suficientes para este grafico.");
      return;
    }

    const theme = getTheme();
    const total = segments.reduce((sum, segment) => sum + segment.value, 0);
    const width = 900;
    const height = 360;
    const cx = 250;
    const cy = 180;
    const radius = 108;
    const strokeWidth = 42;
    let currentAngle = 0;

    let content = '<rect x="0" y="0" width="' + width + '" height="' + height + '" rx="26" fill="' + theme.surfaceSoft + '"/>';
    content += '<circle cx="' + cx + '" cy="' + cy + '" r="' + radius + '" fill="none" stroke="' + theme.accentSoft + '" stroke-width="' + strokeWidth + '" opacity="0.55"/>';

    segments.forEach((segment) => {
      const sweep = (segment.value / total) * 360;
      const endAngle = currentAngle + sweep;
      const segmentColor = segment.color || theme.accent;
      const title =
        "<title>" +
        escapeXml(
          segment.label +
            ": " +
            formatNumber((segment.value / total) * 100, 1) +
            "% (" +
            formatNumber(segment.value, 0) +
            ")"
        ) +
        "</title>";

      if (sweep >= 359.9) {
        content +=
          '<circle cx="' +
          cx +
          '" cy="' +
          cy +
          '" r="' +
          radius +
          '" fill="none" stroke="' +
          segmentColor +
          '" stroke-width="' +
          strokeWidth +
          '">' +
          title +
          "</circle>";
      } else {
        const path = describeArc(cx, cy, radius, currentAngle, endAngle);
        content +=
          '<path d="' +
          path +
          '" fill="none" stroke="' +
          segmentColor +
          '" stroke-width="' +
          strokeWidth +
          '" stroke-linecap="round">' +
          title +
          "</path>";
      }

      currentAngle = endAngle;
    });

    content +=
      '<circle cx="' +
      cx +
      '" cy="' +
      cy +
      '" r="62" fill="' +
      theme.surface +
      '"/>' +
      '<text x="' +
      cx +
      '" y="' +
      (cy - 12) +
      '" text-anchor="middle" fill="' +
      theme.muted +
      '" font-size="14" font-weight="700">' +
      escapeXml(options.centerLabel || "Agendas") +
      "</text>" +
      '<text x="' +
      cx +
      '" y="' +
      (cy + 18) +
      '" text-anchor="middle" fill="' +
      theme.ink +
      '" font-size="34" font-weight="700">' +
      escapeXml(options.centerValue || formatNumber(total, 0)) +
      "</text>";

    const legendX = 510;
    segments.forEach((segment, index) => {
      const y = 90 + index * 54;
      const percentage = total ? (segment.value / total) * 100 : 0;
      content +=
        '<g>' +
        '<rect x="' +
        legendX +
        '" y="' +
        y +
        '" width="16" height="16" rx="6" fill="' +
        (segment.color || theme.accent) +
        '"/>' +
        '<text x="' +
        (legendX + 28) +
        '" y="' +
        (y + 13) +
        '" fill="' +
        theme.ink +
        '" font-size="14" font-weight="700">' +
        escapeXml(segment.label) +
        "</text>" +
        '<text x="' +
        (legendX + 28) +
        '" y="' +
        (y + 34) +
        '" fill="' +
        theme.muted +
        '" font-size="12">' +
        escapeXml(formatNumber(percentage, 1) + "% · " + formatNumber(segment.value, 0) + " agendas") +
        "</text>" +
        "</g>";
    });

    mountSvg(container, width, height, content);
  }

  function renderStackedBarChart(container, options) {
    const items = (options.items || []).filter((item) => Number.isFinite(item.total) && item.total > 0);
    if (!items.length) {
      setEmpty(container, options.emptyMessage || "No hay datos suficientes para este grafico.");
      return;
    }

    const theme = getTheme();
    const width = 960;
    const rowHeight = 58;
    const height = Math.max(360, 132 + items.length * rowHeight);
    const padding = { top: 68, right: 72, bottom: 34, left: 270 };
    const chartWidth = width - padding.left - padding.right;
    const maxTotal = niceMax(Math.max(...items.map((item) => item.total)));
    const legends = options.legend || [];

    let content = '<rect x="0" y="0" width="' + width + '" height="' + height + '" rx="26" fill="' + theme.surfaceSoft + '"/>';

    legends.forEach((legend, index) => {
      const x = padding.left + index * 138;
      content +=
        '<rect x="' +
        x +
        '" y="22" width="18" height="18" rx="7" fill="' +
        legend.color +
        '"/>' +
        '<text x="' +
        (x + 26) +
        '" y="36" fill="' +
        theme.ink +
        '" font-size="13" font-weight="700">' +
        escapeXml(legend.label) +
        "</text>";
    });

    for (let step = 0; step <= 4; step += 1) {
      const value = (maxTotal / 4) * step;
      const x = padding.left + (value / maxTotal) * chartWidth;
      content +=
        '<line x1="' +
        x +
        '" y1="' +
        padding.top +
        '" x2="' +
        x +
        '" y2="' +
        (height - padding.bottom) +
        '" stroke="' +
        theme.line +
        '" stroke-width="1"/>' +
        '<text x="' +
        x +
        '" y="' +
        (height - 10) +
        '" fill="' +
        theme.muted +
        '" font-size="12" text-anchor="middle">' +
        escapeXml(formatNumber(value, 0)) +
        "</text>";
    }

    items.forEach((item, index) => {
      const y = padding.top + index * rowHeight + 6;
      const labelLines = wrapLabel(item.label, 28, 2);
      let currentX = padding.left;

      content += renderTextLines(
        padding.left - 16,
        y + 12,
        labelLines,
        {
          fill: theme.ink,
          "font-size": 13,
          "font-weight": 600,
          "text-anchor": "end"
        },
        14
      );

      content +=
        '<rect x="' +
        padding.left +
        '" y="' +
        y +
        '" width="' +
        chartWidth +
        '" height="18" rx="9" fill="' +
        theme.accentSoft +
        '" opacity="0.45"/>';

      item.segments.forEach((segment) => {
        const widthValue = (segment.value / maxTotal) * chartWidth;
        if (segment.value <= 0) {
          return;
        }

        content +=
          '<rect x="' +
          currentX +
          '" y="' +
          y +
          '" width="' +
          Math.max(widthValue, 2) +
          '" height="18" rx="7" fill="' +
          segment.color +
          '">' +
          "<title>" +
          escapeXml(segment.label + ": " + formatNumber(segment.value, 0)) +
          "</title></rect>";
        currentX += widthValue;
      });

      content +=
        '<text x="' +
        (padding.left + ((item.total / maxTotal) * chartWidth) + 10) +
        '" y="' +
        (y + 14) +
        '" fill="' +
        theme.ink +
        '" font-size="12" font-weight="700">' +
        escapeXml(formatNumber(item.total, 0)) +
        "</text>";
    });

    mountSvg(container, width, height, content);
  }

  async function exportContainerAsPng(container, filename) {
    const svgMarkup = container && container.dataset ? container.dataset.svgMarkup : "";
    if (!svgMarkup) {
      throw new Error("No hay un grafico disponible para exportar.");
    }

    const svgElement = container.querySelector("svg");
    if (!svgElement || !svgElement.viewBox || !svgElement.viewBox.baseVal) {
      throw new Error("No se ha podido preparar el grafico para exportacion.");
    }

    const viewBox = svgElement.viewBox.baseVal;
    const scale = 2;
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(viewBox.width * scale));
    canvas.height = Math.max(1, Math.round(viewBox.height * scale));

    const context = canvas.getContext("2d");
    context.fillStyle = cssVariable("--surface-strong", "#ffffff");
    context.fillRect(0, 0, canvas.width, canvas.height);

    const blob = new Blob([svgMarkup], { type: "image/svg+xml;charset=utf-8" });
    const url = URL.createObjectURL(blob);

    await new Promise((resolve, reject) => {
      const image = new Image();
      image.onload = () => {
        context.drawImage(image, 0, 0, canvas.width, canvas.height);
        URL.revokeObjectURL(url);
        resolve();
      };
      image.onerror = () => {
        URL.revokeObjectURL(url);
        reject(new Error("No se ha podido convertir el grafico a imagen PNG."));
      };
      image.src = url;
    });

    const pngBlob = await new Promise((resolve) => canvas.toBlob(resolve, "image/png"));
    if (!pngBlob) {
      throw new Error("No se ha podido generar el archivo PNG.");
    }

    const downloadUrl = URL.createObjectURL(pngBlob);
    const link = document.createElement("a");
    link.href = downloadUrl;
    link.download = filename || "grafico.png";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(downloadUrl);
  }

  window.DemorasCharts = {
    renderVerticalBarChart,
    renderHorizontalBarChart,
    renderDonutChart,
    renderStackedBarChart,
    setEmpty,
    exportContainerAsPng
  };
})();

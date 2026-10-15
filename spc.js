// Simple SPC logic + wiring for SPC charts

let rawRows = [];
let currentChart = null;   // main I / run chart
let mrChart = null;        // moving range chart
let annotations = [];      // { date: 'YYYY-MM-DD', label: 'text', yAdjust?: number }
let splits = [];   // indices where a new XmR segment starts (split AFTER index)

// Observations/subgroups deliberately left visible but excluded from
// centre-line and control-limit estimation. Stored as { scope, key }.
// P/P′ share a scope, and U/U′ share a scope, so switching between
// the standard and Laney chart preserves the user's explicit decision.
let limitExclusions = [];

const LIMIT_EXCLUSION_SUPPORTED_CHARTS = new Set([
  "xmr", "c", "p", "pprime", "u", "uprime", "xbars"
]);

function getLimitExclusionScope(chartType) {
  if (chartType === "p" || chartType === "pprime") return "p";
  if (chartType === "u" || chartType === "uprime") return "u";
  return chartType;
}

function normaliseLimitExclusions(value) {
  if (!Array.isArray(value)) return [];

  const seen = new Set();
  const out = [];

  value.forEach(item => {
    if (!item || typeof item !== "object") return;

    const scope = String(item.scope || "").trim();
    const key = String(item.key || "").trim();
    if (!scope || !key) return;

    const token = `${scope}::${key}`;
    if (seen.has(token)) return;

    seen.add(token);
    out.push({ scope, key });
  });

  return out;
}

function makeRowLimitExclusionKey(point, fallbackIndex) {
  const rowIndex = Number(point?._rowIndex);

  if (Number.isInteger(rowIndex) && rowIndex >= 0) {
    return `row:${rowIndex}`;
  }

  return `index:${fallbackIndex}`;
}

function makeSubgroupLimitExclusionKey(subgroupLabel) {
  return `subgroup:${String(subgroupLabel ?? "")}`;
}

function isLimitExcluded(scope, key) {
  return limitExclusions.some(
    item =>
      item.scope === scope &&
      item.key === key
  );
}

function setLimitExcluded(scope, key, excluded) {
  if (!scope || !key) return;

  limitExclusions =
    normaliseLimitExclusions(
      limitExclusions
    );

  const idx =
    limitExclusions.findIndex(
      item =>
        item.scope === scope &&
        item.key === key
    );

  if (excluded) {
    if (idx === -1) {
      limitExclusions.push({
        scope,
        key
      });
    }

  } else if (idx !== -1) {
    limitExclusions.splice(
      idx,
      1
    );
  }
}

function markLimitExclusionsOnPoints(
  chartType,
  points
) {
  const scope =
    getLimitExclusionScope(
      chartType
    );

  return (points || []).map(
    (point, index) => {
      const key =
        makeRowLimitExclusionKey(
          point,
          index
        );

      return {
        ...point,
        _limitExclusionKey:
          key,

        _excludeFromLimits:
          isLimitExcluded(
            scope,
            key
          )
      };
    }
  );
}

function getMarkedExcludedIndices(
  points
) {
  const out = [];

  (points || []).forEach(
    (point, index) => {
      if (
        point?._excludeFromLimits
      ) {
        out.push(index);
      }
    }
  );

  return out;
}

function buildRowLimitExclusionKeys(
  points
) {
  return (points || []).map(
    (point, index) =>
      point?._limitExclusionKey ||
      makeRowLimitExclusionKey(
        point,
        index
      )
  );
}

function getCurrentLimitEstimationWindow(
  pointIndex
) {
  const labels =
    currentChart?.data?.labels || [];

  const n =
    labels.length;

  if (
    !Number.isInteger(pointIndex) ||
    pointIndex < 0 ||
    pointIndex >= n
  ) {
    return {
      eligible: false,
      reason:
        "Right-click near a plotted point."
    };
  }

  const effectiveSplits =
    (
      Array.isArray(splits)
        ? splits
        : []
    )
      .filter(
        i =>
          Number.isInteger(i) &&
          i >= 0 &&
          i < n - 1
      )
      .sort(
        (a, b) => a - b
      );

  let segmentStart = 0;
  let segmentEnd =
    n - 1;

  let segmentIndex = 0;

  for (
    let i = 0;
    i < effectiveSplits.length;
    i++
  ) {
    const splitAfter =
      effectiveSplits[i];

    if (
      pointIndex <=
      splitAfter
    ) {
      segmentEnd =
        splitAfter;

      break;
    }

    segmentStart =
      splitAfter + 1;

    segmentIndex =
      i + 1;
  }

  // Later split periods use the whole period
  // to establish their new limits.
  if (segmentIndex > 0) {
    return {
      eligible: true,
      segmentStart,
      segmentEnd,
      segmentIndex
    };
  }

  const baseline =
    Number(
      baselineInput?.value
    );

  if (
    Number.isFinite(
      baseline
    ) &&
    baseline >= 1
  ) {
    const baselineEnd =
      Math.min(
        segmentEnd,
        segmentStart +
        Math.floor(baseline) -
        1
      );

    if (
      pointIndex >
      baselineEnd
    ) {
      return {
        eligible: false,
        segmentStart,
        segmentEnd,
        segmentIndex,

        reason:
          "This point is outside the baseline used to estimate the current limits, so excluding it would not change the limit calculation."
      };
    }
  }

  return {
    eligible: true,
    segmentStart,
    segmentEnd,
    segmentIndex
  };
}

function getCurrentLimitExclusionMenuState(
  pointIndex
) {
  const chartType =
    (
      typeof
        getSelectedChartType_NoSideEffects ===
        "function"
    )
      ? (
          getSelectedChartType_NoSideEffects() ||
          "run"
        )
      : "run";

  const supported =
    LIMIT_EXCLUSION_SUPPORTED_CHARTS.has(
      chartType
    );

  const scope =
    getLimitExclusionScope(
      chartType
    );

  const key =
    currentChart
      ?.$limitExclusionKeys
      ?.[pointIndex] ??
    null;

  const windowInfo =
    getCurrentLimitEstimationWindow(
      pointIndex
    );

  const excluded =
    !!key &&
    isLimitExcluded(
      scope,
      key
    );

  let eligible =
    !!key &&
    !!windowInfo.eligible;

  let reason =
    windowInfo.reason || "";

  // Prevent a new exclusion from making
  // the limit estimate mathematically unusable.
  if (
    supported &&
    key &&
    eligible &&
    !excluded
  ) {
    const keys =
      currentChart
        ?.$limitExclusionKeys ||
      [];

    const start =
      Number.isInteger(
        windowInfo.segmentStart
      )
        ? windowInfo.segmentStart
        : 0;

    let end =
      Number.isInteger(
        windowInfo.segmentEnd
      )
        ? windowInfo.segmentEnd
        : Math.max(
            0,
            keys.length - 1
          );

    if (
      windowInfo.segmentIndex ===
      0
    ) {
      const baseline =
        Number(
          baselineInput?.value
        );

      if (
        Number.isFinite(
          baseline
        ) &&
        baseline >= 1
      ) {
        end =
          Math.min(
            end,
            start +
            Math.floor(
              baseline
            ) -
            1
          );
      }
    }

    const includedFlags = [];

    for (
      let i = start;
      i <= end;
      i++
    ) {
      const k =
        keys[i];

      const wouldBeExcluded =
        i === pointIndex ||
        (
          !!k &&
          isLimitExcluded(
            scope,
            k
          )
        );

      includedFlags.push(
        !wouldBeExcluded
      );
    }

    const includedCount =
      includedFlags
        .filter(Boolean)
        .length;

    const minimumIncluded =
      [
        "xmr",
        "pprime",
        "uprime",
        "xbars"
      ].includes(chartType)
        ? 2
        : 1;

    if (
      includedCount <
      minimumIncluded
    ) {
      eligible = false;

      reason =
        `Excluding this ${
          chartType === "xbars"
            ? "subgroup"
            : "point"
        } would leave too few included observations to estimate the control limits.`;
    }

    // XmR/P′/U′ need at least one pair
    // of consecutive included observations
    // for a moving-range variation estimate.
    if (
      eligible &&
      [
        "xmr",
        "pprime",
        "uprime"
      ].includes(chartType)
    ) {
      let hasAdjacentPair =
        false;

      for (
        let i = 1;
        i <
          includedFlags.length;
        i++
      ) {
        if (
          includedFlags[i - 1] &&
          includedFlags[i]
        ) {
          hasAdjacentPair =
            true;

          break;
        }
      }

      if (!hasAdjacentPair) {
        eligible = false;

        reason =
          "Excluding this point would leave no pair of consecutive included observations for the moving-range variation estimate.";
      }
    }
  }

  return {
    chartType,
    supported,
    scope,
    key,
    excluded,
    eligible,
    reason
  };
}

function attachLimitExclusionContext(
  chart,
  chartType,
  keys,
  excludedIndices
) {
  if (!chart) return;

  /*
   * Store exclusion metadata directly on the Chart instance.
   *
   * IMPORTANT:
   * Do not modify chart.options here. Once Chart.js has created a chart,
   * chart.options is backed by internal resolver/proxy objects. Mutating
   * those objects from here can cause recursive proxy updates and can also
   * interfere with the annotation drag plugin.
   */
  chart.$limitExclusionChartType =
    chartType;

  chart.$limitExclusionScope =
    getLimitExclusionScope(
      chartType
    );

  chart.$limitExclusionKeys =
    Array.isArray(keys)
      ? keys.slice()
      : [];

  chart.$limitExcludedIndices =
    Array.isArray(
      excludedIndices
    )
      ? excludedIndices.slice()
      : [];

  /*
   * The exclusion marker plugin reads the metadata above.
   * draw() is sufficient to paint the marker and does not re-resolve
   * the whole Chart.js options tree.
   */
  if (
    typeof chart.draw ===
    "function"
  ) {
    chart.draw();
  }
}

function attachRowLimitExclusionContext(
  chart,
  chartType,
  points
) {
  const keys =
    buildRowLimitExclusionKeys(
      points
    );

  const excluded =
    getMarkedExcludedIndices(
      points
    );

  attachLimitExclusionContext(
    chart,
    chartType,
    keys,
    excluded
  );
}

function attachSubgroupLimitExclusionContext(
  chart,
  subgroupLabels
) {
  const scope =
    getLimitExclusionScope(
      "xbars"
    );

  const keys =
    (subgroupLabels || [])
      .map(
        makeSubgroupLimitExclusionKey
      );

  const excluded = [];

  keys.forEach(
    (key, index) => {
      if (
        isLimitExcluded(
          scope,
          key
        )
      ) {
        excluded.push(
          index
        );
      }
    }
  );

  attachLimitExclusionContext(
    chart,
    "xbars",
    keys,
    excluded
  );
}

function appendLimitExclusionSummary() {
  if (
    !summaryDiv ||
    !currentChart
  ) {
    return;
  }

  const excluded =
    currentChart
      .$limitExcludedIndices ||
    [];

  if (!excluded.length) {
    return;
  }

  const isSubgroup =
    currentChart
      .$limitExclusionChartType ===
    "xbars";

  const unit =
    isSubgroup
      ? "subgroup"
      : "point";

  const count =
    excluded.length;

  const note =
    document.createElement(
      "p"
    );

  note.className =
    "limit-exclusion-summary";

  note.innerHTML =
    `<strong>Excluded from limit calculations:</strong> ` +
    `${count} ${unit}${count === 1 ? "" : "s"}. ` +
    `${
      count === 1
        ? `The excluded ${unit} remains visible on the chart.`
        : `Excluded ${unit}s remain visible on the chart.`
    }`;

  summaryDiv.appendChild(
    note
  );
}
let lastXmRAnalysis = null;
let lastRunAnalysis = null;
let lastParetoAnalysis = null;
let dataModelDirty = false;
let gridHeaders = ["Date", "Value"];
let lastGridHeadersKey = ""; // track changes

let chartTitleManuallyEdited = false;
let xAxisLabelManuallyEdited = false;
let yAxisLabelManuallyEdited = false;
let yAxisBoundsManuallyEdited = false;

// ------------------------------------------------------------
// Column intelligence (Levels 1–3):
// - Profile columns once per dataset
// - Use profiles for filtering dropdowns + smart default selection
// ------------------------------------------------------------
let allColumns = [];
let columnProfiles = {}; // { [colName]: { ...stats } }



const fileInput         = document.getElementById("fileInput");
const columnSelectors   = document.getElementById("columnSelectors");
const dateSelect        = document.getElementById("dateColumn");
const valueSelect       = document.getElementById("valueColumn");
const dateFormatPreferenceSelect = document.getElementById("dateFormatPreference");
const dateFormatWarning = document.getElementById("dateFormatWarning");

// Settings import/export buttons (Section 1: Data)
const exportSettingsBtn = document.getElementById("exportSettingsBtn");
const importSettingsBtn = document.getElementById("importSettingsBtn");
const importSettingsFileInput = document.getElementById("importSettingsFileInput");

function updateSaveChartButtonState() {
  if (!exportSettingsBtn) return;
  exportSettingsBtn.disabled = !currentChart;
}

// If settings are imported before data is loaded, we store them here and apply after loadRows()
let pendingImportedSettings = null;



const IMPLEMENTED_CHARTS =
  new Set([
    "run",
    "xmr",
    "c",
    "p",
    "pprime",
    "u",
    "uprime",
    "xbars",
    "t",
    "g",
    "pareto"
  ]);

const PARETO_COUNT_ROWS =
  "__pareto_count_rows__";

function isPFamilyChartType(chartType) {
  return chartType === "p" ||
         chartType === "pprime";
}

function isUFamilyChartType(chartType) {
  return chartType === "u" ||
         chartType === "uprime";
}

// -----------------------------
// Shared chart styling (keep charts consistent)
// -----------------------------
const SPC_STYLE = {
  seriesBlue: "#003f87",
  pointNormal: "#003f87",
  pointSpecial: "#ff8c00",
  pointBeyond: "#d73027",
  centreRed: "#e41a1c",
  limitGreen: "#2ca25f",
  targetOrange: "#fdae61"
};

const SPC_STYLE_DEFAULT = {
  seriesBlue: "#003f87",
  pointNormal: "#003f87",
  pointSpecial: "#ff8c00",
  pointBeyond: "#d73027",
  centreRed: "#e41a1c",
  limitGreen: "#2ca25f",
  targetOrange: "#fdae61"
};

const SPC_STYLE_COLOUR_BLIND = {
  seriesBlue: "#0072B2",
  pointNormal: "#0072B2",
  pointSpecial: "#D55E00",
  pointBeyond: "#000000",
  centreRed: "#000000",
  limitGreen: "#009E73",
  targetOrange: "#CC79A7"
};

function applySpcColourTheme(themeName) {
  const theme = themeName === "colourBlind"
    ? SPC_STYLE_COLOUR_BLIND
    : SPC_STYLE_DEFAULT;

  Object.assign(SPC_STYLE, theme);
}

// -----------------------------
// Shared legend styling
// Shows line datasets as line samples in the legend instead of filled boxes
// -----------------------------
const SPC_LEGEND = {
  display: true,
  position: "bottom",
  align: "center",

  labels: {
    usePointStyle: true,
    pointStyle: "line",
    boxWidth: 55,
    boxHeight: 12,
    padding: 14,

    generateLabels(chart) {
      const defaultLabels =
        Chart.defaults.plugins.legend.labels.generateLabels(
          chart
        );

      return defaultLabels.map(
        label => {
          const dataset =
            chart.data.datasets[
              label.datasetIndex
            ] || {};

          return {
            ...label,

            pointStyle:
              "line",

            strokeStyle:
              dataset.borderColor ||
              label.strokeStyle,

            fillStyle:
              "transparent",

            lineWidth:
              Math.max(
                dataset.borderWidth ||
                2,
                3
              ),

            lineDash:
              dataset.borderDash ||
              [],

            lineCap:
              "round"
          };
        }
      );
    }
  }
};

// Helper: build point colours from a boolean “flag” array
function makePointColoursFromFlags(flags) {
  if (!Array.isArray(flags)) return [];
  return flags.map(f => (f ? SPC_STYLE.pointSpecial : SPC_STYLE.pointNormal));
}

// -----------------------------
// Baseline overlay (shade + boundary line)
// Draws a light band behind the first N baseline points and a vertical line where baseline ends.
// Safe: does not mutate chart options; it only draws to the canvas.
// -----------------------------
const baselineOverlayPlugin = {
  id: "baselineOverlay",
  beforeDatasetsDraw(chart, args, pluginOptions) {
    const opts = pluginOptions || {};
    if (opts.enabled === false) return;

    const baselineEl = document.getElementById("baselinePoints");
    const n = baselineEl && baselineEl.value !== "" ? Number(baselineEl.value) : NaN;
    if (!Number.isFinite(n) || n < 2) return;

    const labels = chart?.data?.labels || [];
    if (!Array.isArray(labels) || labels.length < 2) return;

    const count = Math.min(Math.floor(n), labels.length);
    if (count < 2) return;

    const xScale = chart.scales?.x;
    if (!xScale) return;

    const { ctx, chartArea } = chart;
    if (!ctx || !chartArea) return;

    const x0 = xScale.getPixelForValue(0);

    // End boundary: halfway between last baseline point and next point (if it exists),
    // otherwise to the end of the chart area.
    const lastIdx = count - 1;
    const xLast = xScale.getPixelForValue(lastIdx);
    let xBoundary = chartArea.right;

    if (count < labels.length) {
      const xNext = xScale.getPixelForValue(count);
      xBoundary = (xLast + xNext) / 2;
    } else {
      xBoundary = chartArea.right;
    }

    // Clamp to chart area
    const left = Math.max(chartArea.left, Math.min(x0, xBoundary));
    const right = Math.min(chartArea.right, Math.max(x0, xBoundary));
    if (!(right > left)) return;

    ctx.save();

    // Shade
    ctx.fillStyle = opts.fillStyle || "rgba(120, 120, 120, 0.10)";
    ctx.fillRect(left, chartArea.top, right - left, chartArea.bottom - chartArea.top);

    // Boundary line
    const lineX = Math.max(chartArea.left, Math.min(chartArea.right, xBoundary));
    ctx.strokeStyle = opts.lineStyle || "rgba(80, 80, 80, 0.55)";
    ctx.lineWidth = opts.lineWidth || 1;
    ctx.setLineDash(opts.lineDash || [4, 4]);
    ctx.beginPath();
    ctx.moveTo(lineX, chartArea.top);
    ctx.lineTo(lineX, chartArea.bottom);
    ctx.stroke();

    ctx.restore();
  }
};

const limitExclusionMarkerPlugin = {
  id: "limitExclusionMarkers",

  afterDatasetsDraw(chart) {
    const indices =
      Array.isArray(
        chart
          .$limitExcludedIndices
      )
        ? chart
            .$limitExcludedIndices
        : [];

    if (!indices.length) {
      return;
    }

    const meta =
      chart.getDatasetMeta?.(0);

    if (
      !meta ||
      !Array.isArray(
        meta.data
      )
    ) {
      return;
    }

    const ctx =
      chart.ctx;

    ctx.save();

    ctx.lineWidth = 2;
    ctx.strokeStyle =
      "#333333";

    ctx.fillStyle =
      "#ffffff";

    indices.forEach(
      index => {
        const element =
          meta.data[index];

        if (
          !element ||
          !Number.isFinite(
            element.x
          ) ||
          !Number.isFinite(
            element.y
          )
        ) {
          return;
        }

        const x =
          element.x;

        const y =
          element.y;

        const r = 7;

        // Hollow circle
        ctx.beginPath();
        ctx.arc(
          x,
          y,
          r,
          0,
          Math.PI * 2
        );
        ctx.fill();
        ctx.stroke();

        // X through the circle
        ctx.beginPath();

        ctx.moveTo(
          x - 3.5,
          y - 3.5
        );

        ctx.lineTo(
          x + 3.5,
          y + 3.5
        );

        ctx.moveTo(
          x + 3.5,
          y - 3.5
        );

        ctx.lineTo(
          x - 3.5,
          y + 3.5
        );

        ctx.stroke();
      }
    );

    ctx.restore();
  }
};


// Register once (Chart.js v3/v4)
if (typeof Chart !== "undefined" && Chart.register) {
  Chart.register(
  baselineOverlayPlugin,
  limitExclusionMarkerPlugin
);

  // Give charts a little more headroom/footroom so annotation labels stay visible
  Chart.defaults.layout = Chart.defaults.layout || {};
  Chart.defaults.layout.padding = {
    top: 26,
    right: 8,
    bottom: 30,
    left: 8
  };
}


// Dynamic column labels + optional 3rd selector
const xLabelEl = document.getElementById("xLabel");
const yLabelEl = document.getElementById("yLabel");

const thirdColumnRow = document.getElementById("thirdColumnRow");
const thirdLabelEl = document.getElementById("thirdLabel");
const thirdHintEl = document.getElementById("thirdHint");
const thirdSelect = document.getElementById("thirdColumn");
const uRateMultiplierRow = document.getElementById("uRateMultiplierRow");
const uRateMultiplierSelect = document.getElementById("uRateMultiplier");

let lastURateMultiplier = 1;

function getURateMultiplier() {
  const value = Number(uRateMultiplierSelect?.value || 1);

  return [1, 10, 100, 1000, 10000].includes(value)
    ? value
    : 1;
}

function formatURateMultiplier(multiplier) {
  return Number(multiplier).toLocaleString("en-GB");
}

function rescaleNumericInput(input, ratio) {
  if (!input) return;

  const text = String(input.value || "").trim();
  if (text === "") return;

  const value = Number(text);
  if (!Number.isFinite(value)) return;

  input.value = String(value * ratio);
}

if (uRateMultiplierSelect) {
  lastURateMultiplier = getURateMultiplier();

  uRateMultiplierSelect.addEventListener("change", () => {
    const newMultiplier = getURateMultiplier();
    const oldMultiplier = lastURateMultiplier || 1;

    if (newMultiplier !== oldMultiplier) {
      const ratio = newMultiplier / oldMultiplier;

      // These controls use the same visible units as the chart.
      // Keep their meaning unchanged when the display scale changes.
      rescaleNumericInput(targetInput, ratio);
      rescaleNumericInput(yAxisMinInput, ratio);
      rescaleNumericInput(yAxisMaxInput, ratio);
      rescaleNumericInput(yAxisTickStepInput, ratio);
    }

    lastURateMultiplier = newMultiplier;

    if (rawRows && rawRows.length && generateButton) {
      generateButton.click();
    }
  });
}


// Chart chooser / extra columns
const helpChooseChartBtn   = document.getElementById("helpChooseChartBtn");
const chartSetupBtn        = document.getElementById("chartSetupBtn");

const newDataChartPrompt =
  document.getElementById("newDataChartPrompt");

const newDataHelpChooseBtn =
  document.getElementById("newDataHelpChooseBtn");

const newDataChooseManuallyBtn =
  document.getElementById("newDataChooseManuallyBtn");

const extraColumnsWrap     = document.getElementById("extraColumns");
const extraColumns_PU      = document.getElementById("extraColumns_PU");
const extraColumns_XbarS   = document.getElementById("extraColumns_XbarS");
const extraColumns_T       = document.getElementById("extraColumns_T");
const extraColumns_G       = document.getElementById("extraColumns_G");

const numeratorSelect      = document.getElementById("numeratorColumn");
const denominatorSelect    = document.getElementById("denominatorColumn");
const subgroupSelect       = document.getElementById("subgroupColumn");
const eventDateSelect      = document.getElementById("eventDateColumn");
const oppBetweenSelect     = document.getElementById("oppBetweenColumn");



const baselineInput     = document.getElementById("baselinePoints");
const chartTitleInput   = document.getElementById("chartTitle");
const xAxisLabelInput   = document.getElementById("xAxisLabel");
const yAxisLabelInput   = document.getElementById("yAxisLabel");
const chartTitleFontFamilyInput = document.getElementById("chartTitleFontFamily");
const chartTitleFontSizeInput   = document.getElementById("chartTitleFontSize");
const showChartTitleCheckbox    = document.getElementById("showChartTitle");
const chartTitleBoldBtn      = document.getElementById("chartTitleBoldBtn");
const chartTitleItalicBtn    = document.getElementById("chartTitleItalicBtn");
const chartTitleUnderlineBtn = document.getElementById("chartTitleUnderlineBtn");

const xAxisFontFamilyInput = document.getElementById("xAxisFontFamily");
const xAxisFontSizeInput   = document.getElementById("xAxisFontSize");
const xAxisItalicBtn = document.getElementById("xAxisItalicBtn");
const xAxisBoldBtn   = document.getElementById("xAxisBoldBtn");

const yAxisMinInput        = document.getElementById("yAxisMin");
const yAxisMaxInput        = document.getElementById("yAxisMax");
const yAxisFormatInput     = document.getElementById("yAxisFormat");
const yAxisDecimalsInput = document.getElementById("yAxisDecimals");
const yAxisTickStepInput = document.getElementById("yAxisTickStep");
const yAxisFontFamilyInput = document.getElementById("yAxisFontFamily");
const yAxisFontSizeInput   = document.getElementById("yAxisFontSize");
const yAxisItalicBtn = document.getElementById("yAxisItalicBtn");
const yAxisBoldBtn   = document.getElementById("yAxisBoldBtn");

const targetInput       = document.getElementById("targetValue");


const targetDirectionInput = document.getElementById("targetDirection");
const targetDirectionSelect = targetDirectionInput;
const capabilityDiv     = document.getElementById("capability");
const annotationDateInput  = document.getElementById("annotationDate");
const annotationLabelInput = document.getElementById("annotationLabel");
const addAnnotationBtn     = document.getElementById("addAnnotationButton");
const clearAnnotationsBtn  = document.getElementById("clearAnnotationsButton");
const toggleSidebarButton = document.getElementById("toggleSidebarButton");
const splitPointSelect  = document.getElementById("splitPointSelect");
const addSplitButton    = document.getElementById("addSplitButton");
const clearSplitsButton = document.getElementById("clearSplitsButton");
const showMRCheckbox   = document.getElementById("showMRCheckbox");
const mrPanel           = document.getElementById("mrPanel");
const mrChartCanvas     = document.getElementById("mrChartCanvas");
const mrCanvas = mrChartCanvas;
const mrToggleRow = document.getElementById("mrToggleRow");


const generateButton    = document.getElementById("generateButton");
const errorMessage      = document.getElementById("errorMessage");
const chartCanvas       = document.getElementById("spcChart");
const summaryDiv        = document.getElementById("summary");
const downloadBtn       = document.getElementById("downloadPngButton");
const downloadPdfBtn    = document.getElementById("downloadPdfButton");
const openDataEditorButton   = document.getElementById("openDataEditorButton");
const dataEditorOverlay      = document.getElementById("dataEditorOverlay");
const dataEditorTextarea     = document.getElementById("dataEditorTextarea");
const dataEditorApplyButton  = document.getElementById("dataEditorApplyButton");
const dataEditorCancelButton = document.getElementById("dataEditorCancelButton");
const aiQuestionInput   = document.getElementById("aiQuestionInput");
const aiAskButton       = document.getElementById("aiAskButton");
const spcHelperPanel    = document.getElementById("spcHelperPanel");

const spcHelperIntro    = document.getElementById("spcHelperIntro");
const spcHelperChipsGeneral = document.getElementById("spcHelperChipsGeneral");
const spcHelperChipsChart   = document.getElementById("spcHelperChipsChart");
const spcHelperOutput   = document.getElementById("spcHelperOutput");
const spcHelperToggleGeneral = document.getElementById("spcHelperToggleGeneral");
const spcHelperToggleChart = document.getElementById("spcHelperToggleChart");
const spcHelperGeneralSection = document.getElementById("spcHelperGeneralSection");
const spcHelperChartSection = document.getElementById("spcHelperChartSection");

const shiftRulePointsInput = document.getElementById("shiftRulePoints");
const trendRulePointsInput = document.getElementById("trendRulePoints");
const ruleExplainerBtn = document.getElementById("ruleExplainerBtn");

const advancedRulesDetails = document.getElementById("advancedRulesDetails");
const enableAdvancedTrendCheckbox = document.getElementById("enableAdvancedTrend");
const advancedTrendRow = document.getElementById("advancedTrendRow");
const mrScreeningRow = document.getElementById("mrScreeningRow");
const screenLargeMRsCheckbox = document.getElementById("screenLargeMRs");

const ruleTwoOfThreeOuterThirdCheckbox = document.getElementById("ruleTwoOfThreeOuterThird");
const ruleFourOfFiveOneSigmaCheckbox = document.getElementById("ruleFourOfFiveOneSigma");
const zoneRulesSection = document.getElementById("zoneRulesSection");
const advancedContinuousCaution = document.getElementById("advancedContinuousCaution");

const enableRareRunTrendCheckbox = document.getElementById("enableRareRunTrend");
const rareRulesRow = document.getElementById("rareRulesRow");

if (enableRareRunTrendCheckbox) {
  enableRareRunTrendCheckbox.addEventListener("change", () => {
    if (!enableRareRunTrendCheckbox.checked) return;

    const ok = window.confirm(
      "T and G charts are naturally irregular.\n\n" +
      "Run and trend rules may create false signals on this chart type and are off by default.\n\n" +
      "Only enable these rules if you understand this trade-off."
    );

    if (!ok) {
      enableRareRunTrendCheckbox.checked = false;
    }
  });
}

const conservativeRulesMessage = document.getElementById("conservativeRulesMessage");
const chartTypeAvailabilityHint = document.getElementById("chartTypeAvailabilityHint");
const columnCheckWarning = document.getElementById("columnCheckWarning");

const flagSpecialCauseOnChartCheckbox = document.getElementById("flagSpecialCauseOnChart");
const colourBlindModeCheckbox = document.getElementById("colourBlindMode");
const lclClampRow = document.getElementById("lclClampRow");
const clampLclAtZeroCheckbox = document.getElementById("clampLclAtZero");

const savedSpcTheme = localStorage.getItem("spcColourTheme") || "default";
applySpcColourTheme(savedSpcTheme);

if (colourBlindModeCheckbox) {
  colourBlindModeCheckbox.checked = savedSpcTheme === "colourBlind";

  colourBlindModeCheckbox.addEventListener("change", () => {
    const themeName = colourBlindModeCheckbox.checked ? "colourBlind" : "default";
    localStorage.setItem("spcColourTheme", themeName);
    applySpcColourTheme(themeName);

    if (rawRows && rawRows.length && generateButton) {
      generateButton.click();
    }
  });
}

const dataEditorGridEl = document.getElementById("dataEditorGrid");
let dataEditorGrid = null; // jspreadsheet instance
const dataEditorHasHeaders = document.getElementById("dataEditorHasHeaders");
const dataEditorDetectHeadersButton = document.getElementById("dataEditorDetectHeadersButton");
const dataEditorHeaderStatus = document.getElementById("dataEditorHeaderStatus");
const dataEditorWorkbookBar = document.getElementById("dataEditorWorkbookBar");
const dataEditorSheetSelect = document.getElementById("dataEditorSheetSelect");
const dataEditorWorkbookStatus = document.getElementById("dataEditorWorkbookStatus");
const dataEditorDeleteHelpBtn = document.getElementById("dataEditorDeleteHelpBtn");
const dataEditorDeleteHelpPopup = document.getElementById("dataEditorDeleteHelpPopup");

let dataEditorSourceMode = "manual"; // "manual" | "excel"
let dataEditorWorkbook = null;
let dataEditorWorkbookSheetNames = [];
let dataEditorCurrentSheetName = "";

const sheetPickerOverlay = document.getElementById("sheetPickerOverlay");
const sheetPickerSelect = document.getElementById("sheetPickerSelect");
const sheetPickerConfirmButton = document.getElementById("sheetPickerConfirmButton");
const sheetPickerCancelButton = document.getElementById("sheetPickerCancelButton");


/* ============================================================
   PROJECT SAVE/LOAD + SETTINGS
   - "Project" = data + chart settings in one JSON file
   - Also remains backward-compatible with old settings-only files
   ============================================================ */

function downloadTextFile(filename, text) {
  const blob = new Blob([text], { type: "application/json;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function getCheckedRadioValue(name) {
  const el = document.querySelector(`input[name="${name}"]:checked`);
  return el ? el.value : "";
}

function setCheckedRadioValue(name, value) {
  if (!value) return;
  const el = document.querySelector(`input[name="${name}"][value="${value}"]`);
  if (el) el.checked = true;
}

function collectToolSettings() {
  const chartType = (typeof getSelectedChartType_NoSideEffects === "function")
    ? getSelectedChartType_NoSideEffects()
    : (typeof getSelectedChartType === "function" ? getSelectedChartType() : "run");

  const axisType = getCheckedRadioValue("axisType");

  const baselinePoints = baselineInput?.value ?? "";
  const targetValue = targetInput?.value ?? "";
  const targetDirection = targetDirectionSelect?.value ?? "";
  const title = chartTitleInput?.value ?? "";
  const xLabel = xAxisLabelInput?.value ?? "";
  const yLabel = yAxisLabelInput?.value ?? "";

  const shiftRule = shiftRulePointsInput?.value ?? "";
  const trendRule = trendRulePointsInput?.value ?? "";

  const flagSpecial = flagSpecialCauseOnChartCheckbox?.checked ?? true;
  const clampLcl = clampLclAtZeroCheckbox?.checked ?? false;

  const selectedColumns = {
    x: dateSelect?.value ?? "",
    y: valueSelect?.value ?? "",
    third: thirdSelect?.value ?? ""
  };

  return {
    tool: "Simple SPC Web Tool",
    settingsVersion: 1,
    savedAt: new Date().toISOString(),

    chartType,
    axisType,

    selectedColumns,
dateFormatPreference: getDateFormatPreference(),

appearance: {
  colourTheme: colourBlindModeCheckbox?.checked
    ? "colourBlind"
    : "default"
},

uRateMultiplier: String(getURateMultiplier()),

baselinePoints,
    target: {
      value: targetValue,
      direction: targetDirection,
      enabled: (typeof targetEnabled !== "undefined") ? !!targetEnabled : true
    },

  xmrMeanLabels: {
  show:
    (typeof showXmRMeanLabels !== "undefined")
      ? !!showXmRMeanLabels
      : false,

  positions:
    (typeof xmrMeanLabelPositions !== "undefined" &&
     xmrMeanLabelPositions)
      ? { ...xmrMeanLabelPositions }
      : {},

  verticalPositions:
    (typeof xmrMeanLabelVerticalPositions !== "undefined" &&
     xmrMeanLabelVerticalPositions)
      ? { ...xmrMeanLabelVerticalPositions }
      : {},

  hidden:
    (typeof xmrHiddenMeanLabels !== "undefined" &&
     xmrHiddenMeanLabels instanceof Set)
      ? Array.from(xmrHiddenMeanLabels)
      : []
},
    labels: {
  title,
  xLabel,
  yLabel,
  titleDisplay: showChartTitleCheckbox?.checked ?? true,
  titleFont: {
    family: chartTitleFontFamilyInput?.value ?? "",
    size: chartTitleFontSizeInput?.value ?? "16",
    bold: isPressed(chartTitleBoldBtn),
    italic: isPressed(chartTitleItalicBtn),
    underline: isPressed(chartTitleUnderlineBtn)
  }
},

    axes: {
      x: {
        font: {
          family: xAxisFontFamilyInput?.value ?? "",
          size: xAxisFontSizeInput?.value ?? "",
          style: isPressed(xAxisItalicBtn) ? "italic" : "normal",
          weight: isPressed(xAxisBoldBtn) ? "bold" : "normal"
        }
      },
      y: {
        min: yAxisMinInput?.value ?? "",
        max: yAxisMaxInput?.value ?? "",
        format: yAxisFormatInput?.value ?? "auto",
	decimals: yAxisDecimalsInput?.value ?? "auto",
        stepSize: yAxisTickStepInput?.value ?? "",
        font: {
          family: yAxisFontFamilyInput?.value ?? "",
          size: yAxisFontSizeInput?.value ?? "",
          style: isPressed(yAxisItalicBtn) ? "italic" : "normal",
          weight: isPressed(yAxisBoldBtn) ? "bold" : "normal"
        }
      }
    },

    rules: {
      shiftRulePoints: shiftRule,
      trendRulePoints: trendRule,
      enableAdvancedTrend: enableAdvancedTrendCheckbox?.checked ?? false,
      enableRareRunTrend: enableRareRunTrendCheckbox?.checked ?? false,
      ruleTwoOfThreeOuterThird: ruleTwoOfThreeOuterThirdCheckbox?.checked ?? false,
      ruleFourOfFiveOneSigma: ruleFourOfFiveOneSigmaCheckbox?.checked ?? false,
      flagSpecialCauseOnChart: flagSpecial,
      clampLclAtZero: clampLcl,
      screenLargeMRs: screenLargeMRsCheckbox?.checked ?? true
    },

        splits:
      Array.isArray(splits)
        ? splits.slice()
        : [],

    annotations:
      Array.isArray(annotations)
        ? annotations.slice()
        : [],

    limitExclusions:
      normaliseLimitExclusions(
        limitExclusions
      )
  };
}

function applyToolSettings(settings, { silent = true } = {}) {
  if (!settings || typeof settings !== "object") return;

// Restore the colour theme saved with this chart.
// Do not write it to localStorage: opening somebody else's saved
// chart should not permanently change this browser's preference.
const savedColourTheme = settings.appearance?.colourTheme;

if (
  savedColourTheme === "default" ||
  savedColourTheme === "colourBlind"
) {
  applySpcColourTheme(savedColourTheme);

  if (colourBlindModeCheckbox) {
    colourBlindModeCheckbox.checked =
      savedColourTheme === "colourBlind";
  }
}

  if (settings.chartType) setCheckedRadioValue("chartType", settings.chartType);
  if (settings.axisType) setCheckedRadioValue("axisType", settings.axisType);

  if (typeof updateUIForChartType === "function" && settings.chartType) {
    updateUIForChartType(settings.chartType);
  }

  if (uRateMultiplierSelect) {
  const savedMultiplier = String(settings.uRateMultiplier ?? "1");

  const validMultiplier = ["1", "10", "100", "1000", "10000"]
    .includes(savedMultiplier)
      ? savedMultiplier
      : "1";

  uRateMultiplierSelect.value = validMultiplier;
  lastURateMultiplier = Number(validMultiplier);
}

  if (shiftRulePointsInput && settings.rules?.shiftRulePoints !== undefined) {
    shiftRulePointsInput.value = settings.rules.shiftRulePoints;
  }
  if (trendRulePointsInput && settings.rules?.trendRulePoints !== undefined) {
    trendRulePointsInput.value = settings.rules.trendRulePoints;
  }

  if (enableAdvancedTrendCheckbox && settings.rules?.enableAdvancedTrend !== undefined) {
    enableAdvancedTrendCheckbox.checked = !!settings.rules.enableAdvancedTrend;
  }

  if (enableRareRunTrendCheckbox && settings.rules?.enableRareRunTrend !== undefined) {
    enableRareRunTrendCheckbox.checked = !!settings.rules.enableRareRunTrend;
  }

  if (flagSpecialCauseOnChartCheckbox && settings.rules?.flagSpecialCauseOnChart !== undefined) {
    flagSpecialCauseOnChartCheckbox.checked = !!settings.rules.flagSpecialCauseOnChart;
  }

  if (ruleTwoOfThreeOuterThirdCheckbox && settings.rules?.ruleTwoOfThreeOuterThird !== undefined) {
    ruleTwoOfThreeOuterThirdCheckbox.checked = !!settings.rules.ruleTwoOfThreeOuterThird;
  }

  if (ruleFourOfFiveOneSigmaCheckbox && settings.rules?.ruleFourOfFiveOneSigma !== undefined) {
    ruleFourOfFiveOneSigmaCheckbox.checked = !!settings.rules.ruleFourOfFiveOneSigma;
  }

  if (clampLclAtZeroCheckbox && settings.rules?.clampLclAtZero !== undefined) {
    clampLclAtZeroCheckbox.checked = !!settings.rules.clampLclAtZero;
  }

  if (screenLargeMRsCheckbox) {
  if (settings.rules?.screenLargeMRs !== undefined) {
    // Newer saved chart: restore the user's saved choice.
    screenLargeMRsCheckbox.checked =
      !!settings.rules.screenLargeMRs;
  } else {
    // Older saved chart: preserve the legacy XmR calculation
    // rather than silently changing its historical limits.
    screenLargeMRsCheckbox.checked = false;
  }
}

  if (baselineInput && settings.baselinePoints !== undefined) baselineInput.value = settings.baselinePoints;

  if (targetInput && settings.target?.value !== undefined) targetInput.value = settings.target.value;
  if (targetDirectionSelect && settings.target?.direction) targetDirectionSelect.value = settings.target.direction;

  if (typeof targetEnabled !== "undefined" && settings.target?.enabled !== undefined) {
    targetEnabled = !!settings.target.enabled;
    if (typeof updateTargetToggleBtn === "function") updateTargetToggleBtn();
    if (typeof updateTargetToggleVisibility === "function") updateTargetToggleVisibility();
  }

  // Restore XmR mean-label display settings.
// Older saved charts do not contain these settings, so default to labels off.
if (settings.xmrMeanLabels && typeof settings.xmrMeanLabels === "object") {
  showXmRMeanLabels =
    !!settings.xmrMeanLabels.show;

  xmrMeanLabelPositions = {};

  Object.entries(
    settings.xmrMeanLabels.positions || {}
  ).forEach(([key, value]) => {
    const idx = Number(key);
    const pos = Number(value);

    if (
      Number.isInteger(idx) &&
      idx >= 0 &&
      Number.isFinite(pos)
    ) {
      xmrMeanLabelPositions[idx] =
        Math.max(0, Math.min(1, pos));
    }
  });

  xmrMeanLabelVerticalPositions = {};

  Object.entries(
    settings.xmrMeanLabels.verticalPositions || {}
  ).forEach(([key, value]) => {
    const idx = Number(key);

    if (
      Number.isInteger(idx) &&
      idx >= 0 &&
      (value === "above" || value === "below")
    ) {
      xmrMeanLabelVerticalPositions[idx] =
        value;
    }
  });

  const hidden =
    Array.isArray(settings.xmrMeanLabels.hidden)
      ? settings.xmrMeanLabels.hidden
      : [];

  xmrHiddenMeanLabels = new Set(
    hidden
      .map(Number)
      .filter(
        idx =>
          Number.isInteger(idx) &&
          idx >= 0
      )
  );
} else {
  // Backward compatibility with saved charts created
  // before mean labels were introduced.
  showXmRMeanLabels = false;
  xmrMeanLabelPositions = {};
  xmrMeanLabelVerticalPositions = {};
  xmrHiddenMeanLabels = new Set();
}

  if (chartTitleInput && settings.labels?.title !== undefined) {
  chartTitleInput.value = settings.labels.title;
  chartTitleManuallyEdited = String(settings.labels.title).trim() !== "";
}

if (xAxisLabelInput && settings.labels?.xLabel !== undefined) {
  xAxisLabelInput.value = settings.labels.xLabel;
  xAxisLabelManuallyEdited = String(settings.labels.xLabel).trim() !== "";
}

if (yAxisLabelInput && settings.labels?.yLabel !== undefined) {
  yAxisLabelInput.value = settings.labels.yLabel;
  yAxisLabelManuallyEdited = String(settings.labels.yLabel).trim() !== "";
}

if (
  showChartTitleCheckbox &&
  settings.labels?.titleDisplay !== undefined
) {
  showChartTitleCheckbox.checked = !!settings.labels.titleDisplay;
}

if (
  chartTitleFontFamilyInput &&
  settings.labels?.titleFont?.family !== undefined
) {
  chartTitleFontFamilyInput.value = settings.labels.titleFont.family;
}

if (
  chartTitleFontSizeInput &&
  settings.labels?.titleFont?.size !== undefined
) {
  chartTitleFontSizeInput.value = settings.labels.titleFont.size;
}

if (settings.labels?.titleFont?.bold !== undefined) {
  setPressed(
    chartTitleBoldBtn,
    !!settings.labels.titleFont.bold
  );
}

if (settings.labels?.titleFont?.italic !== undefined) {
  setPressed(
    chartTitleItalicBtn,
    !!settings.labels.titleFont.italic
  );
}

if (settings.labels?.titleFont?.underline !== undefined) {
  setPressed(
    chartTitleUnderlineBtn,
    !!settings.labels.titleFont.underline
  );
}

// Older project files pre-date title style controls.
// Their titles were bold by default.
if (settings.labels?.titleFont?.bold === undefined) {
  setPressed(chartTitleBoldBtn, true);
}

  if (xAxisFontFamilyInput && settings.axes?.x?.font?.family !== undefined) xAxisFontFamilyInput.value = settings.axes.x.font.family;
  if (xAxisFontSizeInput && settings.axes?.x?.font?.size !== undefined) xAxisFontSizeInput.value = settings.axes.x.font.size;
  setPressed(xAxisItalicBtn, settings.axes?.x?.font?.style === "italic");
  setPressed(xAxisBoldBtn, settings.axes?.x?.font?.weight === "bold");

  if (yAxisMinInput && settings.axes?.y?.min !== undefined) {
  yAxisMinInput.value = settings.axes.y.min;
}

if (yAxisMaxInput && settings.axes?.y?.max !== undefined) {
  yAxisMaxInput.value = settings.axes.y.max;
}

const hasSavedYMin =
  settings.axes?.y?.min !== undefined &&
  String(settings.axes.y.min).trim() !== "";

const hasSavedYMax =
  settings.axes?.y?.max !== undefined &&
  String(settings.axes.y.max).trim() !== "";

yAxisBoundsManuallyEdited = hasSavedYMin || hasSavedYMax;

if (yAxisFormatInput && settings.axes?.y?.format !== undefined) {
  yAxisFormatInput.value = settings.axes.y.format;
}

if (yAxisDecimalsInput && settings.axes?.y?.decimals !== undefined) {
  yAxisDecimalsInput.value = settings.axes.y.decimals;
}

if (yAxisTickStepInput && settings.axes?.y?.stepSize !== undefined) {
  yAxisTickStepInput.value = settings.axes.y.stepSize;
}

if (typeof updateYAxisInputStep === "function") {
  updateYAxisInputStep();
}

if (yAxisFontFamilyInput && settings.axes?.y?.font?.family !== undefined) {
  yAxisFontFamilyInput.value = settings.axes.y.font.family;
}

if (yAxisFontSizeInput && settings.axes?.y?.font?.size !== undefined) {
  yAxisFontSizeInput.value = settings.axes.y.font.size;
}

setPressed(yAxisItalicBtn, settings.axes?.y?.font?.style === "italic");
setPressed(yAxisBoldBtn, settings.axes?.y?.font?.weight === "bold");

  if (Array.isArray(settings.splits)) splits = settings.splits.slice();
  if (Array.isArray(settings.annotations)) annotations = settings.annotations.slice();
  limitExclusions =
    normaliseLimitExclusions(
      settings.limitExclusions
    );

  const missing = [];

  function setSelectIfOptionExists(selectEl, value, labelForMissing) {
    if (!selectEl || !value) return;
    const exists = Array.from(selectEl.options).some(o => o.value === value && !o.disabled);
    if (exists) {
      selectEl.value = value;
    } else {
      missing.push(labelForMissing || value);
    }
  }

  if (dateFormatPreferenceSelect && settings.dateFormatPreference) {
    dateFormatPreferenceSelect.value = settings.dateFormatPreference;
  }

  const cols = settings.selectedColumns || {};
  setSelectIfOptionExists(dateSelect, cols.x, `X-axis column "${cols.x}"`);
  setSelectIfOptionExists(valueSelect, cols.y, `Value column "${cols.y}"`);
  setSelectIfOptionExists(thirdSelect, cols.third, `Third column "${cols.third}"`);

  const chartTypeNow = (typeof getSelectedChartType_NoSideEffects === "function")
    ? getSelectedChartType_NoSideEffects()
    : (typeof getSelectedChartType === "function" ? getSelectedChartType() : "run");

  if (rawRows && rawRows.length && typeof applyColumnIntelligence === "function") {
    applyColumnIntelligence(chartTypeNow);
  }

// Reapply the saved X-axis type after column intelligence.
// X̄–S is always categorical by subgroup, so old saved projects
// that used Date / time are safely normalised to Sequence / category.
if (
  chartTypeNow === "xbars" ||
  chartTypeNow === "pareto"
) {
  if (typeof setAxisType === "function") {
    setAxisType("sequence");
  } else {
    setCheckedRadioValue(
      "axisType",
      "sequence"
    );
  }
} else if (settings.axisType) {
  setCheckedRadioValue(
    "axisType",
    settings.axisType
  );
}

if (
  clampLclAtZeroCheckbox &&
  settings.rules?.clampLclAtZero !== undefined
) {
  clampLclAtZeroCheckbox.checked = !!settings.rules.clampLclAtZero;
}

  if (missing.length && typeof showError === "function" && !silent) {
    showError(
      "Imported settings/project applied, but some saved columns were not found in your current data. " +
      "Please reselect: " + missing.join(", ")
    );
  }

  updateDateFormatWarning();
  if (typeof updateDateControlsState === "function") updateDateControlsState();
}

function collectProjectFile() {
  return {
    tool: "Simple SPC Web Tool",
    projectVersion: 1,
    savedAt: new Date().toISOString(),
    data: {
      rawRows: Array.isArray(rawRows) ? rawRows : []
    },
    settings: collectToolSettings()
  };
}

function exportProjectNow() {
  const project = collectProjectFile();
  const safeDate = new Date().toISOString().slice(0, 10);
  const filename = `spc-project-${safeDate}.json`;
  downloadTextFile(filename, JSON.stringify(project, null, 2));
}

function resetAnnotationsAndSplitsForNewData() {
  annotations = [];
  limitExclusions = [];
  if (annotationDateInput) annotationDateInput.value = "";
  if (annotationLabelInput) annotationLabelInput.value = "";

  splits = [];
  if (splitPointSelect) splitPointSelect.innerHTML = "";

  resetXmRMeanLabelState();
}

function loadProjectObject(projectObj) {
  if (!projectObj || typeof projectObj !== "object") {
    alert("That file could not be read as an SPC project.");
    return;
  }

  const rows = projectObj?.data?.rawRows;
  const settings = projectObj?.settings;

  if (!Array.isArray(rows) || rows.length === 0) {
    alert("This project file does not contain any saved data.");
    return;
  }

  resetAnnotationsAndSplitsForNewData();

  const ok = loadRows(rows);
  if (!ok) {
    alert("The project data could not be loaded.");
    return;
  }

  if (settings) {
    applyToolSettings(settings, { silent: false });
  }

  /*
   * The data and restored settings are now in place.
   * Mark the model dirty before generating the final chart.
   */
  markDataModelDirty();

  /*
   * Wait until the browser has completed the restored-control updates,
   * then generate the chart once using the final saved settings.
   */
  if (generateButton) {
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {

        /*
         * Reassert saved user choices that automatic data/column logic
         * may have changed during loading.
         */
        if (settings?.axisType) {
          setCheckedRadioValue("axisType", settings.axisType);
        }

        if (
          clampLclAtZeroCheckbox &&
          settings?.rules?.clampLclAtZero !== undefined
        ) {
          clampLclAtZeroCheckbox.checked =
            !!settings.rules.clampLclAtZero;
        }

        if (typeof updateDateFormatWarning === "function") {
          updateDateFormatWarning();
        }

        if (typeof updateDateControlsState === "function") {
          updateDateControlsState();
        }

        if (typeof lastGenerateWasManual !== "undefined") {
          lastGenerateWasManual = false;
        }

        generateButton.click();
      });
    });
  }
}

function importSettingsOrProjectFromFile(file) {
  if (!file) return;

  const reader = new FileReader();

  reader.onload = () => {
    try {
      const text = String(reader.result || "");
      const parsed = JSON.parse(text);

      // New full project file
      if (
        parsed &&
        typeof parsed === "object" &&
        parsed.projectVersion === 1 &&
        parsed.data?.rawRows
      ) {
        loadProjectObject(parsed);
        return;
      }

      // Old settings-only file (backward compatibility)
      if (
        parsed &&
        typeof parsed === "object" &&
        parsed.settingsVersion === 1
      ) {
        if (rawRows && rawRows.length) {
          applyToolSettings(parsed, { silent: false });

          markDataModelDirty();

          requestAnimationFrame(() => {
            requestAnimationFrame(() => {
              if (typeof lastGenerateWasManual !== "undefined") {
                lastGenerateWasManual = false;
              }

              if (generateButton) {
                generateButton.click();
              }
            });
          });

        } else {
          pendingImportedSettings = parsed;

          alert(
            "Settings loaded. Now upload your CSV or Excel data and the tool will apply these settings automatically."
          );
        }

        return;
      }

      alert(
        "That file doesn’t look like a supported SPC project or settings file."
      );

    } catch (e) {
      alert(
        "Could not read that JSON file. Please check it is a valid export from this tool."
      );
    }
  };

  reader.readAsText(file);
}

// Wire up the buttons
if (exportSettingsBtn) {
  exportSettingsBtn.addEventListener("click", () => {
    exportProjectNow();
  });
}

if (importSettingsBtn && importSettingsFileInput) {
  importSettingsBtn.addEventListener("click", () => {
    importSettingsFileInput.value = "";
    importSettingsFileInput.click();
  });

  importSettingsFileInput.addEventListener("change", () => {
    const file = importSettingsFileInput.files && importSettingsFileInput.files[0];
    if (file) importSettingsOrProjectFromFile(file);
  });
}

function openSheetPicker(sheetNames) {
  return new Promise((resolve) => {
    if (!sheetPickerOverlay || !sheetPickerSelect || !sheetPickerConfirmButton || !sheetPickerCancelButton) {
      resolve(sheetNames[0] || null);
      return;
    }

    sheetPickerSelect.innerHTML = "";
    (sheetNames || []).forEach(name => {
      const opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      sheetPickerSelect.appendChild(opt);
    });

    sheetPickerOverlay.style.display = "flex";
    sheetPickerSelect.focus();

    function cleanup(result) {
      sheetPickerOverlay.style.display = "none";
      sheetPickerConfirmButton.removeEventListener("click", onConfirm);
      sheetPickerCancelButton.removeEventListener("click", onCancel);
      sheetPickerOverlay.removeEventListener("click", onOverlayClick);
      document.removeEventListener("keydown", onKeyDown);
      resolve(result);
    }

    function onConfirm() {
      cleanup(sheetPickerSelect.value || sheetNames[0] || null);
    }

    function onCancel() {
      cleanup(null);
    }

    function onOverlayClick(e) {
      if (e.target === sheetPickerOverlay) {
        cleanup(null);
      }
    }

    function onKeyDown(e) {
      if (e.key === "Escape") cleanup(null);
      if (e.key === "Enter") cleanup(sheetPickerSelect.value || sheetNames[0] || null);
    }

    sheetPickerConfirmButton.addEventListener("click", onConfirm);
    sheetPickerCancelButton.addEventListener("click", onCancel);
    sheetPickerOverlay.addEventListener("click", onOverlayClick);
    document.addEventListener("keydown", onKeyDown);
  });
}

function normalizeWorkbookCellValue(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }
  return value;
}

function sheetToRowObjects(worksheet) {
  if (!worksheet) return [];

  const rows = XLSX.utils.sheet_to_json(worksheet, {
    defval: "",
    raw: false
  });

  return rows.map(row => {
    const out = {};
    Object.keys(row).forEach(key => {
      out[key] = normalizeWorkbookCellValue(row[key]);
    });
    return out;
  });
}

async function readExcelWorkbook(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, {
    type: "array",
    cellDates: true
  });

  const sheetNames = workbook.SheetNames || [];
  if (!sheetNames.length) {
    throw new Error("No worksheets were found in the Excel file.");
  }

  return workbook;
}

function resetStateAfterDataLoad() {
  markDataModelDirty();

  annotations = [];
  if (annotationDateInput) annotationDateInput.value = "";
  if (annotationLabelInput) annotationLabelInput.value = "";

  splits = [];
  if (splitPointSelect) splitPointSelect.innerHTML = "";

  resetXmRMeanLabelState();
}

function prepareForNewImportedChart(rows) {
  // Validate the new data before destroying the current chart.
  if (!rows || rows.length === 0) {
    showError("No rows found in the data.");
    return false;
  }

  const firstRow = rows[0];
  const columns =
    firstRow && typeof firstRow === "object"
      ? Object.keys(firstRow)
      : [];

  if (columns.length === 0) {
    showError("Could not detect any columns in the data.");
    return false;
  }

  const hasExistingWork =
    (Array.isArray(rawRows) && rawRows.length > 0) ||
    !!currentChart ||
    !!dataModelDirty;

  if (hasExistingWork) {
    const ok = window.confirm(
      "Load these data as a new chart?\n\n" +
      "This will clear the current data and chart settings. " +
      "Save the current chart first if you want to reopen it later."
    );

    if (!ok) {
      // Allow the same file to be selected again later.
      if (fileInput) fileInput.value = "";
      return false;
    }
  }

  // We now know the replacement data are usable.
  resetAll();
  return true;
}

function guessColumns(rows) {
  if (!rows || rows.length === 0) return { dateCol: null, valueCol: null, hasDateCandidate: false };

  const sample = rows.slice(0, Math.min(rows.length, 200));
  const cols = Object.keys(sample[0] || {});
  if (cols.length === 0) return { dateCol: null, valueCol: null, hasDateCandidate: false };

  const norm = (v) => String(v ?? "").trim();
  const lowerName = (c) => String(c || "").toLowerCase();

  function isDateLikeValue(v) {
    const s = norm(v);
    if (!s) return false;
    const d = parseDateValue(s);
    return !!d && isFinite(d.getTime());
  }

  function numericValue(v) {
    const n = toNumericValue(v);
    return Number.isFinite(n) ? n : NaN;
  }

  function profileColumn(col) {
    const vals = sample
      .map(r => r[col])
      .filter(v => v !== null && v !== undefined && norm(v) !== "");

    const maxTake = Math.min(vals.length, 200);
    const taken = vals.slice(0, maxTake);

    let dateLike = 0;
    let numeric = 0;
    let integerish = 0;

    const nums = [];
    const uniques = new Set();

    for (const v of taken) {
      const s = norm(v);
      uniques.add(s);

      if (isDateLikeValue(v)) dateLike++;

      const n = numericValue(v);
      if (Number.isFinite(n)) {
        numeric++;
        nums.push(n);
        if (isIntegerish(n)) integerish++;
      }
    }

    const total = taken.length || 1;
    const numericFrac = numeric / total;
    const dateFrac = dateLike / total;
    const intFrac = numeric ? (integerish / numeric) : 0;
    const uniqueFrac = uniques.size / total;

    // Detect monotonic increasing (common in index columns like Week_Number)
    let monotonicScore = 0;
    if (nums.length >= 6) {
      let inc = 0;
      for (let i = 1; i < nums.length; i++) {
        if (nums[i] >= nums[i - 1]) inc++;
      }
      monotonicScore = inc / (nums.length - 1);
    }

    // Rough variability (std dev)
    let sd = 0;
    if (nums.length >= 3) {
      const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
      const v = nums.reduce((a, b) => a + (b - mean) * (b - mean), 0) / (nums.length - 1);
      sd = Math.sqrt(Math.max(0, v));
    }

    const name = lowerName(col);
    const nameLooksLikeId =
      name.includes("week") ||
      name.includes("id") ||
      name.includes("index") ||
      name.includes("number") ||
      name.includes("seq") ||
      name.includes("row");

    const nameLooksLikeMeasure =
      name.includes("value") ||
      name.includes("measure") ||
      name.includes("mean") ||
      name.includes("rate") ||
      name.includes("time") ||
      name.includes("score") ||
      name.includes("count");

    const looksIndexLike =
      numericFrac > 0.9 &&
      intFrac > 0.9 &&
      uniqueFrac > 0.85 &&
      monotonicScore > 0.85;

    return {
      col,
      numericFrac,
      dateFrac,
      intFrac,
      uniqueFrac,
      monotonicScore,
      sd,
      nameLooksLikeId,
      nameLooksLikeMeasure,
      looksIndexLike
    };
  }

  const profiles = cols.map(profileColumn);

  // ---- Pick date column (prefer truly date-like columns that are NOT strongly numeric) ----
  const bestDate = profiles
    .filter(p => p.dateFrac > 0.4)
    .sort((a, b) => {
      if (b.dateFrac !== a.dateFrac) return b.dateFrac - a.dateFrac;
      // tie-break: name contains "date" or "time"
      const aName = lowerName(a.col);
      const bName = lowerName(b.col);
      const aHas = aName.includes("date") || aName.includes("time");
      const bHas = bName.includes("date") || bName.includes("time");
      return (bHas ? 1 : 0) - (aHas ? 1 : 0);
    })[0];

  let dateCol = bestDate ? bestDate.col : null;
  const hasDateCandidate = !!dateCol;

  // If we did NOT find a real date column, default X to the first column (sequence/category label)
  if (!dateCol) dateCol = cols[0];

  // ---- Pick value column (prefer measurement-like, avoid index-like) ----
  const candidates = profiles
    .filter(p => p.col !== dateCol)
    .filter(p => p.numericFrac > 0.4)
    .sort((a, b) => {
      const score = (p) => {
        let s = 0;

        // numeric quality
        s += p.numericFrac * 2;

        // avoid date-like
        s -= p.dateFrac * 3;

        // avoid index-like columns hard
        if (p.looksIndexLike) s -= 3;

        // mild preference for continuous measures (not purely integer IDs)
        s += (1 - p.intFrac) * 0.8;

        // prefer variability (avoid flat or IDs)
        s += Math.min(1, p.sd / 10) * 0.7;

        // name hints (tie-breakers)
        if (p.nameLooksLikeMeasure) s += 0.5;
        if (p.nameLooksLikeId) s -= 0.8;

        return s;
      };

      return score(b) - score(a);
    });

  let valueCol = candidates.length ? candidates[0].col : null;

  // last-resort fallback to avoid null
  if (!valueCol) valueCol = dateCol;

  return { dateCol, valueCol, hasDateCandidate };
}



function updateMrToggleVisibility() {
  if (!showMRCheckbox || !mrPanel) return;

  const chartType = getSelectedChartType_NoSideEffects();
  const mrDisplayOptions = document.getElementById("mrDisplayOptions");
  const toggleLabel = document.getElementById("secondaryChartToggleLabel");
  const displayModeLabel = document.getElementById("mrDisplayModeLabel");
  const displaySubhint = document.getElementById("mrDisplaySubhint");

  const showSecondaryToggle = (chartType === "xmr" || chartType === "xbars");
  const showMR = !!showMRCheckbox.checked;

  // Show the checkbox row for XmR and X̄–S
  if (mrToggleRow) {
    mrToggleRow.style.display = showSecondaryToggle ? "block" : "none";
  }

  // Relabel based on chart type
  if (chartType === "xmr") {
    if (toggleLabel) toggleLabel.textContent = "Show MR";
    if (displayModeLabel) displayModeLabel.textContent = "MR display";
    if (displaySubhint) displaySubhint.textContent = "(includes splits)";
    if (mrDisplayOptions) mrDisplayOptions.style.display = showMR ? "flex" : "none";
  } else if (chartType === "xbars") {
    if (toggleLabel) toggleLabel.textContent = "Show S chart";
    if (displayModeLabel) displayModeLabel.textContent = "S chart display";
    if (displaySubhint) displaySubhint.textContent = "(includes splits)";
    if (mrDisplayOptions) mrDisplayOptions.style.display = showMR ? "flex" : "none";
  } else {
    if (mrDisplayOptions) mrDisplayOptions.style.display = "none";
    hideMrPanelNow();
  }
}

function isProbablyHeaderRow(row) {
  // Heuristic: headers tend to be non-numeric strings; data tends to be numeric/date-ish.
  // We’ll score each cell and decide.
  let headerish = 0;
  let datish = 0;

  for (const cell of row) {
    const s = String(cell ?? "").trim();
    if (!s) continue;

    const looksNumeric = /^-?\d+(\.\d+)?%?$/.test(s.replace(/,/g, ""));
    const looksDate = /^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(s) || /^\d{4}-\d{2}-\d{2}/.test(s);

    if (looksNumeric || looksDate) datish++;
    else headerish++;
  }

  // If mostly text labels => header row
  return headerish >= datish && headerish > 0;
}



if (dataEditorSheetSelect) {
  dataEditorSheetSelect.addEventListener("change", () => {
    const nextSheet = dataEditorSheetSelect.value;
    if (!nextSheet || nextSheet === dataEditorCurrentSheetName) return;
    loadWorkbookSheetIntoDataEditor(nextSheet);
  });
}

function getNonBlankGridRows() {
  if (!dataEditorGrid) return [];
  const data2D = dataEditorGrid.getData();
  return data2D.filter(row => row.some(cell => String(cell ?? "").trim() !== ""));
}


function detectHeadersFromGrid() {
  const rows = getNonBlankGridRows();

  // When editing an existing dataset, the column headings
  // are already stored separately in the spreadsheet titles.
  // Do not treat the first observation as a header row.
  if (
    dataEditorSourceMode === "manual" &&
    Array.isArray(rawRows) &&
    rawRows.length > 0
  ) {
    return false;
  }

  // For a new, empty editor or an Excel worksheet,
  // retain the existing header-detection behaviour.
  if (rows.length < 2) return false;

  return isProbablyHeaderRow(rows[0]);
}

function renderHeaderStatus() {
  if (!dataEditorHeaderStatus || !dataEditorHasHeaders) return;

  const mode = dataEditorHasHeaders.checked ? "headings" : "data";

  dataEditorHeaderStatus.innerHTML =
    `Apply will treat the <strong>first row of the grid</strong> as <strong>${mode}</strong>. ` +
    `<br><small>Tip: If you loaded a CSV normally, your headings are already the column titles — so usually leave this OFF.</small>`;
}




function hideMrPanelNow() {
  if (mrChart) {
    mrChart.destroy();
    mrChart = null;
  }
  if (mrPanel) {
    mrPanel.style.display = "none";
  }
}

if (showMRCheckbox) {
  showMRCheckbox.addEventListener("change", () => {
    updateMrToggleVisibility();
    const chartType = getSelectedChartType_NoSideEffects();

    // Secondary chart only relevant for XmR and X̄–S
    if (chartType !== "xmr" && chartType !== "xbars") {
      hideMrPanelNow();
      return;
    }

    if (currentChart) {
      generateButton.click();
    } else {
      hideMrPanelNow();
    }
  });
}

// Redraw MR chart when MR display mode changes
document.querySelectorAll("input[name='mrDisplayMode']").forEach(r => {
  r.addEventListener("change", () => {
    updateMrToggleVisibility();
    const chartType = getSelectedChartType ? getSelectedChartType_NoSideEffects() : "run";

    // Only relevant for XmR and X̄–S charts
    if (chartType !== "xmr" && chartType !== "xbars") return;

    if (rawRows && rawRows.length && currentChart) {
      generateButton.click();
    }
  });
});


const targetToggleBtn = document.getElementById("targetToggleBtn");
let targetEnabled = true;

function updateTargetToggleBtn() {
  if (!targetToggleBtn) return;
  targetToggleBtn.textContent = targetEnabled ? "Hide target line" : "Show target line";
}

function getChartTitleSettings() {
  const size = Number(chartTitleFontSizeInput?.value);

  return {
    show: showChartTitleCheckbox ? showChartTitleCheckbox.checked : true,

    underline: isPressed(chartTitleUnderlineBtn),

    font: {
      family: (chartTitleFontFamilyInput?.value || "").trim(),
      size: Number.isFinite(size) && size > 0 ? size : 16,
      style: isPressed(chartTitleItalicBtn) ? "italic" : "normal",
      weight: isPressed(chartTitleBoldBtn) ? "bold" : "normal"
    }
  };
}

function buildChartTitleConfig(title) {
  const settings = getChartTitleSettings();

  return {
    display: settings.show && !!String(title || "").trim(),
    text: title,
    font: cleanFontOptions(settings.font)
  };
}

function applyPresentationEditsLive() {
  if (!currentChart) return;

  const title = (chartTitleInput?.value || "").trim();
  const xLabel = (xAxisLabelInput?.value || "").trim();
  const yLabel = (yAxisLabelInput?.value || "").trim();
  const axisSettings = (typeof getAxisSettings === "function") ? getAxisSettings() : null;

  // Title
  // Title
if (currentChart.options?.plugins?.title) {
  const titleSettings = getChartTitleSettings();

  currentChart.options.plugins.title.display =
    titleSettings.show && !!title;

  currentChart.options.plugins.title.text = title;
  currentChart.options.plugins.title.font =
    cleanFontOptions(titleSettings.font);
}

  // X axis
  if (currentChart.options?.scales?.x) {
    if (currentChart.options.scales.x.title) {
      currentChart.options.scales.x.title.display = !!xLabel;
      currentChart.options.scales.x.title.text = xLabel;
      if (axisSettings?.x?.font) {
        currentChart.options.scales.x.title.font = cleanFontOptions(axisSettings.x.font);
      }
    }

    if (!currentChart.options.scales.x.ticks) {
      currentChart.options.scales.x.ticks = {};
    }

    currentChart.options.scales.x.ticks.font = cleanFontOptions(axisSettings?.x?.font);
  }

  // Y axis
  if (currentChart.options?.scales?.y) {
    if (currentChart.options.scales.y.title) {
      currentChart.options.scales.y.title.display = !!yLabel;
      currentChart.options.scales.y.title.text = yLabel;
      if (axisSettings?.y?.font) {
        currentChart.options.scales.y.title.font = cleanFontOptions(axisSettings.y.font);
      }
    }

    if (!currentChart.options.scales.y.ticks) {
      currentChart.options.scales.y.ticks = {};
    }

    currentChart.options.scales.y.ticks.font = cleanFontOptions(axisSettings?.y?.font);
    currentChart.options.scales.y.ticks.callback = buildTickFormatter(axisSettings?.y?.format);
    currentChart.options.scales.y.min = axisSettings?.y?.min;
    currentChart.options.scales.y.max = axisSettings?.y?.max;
  }

  currentChart.update("none");
}

function hasValidTargetInput() {
  if (!targetInput) return false;
  const v = targetInput.value.trim();
  if (v === "") return false;
  const num = Number(v);
  return isFinite(num);
}

function updateTargetToggleVisibility() {
  if (!targetToggleBtn) return;

  if (hasValidTargetInput()) {
    targetToggleBtn.style.display = "inline-flex";
  } else {
    // No target defined: hide the button, but keep the user's
    // current show/hide preference for when a target is entered.
    targetToggleBtn.style.display = "none";
  }

  if (typeof updateTargetToggleBtn === "function") {
    updateTargetToggleBtn();
  }
}

// When user types target value: show/hide button and (optionally) redraw
if (targetInput) {
  targetInput.addEventListener("input", () => {
    updateTargetToggleVisibility();

    if (rawRows && rawRows.length && generateButton) {
      lastGenerateWasManual = false;
      generateButton.click();
    }
  });
}

// When target direction changes, immediately recalculate
// target performance and capability analysis.
if (targetDirectionInput) {
  targetDirectionInput.addEventListener("change", () => {
    if (rawRows && rawRows.length && generateButton) {
      lastGenerateWasManual = false;
      generateButton.click();
    }
  });
}



// Call once on load
updateTargetToggleVisibility();
updateYAxisInputStep();	

function debounce(fn, ms = 80) {
  let t;
  return (...args) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...args), ms);
  };
}

const applyPresentationEditsLiveDebounced = debounce(applyPresentationEditsLive, 60);

if (chartTitleInput) {
  chartTitleInput.addEventListener("input", () => {
    chartTitleManuallyEdited = true;
    applyPresentationEditsLiveDebounced();
  });
}

[
  chartTitleFontFamilyInput,
  chartTitleFontSizeInput,
  showChartTitleCheckbox
].forEach(el => {
  if (!el) return;

  el.addEventListener("input", () => {
    applyPresentationEditsLiveDebounced();
  });

  el.addEventListener("change", () => {
    applyPresentationEditsLiveDebounced();
  });
});

if (xAxisLabelInput) {
  xAxisLabelInput.addEventListener("input", () => {
    xAxisLabelManuallyEdited = true;
    applyPresentationEditsLiveDebounced();
  });
}

if (yAxisLabelInput) {
  yAxisLabelInput.addEventListener("input", () => {
    yAxisLabelManuallyEdited = true;
    applyPresentationEditsLiveDebounced();
  });
}

function handleAxisControlChanged({ livePreview = false } = {}) {
  markDataModelDirty();

  if (!rawRows || !rawRows.length) return;

  if (livePreview && currentChart) {
    applyPresentationEditsLiveDebounced();
  }
}

[
  yAxisMinInput,
  yAxisMaxInput
].forEach(el => {
  if (!el) return;
  el.addEventListener("input", () => {
    yAxisBoundsManuallyEdited = true;
    handleAxisControlChanged({ livePreview: false });
  });
  el.addEventListener("change", () => {
    yAxisBoundsManuallyEdited = true;
    handleAxisControlChanged({ livePreview: false });
  });
});

[
  yAxisFormatInput,
  yAxisDecimalsInput,
  yAxisTickStepInput,
  xAxisFontFamilyInput,
  yAxisFontFamilyInput,
  xAxisFontSizeInput,
  yAxisFontSizeInput,
  xAxisItalicBtn,
  xAxisBoldBtn,
  yAxisItalicBtn,
  yAxisBoldBtn
].forEach(el => {
  if (!el) return;
  el.addEventListener("input", () => {
  if (el === yAxisFormatInput && typeof updateYAxisInputStep === "function") {
  updateYAxisInputStep();
}
  handleAxisControlChanged({ livePreview: true });
});

el.addEventListener("change", () => {
  if (el === yAxisFormatInput && typeof updateYAxisInputStep === "function") {
  updateYAxisInputStep();
}
  handleAxisControlChanged({ livePreview: true });
});
  el.addEventListener("click", () => handleAxisControlChanged({ livePreview: true }));
});

if (valueSelect) {
  valueSelect.addEventListener("change", () => {
    yAxisBoundsManuallyEdited = false;
    applyDefaultYBoundsForSelectedColumn();

    if (rawRows && rawRows.length) {
      markDataModelDirty();
    }
  });
}

if (dateSelect) {
  dateSelect.addEventListener("change", () => {
    updateDateFormatWarning();
  });
}

if (dateFormatPreferenceSelect) {
  dateFormatPreferenceSelect.addEventListener("change", () => {
    updateDateFormatWarning();

    if (rawRows && rawRows.length && generateButton) {
      lastGenerateWasManual = false;
      generateButton.click();
    }
  });
}

document.querySelectorAll("input[name='axisType']").forEach(r => {
  r.addEventListener("change", () => {
    updateDateFormatWarning();
    updateDateControlsState();
  });
});

function loadRows(rows) {
  if (!rows || rows.length === 0) {
    showError("No rows found in the data.");
    return false;
  }

  rawRows = rows;
  applyDefaultYBoundsForSelectedColumn();

  const firstRow = rows[0];
  const columns = firstRow ? Object.keys(firstRow) : [];

  if (!columns || columns.length === 0) {
    showError("Could not detect any columns in the data.");
    return false;
  }

  // Ensure global column list is updated (used by column intelligence)
  allColumns = columns.slice();

// --- Reset dropdowns ---
if (dateSelect) dateSelect.innerHTML = "";
if (valueSelect) valueSelect.innerHTML = "";
if (thirdSelect) thirdSelect.innerHTML = "";
if (splitPointSelect) splitPointSelect.innerHTML = "";

  // (Optional extra selects - keep safe; these may exist in older versions)
  if (typeof numeratorSelect !== "undefined" && numeratorSelect) numeratorSelect.innerHTML = "";
  if (typeof denominatorSelect !== "undefined" && denominatorSelect) denominatorSelect.innerHTML = "";
  if (typeof subgroupSelect !== "undefined" && subgroupSelect) subgroupSelect.innerHTML = "";
  if (typeof eventDateSelect !== "undefined" && eventDateSelect) eventDateSelect.innerHTML = "";
  if (typeof oppBetweenSelect !== "undefined" && oppBetweenSelect) oppBetweenSelect.innerHTML = "";

  // ------------------------------------------------------------
  // LEVEL 3 FOUNDATION: profile the dataset columns once
  // ------------------------------------------------------------
  if (typeof profileColumns === "function") {
    profileColumns(rows);
  } else {
    // If profiling is missing for some reason, fall back to basic profiles
    columnProfiles = {};
  }

  // Determine the chart type currently selected (without changing state)
  const chartTypeNow =
    (typeof getSelectedChartType_NoSideEffects === "function")
      ? getSelectedChartType_NoSideEffects()
      : (typeof getSelectedChartType === "function")
        ? getSelectedChartType()
        : "run";

  // ------------------------------------------------------------
  // LEVELS 1–2: populate dropdowns with filtering + smart defaults
  // ------------------------------------------------------------
    if (typeof applyColumnIntelligence === "function") {
    applyColumnIntelligence(chartTypeNow);
    if (typeof applyChartTypeAvailability === "function") {
      applyChartTypeAvailability();
    }
  } else {
    // Fallback: populate all dropdowns with all columns (old behaviour)
    columns.forEach((col) => {
      if (dateSelect) {
        const opt1 = document.createElement("option");
        opt1.value = col;
        opt1.textContent = col;
        dateSelect.appendChild(opt1);
      }
      if (valueSelect) {
        const opt2 = document.createElement("option");
        opt2.value = col;
        opt2.textContent = col;
        valueSelect.appendChild(opt2);
      }
      if (thirdSelect) {
        const opt3 = document.createElement("option");
        opt3.value = col;
        opt3.textContent = col;
        thirdSelect.appendChild(opt3);
      }
    });
  }



  // ------------------------------------------------------------
  // Optional: keep your older "guessColumns" logic ONLY as a fallback
  // (i.e., if defaults didn't get set by column intelligence)
  // ------------------------------------------------------------
  const needDateDefault = dateSelect && !dateSelect.value;
  const needValueDefault = valueSelect && !valueSelect.value;

  if ((needDateDefault || needValueDefault) && typeof guessColumns === "function") {
    const guessed = guessColumns(rows);

    // Guess X
    if (needDateDefault && guessed && guessed.dateCol && dateSelect) {
      dateSelect.value = guessed.dateCol;
    } else if (needDateDefault && dateSelect) {
      // Prefer date-like if present, otherwise first column
      const best = (typeof getBestXAxisColumn === "function") ? getBestXAxisColumn() : (columns[0] || "");
      if (best) dateSelect.value = best;
    }

    // Guess Y
    if (needValueDefault && guessed && guessed.valueCol && valueSelect) {
      valueSelect.value = guessed.valueCol;
    } else if (needValueDefault && valueSelect) {
      // Fall back to first available option after blank
      const opts = Array.from(valueSelect.options).filter(o => o.value);
      if (opts[0]) valueSelect.value = opts[0].value;
    }
  }

  updateDateFormatWarning();

  // Optional: if no date-like column, don't nag, but you can keep your tip
  // (only show if you want — comment out if noisy)
  if (dateSelect && dateSelect.value) {
    const p = getProfile ? getProfile(dateSelect.value) : null;
    if (p && !p.looksLikeDate) {
      // Do nothing by default. If you WANT the old tip, uncomment below:
      // if (typeof setAxisType === "function") setAxisType("sequence");
      // showError("Tip: No date column detected. I’ll treat the data as a simple sequence (run chart by order).");
    }
  }

  // Show selectors safely
  if (columnSelectors) {
    columnSelectors.style.display = "block";
  }

  // Hide "load data first" hint safely (if present)
  const hint = document.getElementById("noDataYetHint");
  if (hint) hint.style.display = "none";

    // If user imported settings before loading data, apply them now
  if (pendingImportedSettings) {
    const toApply = pendingImportedSettings;
    pendingImportedSettings = null;
    applyToolSettings(toApply, { silent: false });
  }

  return true;

}

function getDefaultYBoundsForSelectedColumn() {
  return { min: "", max: "" };
}

function applyDefaultYBoundsForSelectedColumn() {
  const { min, max } = getDefaultYBoundsForSelectedColumn();

  if (yAxisMinInput) yAxisMinInput.value = min;
  if (yAxisMaxInput) yAxisMaxInput.value = max;
}

function updateYAxisInputStep() {
  if (!yAxisMinInput || !yAxisMaxInput || !yAxisFormatInput) return;

  const format = (yAxisFormatInput.value || "").toLowerCase();

  let stepValue = "1";

  // Percent values are stored as proportions:
  // 0.01 = 1.0%, 0.001 = 0.1%, 0.05 = 5.0%
  if (format === "percent") {
    const minRaw = parseFloat(yAxisMinInput.value);
    const maxRaw = parseFloat(yAxisMaxInput.value);

    let rangeRaw = NaN;

    if (Number.isFinite(minRaw) && Number.isFinite(maxRaw) && maxRaw > minRaw) {
      rangeRaw = maxRaw - minRaw;
    } else if (currentChart?.scales?.y) {
      const yScale = currentChart.scales.y;
      if (Number.isFinite(yScale.min) && Number.isFinite(yScale.max) && yScale.max > yScale.min) {
        rangeRaw = yScale.max - yScale.min;
      }
    }

    // Default: 1 percentage point
    stepValue = "0.01";

    // Small percent ranges: use 0.1 percentage point steps
    if (Number.isFinite(rangeRaw) && rangeRaw <= 0.05) {
      stepValue = "0.001";
    }

    // Large percent ranges: use 5 percentage point steps
    if (Number.isFinite(rangeRaw) && rangeRaw >= 0.25) {
      stepValue = "0.05";
    }
  }

  yAxisMinInput.step = stepValue;
  yAxisMaxInput.step = stepValue;
}

function applyCurrentChartYBoundsToInputs(chart = currentChart) {
  if (!chart?.scales?.y) return;

  const yScale = chart.scales.y;

  if (yAxisMinInput) yAxisMinInput.value = String(yScale.min);
  if (yAxisMaxInput) yAxisMaxInput.value = String(yScale.max);
}

function applyCurrentChartYBoundsToInputs(chart = currentChart) {
  if (!chart?.scales?.y) return;

  const yScale = chart.scales.y;

  if (yAxisMinInput) yAxisMinInput.value = String(yScale.min);
  if (yAxisMaxInput) yAxisMaxInput.value = String(yScale.max);
}

function showError(msg) {
  if (errorMessage) errorMessage.textContent = msg;
}

function showChartMessage(msg) {
  // Safe alias used by some newer validation code
  showError(msg);
}

function clearError() {
  if (errorMessage) errorMessage.textContent = "";
}

function getTargetValue() {
  if (!targetInput) return null;

  const v = targetInput.value.trim();
  if (v === "") return null;

  const num = Number(v);
  return isFinite(num) ? num : null;
}



if (targetToggleBtn) {
  updateTargetToggleBtn();
  targetToggleBtn.addEventListener("click", () => {
    targetEnabled = !targetEnabled;
    updateTargetToggleBtn();
    if (currentChart) generateButton.click();
  });
}

const debouncedRegen = debounce(() => {
  if (rawRows && rawRows.length) {
    lastGenerateWasManual = false;
    generateButton.click();
  }
}, 250);


if (baselineInput) {
  baselineInput.addEventListener("input", debouncedRegen);
  baselineInput.addEventListener("change", debouncedRegen);
}

if (shiftRulePointsInput) {
  shiftRulePointsInput.addEventListener("input", debouncedRegen);
  shiftRulePointsInput.addEventListener("change", debouncedRegen);
}
if (trendRulePointsInput) {
  trendRulePointsInput.addEventListener("input", debouncedRegen);
  trendRulePointsInput.addEventListener("change", debouncedRegen);
}
if (enableAdvancedTrendCheckbox) {
  enableAdvancedTrendCheckbox.addEventListener("change", () => {
    const chartType = getSelectedChartType_NoSideEffects();

    if (typeof updateRuleUIForChartType === "function") {
      updateRuleUIForChartType(chartType);
    }

    if (typeof debouncedRegen === "function") {
      debouncedRegen();
    }
  });
}
if (flagSpecialCauseOnChartCheckbox) {
  flagSpecialCauseOnChartCheckbox.addEventListener("change", () => {
    if (rawRows && rawRows.length) generateButton.click();
  });
}
if (clampLclAtZeroCheckbox) {
  clampLclAtZeroCheckbox.addEventListener("change", () => {
    if (rawRows && rawRows.length) generateButton.click();
  });
}

if (screenLargeMRsCheckbox) {
  screenLargeMRsCheckbox.addEventListener("change", () => {
    const chartType = getSelectedChartType_NoSideEffects();

    if (
      chartType === "xmr" &&
      rawRows &&
      rawRows.length &&
      generateButton
    ) {
      generateButton.click();
    }
  });
}

if (ruleTwoOfThreeOuterThirdCheckbox) {
  ruleTwoOfThreeOuterThirdCheckbox.addEventListener("change", debouncedRegen);
}

if (ruleFourOfFiveOneSigmaCheckbox) {
  ruleFourOfFiveOneSigmaCheckbox.addEventListener("change", debouncedRegen);
}

if (enableRareRunTrendCheckbox) {
  enableRareRunTrendCheckbox.addEventListener("change", () => {
    const chartType = getSelectedChartType_NoSideEffects();

    if (!confirmEnableRareRunTrendOnce(chartType)) {
      if (typeof updateRuleUIForChartType === "function") {
        updateRuleUIForChartType(chartType);
      }
      return;
    }

    if (typeof updateRuleUIForChartType === "function") {
      updateRuleUIForChartType(chartType);
    }

    if (typeof debouncedRegen === "function") {
      debouncedRegen();
    }
  });
}

const recalcPrompt = document.getElementById("recalcPrompt");
const firstRunGuide = document.getElementById("firstRunGuide");

// Treat the Quick start panel as useful empty-state guidance,
// rather than something shown only once per browser.
function updateFirstRunGuideVisibility() {
  if (!firstRunGuide) return;

  firstRunGuide.style.display = currentChart ? "none" : "block";
}

function markFirstRunComplete() {
  if (!firstRunGuide) return;

  firstRunGuide.style.display = "none";
}

// Kept for compatibility with existing reset code.
function clearFirstRunFlag() {
  updateFirstRunGuideVisibility();
}

function getSelectedChartType() {
  const el = document.querySelector('input[name="chartType"]:checked');
  return el ? el.value : "run";
}


// On initial load
updateFirstRunGuideVisibility();


// --- Replace the recalc prompt with a red button state ---
function setGenerateNeedsRecalc(needs) {
  if (!generateButton) return;
  generateButton.classList.toggle("needs-recalc", !!needs);
  generateButton.title = needs ? "Changes saved — click Generate / Recalculate" : "";
}

function markDataModelDirty() {
  dataModelDirty = true;

  // Don’t show extra red text; just make the button obvious
  setGenerateNeedsRecalc(true);

  // Optional: keep errors for *real* errors only (recommended)
  // (so don't call showError here)
}

function clearDataModelDirty() {
  dataModelDirty = false;
  setGenerateNeedsRecalc(false);

  if (typeof updateSaveChartButtonState === "function") {
    updateSaveChartButtonState();
  }

  // don’t clearError() automatically; user may still want to see tips
}

// On initial load, show guide only until first successful generate
updateFirstRunGuideVisibility();

window.addEventListener("beforeunload", (e) => {
  // If you have a boolean dirty flag, use it here.
  // Fallback: warn if a chart exists (user did work)
  const shouldWarn =
    (typeof isDataModelDirty === "function" && isDataModelDirty()) ||
    !!currentChart;

  if (!shouldWarn) return;

  e.preventDefault();
  e.returnValue = "";
});



//---- Add annotations button

if (addAnnotationBtn) {
  addAnnotationBtn.addEventListener("click", () => {
    if (!annotationDateInput || !annotationLabelInput) return;

    const dateVal = annotationDateInput.value;
    const labelVal = annotationLabelInput.value.trim();

    if (!dateVal || !labelVal) {
      alert("Please enter both a date and a label for the annotation.");
      return;
    }

    // Dates from <input type="date"> are already 'YYYY-MM-DD'
    annotations.push({ date: dateVal, label: labelVal, yAdjust: null });

	// Clear just the label field, keep the date selection
	annotationLabelInput.value = "";

    // Re-generate the chart with the new annotation
    generateButton.click();
  });
}

//---- Clear annotations button
if (clearAnnotationsBtn) {
  clearAnnotationsBtn.addEventListener("click", () => {
    annotations = [];

    if (annotationDateInput) annotationDateInput.value = "";
    if (annotationLabelInput) annotationLabelInput.value = "";

    // If a chart already exists, re-generate it to remove the lines
    if (currentChart) {
      generateButton.click();
    }
  });
}

// ---- Toggle sidebar button ----
if (toggleSidebarButton) {
  toggleSidebarButton.addEventListener("click", () => {
    const collapsed = document.body.classList.toggle("sidebar-collapsed");
    toggleSidebarButton.textContent = collapsed ? "Show controls" : "Hide controls";
  });
}

// ---- CSV / Excel upload & column selection ----
fileInput.addEventListener("change", async () => {
  const file = fileInput.files[0];
  if (!file) return;

  clearError();
  if (summaryDiv) summaryDiv.innerHTML = "";
  if (capabilityDiv) capabilityDiv.innerHTML = "";

  const lowerName = String(file.name || "").toLowerCase();
  const isExcel = lowerName.endsWith(".xlsx") || lowerName.endsWith(".xls");
  const isCsv = lowerName.endsWith(".csv");

  try {
        if (isExcel) {
      const workbook = await readExcelWorkbook(file);
      openExcelWorkbookInDataEditor(workbook, workbook.SheetNames[0] || "");
      return;
    }

    if (!isCsv) {
      showError("Please upload a CSV or Excel file (.csv, .xlsx, .xls).");
      return;
    }

    const text = await file.text();
    const parsed = parseTabularTextWithHeaderDetection(text);

    if (!parsed.ok) {
      showError("Error parsing CSV: " + parsed.message);
      return;
    }

    if (!parsed.hadHeader && parsed.rows2D && parsed.rows2D.length >= 2) {
      const r0 = parsed.rows2D[0];
      const r1 = parsed.rows2D[1];

      const score0 = rowDataLikenessScore(r0);
      const duplicateHeaderRow = rowsEqualNormalized(r0, r1) && score0 <= 0.2;

      if (duplicateHeaderRow) {
        const results = Papa.parse(text, { header: true, dynamicTyping: true, skipEmptyLines: true });

        if (results.errors && results.errors.length > 0) {
          console.error(results.errors);
          showError("Error parsing CSV: " + results.errors[0].message);
          return;
        }

        let rows = results.data || [];
        const headers = results.meta && results.meta.fields ? results.meta.fields : null;
        rows = stripDuplicateHeaderRow(rows, headers);

	if (!prepareForNewImportedChart(rows)) return;
if (!loadRows(rows)) return;

resetStateAfterDataLoad();
showNewDataChartPrompt();

return;
      }
    }

    if (parsed.hadHeader) {
  if (!prepareForNewImportedChart(parsed.rows)) return;
  if (!loadRows(parsed.rows)) return;
} else {
      const ok = confirm(
        "It looks like your CSV does not include column headings.\n\n" +
        "Click OK to treat the first row as DATA (I will create Column1, Column2...).\n" +
        "Click Cancel if the first row IS a header row (then add headings and upload again)."
      );

      if (!ok) {
        showError("Please add a header row (e.g. Date,Value) and upload again.");
        return;
      }

      const data2D = parsed.rows2D;
      const colCount = Math.max(...data2D.map(r => r.length));
      const headers = Array.from({ length: colCount }, (_, i) => `Column${i + 1}`);

      const objRows = data2D.map(r => {
        const o = {};
        headers.forEach((h, i) => (o[h] = r[i]));
        return o;
      });

      if (!prepareForNewImportedChart(objRows)) return;
      if (!loadRows(objRows)) return;
    }

    resetStateAfterDataLoad();
showNewDataChartPrompt();
  } catch (err) {
    console.error(err);
    showError(
      isExcel
        ? "Unexpected error reading the Excel file."
        : "Unexpected error reading the CSV file."
    );
  }
});

function getMrDisplayMode() {
  const el = document.querySelector("input[name='mrDisplayMode']:checked");
  return el ? el.value : "last";
}


// -----------------------------
// Math helpers for X̄–S constants
// -----------------------------

function gammaLanczos(z) {
  // Lanczos approximation for Gamma(z)
  // Good enough for SPC constants.
  const p = [
    676.5203681218851,
    -1259.1392167224028,
    771.32342877765313,
    -176.61502916214059,
    12.507343278686905,
    -0.13857109526572012,
    0.0000099843695780195716,
    0.00000015056327351493116
  ];
  const g = 7;

  if (z < 0.5) {
    return Math.PI / (Math.sin(Math.PI * z) * gammaLanczos(1 - z));
  }

  z -= 1;
  let x = 0.99999999999980993;
  for (let i = 0; i < p.length; i++) {
    x += p[i] / (z + i + 1);
  }
  const t = z + g + 0.5;
  return Math.sqrt(2 * Math.PI) * Math.pow(t, z + 0.5) * Math.exp(-t) * x;
}

function c4Constant(n) {
  if (!Number.isFinite(n) || n < 2) return NaN;
  // c4 = sqrt(2/(n-1)) * Gamma(n/2) / Gamma((n-1)/2)
  return Math.sqrt(2 / (n - 1)) * (gammaLanczos(n / 2) / gammaLanczos((n - 1) / 2));
}

function xbarSConstants(n) {
  const c4 = c4Constant(n);
  if (!isFinite(c4) || c4 <= 0) return null;

  const term = Math.sqrt(Math.max(1 - c4 * c4, 0)) / c4;

  const A3 = 3 / (c4 * Math.sqrt(n));
  const B3 = Math.max(0, 1 - 3 * term);
  const B4 = 1 + 3 * term;

  return { c4, A3, B3, B4 };
}

// Estimate process sigma for an X̄–S chart from subgroup standard deviations.
// This is the variable-subgroup-size extension of the existing S-bar method.
// When all subgroup sizes are equal, it reduces to s-bar / c4(n).
function estimateXbarSSigmaFromSbar(sValues, subgroupSizes) {
  if (
    !Array.isArray(sValues) ||
    !Array.isArray(subgroupSizes) ||
    sValues.length !== subgroupSizes.length ||
    sValues.length === 0
  ) {
    return NaN;
  }

  let numerator = 0;
  let denominator = 0;

  for (let i = 0; i < sValues.length; i++) {
    const s = Number(sValues[i]);
    const n = Number(subgroupSizes[i]);

    if (!Number.isFinite(s) || !Number.isFinite(n) || n < 2) continue;

    const c4 = c4Constant(n);
    if (!Number.isFinite(c4) || c4 <= 0 || c4 >= 1) continue;

    const varianceTerm = 1 - c4 * c4;
    if (!(varianceTerm > 0)) continue;

    const h = (c4 * c4) / varianceTerm;

    numerator += h * (s / c4);
    denominator += h;
  }

  return denominator > 0
    ? numerator / denominator
    : NaN;
}

// -----------------------------
// Draw a second chart in the MR panel (re-uses the existing mrPanel UI)
// -----------------------------
function buildXbarSSplitMarkerPlugin(splitAfterIndices = []) {
  const safeSplits = Array.isArray(splitAfterIndices)
    ? splitAfterIndices.filter(
        i =>
          Number.isInteger(i) &&
          i >= 0
      )
    : [];

  return {
    id: "xbarSSplitMarkers",

    afterDraw(chart) {
      if (!safeSplits.length) return;

      const ctx = chart.ctx;
      const chartArea = chart.chartArea;
      const xScale = chart.scales?.x;
      const labels =
        chart.data?.labels || [];

      if (
        !ctx ||
        !chartArea ||
        !xScale ||
        labels.length < 2
      ) {
        return;
      }

      ctx.save();

      ctx.strokeStyle =
        "rgba(80,80,80,0.65)";

      ctx.fillStyle =
        "rgba(80,80,80,0.85)";

      ctx.lineWidth = 1;
      ctx.setLineDash([5, 4]);

      ctx.font =
        "600 10px Arial, sans-serif";

      ctx.textAlign = "center";
      ctx.textBaseline = "bottom";

      safeSplits.forEach(
        (splitAfterIndex) => {
          if (
            splitAfterIndex >=
            labels.length - 1
          ) {
            return;
          }

          const leftX =
            xScale.getPixelForValue(
              splitAfterIndex
            );

          const rightX =
            xScale.getPixelForValue(
              splitAfterIndex + 1
            );

          if (
            !Number.isFinite(leftX) ||
            !Number.isFinite(rightX)
          ) {
            return;
          }

          const x =
            (leftX + rightX) / 2;

          ctx.beginPath();

          ctx.moveTo(
            x,
            chartArea.top
          );

          ctx.lineTo(
            x,
            chartArea.bottom
          );

          ctx.stroke();

          ctx.fillText(
            "Split",
            x,
            chartArea.top - 3
          );
        }
      );

      ctx.restore();
    }
  };
}

function buildXbarSSplitSegmentStyle(
  baseColour,
  splitAfterIndices = []
) {
  const splitSet =
    new Set(
      Array.isArray(
        splitAfterIndices
      )
        ? splitAfterIndices.filter(
            i =>
              Number.isInteger(i) &&
              i >= 0
          )
        : []
    );

  if (!splitSet.size) {
    return undefined;
  }

  return {
    borderColor: (ctx) =>
      splitSet.has(
        ctx.p0DataIndex
      )
        ? "rgba(0,0,0,0)"
        : baseColour
  };
}

function drawSecondarySPCChart({
  canvas,
  labels,
  values,
  pointColours,
  cl,
  ucl,
  lcl,
  title,
  xLabel,
  yLabel,
  suggestedMin,
  suggestedMax,
  splitAfterIndices = []
}) {
  if (!canvas) return null;

  const datasets = [
    {
      label: "Value",
      data: values,
      borderColor:
        SPC_STYLE.seriesBlue,
      borderWidth: 2,
      fill: false,
      pointRadius: 4,
      pointBackgroundColor:
        pointColours,
      pointBorderColor:
        pointColours,
      tension: 0.1
    },

    {
      label: "Centre line",
      data: cl,
      borderColor:
        SPC_STYLE.centreRed,
      borderDash: [6, 4],
      borderWidth: 2,
      pointRadius: 0,
      pointHoverRadius: 0,

      segment:
        buildXbarSSplitSegmentStyle(
          SPC_STYLE.centreRed,
          splitAfterIndices
        )
    },

    {
      label: "UCL",
      data: ucl,
      borderColor:
        SPC_STYLE.limitGreen,
      borderDash: [4, 4],
      borderWidth: 2,
      pointRadius: 0,
      pointHoverRadius: 0,

      segment:
        buildXbarSSplitSegmentStyle(
          SPC_STYLE.limitGreen,
          splitAfterIndices
        )
    },

    {
      label: "LCL",
      data: lcl,
      borderColor:
        SPC_STYLE.limitGreen,
      borderDash: [4, 4],
      borderWidth: 2,
      pointRadius: 0,
      pointHoverRadius: 0,

      segment:
        buildXbarSSplitSegmentStyle(
          SPC_STYLE.limitGreen,
          splitAfterIndices
        )
    }
  ];

  const splitMarkerPlugin =
    buildXbarSSplitMarkerPlugin(
      splitAfterIndices
    );

  return new Chart(
    canvas.getContext("2d"),
    {
      type: "line",

      data: {
        labels,
        datasets
      },

      plugins: [
        splitMarkerPlugin
      ],

      options: {
        responsive: true,
        maintainAspectRatio: false,

        plugins: {
          title:
            buildChartTitleConfig(
              title
            ),

          legend:
            SPC_LEGEND,

          annotation: {
            annotations:
              (
                typeof buildAnnotationConfig ===
                "function"
              )
                ? buildAnnotationConfig(
                    labels
                  )
                : {}
          }
        },

        elements: {
          point: {
            radius: 0,
            hoverRadius: 0
          }
        },

        scales: (() => {
          const axisSettings =
            getAxisSettings();

          return {
            x:
              buildCategoryXAxisConfig(
                xLabel,
                axisSettings.x,
                labels
              ),

            y:
              buildAxisConfig(
                yLabel,
                withoutAxisBounds(
                  axisSettings.y
                ),
                {
                  suggestedMin:
                    isFinite(
                      suggestedMin
                    )
                      ? suggestedMin
                      : undefined,

                  suggestedMax:
                    isFinite(
                      suggestedMax
                    )
                      ? suggestedMax
                      : undefined
                }
              )
          };
        })()
      }
    }
  );
}

// -----------------------------
// X̄–S combined chart renderer
// - Main canvas: X̄ chart
// - MR panel canvas: S chart (re-uses existing UI)
// -----------------------------
function drawXbarSCombinedChart({
  labels,
  xbarVals,
  sVals,
  pointColoursX,
  pointColoursS,
  clX,
  uclXArr,
  lclXArr,
  clS,
  uclSArr,
  lclSArr,
  splitAfterIndices = []
}) {
  if (!chartCanvas) return;

  if (
    typeof populateAnnotationDateOptions ===
    "function"
  ) {
    populateAnnotationDateOptions(
      labels
    );
  }

  if (
    typeof populateSplitOptions ===
    "function"
  ) {
    populateSplitOptions(
      labels
    );
  }

  // -------------------------
  // 1) Main chart: X̄ chart
  // -------------------------
  if (currentChart) {
    currentChart.destroy();
    currentChart = null;
  }

  const mainLabels =
    getChartLabels(
      "X̄ chart",
      "Subgroup",
      "X̄"
    );

  const mainDatasets = [
    {
      label: "X̄",
      data: xbarVals,
      borderColor:
        SPC_STYLE.seriesBlue,
      borderWidth: 2,
      fill: false,
      pointRadius: 4,
      pointBackgroundColor:
        pointColoursX,
      pointBorderColor:
        pointColoursX,
      tension: 0.1
    },

    {
      label: "Centre line",
      data: clX,
      borderColor:
        SPC_STYLE.centreRed,
      borderDash: [6, 4],
      borderWidth: 2,
      pointRadius: 0,
      pointHoverRadius: 0,

      segment:
        buildXbarSSplitSegmentStyle(
          SPC_STYLE.centreRed,
          splitAfterIndices
        )
    },

    {
      label: "UCL",
      data: uclXArr,
      borderColor:
        SPC_STYLE.limitGreen,
      borderDash: [4, 4],
      borderWidth: 2,
      pointRadius: 0,
      pointHoverRadius: 0,

      segment:
        buildXbarSSplitSegmentStyle(
          SPC_STYLE.limitGreen,
          splitAfterIndices
        )
    },

    {
      label: "LCL",
      data: lclXArr,
      borderColor:
        SPC_STYLE.limitGreen,
      borderDash: [4, 4],
      borderWidth: 2,
      pointRadius: 0,
      pointHoverRadius: 0,

      segment:
        buildXbarSSplitSegmentStyle(
          SPC_STYLE.limitGreen,
          splitAfterIndices
        )
    }
  ];

  const centreLineLabelPlugin =
    buildXmRMeanLabelPlugin();

  const splitMarkerPlugin =
    buildXbarSSplitMarkerPlugin(
      splitAfterIndices
    );

  currentChart =
    new Chart(
      chartCanvas.getContext("2d"),
      {
        type: "line",

        data: {
          labels,
          datasets:
            mainDatasets
        },

        plugins: [
          centreLineLabelPlugin,
          splitMarkerPlugin
        ],

        options: {
          responsive: true,
          maintainAspectRatio: false,

          plugins: {
            title:
              buildChartTitleConfig(
                mainLabels.title
              ),

            legend:
              SPC_LEGEND,

            annotation: {
              annotations:
                (
                  typeof buildAnnotationConfig ===
                  "function"
                )
                  ? buildAnnotationConfig(
                      labels
                    )
                  : {}
            }
          },

          elements: {
            point: {
              radius: 0,
              hoverRadius: 0
            }
          },

          scales: (() => {
            const axisSettings =
              getAxisSettings();

            return {
              x:
                buildCategoryXAxisConfig(
                  mainLabels.xLabel,
                  axisSettings.x,
                  labels
                ),

              y:
                buildAxisConfig(
                  mainLabels.yLabel,
                  axisSettings.y
                )
            };
          })()
        }
      }
    );

  // -------------------------
  // 2) Secondary chart: S chart
  // -------------------------
  if (mrChart) {
    mrChart.destroy();
    mrChart = null;
  }

  const showSecondary =
    showMRCheckbox
      ? showMRCheckbox.checked
      : true;

  if (!showSecondary) {
    if (mrPanel) {
      mrPanel.style.display =
        "none";
    }

    return;
  }

  if (mrPanel) {
    mrPanel.style.display =
      "block";

    const strong =
      mrPanel.querySelector(
        "strong"
      );

    if (strong) {
      strong.textContent =
        "S chart:";
    }
  }

  if (!mrCanvas) return;

  
  // The secondary S chart needs its own default labels,
  // but must not overwrite the shared Labels-tab fields.
  // Preserve any labels deliberately entered by the user.
  const sLabels = {
    title: chartTitleManuallyEdited
      ? mainLabels.title
      : "S chart",

    xLabel: xAxisLabelManuallyEdited
      ? mainLabels.xLabel
      : "Subgroup",

    yLabel: yAxisLabelManuallyEdited
      ? mainLabels.yLabel
      : "S"
  };


  mrChart =
    drawSecondarySPCChart({
      canvas:
        mrCanvas,

      labels,

      values:
        sVals,

      pointColours:
        pointColoursS,

      cl:
        clS,

      ucl:
        uclSArr,

      lcl:
        lclSArr,

      title:
        sLabels.title,

      xLabel:
        sLabels.xLabel,

      yLabel:
        sLabels.yLabel,

      suggestedMin:
        0,

      suggestedMax:
        undefined,

      splitAfterIndices
    });
}



function resetAll() {
  // --- Clear stored data ---
  rawRows = [];
annotations = [];
splits = [];
limitExclusions = [];
resetXmRMeanLabelState();

// Clear all cached analysis from the previous chart
lastXmRAnalysis = null;
lastRunAnalysis = null;
lastAttributeAnalysis = null;
lastRareAnalysis = null;
lastXbarSAnalysis = null;

// A new chart is no longer carrying unsaved changes
dataModelDirty = false;

  // --- Reset file input ---
  if (fileInput) fileInput.value = "";

  // --- Hide column selectors ---
  if (columnSelectors) columnSelectors.style.display = "none";

    // --- Reset dropdowns ---
  if (dateSelect) dateSelect.innerHTML = "";
  if (valueSelect) valueSelect.innerHTML = "";
  if (thirdSelect) thirdSelect.innerHTML = "";
  if (splitPointSelect) splitPointSelect.innerHTML = "";

  // --- Reset optional third-column UI ---
  if (thirdColumnRow) thirdColumnRow.style.display = "none";
  if (thirdLabelEl) thirdLabelEl.textContent = "Denominator / opportunities";
  if (thirdHintEl) thirdHintEl.textContent = "";
  if (uRateMultiplierSelect) {
  uRateMultiplierSelect.value = "1";
}

lastURateMultiplier = 1;

if (uRateMultiplierRow) {
  uRateMultiplierRow.style.display = "none";
}

  // --- Reset text inputs ---
  if (baselineInput) baselineInput.value = "";
  if (chartTitleInput) chartTitleInput.value = "";
  if (xAxisLabelInput) xAxisLabelInput.value = "";
  if (yAxisLabelInput) yAxisLabelInput.value = "";
  if (chartTitleFontFamilyInput) chartTitleFontFamilyInput.value = "";
  if (chartTitleFontSizeInput) chartTitleFontSizeInput.value = "16";
  if (showChartTitleCheckbox) showChartTitleCheckbox.checked = true;
  setPressed(chartTitleBoldBtn, true);
  setPressed(chartTitleItalicBtn, false);
  setPressed(chartTitleUnderlineBtn, false);
  if (targetInput) targetInput.value = "";

// A fresh/reset chart should show a target by default once one is entered.
targetEnabled = true;
if (typeof updateTargetToggleBtn === "function") {
  updateTargetToggleBtn();
}
if (typeof updateTargetToggleVisibility === "function") {
  updateTargetToggleVisibility();
}

if (annotationDateInput) annotationDateInput.value = "";
if (annotationLabelInput) annotationLabelInput.value = "";

  chartTitleManuallyEdited = false;
  xAxisLabelManuallyEdited = false;
  yAxisLabelManuallyEdited = false;

  // --- Reset axis controls ---
  if (xAxisFontFamilyInput) xAxisFontFamilyInput.value = "";
  if (xAxisFontSizeInput) xAxisFontSizeInput.value = "11";
  if (typeof setPressed === "function") {
    setPressed(xAxisItalicBtn, false);
    setPressed(xAxisBoldBtn, false);
  }

  if (yAxisMinInput) yAxisMinInput.value = "";
  if (yAxisMaxInput) yAxisMaxInput.value = "";
  if (yAxisFormatInput) yAxisFormatInput.value = "auto";
  if (yAxisDecimalsInput) yAxisDecimalsInput.value = "auto";
  if (yAxisTickStepInput) yAxisTickStepInput.value = "";
  if (dateFormatPreferenceSelect) dateFormatPreferenceSelect.selectedIndex = 0;
  if (yAxisFontFamilyInput) yAxisFontFamilyInput.value = "";
  if (yAxisFontSizeInput) yAxisFontSizeInput.value = "11";
  if (typeof setPressed === "function") {
    setPressed(yAxisItalicBtn, false);
    setPressed(yAxisBoldBtn, false);
  }

  // --- Reset target direction dropdown ---
  if (targetDirectionInput) targetDirectionInput.value = "above";

  // --- Reset axis + chart type radios to defaults (match first load HTML) ---
  const axisDateRadio = document.querySelector("input[name='axisType'][value='date']");
  const axisSeqRadio = document.querySelector("input[name='axisType'][value='sequence']");
  if (axisDateRadio) axisDateRadio.checked = true;
  if (axisSeqRadio) axisSeqRadio.checked = false;

  const runRadio = document.querySelector("input[name='chartType'][value='run']");
  const xmrRadio = document.querySelector("input[name='chartType'][value='xmr']");
  if (runRadio) runRadio.checked = true;
  if (xmrRadio) xmrRadio.checked = false;

  if (typeof updateUIForChartType === "function") {
    updateUIForChartType("run");
  }

  const moreChartTypesDetails = document.getElementById("moreChartTypesDetails");
  if (moreChartTypesDetails) moreChartTypesDetails.open = false;

  // MR toggle default (match first load)
  const showMRCheckbox = document.getElementById("showMRCheckbox");
  if (showMRCheckbox) showMRCheckbox.checked = true;

  // --- Reset Rules & interpretation defaults ---
  const shiftRulePointsInput = document.getElementById("shiftRulePoints");
  const trendRulePointsInput = document.getElementById("trendRulePoints");
  const flagSpecialCauseOnChart = document.getElementById("flagSpecialCauseOnChart");
  const clampLclAtZero = document.getElementById("clampLclAtZero");
  const lclClampRow = document.getElementById("lclClampRow");

  if (shiftRulePointsInput) shiftRulePointsInput.value = "8";
if (trendRulePointsInput) trendRulePointsInput.value = "6";
if (flagSpecialCauseOnChart) flagSpecialCauseOnChart.checked = true;
if (clampLclAtZero) clampLclAtZero.checked = false;
if (lclClampRow) lclClampRow.style.display = "none";

// Advanced rules are off by default on a new chart
if (enableAdvancedTrendCheckbox) enableAdvancedTrendCheckbox.checked = false;
if (enableRareRunTrendCheckbox) enableRareRunTrendCheckbox.checked = false;
if (ruleTwoOfThreeOuterThirdCheckbox) ruleTwoOfThreeOuterThirdCheckbox.checked = false;
if (ruleFourOfFiveOneSigmaCheckbox) ruleFourOfFiveOneSigmaCheckbox.checked = false;
if (screenLargeMRsCheckbox) {
  screenLargeMRsCheckbox.checked = true;
}

if (advancedRulesDetails) advancedRulesDetails.open = false;

if (typeof updateRuleUIForChartType === "function") {
  updateRuleUIForChartType("run");
}

  // --- Clear any error message ---
  if (errorMessage) errorMessage.textContent = "";

if (typeof updateColumnCheckWarning === "function") {
  updateColumnCheckWarning("run");
}

  // --- Clear summary & capability output ---
  if (summaryDiv) summaryDiv.innerHTML = "";
  if (capabilityDiv) capabilityDiv.innerHTML = "";

  // --- Destroy main chart ---
  if (currentChart) {
    currentChart.destroy();
    currentChart = null;
  }

  updateSaveChartButtonState();

  // --- Destroy MR chart ---
  if (mrChart) {
    mrChart.destroy();
    mrChart = null;
  }

  // --- Hide MR panel ---
  if (mrPanel) mrPanel.style.display = "none";

  // --- Reset data editor ---
  if (dataEditorTextarea) dataEditorTextarea.value = "";
  if (dataEditorOverlay) dataEditorOverlay.style.display = "none";

  // Close help modal if open (don’t hide the section itself)
const helpModal = document.getElementById("helpModal");
if (helpModal) {
  helpModal.classList.remove("visible");
  helpModal.setAttribute("aria-hidden", "true");
}
document.body.classList.remove("modal-open");


  // --- Reset SPC helper (first-load behaviour) ---
  if (aiQuestionInput) aiQuestionInput.value = "";
  if (spcHelperOutput) spcHelperOutput.innerHTML = "";
  if (spcHelperPanel) spcHelperPanel.classList.remove("visible");

  // Re-render chip suggestions (safe even if helper never opened)
  if (typeof renderHelperState === "function") renderHelperState();

  // --- Ensure sidebar is visible (not collapsed) like first load ---
  document.body.classList.remove("sidebar-collapsed");
  const toggleBtn = document.getElementById("toggleSidebarButton");
  if (toggleBtn) toggleBtn.textContent = " Hide controls";

  const firstTab = document.querySelector(".tab-btn");
if (firstTab) firstTab.click();

  // Keep MR toggle visibility consistent with chart type default
  if (typeof updateMrToggleVisibility === "function") {
    updateMrToggleVisibility();
  }
  
if (typeof updateDateFormatWarning === "function") {
  updateDateFormatWarning();
}

if (typeof updateDateControlsState === "function") {
  updateDateControlsState();
}

// Keep the user's first-run/help history when starting another chart.
if (typeof updateFirstRunGuideVisibility === "function") {
  updateFirstRunGuideVisibility();
}

if (typeof setGenerateNeedsRecalc === "function") setGenerateNeedsRecalc(false);

  console.log("All elements reset.");
}

function resetChartKeepData() {
  // If there is no dataset to preserve, a normal full reset is enough.
  if (!Array.isArray(rawRows) || rawRows.length === 0) {
    resetAll();
    return;
  }

  // Keep the current dataset in memory.
  // resetAll() replaces rawRows with a new empty array, so a shallow
  // copy of the existing array is sufficient here.
  const savedRows = rawRows.slice();

  // Keep data-related choices rather than chart-specific settings.
  const savedXColumn = dateSelect?.value || "";
  const savedYColumn = valueSelect?.value || "";

  const savedAxisType =
    document.querySelector("input[name='axisType']:checked")?.value || "date";

  const savedDateFormat =
    dateFormatPreferenceSelect?.value || "";

  // Perform the normal, thoroughly defined clean reset.
  resetAll();

  // Put the existing dataset back.
  if (!loadRows(savedRows)) {
    showError("The existing data could not be restored after resetting the chart.");
    return;
  }

  // Restore X and Y column selections where those columns are still available.
  function restoreSelectValue(selectEl, value) {
    if (!selectEl || !value) return;

    const exists = Array.from(selectEl.options)
      .some(option => option.value === value && !option.disabled);

    if (exists) {
      selectEl.value = value;
    }
  }

  restoreSelectValue(dateSelect, savedXColumn);
  restoreSelectValue(valueSelect, savedYColumn);

  // Keep the user's interpretation of the X axis.
  const axisRadio = document.querySelector(
    `input[name="axisType"][value="${savedAxisType}"]`
  );

  if (axisRadio) {
    axisRadio.checked = true;
  }

  // Keep the date parsing preference because this belongs to the dataset,
  // rather than to the appearance/statistical setup of the chart.
  if (dateFormatPreferenceSelect && savedDateFormat) {
    const exists = Array.from(dateFormatPreferenceSelect.options)
      .some(option => option.value === savedDateFormat);

    if (exists) {
      dateFormatPreferenceSelect.value = savedDateFormat;
    }
  }

  if (typeof updateDateFormatWarning === "function") {
    updateDateFormatWarning();
  }

  if (typeof updateDateControlsState === "function") {
    updateDateControlsState();
  }

  // Y-axis limits have just been reset to automatic.
  yAxisBoundsManuallyEdited = false;

  // Generate the clean Run chart immediately from the retained data.
  if (generateButton) {
    generateButton.click();
  }
}

function validateBeforeGenerate() {
  if (!rawRows || rawRows.length === 0) {
    showError("No data loaded yet. Upload a CSV or use the data editor first.");
    return false;
  }

  const chartType = (typeof getSelectedChartType_NoSideEffects === "function")
    ? getSelectedChartType_NoSideEffects()
    : "run";

  const axisType = (typeof getAxisType === "function")
    ? getAxisType()
    : "date";

  const dateCol = dateSelect?.value;
  const valueCol = valueSelect?.value;

  if (!dateCol || !valueCol) {
    showError("Please choose both an X-axis column and a value column.");
    return false;
  }

  if (chartType === "pareto") {
    const paretoCheck =
      computeParetoData(
        rawRows,
        dateCol,
        valueCol
      );

    if (!paretoCheck.ok) {
      showError(
        paretoCheck.error
      );

      return false;
    }

    // Pareto is categorical, so the ordinary numeric-value
    // and date-axis safety checks below do not apply.
    return true;
  }

  // Check at least 3 valid numeric points in the selected value column.
  // (T chart event-dates mode still uses a value selector in the UI for now,
  // so keep this behaviour consistent until that workflow is changed.)
  let good = 0;
  for (const row of rawRows) {
    const y = toNumericValue(row[valueCol]);
    if (isFinite(y)) good++;
  }

  if (good < 3 && !(chartType === "t" && typeof tChartInputMode !== "undefined" && tChartInputMode === "eventDates")) {
    showError(
      "I can’t create a chart yet: I need at least 3 numeric values in the selected value column. " +
      "Check the column selection and make sure the values are numbers (e.g. 12.3 not '12,3' or text)."
    );
    return false;
  }

  // Safety warning for suspicious selections
  const safetyResult = validateColumnSelectionSafety({
    chartType,
    dateCol,
    valueCol,
    axisType
  });

  if (!handleValidationResult(safetyResult, { manual: lastGenerateWasManual })) {
    return false;
  }

  return true;
}


// ---- Helpers ----

function isPressed(btn) {
  return btn?.getAttribute("aria-pressed") === "true";
}

function setPressed(btn, pressed) {
  if (!btn) return;
  btn.setAttribute("aria-pressed", pressed ? "true" : "false");
}

function wireToggleButton(btn) {
  if (!btn) return;
  btn.addEventListener("click", () => {
    const next = !isPressed(btn);
    setPressed(btn, next);
    markDataModelDirty();
    if (rawRows && rawRows.length && currentChart) {
      applyPresentationEditsLiveDebounced();
    }
  });
}

wireToggleButton(chartTitleBoldBtn);
wireToggleButton(chartTitleItalicBtn);
wireToggleButton(chartTitleUnderlineBtn);

wireToggleButton(xAxisItalicBtn);
wireToggleButton(xAxisBoldBtn);
wireToggleButton(yAxisItalicBtn);
wireToggleButton(yAxisBoldBtn);

function getSelectedChartType_NoSideEffects() {
  const radios = document.querySelectorAll("input[name='chartType']");
  for (const r of radios) {
    if (r.checked) return r.value;
  }
  return "run";
}


function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function isRareChartType(chartType) {
  return chartType === "t" || chartType === "g";
}

function getRareRunTrendEnabled(chartType) {
  if (!isRareChartType(chartType)) return false;
  return !!enableRareRunTrendCheckbox?.checked;
}

function confirmEnableRareRunTrendOnce(chartType) {
  if (!isRareChartType(chartType)) return true;

  // Only needed if user is trying to enable it
  if (!enableRareRunTrendCheckbox?.checked) return true;

  const key = `spc_confirmedRareRunTrend_${chartType}`;
  let alreadyConfirmed = false;
  try { alreadyConfirmed = localStorage.getItem(key) === "true"; } catch {}

  if (alreadyConfirmed) return true;

  const msg =
    `Advanced option: Run & trend rules on ${chartType.toUpperCase()} charts\n\n` +
    `These patterns often happen by chance on T/G charts and can increase false alerts.\n\n` +
    `Enable anyway?`;

  const ok = confirm(msg);

  if (!ok) {
    // revert checkbox
    enableRareRunTrendCheckbox.checked = false;
    return false;
  }

  try { localStorage.setItem(key, "true"); } catch {}
  return true;
}


/* ============================================================
   COLUMN INTELLIGENCE (Levels 1–3)
   Level 3: profileColumns(rows) -> builds columnProfiles
   Level 2: populate dropdowns using profiles (filter/sort)
   Level 1: auto-select sensible defaults when chartType changes
   ============================================================ */

function looksLikeDateString(value) {
  if (value === null || value === undefined) return false;
  const s = String(value).trim();
  if (!s) return false;

  // IMPORTANT: Pure numbers are NOT dates.
  // Without this guard, Date.parse("1") etc can be treated as valid dates.
  if (/^[+-]?\d+(\.\d+)?$/.test(s)) return false;

  // Obvious date patterns (keep these simple and robust)
  const iso = /^\d{4}-\d{2}-\d{2}([T\s].*)?$/.test(s);      // 2024-01-07 or 2024-01-07T...
  const uk  = /^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(s);        // 07/01/2024
  const dash= /^\d{1,2}-\d{1,2}-\d{2,4}$/.test(s);          // 07-01-2024

  // If it contains common date separators, treat as date-ish
  if (iso || uk || dash) return true;

  // Light additional cue: strings with ':' often represent times (but not always)
  if (s.includes(":") && /\d/.test(s)) return true;

  // Do NOT fall back to Date.parse() — it is too permissive for our use case.
  return false;
}


function profileColumns(rows) {
  const profiles = {};
  if (!rows || rows.length === 0) return profiles;

  const firstRow = rows[0];
  const cols = firstRow ? Object.keys(firstRow) : [];
  allColumns = cols.slice();

  const nRows = rows.length;
  const maxSample = Math.min(nRows, 500); // keep profiling cheap
  const sampleIdx = [];
  // Evenly sample across rows
  for (let i = 0; i < maxSample; i++) {
    const idx = Math.floor(i * (nRows - 1) / Math.max(1, (maxSample - 1)));
    sampleIdx.push(idx);
  }

  for (const col of cols) {
    let nonEmpty = 0;
    let numericCount = 0;
    let intLikeCount = 0;
    let hasNeg = false;
    let hasZero = false;
    let min = Infinity;
    let max = -Infinity;

    let dateLikeCount = 0;

    const seen = new Set();
    const numericVals = [];

    for (const idx of sampleIdx) {
      const vRaw = rows[idx]?.[col];
      if (vRaw === null || vRaw === undefined) continue;

      const s = String(vRaw).trim();
      if (!s) continue;

      nonEmpty++;
      seen.add(s);

      if (looksLikeDateString(s)) dateLikeCount++;

      const num = toNumericValue(vRaw);
      if (Number.isFinite(num)) {
        numericCount++;
        numericVals.push(num);

        if (Math.abs(num - Math.round(num)) < 1e-9) intLikeCount++;
        if (num < 0) hasNeg = true;
        if (num === 0) hasZero = true;
        if (num < min) min = num;
        if (num > max) max = num;
      }
    }

    const numericFraction = nonEmpty > 0 ? numericCount / nonEmpty : 0;
    const intFraction = numericCount > 0 ? intLikeCount / numericCount : 0;
    const uniqueRatio = nonEmpty > 0 ? seen.size / nonEmpty : 1;
    const dateLikeFraction = nonEmpty > 0 ? dateLikeCount / nonEmpty : 0;

    profiles[col] = {
      col,
      nonEmpty,
      numericCount,
      numericFraction,
      isNumeric: numericFraction >= 0.8,      // tolerant of occasional blanks/text
      intFraction,
      isMostlyInteger: intFraction >= 0.9,    // “count-like”
      hasNeg,
      hasZero,
      min: min === Infinity ? NaN : min,
      max: max === -Infinity ? NaN : max,
      uniqueRatio,
      repeatsOften: uniqueRatio <= 0.6,       // useful for subgroup candidates
      looksLikeDate: dateLikeFraction >= 0.6, // likely date/time column
      dateLikeFraction
    };
  }

  columnProfiles = profiles;
  return profiles;
}

function getProfile(col) {
  return columnProfiles && col ? columnProfiles[col] : null;
}

function buildOption(label, value, { disabled = false, hint = "" } = {}) {
  const opt = document.createElement("option");
  opt.value = value;
  opt.textContent = hint ? `${label} ${hint}` : label;
  if (disabled) opt.disabled = true;
  return opt;
}

function setSelectOptions(selectEl, colList, { includeBlank = true, blankLabel = "(select)" } = {}) {
  if (!selectEl) return;

  const prev = selectEl.value;
  selectEl.innerHTML = "";

  if (includeBlank) {
    const blank = document.createElement("option");
    blank.value = "";
    blank.textContent = blankLabel;
    selectEl.appendChild(blank);
  }

  for (const item of colList) {
    // item can be string or {col, disabled, hint}
    if (typeof item === "string") {
      selectEl.appendChild(buildOption(item, item));
    } else {
      selectEl.appendChild(buildOption(item.col, item.col, { disabled: !!item.disabled, hint: item.hint || "" }));
    }
  }

  // restore previous if still present
  if (prev && Array.from(selectEl.options).some(o => o.value === prev && !o.disabled)) {
    selectEl.value = prev;
  }
}

function getNumericColumnsSorted() {
  // numeric columns first, then non-numeric
  const cols = allColumns.slice();
  cols.sort((a, b) => {
    const pa = getProfile(a);
    const pb = getProfile(b);
    const na = pa?.isNumeric ? 1 : 0;
    const nb = pb?.isNumeric ? 1 : 0;
    if (na !== nb) return nb - na;

    // then date-like last for numeric charts (but not for x-axis)
    const da = pa?.looksLikeDate ? 1 : 0;
    const db = pb?.looksLikeDate ? 1 : 0;
    if (da !== db) return da - db;

    // then more non-empty first
    const ea = pa?.nonEmpty ?? 0;
    const eb = pb?.nonEmpty ?? 0;
    return eb - ea;
  });
  return cols;
}

function getCandidatesForValue(chartType) {
  // returns array of items for valueSelect (strings or {col,disabled,hint})
  const cols = getNumericColumnsSorted();

  // For most charts, require numeric; for x-axis we do something else.
  const items = [];
  for (const c of cols) {
    const p = getProfile(c);
    if (!p) continue;

    // Hide clearly non-numeric for value roles
    if (!p.isNumeric) continue;

    // Soft-hints for suspicious columns
    let hint = "";
    let disabled = false;

    if ((chartType === "c" || isPFamilyChartType(chartType) || isUFamilyChartType(chartType)) && p.hasNeg) {
      hint = "(has negatives)";
    }
    if ((chartType === "c" || isPFamilyChartType(chartType) || isUFamilyChartType(chartType)) && p.looksLikeDate) {
      hint = "(looks like date)";
    }

    // For G: must be >= 1 (we don't hard-block; we hint)
    if (chartType === "g" && Number.isFinite(p.min) && p.min < 1) {
      hint = "(min < 1)";
    }

    // Prefer count-like columns for C/P/U numerators by sorting (done elsewhere),
    // but do not disable here.
    items.push({ col: c, disabled, hint: hint ? `— ${hint}` : "" });
  }

  return items;
}

function getCandidatesForThird(chartType) {
  // thirdSelect: denominator/opportunities (P/P′/U), subgroup (Xbars)
  if (isPFamilyChartType(chartType) || isUFamilyChartType(chartType)) {
    const cols = getNumericColumnsSorted();
    const items = [];
    for (const c of cols) {
      const p = getProfile(c);
      if (!p || !p.isNumeric) continue;

      let hint = "";
      // Denominators should be > 0
      if (Number.isFinite(p.min) && p.min <= 0) hint = "(min ≤ 0)";
      if (p.hasNeg) hint = "(has negatives)";

      items.push({ col: c, disabled: false, hint: hint ? `— ${hint}` : "" });
    }
    return items;
  }

  if (chartType === "xbars") {
    // subgroup id can be text or numeric; date-like is usually NOT subgroup
    const items = [];
    for (const c of allColumns) {
      const p = getProfile(c);
      if (!p) continue;

      // exclude obvious date columns
      if (p.looksLikeDate) continue;

      // subgroup candidates tend to repeat
      let hint = "";
      if (p.repeatsOften) hint = "(repeats — good subgroup)";
      else hint = "(many unique values)";

      // allow non-numeric too
      items.push({ col: c, disabled: false, hint: hint ? `— ${hint}` : "" });
    }

    // sort: repeats first
    items.sort((a, b) => {
      const pa = getProfile(a.col);
      const pb = getProfile(b.col);
      const ra = pa?.repeatsOften ? 1 : 0;
      const rb = pb?.repeatsOften ? 1 : 0;
      if (ra !== rb) return rb - ra;
      return (pb?.nonEmpty ?? 0) - (pa?.nonEmpty ?? 0);
    });

    return items;
  }

  return [];
}

function getBestXAxisColumn() {
  // Prefer date-like columns; otherwise first column
  const dateLike = allColumns.filter(c => getProfile(c)?.looksLikeDate);
  if (dateLike.length) return dateLike[0];
  return allColumns[0] || "";
}

function scorePChartPair(numerCol, denomCol) {
  // Score based on how often 0 <= numer <= denom and denom > 0
  if (!rawRows || rawRows.length === 0) return -Infinity;

  let ok = 0;
  let total = 0;

  const maxSample = Math.min(rawRows.length, 600);
  for (let i = 0; i < maxSample; i++) {
    const row = rawRows[i];
    const n = toNumericValue(row[numerCol]);
    const d = toNumericValue(row[denomCol]);
    if (!Number.isFinite(n) || !Number.isFinite(d)) continue;
    total++;
    if (d > 0 && n >= 0 && n <= d) ok++;
  }

  if (total < 5) return -Infinity;
  return ok / total;
}

function chooseDefaultsForChart(chartType) {
  // returns { xCol, yCol, thirdCol } (any may be "")
  if (!rawRows || rawRows.length === 0) return { xCol: "", yCol: "", thirdCol: "" };

  const xCol = getBestXAxisColumn();

  // candidates for numeric roles
  const numericCols = allColumns.filter(c => getProfile(c)?.isNumeric && !getProfile(c)?.looksLikeDate);

  if (chartType === "pareto") {
    const categoryCandidates =
      allColumns
        .filter(
          c =>
            getProfile(c)?.nonEmpty > 0
        )
        .slice()
        .sort(
          (a, b) => {
            function score(col) {
              const p =
                getProfile(col);

              const name =
                String(col || "")
                  .toLowerCase();

              let s = 0;

              if (!p?.looksLikeDate) {
                s += 40;
              } else {
                s -= 50;
              }

              if (!p?.isNumeric) {
                s += 25;
              }

              if (p?.repeatsOften) {
                s += 20;
              }

              if (
                Number.isFinite(
                  p?.uniqueRatio
                ) &&
                p.uniqueRatio < 0.75
              ) {
                s += 10;
              }

              if (
                /category|type|reason|cause|group|class|location|ward|unit|issue|complaint|failure/.test(
                  name
                )
              ) {
                s += 30;
              }

              return s;
            }

            return (
              score(b) -
              score(a)
            );
          }
        );

    return {
      xCol:
        categoryCandidates[0] ||
        allColumns[0] ||
        "",

      yCol:
        PARETO_COUNT_ROWS,

      thirdCol: ""
    };
  }

  // Helper for count-like numeric columns
  const countLike = numericCols
    .slice()
    .sort((a, b) => (getProfile(b)?.intFraction ?? 0) - (getProfile(a)?.intFraction ?? 0));

  // Scoring helper for "measure-like" columns
  function scoreMeasureColumn(col, { preferCounts = false } = {}) {
    const p = getProfile(col);
    if (!p || !p.isNumeric || p.looksLikeDate) return -Infinity;

    const name = String(col || "").toLowerCase();

    let score = 0;

    // Prefer columns that are not the chosen x-axis
    if (col !== xCol) score += 3;
    else score -= 4;

    // Prefer columns with some variability / not just repeated IDs
    if (!p.repeatsOften) score += 1;
    if (p.uniqueRatio >= 0.6 && p.uniqueRatio <= 1) score += 1;

    // Prefer typical outcome/measure names
    if (
      name.includes("value") ||
      name.includes("measure") ||
      name.includes("metric") ||
      name.includes("rate") ||
      name.includes("score") ||
      name.includes("count") ||
      name.includes("result") ||
      name.includes("time")
    ) {
      score += 3;
    }

    // Penalise likely index / identifier columns
    if (
      name.includes("id") ||
      name.includes("index") ||
      name.includes("seq") ||
      name.includes("sequence") ||
      name.includes("week") ||
      name.includes("day") ||
      name.includes("number") ||
      name.includes("row")
    ) {
      score -= 4;
    }

    // For non-count charts, prefer less "pure count-like" columns slightly
    if (!preferCounts && !p.isMostlyInteger) score += 1;

    // For count charts, prefer integer-like non-negative columns
    if (preferCounts) {
      if (p.isMostlyInteger) score += 2;
      if (!p.hasNeg) score += 1;
    }

    return score;
  }

  if (chartType === "run" || chartType === "xmr") {
    const yCol =
      numericCols
        .slice()
        .sort((a, b) => scoreMeasureColumn(b) - scoreMeasureColumn(a))[0] || "";
    return { xCol, yCol, thirdCol: "" };
  }

  if (chartType === "c") {
    const yCol =
      countLike
        .slice()
        .sort((a, b) => scoreMeasureColumn(b, { preferCounts: true }) - scoreMeasureColumn(a, { preferCounts: true }))[0] ||
      numericCols[0] ||
      "";
    return { xCol, yCol, thirdCol: "" };
  }

  if (chartType === "g") {
    const yCol =
      countLike.find(c => {
        const p = getProfile(c);
        return !p?.hasNeg && Number.isFinite(p?.min) && p.min >= 1;
      }) ||
      countLike
        .slice()
        .sort((a, b) => scoreMeasureColumn(b, { preferCounts: true }) - scoreMeasureColumn(a, { preferCounts: true }))[0] ||
      numericCols[0] ||
      "";
    return { xCol, yCol, thirdCol: "" };
  }

  if (isPFamilyChartType(chartType)) {
    // choose best (numer, denom) pair among integer-ish columns
    const intCols = numericCols.filter(c => getProfile(c)?.isMostlyInteger && !getProfile(c)?.hasNeg);
    let best = { numer: "", denom: "", score: -Infinity };

    for (const numer of intCols) {
      for (const denom of intCols) {
        if (numer === denom) continue;
        const s = scorePChartPair(numer, denom);
        if (s > best.score) best = { numer, denom, score: s };
      }
    }

    let yCol = best.numer || intCols[0] || numericCols[0] || "";
    let thirdCol = best.denom || intCols.find(c => c !== yCol) || numericCols.find(c => c !== yCol) || "";

    if (yCol && thirdCol) {
      const py = getProfile(yCol);
      const pt = getProfile(thirdCol);
      const yMeanish = (Number.isFinite(py?.min) && Number.isFinite(py?.max)) ? (py.min + py.max) / 2 : NaN;
      const tMeanish = (Number.isFinite(pt?.min) && Number.isFinite(pt?.max)) ? (pt.min + pt.max) / 2 : NaN;
      if (Number.isFinite(yMeanish) && Number.isFinite(tMeanish) && yMeanish > tMeanish) {
        [yCol, thirdCol] = [thirdCol, yCol];
      }
    }

    return { xCol, yCol, thirdCol };
  }

  if (isUFamilyChartType(chartType)) {
  // choose count/opportunity pair:
  // numerator = count-like column
  // denominator = positive integer-like opportunities column, usually larger
  const intCols = numericCols.filter(c => {
    const p = getProfile(c);
    return p?.isMostlyInteger && !p?.hasNeg;
  });

  let yCol =
    intCols
      .slice()
      .sort((a, b) =>
        scoreMeasureColumn(b, { preferCounts: true }) -
        scoreMeasureColumn(a, { preferCounts: true })
      )[0] ||
    numericCols[0] ||
    "";

  const positiveOpportunityCols = intCols.filter(c => {
    const p = getProfile(c);
    return c !== yCol && Number.isFinite(p?.min) && p.min > 0;
  });

  let thirdCol =
    positiveOpportunityCols.find(c =>
      (getProfile(c)?.max ?? -Infinity) >= (getProfile(yCol)?.max ?? Infinity)
    ) ||
    positiveOpportunityCols[0] ||
    intCols.find(c => c !== yCol && Number.isFinite(getProfile(c)?.min) && getProfile(c).min > 0) ||
    intCols.find(c => c !== yCol) ||
    numericCols.find(c => c !== yCol) ||
    "";

  return { xCol, yCol, thirdCol };
}

  if (chartType === "xbars") {
    // subgroup: prefer repeated non-date columns
    const subgroup =
      allColumns
        .filter(c => c !== xCol)
        .filter(c => !getProfile(c)?.looksLikeDate)
        .sort((a, b) => {
          const pa = getProfile(a), pb = getProfile(b);
          const ra = pa?.repeatsOften ? 1 : 0;
          const rb = pb?.repeatsOften ? 1 : 0;
          if (ra !== rb) return rb - ra;
          return (pb?.nonEmpty ?? 0) - (pa?.nonEmpty ?? 0);
        })[0] || "";

    const yCol =
      numericCols
        .filter(c => c !== subgroup)
        .sort((a, b) => scoreMeasureColumn(b) - scoreMeasureColumn(a))[0] ||
      numericCols[0] ||
      "";

    return { xCol, yCol, thirdCol: subgroup };
  }

  if (chartType === "t") {
    // Prefer event date/time in xCol; if there is a separate numeric gap column use it as y,
    // otherwise keep first numeric fallback.
    const yCol =
      numericCols
        .filter(c => c !== xCol)
        .sort((a, b) => scoreMeasureColumn(b) - scoreMeasureColumn(a))[0] ||
      numericCols[0] ||
      "";
    return { xCol, yCol, thirdCol: "" };
  }

  return {
    xCol,
    yCol:
      numericCols
        .slice()
        .sort((a, b) => scoreMeasureColumn(b) - scoreMeasureColumn(a))[0] || "",
    thirdCol: ""
  };
}
function applyColumnIntelligence(chartType) {
  // Level 2: filter option lists
  if (dateSelect) {
    // x-axis can be anything; prefer date-like first in ordering
    const cols = allColumns.slice().sort((a, b) => {
      const pa = getProfile(a), pb = getProfile(b);
      const da = pa?.looksLikeDate ? 1 : 0;
      const db = pb?.looksLikeDate ? 1 : 0;
      if (da !== db) return db - da;
      return (pb?.nonEmpty ?? 0) - (pa?.nonEmpty ?? 0);
    });
    setSelectOptions(dateSelect, cols, { includeBlank: true, blankLabel: "(select x-axis)" });
  }

  const valueItems = getCandidatesForValue(chartType);
  setSelectOptions(valueSelect, valueItems, { includeBlank: true, blankLabel: "(select value)" });

  if (
    chartType === "pareto" &&
    valueSelect
  ) {
    const countRowsOption =
      document.createElement(
        "option"
      );

    countRowsOption.value =
      PARETO_COUNT_ROWS;

    countRowsOption.textContent =
      "(count each row as 1)";

    valueSelect.insertBefore(
      countRowsOption,
      valueSelect.options[1] ||
      null
    );
  }

  const thirdItems = getCandidatesForThird(chartType);
  if (thirdSelect) {
    if (thirdItems.length) {
      setSelectOptions(thirdSelect, thirdItems, { includeBlank: true, blankLabel: "(select)" });
    } else {
      // keep blank if not needed
      setSelectOptions(thirdSelect, [], { includeBlank: true, blankLabel: "(not needed)" });
    }
  }

  // Level 1: apply sensible defaults if current selections are empty or invalid
  const defaults = chooseDefaultsForChart(chartType);

  // only set defaults when current selection is empty OR no longer valid in the options
  function setIfEmptyOrMissing(selectEl, newVal) {
    if (!selectEl || !newVal) return;
    const has = Array.from(selectEl.options).some(o => o.value === newVal && !o.disabled);
    if (!has) return;

    const current = selectEl.value;
    const currentStillValid = current && Array.from(selectEl.options).some(o => o.value === current && !o.disabled);
    if (!currentStillValid) {
      selectEl.value = newVal;
      return;
    }

    if (!current) selectEl.value = newVal;
  }

  setIfEmptyOrMissing(dateSelect, defaults.xCol);

// For count-style charts, be a bit more proactive:
// if the current value column is numeric but clearly a poor fit
// (e.g. mostly decimal or negative), switch to the better default.
const currentValueProfile =
  valueSelect ? getProfile(valueSelect.value) : null;

const currentThirdProfile =
  thirdSelect ? getProfile(thirdSelect.value) : null;

const isPFamily =
  isPFamilyChartType(chartType);

const isUFamily =
  isUFamilyChartType(chartType);

const currentValuePoorForCountChart =
  !!(
    currentValueProfile &&
    (
      chartType === "c" ||
      isPFamily ||
      isUFamily
    ) &&
    (
      !currentValueProfile.isMostlyInteger ||
      currentValueProfile.hasNeg
    )
  );

const currentThirdPoorForDenominator =
  !!(
    currentThirdProfile &&
    (
      isPFamily ||
      isUFamily
    ) &&
    (
      !currentThirdProfile.isMostlyInteger ||
      currentThirdProfile.hasNeg ||
      (
        Number.isFinite(currentThirdProfile.min) &&
        currentThirdProfile.min <= 0
      )
    )
  );

if (currentValuePoorForCountChart && defaults.yCol) {
  valueSelect.value = defaults.yCol;
} else {
  setIfEmptyOrMissing(
    valueSelect,
    defaults.yCol
  );
}

if (
  isPFamily ||
  isUFamily ||
  chartType === "xbars"
) {
  if (
    isPFamily ||
    isUFamily
  ) {
    if (
      currentThirdPoorForDenominator &&
      defaults.thirdCol
    ) {
      thirdSelect.value =
        defaults.thirdCol;
    } else {
      setIfEmptyOrMissing(
        thirdSelect,
        defaults.thirdCol
      );
    }
  } else {
    setIfEmptyOrMissing(
      thirdSelect,
      defaults.thirdCol
    );
  }
}

// avoid third == y if needed
if (
  (
    isPFamily ||
    isUFamily ||
    chartType === "xbars"
  ) &&
  thirdSelect &&
  valueSelect
) {
    if (thirdSelect.value && valueSelect.value && thirdSelect.value === valueSelect.value) {
      const alt = Array.from(thirdSelect.options)
        .map(o => o.value)
        .find(v => v && v !== valueSelect.value);
      if (alt) thirdSelect.value = alt;
    }
  }

// ---- Auto-adjust axis type if user has not manually chosen ----
if (!axisTypeManuallyChanged) {
  const xCol = dateSelect?.value;
  const p = (typeof getProfile === "function") ? getProfile(xCol) : null;

  const xLooksLikeNumericSequence =
    !!(p &&
       p.isNumeric &&
       p.isMostlyInteger &&
       p.uniqueRatio >= 0.85 &&
       !p.repeatsOften &&
       !p.looksLikeDate);

  const xLooksLikeNonDateLabel =
    !!(p &&
       !p.looksLikeDate &&
       (
         !p.isNumeric ||
         p.repeatsOften ||
         p.uniqueRatio > 0
       ));

  const shouldUseSequenceAxis =
    xLooksLikeNumericSequence ||
    (chartType === "xbars" && xLooksLikeNonDateLabel) ||
    (
  (
    chartType === "run" ||
    chartType === "xmr" ||
    chartType === "c" ||
    isPFamily ||
    isUFamily ||
    chartType === "g"
  ) &&
  xLooksLikeNonDateLabel
);

  const seqRadio = document.querySelector("input[name='axisType'][value='sequence']");
  const dateRadio = document.querySelector("input[name='axisType'][value='date']");

  if (shouldUseSequenceAxis) {
    if (seqRadio) seqRadio.checked = true;
  } else if (p && p.looksLikeDate) {
    if (dateRadio) dateRadio.checked = true;
  }
}

}

function getChartTypeRadioInput(chartType) {
  return document.querySelector(`input[name='chartType'][value='${chartType}']`);
}

function getChartTypeRadioLabel(chartType) {
  const input = getChartTypeRadioInput(chartType);
  return input ? input.closest('label') : null;
}

function getFirstEnabledChartType() {
  const radios = Array.from(document.querySelectorAll("input[name='chartType']"));
  const firstEnabled = radios.find(r => !r.disabled);
  return firstEnabled ? firstEnabled.value : "run";
}

function countValidNumericPointsForColumn(colName) {
  if (!rawRows || !rawRows.length || !colName) return 0;
  let n = 0;
  for (const row of rawRows) {
    const y = toNumericValue(row[colName]);
    if (Number.isFinite(y)) n++;
  }
  return n;
}

function getSelectedOrSuggestedValueColumn(chartType) {
  const selected = valueSelect?.value || "";
  if (selected) return selected;
  if (typeof chooseDefaultsForChart === "function") {
    return chooseDefaultsForChart(chartType)?.yCol || "";
  }
  return "";
}

function getChartAvailability() {
  const out = {};
  const noDataReason = "Load data to assess which chart types are available.";

  for (const chartType of ["run", "xmr", "c", "p", "pprime", "u","uprime", "xbars", "t", "g", "pareto"]) {
    out[chartType] = { enabled: true, reason: "" };
  }

  if (!rawRows || !rawRows.length) {
    for (const key of Object.keys(out)) {
      if (key === "run") continue;
      out[key] = { enabled: false, reason: noDataReason };
    }
    return out;
  }

  const numericNonDateCols = allColumns.filter(c => {
    const p = getProfile(c);
    return p?.isNumeric && !p?.looksLikeDate;
  });

  const countLikeNonNegCols = numericNonDateCols.filter(c => {
    const p = getProfile(c);
    return p?.isMostlyInteger && !p?.hasNeg;
  });

  const positiveCountLikeCols = countLikeNonNegCols.filter(c => {
    const p = getProfile(c);
    return Number.isFinite(p?.min) && p.min > 0;
  });

  const repeatingSubgroupCols = allColumns.filter(c => {
    const p = getProfile(c);
    return p && !p.looksLikeDate && p.repeatsOften;
  });

  const dateLikeCols = allColumns.filter(c => getProfile(c)?.looksLikeDate);
  const gapLikeCols = numericNonDateCols.filter(c => countValidNumericPointsForColumn(c) >= 3);

  const xmrValueCol = getSelectedOrSuggestedValueColumn("xmr");
  const xmrValidPoints = countValidNumericPointsForColumn(xmrValueCol);
  if (xmrValidPoints < 12) {
    out.xmr = {
      enabled: false,
      reason: `Less than 12 data points available in "${xmrValueCol || "the selected value column"}" — this is a minimum requirement for a valid XmR chart.`
    };
  }

  if (!countLikeNonNegCols.length) {
    out.c = {
      enabled: false,
      reason: "No non-negative whole-number count column is available yet — C charts need counts per time period."
    };
  }

  let hasValidPPair = false;
  if (countLikeNonNegCols.length >= 2) {
    for (const numer of countLikeNonNegCols) {
      for (const denom of countLikeNonNegCols) {
        if (numer === denom) continue;
        if (scorePChartPair(numer, denom) > -Infinity) {
          hasValidPPair = true;
          break;
        }
      }
      if (hasValidPPair) break;
    }
  }
  if (!hasValidPPair) {
  out.p = {
    enabled: false,
    reason: "No suitable numerator/denominator pair is available yet — P charts need two whole-number columns where the numerator is part of the denominator."
  };

  out.pprime = {
    enabled: false,
    reason: "No suitable numerator/denominator pair is available yet — P′ charts need two whole-number columns where the numerator is part of the denominator."
  };
}

  if (
  countLikeNonNegCols.length < 2 ||
  !positiveCountLikeCols.length
) {
  out.u = {
    enabled: false,
    reason:
      "No suitable count/opportunities pair is available yet — U charts need a whole-number count column plus a positive opportunities column."
  };

  out.uprime = {
    enabled: false,
    reason:
      "No suitable count/opportunities pair is available yet — U′ charts need a whole-number count column plus a positive opportunities column."
  };
}

  if (!numericNonDateCols.length || !repeatingSubgroupCols.length) {
    out.xbars = {
      enabled: false,
      reason: "X̄–S charts need a numeric measurement column and a subgroup column with repeated subgroup labels."
    };
  }

  if (!dateLikeCols.length && !gapLikeCols.length) {
    out.t = {
      enabled: false,
      reason: "T charts need either event dates or a numeric time-between-events column."
    };
  }

  if (!positiveCountLikeCols.length) {
    out.g = {
      enabled: false,
      reason: "G charts need a whole-number column where all values are at least 1, because they plot opportunities between rare events."
    };
  }

  return out;
}

function applyChartTypeAvailability() {
  const availability = getChartAvailability();
  let selectedInput = document.querySelector("input[name='chartType']:checked");
  let selectedType = selectedInput?.value || "run";

  Object.entries(availability).forEach(([chartType, state]) => {
    const input = getChartTypeRadioInput(chartType);
    const label = getChartTypeRadioLabel(chartType);
    if (!input || !label) return;

    const reason = state.reason || "";
    input.disabled = !state.enabled;
    input.title = reason;
    label.title = reason;
    label.classList.toggle("chart-disabled", !state.enabled);
    label.setAttribute("aria-disabled", state.enabled ? "false" : "true");
  });

  if (selectedInput && selectedInput.disabled) {
    const fallbackType = availability.run?.enabled ? "run" : getFirstEnabledChartType();
    const fallbackInput = getChartTypeRadioInput(fallbackType);
    if (fallbackInput) {
      fallbackInput.checked = true;
      selectedInput = fallbackInput;
      selectedType = fallbackType;
    }
  }

  if (chartTypeAvailabilityHint) {
    const state = availability[selectedType];
    chartTypeAvailabilityHint.textContent = state?.enabled
      ? ""
      : (state.reason || "This chart type is not available for the current data.");
  }

  return availability;
}

function refreshChartTypeAvailability() {
  const availability = applyChartTypeAvailability();
  const currentType = getSelectedChartType_NoSideEffects();
  if (rawRows && rawRows.length) {
    updateUIForChartType(currentType);
  }
  return availability;
}

function isSelectedChartTypeAvailable() {
  const availability = getChartAvailability();
  const currentType = getSelectedChartType_NoSideEffects();
  return !!availability[currentType]?.enabled;
}

function getSelectedChartTypeUnavailableReason() {
  const availability = getChartAvailability();
  const currentType = getSelectedChartType_NoSideEffects();
  return availability[currentType]?.reason || "";
}

function syncChartTypeAvailabilityMessage() {
  if (!chartTypeAvailabilityHint) return;
  const reason = getSelectedChartTypeUnavailableReason();
  chartTypeAvailabilityHint.textContent = reason || "";
}

function getRuleSettings() {
  const shift = shiftRulePointsInput ? parseInt(shiftRulePointsInput.value, 10) : NaN;
  const trend = trendRulePointsInput ? parseInt(trendRulePointsInput.value, 10) : NaN;

  return {
    shiftLength: Number.isFinite(shift) && shift >= 3 ? shift : 8,
    trendLength: Number.isFinite(trend) && trend >= 3 ? trend : 6,

    enableAdvancedTrend: enableAdvancedTrendCheckbox ? !!enableAdvancedTrendCheckbox.checked : false,
    enableRareRunTrend: enableRareRunTrendCheckbox ? !!enableRareRunTrendCheckbox.checked : false,

    ruleTwoOfThreeOuterThird: ruleTwoOfThreeOuterThirdCheckbox ? !!ruleTwoOfThreeOuterThirdCheckbox.checked : false,
    ruleFourOfFiveOneSigma: ruleFourOfFiveOneSigmaCheckbox ? !!ruleFourOfFiveOneSigmaCheckbox.checked : false
  };
}


function shouldFlagSpecialCauseOnChart() {
  return flagSpecialCauseOnChartCheckbox ? !!flagSpecialCauseOnChartCheckbox.checked : true;
}

function shouldClampLclAtZero() {
  // The saved checkbox value determines whether clamping is enabled.
  // Do not depend on the visibility of the UI row, because visibility
  // is determined during chart calculation.
  return clampLclAtZeroCheckbox
    ? !!clampLclAtZeroCheckbox.checked
    : false;
}

function setLclClampVisibility(shouldShow) {
  if (!lclClampRow) return;
  lclClampRow.style.display = shouldShow ? "block" : "none";

  // if the option disappears, clear it to avoid “sticky” state
  if (!shouldShow && clampLclAtZeroCheckbox) clampLclAtZeroCheckbox.checked = false;
}

function findLongRunRanges(values, centre, runLength) {
  const ranges = [];
  let start = 0;

  while (start < values.length) {
    const v = values[start];
    const side = v > centre ? "above" : v < centre ? "below" : "on";
    if (side === "on") { start++; continue; }

    let end = start + 1;
    while (end < values.length) {
      const v2 = values[end];
      const side2 = v2 > centre ? "above" : v2 < centre ? "below" : "on";
      if (side2 !== side) break;
      end++;
    }

    const len = end - start;
    if (len >= runLength) ranges.push({ start, end: end - 1, side, len });

    start = end;
  }
  return ranges;
}

function flagFromRanges(n, ranges) {
  const flags = new Array(n).fill(false);
  ranges.forEach(r => {
    for (let i = r.start; i <= r.end; i++) flags[i] = true;
  });
  return flags;
}

function findTrendRanges(values, length) {
  const ranges = [];
  if (values.length < length) return ranges;

  let incStart = 0, incLen = 1;
  let decStart = 0, decLen = 1;

  for (let i = 1; i < values.length; i++) {
    if (values[i] > values[i - 1]) {
      incLen++; decLen = 1; decStart = i;
    } else if (values[i] < values[i - 1]) {
      decLen++; incLen = 1; incStart = i;
    } else {
      incLen = 1; decLen = 1; incStart = i; decStart = i;
    }

    if (incLen >= length) {
      const start = i - incLen + 1;
      ranges.push({ start, end: i, direction: "increasing", len: incLen });
      incLen = 1; // avoid overlapping spam; simple approach
      incStart = i;
    }
    if (decLen >= length) {
      const start = i - decLen + 1;
      ranges.push({ start, end: i, direction: "decreasing", len: decLen });
      decLen = 1;
      decStart = i;
    }
  }

  return ranges;
}

function updateRuleUIForChartType(chartType) {
  const policy = getRulePolicy(chartType);
  const isRare = isRareChartType(chartType);
  const isAdvancedContinuous = isAdvancedContinuousChartType(chartType);
  const trendRulePointsInput = document.getElementById("trendRulePoints");
  const enableAdvancedTrendCheckbox = document.getElementById("enableAdvancedTrend");
  const enableRareRunTrendCheckbox = document.getElementById("enableRareRunTrend");

  // Advanced section is only useful when there is something advanced to show
  const showAdvancedSection =
    isRare ||
    policy.trend === "optional" ||
    policy.zone23 === "optional" ||
    policy.zone45 === "optional";

  if (advancedRulesDetails) {
    advancedRulesDetails.style.display = showAdvancedSection ? "block" : "none";
    if (!showAdvancedSection) advancedRulesDetails.open = false;
  }

  // Advanced trend for X-MR / XbarS / Run only
  if (advancedTrendRow) {
    advancedTrendRow.style.display = (policy.trend === "optional") ? "block" : "none";
  }

  if (enableAdvancedTrendCheckbox) {
    enableAdvancedTrendCheckbox.disabled = !(policy.trend === "optional");
    if (policy.trend !== "optional") enableAdvancedTrendCheckbox.checked = false;
    enableAdvancedTrendCheckbox.title =
      policy.trend === "optional" ? "" : "Trend rule is not offered for this chart type.";
  }

  if (trendRulePointsInput) {
  let trendInputEnabled = false;
  let trendInputTitle = "";

  if (policy.trend === "optional") {
    // Run / XmR / XbarS: only enabled when user ticks the advanced trend checkbox
    trendInputEnabled = !!enableAdvancedTrendCheckbox?.checked;
    trendInputTitle = trendInputEnabled
      ? ""
      : "Tick 'Enable trend rule' to use this setting.";
  } else if (policy.trend === "warn") {
    // T / G: only enabled when user explicitly enables rare-chart run/trend rules
    trendInputEnabled = !!enableRareRunTrendCheckbox?.checked;
    trendInputTitle = trendInputEnabled
      ? ""
      : "Enable run & trend rules for this chart type to use this setting.";
  } else {
    // P / U / C and any chart type where trend is blocked
    trendInputEnabled = false;
    trendInputTitle = "Trend rule is not offered for this chart type.";
  }

  trendRulePointsInput.disabled = !trendInputEnabled;
  trendRulePointsInput.title = trendInputTitle;
}

  // Rare chart advanced row
  if (rareRulesRow) {
    rareRulesRow.style.display = isRare ? "block" : "none";
  }

  if (enableRareRunTrendCheckbox) {
    enableRareRunTrendCheckbox.disabled = !isRare;
    if (!isRare) enableRareRunTrendCheckbox.checked = false;
    enableRareRunTrendCheckbox.title = isRare ? "" : "Only used for T and G charts.";
  }

  // Zone rules only for X-MR / XbarS
  const showZoneRules = policy.zone23 === "optional" || policy.zone45 === "optional";
  if (zoneRulesSection) {
    zoneRulesSection.style.display = showZoneRules ? "block" : "none";
  }

  if (ruleTwoOfThreeOuterThirdCheckbox) {
    const blocked = policy.zone23 === "blocked";
    ruleTwoOfThreeOuterThirdCheckbox.disabled = blocked;
    if (blocked) ruleTwoOfThreeOuterThirdCheckbox.checked = false;
    ruleTwoOfThreeOuterThirdCheckbox.title = blocked
      ? "Zone rules are not available for this chart type because they may create misleading alerts."
      : "";
  }

  if (ruleFourOfFiveOneSigmaCheckbox) {
    const blocked = policy.zone45 === "blocked";
    ruleFourOfFiveOneSigmaCheckbox.disabled = blocked;
    if (blocked) ruleFourOfFiveOneSigmaCheckbox.checked = false;
    ruleFourOfFiveOneSigmaCheckbox.title = blocked
      ? "Zone rules are not available for this chart type because they may create misleading alerts."
      : "";
  }

  // Show caution text specifically for X-MR
  if (advancedContinuousCaution) {
    advancedContinuousCaution.style.display = (chartType === "xmr") ? "block" : "none";
  }

  // XmR-specific limit calculation option
if (mrScreeningRow) {
  mrScreeningRow.style.display =
    chartType === "xmr" ? "block" : "none";
}

  // Conservative message for chart types with no advanced offering
  if (conservativeRulesMessage) {
    const showConservativeMessage =
  chartType === "c" ||
  isPFamilyChartType(chartType) ||
  isUFamilyChartType(chartType);
    conservativeRulesMessage.style.display = showConservativeMessage ? "block" : "none";
  }
}

function updateColumnCheckWarning(chartType) {
  if (!columnCheckWarning) return;

  const warnings = {
    c: `
      <strong>Check your selected columns.</strong>
      For a C chart, the value column should be a count of events/defects per time period.
      Do not use a denominator column here.
    `,
    p: `
      <strong>Check your selected columns.</strong>
      For a P chart, the value column should be the numerator, e.g. number of cases/events,
      and the third column should be the denominator, e.g. total patients/episodes.
    `,
     pprime: `
  		<strong>Check your selected columns.</strong>
  		For a P′ chart, the value column should be the numerator, e.g. number of cases/events,
  		and the third column should be the denominator, e.g. total patients/episodes.
  		P′ uses the same data structure as a P chart but adjusts the control limits using the Laney method.
    `,
    u: `
      <strong>Check your selected columns.</strong>
      For a U chart, the value column should be the numerator, e.g. number of defects/events,
      and the third column should be the opportunities/denominator, e.g. bed days, attendances, or procedures.
    `,
    uprime: `
    <strong>Check your selected columns.</strong>
 	 For a U′ chart, the value column should be the numerator,
  	e.g. number of defects/events, and the third column should
  	be the opportunities/denominator, e.g. bed days,
 	 attendances, procedures or inspections.
 	 U′ uses the same data as a U chart but adjusts the
 	 control limits using the Laney method.
    `,
    xbars: `
      <strong>Check your selected columns.</strong>
      For an X̄–S chart, the value column should contain the individual measurements,
      and the third column should identify the subgroup, e.g. day, week, batch, or sample.
    `,
     
    
    t: tChartInputMode === "gaps"
      ? `
        <strong>T chart: using numeric time gaps.</strong><br>
        Select the column containing the time between
        consecutive events (e.g. days) as the
        <strong>Value / Y-axis column</strong>.
        The X-axis column provides the observation labels.

        <br><br>
        <strong>Want to use event dates instead?</strong>
        Click <strong>Chart setup</strong> above and select
        <strong>I have event dates</strong>.
      `
      : `
        <strong>T chart: using event dates.</strong><br>
        Select the dates or times when events occurred in
        the <strong>Date / X-axis column</strong>.
        SimpleSPC calculates the intervals automatically,
        so the Value / Y-axis column is disabled.

        <br><br>
        <strong>Already have the time gaps calculated?</strong>
        Click <strong>Chart setup</strong> above and select
        <strong>I already have the gaps</strong>.
        This will enable the Value / Y-axis column.
      `,


    g: `
      <strong>Check your selected columns.</strong>
      For a G chart, the value column should usually represent the number of opportunities,
      cases, or observations between rare events. Check this is not a rate or percentage.
    `,

        pareto: `
      <strong>Choose how you want to group the occurrences.</strong>
      Select the category you want to compare. If every row is one occurrence,
      choose <strong>(count each row as 1)</strong> for the frequency column.
      If your data are already summarised, choose the numeric count column instead.
      Pareto categories help focus investigation but should not automatically be assumed to be causes.
    `
  };

  const message = warnings[chartType];

  if (!message) {
    columnCheckWarning.style.display = "none";
    columnCheckWarning.innerHTML = "";
    return;
  }

  columnCheckWarning.innerHTML = message;
  columnCheckWarning.style.display = "block";
}

function updateUIForChartType(chartType) {
  if (!xLabelEl || !yLabelEl || !thirdColumnRow) return;

updateColumnCheckWarning(chartType);

  // ---- Default UI state (safe baseline) ----
  xLabelEl.textContent = "Date / X-axis column";
  yLabelEl.textContent = "Value / Y-axis column";

  thirdColumnRow.style.display = "none";
  thirdLabelEl.textContent = "";
  thirdHintEl.textContent = "";

  if (uRateMultiplierRow) {
  uRateMultiplierRow.style.display =
  isUFamilyChartType(chartType)
    ? "block"
    : "none";
}

  // ---- Chart-specific UI definitions ----
  const chartUI = {
    run: {
      // defaults are fine
    },

    xmr: {
      yLabel: "Measure (used for XmR limits)"
    },

    c: {
      yLabel: "Count (c) per time period"
    },

    p: {
      yLabel: "Numerator: defectives (d)",
      thirdLabel: "Denominator: total (n)",
      thirdHint: "P chart plots a proportion: d out of n.",
      needsThird: true
    },

    pprime: {
          yLabel: "Numerator: defectives (d)",
         thirdLabel: "Denominator: total (n)",
         thirdHint: "P′ chart plots a proportion: d out of n, with Laney-adjusted control limits.",
          needsThird: true
          },

    u: {
      yLabel: "Numerator: defects (c)",
      thirdLabel: "Denominator: opportunities (n)",
      thirdHint: "U chart plots defects per opportunity: c per n.",
      needsThird: true
    },

    uprime: {
 	 yLabel: "Numerator: defects (c)",
 	 thirdLabel: "Denominator: opportunities (n)",
 	 thirdHint:
   	 "U′ chart plots defects per opportunity, with Laney-adjusted control limits.",
  	needsThird: true
	},
   
    xbars: {
  yLabel: "Measurement value",
  thirdLabel: "Subgroup ID (e.g. day / week / sample)",
  thirdHint:
    "X̄–S needs multiple measurements per subgroup. The horizontal axis uses subgroup order automatically.",
  needsThird: true
},

    t: {
          // Labels depend on T chart input mode (event dates vs gaps)
         yLabel: (tChartInputMode === "gaps") ? "Time between events (e.g. days)" : "Value column not used (T chart uses dates)",
         thirdHint: (tChartInputMode === "gaps")
         ? "T chart plots time between rare events (using your numeric gaps)."
          : "T chart plots time between rare events (calculated from event dates)."
         },


    g: {
      yLabel: "Opportunities between events",
      thirdHint: "G chart plots opportunities between rare events."
    },

        pareto: {
      		xLabel:   "Category column",
     		 yLabel:  "Frequency / count column"
    	 }

  };

  // ---- Apply config (if defined) ----
  const cfg = chartUI[chartType];
  if (!cfg) return;

if (cfg.xLabel) {
  xLabelEl.textContent =
    cfg.xLabel;
}

  if (cfg.yLabel) {
    yLabelEl.textContent = cfg.yLabel;
  }

  if (cfg.needsThird) {
    thirdColumnRow.style.display = "block";
  }

  if (cfg.thirdLabel) {
    thirdLabelEl.textContent = cfg.thirdLabel;
  }

  if (cfg.thirdHint) {
    thirdHintEl.textContent = cfg.thirdHint;
  }

  // ------------------------------------------------------------
  // Levels 1–3 glue:
  // Whenever the chart type changes, rebuild dropdown options
  // (filtering) and apply sensible defaults for this chart type.
  // ------------------------------------------------------------
    if (rawRows && rawRows.length && typeof applyColumnIntelligence === "function") {
    applyColumnIntelligence(chartType);
  }
  if (typeof syncChartTypeAvailabilityMessage === "function") {
    syncChartTypeAvailabilityMessage();
  }

    // ---- Optional UX polish: avoid third == y by default ----
  if (cfg.needsThird && thirdSelect && valueSelect) {
    if (thirdSelect.value && valueSelect.value && thirdSelect.value === valueSelect.value) {
      const alt = Array.from(thirdSelect.options)
        .map(o => o.value)
        .find(v => v && v !== valueSelect.value);
      if (alt) thirdSelect.value = alt;
    }
  }

  // ---- Reset Value column enabled state for non-T charts ----
  if (valueSelect) {
    valueSelect.disabled = false;
    valueSelect.title = "";
  }

if (typeof updateRuleUIForChartType === "function") {
  updateRuleUIForChartType(chartType);
}

// ---- T chart UX: enable/disable Value column depending on input mode ----
if (chartType === "t" && valueSelect) {
  const shouldDisableValue = (tChartInputMode === "eventDates");
  valueSelect.disabled = shouldDisableValue;

  // Soft hint if disabled
  if (shouldDisableValue) {
    valueSelect.title = "Not used for T chart when using event dates.";
  }
}

// ---- X̄–S axis policy ----
// X̄–S is plotted by subgroup, so its horizontal axis is categorical.
// Do not make the calculation depend on individual rows having parseable dates.
const axisDateRadio =
  document.querySelector(
    "input[name='axisType'][value='date']"
  );

const axisSequenceRadio =
  document.querySelector(
    "input[name='axisType'][value='sequence']"
  );

if (chartType === "xbars" ||
  chartType === "pareto") {
  // Force the appropriate axis mode.
  if (typeof setAxisType === "function") {
    setAxisType("sequence");
  } else {
    if (axisSequenceRadio) {
      axisSequenceRadio.checked = true;
    }

    if (axisDateRadio) {
      axisDateRadio.checked = false;
    }
  }

  // Prevent an accidental switch back to Date / time.
  if (axisDateRadio) {
    axisDateRadio.disabled = true;
    axisDateRadio.title =
      "This chart uses categories on the horizontal axis."
  }

  if (axisSequenceRadio) {
    axisSequenceRadio.disabled = false;
    axisSequenceRadio.title =
      "X̄–S charts use subgroup order on the horizontal axis.";
  }

} else {
  // Restore normal axis choices for every other chart type.
  if (axisDateRadio) {
    axisDateRadio.disabled = false;
    axisDateRadio.title = "";
  }

  if (axisSequenceRadio) {
    axisSequenceRadio.disabled = false;
    axisSequenceRadio.title = "";
  }
}

// Keep the date-format control consistent with the forced axis mode.
if (typeof updateDateControlsState === "function") {
  updateDateControlsState();
}

if (typeof updateDateFormatWarning === "function") {
  updateDateFormatWarning();
}

// Keep MR / S-chart toggle synchronized with the selected chart type
if (typeof updateMrToggleVisibility === "function") {
  updateMrToggleVisibility();
}
}



function parseTabularTextWithHeaderDetection(text) {
  const preview = Papa.parse(text, {
    header: false,
    dynamicTyping: false,
    skipEmptyLines: true
  });

  if (preview.errors && preview.errors.length) {
    return { ok: false, message: preview.errors[0].message };
  }

  const rows2D = preview.data || [];
  if (rows2D.length < 2) {
    return { ok: false, message: "Please provide at least 2 rows." };
  }

  const r0 = rows2D[0];
  const r1 = rows2D[1];

  // Same scoring functions you already added for the data editor:
  const score0 = rowDataLikenessScore(r0);
  const score1 = rowDataLikenessScore(r1);
  const looksLikeHeader = (score1 - score0) >= 0.35;

  if (looksLikeHeader) {
    const results = Papa.parse(text, { header: true, dynamicTyping: true, skipEmptyLines: true });
    if (results.errors && results.errors.length) {
      return { ok: false, message: results.errors[0].message };
    }

    let rows = results.data || [];
    const headers = results.meta && results.meta.fields ? results.meta.fields : null;
    rows = stripDuplicateHeaderRow(rows, headers);

    return { ok: true, rows, hadHeader: true };
  }

  // No header detected
  return { ok: true, rows2D, hadHeader: false };
}

function computeMAD(values, centre) {
  const absDevs = values.map(v => Math.abs(v - centre));
  return computeMedian(absDevs);
}

/**
 * Astronomical point detection using modified z-score (MAD-based).
 * Common robust rule of thumb: |z| > 3.5
 * Returns { indices: number[], flags: boolean[] }
 */
function findAstronomicalPoints(
  values,
  centre,
  referenceValues = null,
  threshold = 3.5
) {
  const source =
    (
      Array.isArray(
        referenceValues
      ) &&
      referenceValues.length >= 3
    )
      ? referenceValues
      : values;

  const ref =
    (source || [])
      .filter(
        v =>
          Number.isFinite(v)
      );

  const flags =
    new Array(
      values.length
    ).fill(false);

  const indices = [];

  if (
    ref.length < 3 ||
    !Number.isFinite(
      centre
    )
  ) {
    return {
      indices,
      flags,
      mad: 0
    };
  }

  const refMedian =
    centre;

  const mad =
    computeMAD(
      ref,
      refMedian
    );

  if (
    !mad ||
    mad === 0 ||
    !Number.isFinite(mad)
  ) {
    return {
      indices,
      flags,
      mad: 0
    };
  }

  const c =
    0.6745;

  for (
    let i = 0;
    i < values.length;
    i++
  ) {
    const value =
      values[i];

    if (
      !Number.isFinite(
        value
      )
    ) {
      continue;
    }

    const z =
      (
        c *
        (
          value -
          refMedian
        )
      ) /
      mad;

    if (
      Math.abs(z) >
      threshold
    ) {
      flags[i] = true;
      indices.push(i);
    }
  }

  return {
    indices,
    flags,
    mad
  };
}

function computeMedian(values) {
  const sorted = [...values].sort((a, b) => a - b);
  const n = sorted.length;
  if (n === 0) return NaN;
  if (n % 2 === 1) return sorted[(n - 1) / 2];
  return (sorted[n / 2 - 1] + sorted[n / 2]) / 2;
}

/**
 * Detect runs of >= runLength points on the same side of the centre line.
 */
function detectLongRuns(values, centre, runLength = 8) {
  const flags = new Array(values.length).fill(false);

  let start = 0;
  while (start < values.length) {
    const v = values[start];
    const side = v > centre ? "above" : v < centre ? "below" : "on";

    if (side === "on") {
      start++;
      continue;
    }

    // extend this run while points stay on the same side
    let end = start + 1;
    while (end < values.length) {
      const v2 = values[end];
      const side2 = v2 > centre ? "above" : v2 < centre ? "below" : "on";
      if (side2 !== side) break;
      end++;
    }

    const length = end - start;
    if (length >= runLength) {
      for (let i = start; i < end; i++) {
        flags[i] = true;
      }
    }

    start = end;
  }

  return flags;
}

/**
 * Detect simple trend: >= length points all increasing or all decreasing
 */
function detectTrend(values, length = 6) {
  if (values.length < length) return false;

  let incRun = 1;
  let decRun = 1;

  for (let i = 1; i < values.length; i++) {
    if (values[i] > values[i - 1]) {
      incRun++;
      decRun = 1;
    } else if (values[i] < values[i - 1]) {
      decRun++;
      incRun = 1;
    } else {
      incRun = 1;
      decRun = 1;
    }

    if (incRun >= length || decRun >= length) {
      return true;
    }
  }
  return false;
}

function populateSplitOptions(labels) {
  if (!splitPointSelect) return;

  splitPointSelect.innerHTML = "";

  if (!labels || labels.length <= 1) {
    const opt = document.createElement("option");
    opt.value = "";
    opt.textContent = "Not enough points to split";
    splitPointSelect.appendChild(opt);
    splitPointSelect.disabled = true;
    if (addSplitButton) addSplitButton.disabled = true;
    return;
  }

  splitPointSelect.disabled = false;
  if (addSplitButton) addSplitButton.disabled = false;

  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "Select point…";
  splitPointSelect.appendChild(placeholder);

  // You can split after any point except the last one
  for (let i = 0; i < labels.length - 1; i++) {
    const opt = document.createElement("option");
    opt.value = String(i); // index of the point AFTER which we split
    opt.textContent = `After ${labels[i]} (point ${i + 1})`;
    splitPointSelect.appendChild(opt);
  }
}

const MIN_XMR_LIMIT_POINTS = 12;

/**
 * Compute XmR statistics and MR values.
 *
 * By default, unusually large baseline moving ranges are screened
 * when estimating routine variation using the Nelson method:
 *
 *   screening threshold = 3.5 × initial average moving range
 *
 * Moving ranges above that threshold are excluded from the
 * variation estimate only. The original observations and moving
 * ranges remain in the chart.
 */
function computeXmR(
  points,
  baselineCount,
  clampLclAtZero = false,
  screenLargeMRs = true
) {
  const pts =
    [...points].sort(
      (a, b) =>
        a.x - b.x
    );

  let baselineCountUsed;

  if (
    baselineCount &&
    baselineCount >= 2
  ) {
    baselineCountUsed =
      Math.min(
        baselineCount,
        pts.length
      );
  } else {
    baselineCountUsed =
      pts.length;
  }

  // Keep the baseline WINDOW fixed.
  // Excluded points inside it are omitted;
  // later observations are not pulled in.
  const baseline =
    pts.slice(
      0,
      baselineCountUsed
    );

  const eligibleBaseline =
    baseline.filter(
      p =>
        !p._excludeFromLimits
    );

  const mean =
    eligibleBaseline.length
      ? eligibleBaseline.reduce(
          (sum, p) =>
            sum + p.y,
          0
        ) /
        eligibleBaseline.length
      : NaN;

  // An excluded XmR observation removes
  // both adjacent moving ranges.
  // We do not bridge across it.
  const baselineMRs = [];

  let manualAdjacentMRExclusions =
    0;

  for (
    let i = 1;
    i < baseline.length;
    i++
  ) {
    if (
      baseline[i - 1]
        ?._excludeFromLimits ||
      baseline[i]
        ?._excludeFromLimits
    ) {
      manualAdjacentMRExclusions +=
        1;

      continue;
    }

    baselineMRs.push({
      value:
        Math.abs(
          baseline[i].y -
          baseline[i - 1].y
        ),

      fromPointIndex:
        i - 1,

      toPointIndex:
        i
    });
  }

  const baselineMRValues =
    baselineMRs.map(
      item =>
        item.value
    );

  const initialAvgMR =
    baselineMRValues.length
      ? baselineMRValues.reduce(
          (sum, v) =>
            sum + v,
          0
        ) /
        baselineMRValues.length
      : 0;

  const mrScreeningThreshold =
    screenLargeMRs &&
    initialAvgMR > 0
      ? 3.5 *
        initialAvgMR
      : null;

  const includedBaselineMRs =
    [];

  const excludedMRs =
    [];

  baselineMRs.forEach(
    (
      mrInfo,
      mrIndex
    ) => {
      const mr =
        mrInfo.value;

      const shouldExclude =
        screenLargeMRs &&
        Number.isFinite(
          mrScreeningThreshold
        ) &&
        mr >
          mrScreeningThreshold;

      if (shouldExclude) {
        excludedMRs.push({
          mrIndex,

          fromPointIndex:
            mrInfo
              .fromPointIndex,

          toPointIndex:
            mrInfo
              .toPointIndex,

          fromLabel:
            baseline[
              mrInfo
                .fromPointIndex
            ]?.label ??
            null,

          toLabel:
            baseline[
              mrInfo
                .toPointIndex
            ]?.label ??
            null,

          value: mr
        });

      } else {
        includedBaselineMRs.push(
          mr
        );
      }
    }
  );

  const avgMR =
    includedBaselineMRs.length
      ? includedBaselineMRs.reduce(
          (sum, v) =>
            sum + v,
          0
        ) /
        includedBaselineMRs.length
      : initialAvgMR;

  const sigma =
    avgMR === 0
      ? 0
      : avgMR / 1.128;

  const ucl =
    mean +
    3 * sigma;

  const rawLcl =
    mean -
    3 * sigma;

  const lcl =
    (
      clampLclAtZero &&
      rawLcl < 0
    )
      ? 0
      : rawLcl;

  // Keep all observed moving ranges
  // for display on the MR chart.
  const mrValues = [];

  for (
    let i = 1;
    i < pts.length;
    i++
  ) {
    mrValues.push(
      Math.abs(
        pts[i].y -
        pts[i - 1].y
      )
    );
  }

  const flagged =
    pts.map(
      p => ({
        ...p,

        beyondLimits:
          !p._excludeFromLimits &&
          sigma > 0 &&
          (
            p.y > ucl ||
            p.y < lcl
          )
      })
    );

  return {
    points: flagged,

    mean,
    ucl,
    lcl,
    rawLcl,
    sigma,
    avgMR,

    initialAvgMR,

    mrScreeningEnabled:
      !!screenLargeMRs,

    mrScreeningThreshold,

    excludedMRCount:
      excludedMRs.length,

    excludedMRs,

    baselineMRCount:
      baselineMRValues.length,

    effectiveMRCount:
      includedBaselineMRs.length,

    manualExcludedPointCount:
      baseline.filter(
        p =>
          p._excludeFromLimits
      ).length,

    manualAdjacentMRExclusions,

    baselineCountUsed,

    effectiveBaselinePointCount:
      eligibleBaseline.length,

    mrValues
  };
}


// -----------------------------
// Attribute chart calculations (C / P / U)
// -----------------------------

function computeC(points, baselineCount = null) {
  const n = points.length;
  const baseN = baselineCount && baselineCount >= 1 ? Math.min(baselineCount, n) : n;

  const baseVals =
  points
    .slice(0, baseN)
    .filter(
      p =>
        !p._excludeFromLimits
    )
    .map(
      p => p.y
    )
    .filter(
      v => isFinite(v)
    );
  const cbar = baseVals.reduce((a, b) => a + b, 0) / baseVals.length;

  const sigma = Math.sqrt(Math.max(cbar, 0));
  const ucl = cbar + 3 * sigma;
  const lcl = Math.max(0, cbar - 3 * sigma); // ✅ always clamp

  const beyond =
  points.map(
    p =>
      !p._excludeFromLimits &&
      isFinite(p.y) &&
      (
        p.y > ucl ||
        p.y < lcl
      )
  );
  return { cbar, ucl, lcl, beyond };
}


// For P and U we expect points like: {x, y: numerator, n: denominator/opportunities}
function computeP(points, baselineCount = null, clampLclAtZero = false) {
  const nPts = points.length;
  const baseN = baselineCount && baselineCount >= 1 ? Math.min(baselineCount, nPts) : nPts;

  const base =
  points
    .slice(0, baseN)
    .filter(
      p =>
        !p._excludeFromLimits &&
        isFinite(p.y) &&
        isFinite(p.n) &&
        p.n > 0
    );
  const sumD = base.reduce((acc, p) => acc + p.y, 0);
  const sumN = base.reduce((acc, p) => acc + p.n, 0);

  const pbar = sumN > 0 ? (sumD / sumN) : NaN;

  const pVals = new Array(nPts).fill(NaN);
  const ucl = new Array(nPts).fill(NaN);
  const lcl = new Array(nPts).fill(NaN);
  const rawLcl = new Array(nPts).fill(NaN);
  const beyond = new Array(nPts).fill(false);

  for (let i = 0; i < nPts; i++) {
    const d = points[i].y;
    const ni = points[i].n;

    if (!isFinite(d) || !isFinite(ni) || ni <= 0 || !isFinite(pbar)) continue;

    const pi = d / ni;
    pVals[i] = pi;

    const sigma = Math.sqrt(Math.max(pbar * (1 - pbar) / ni, 0));
    const u = pbar + 3 * sigma;
    const lRaw = pbar - 3 * sigma;
    const l = clampLclAtZero ? Math.max(0, lRaw) : lRaw;

    // P chart limits should not exceed [0,1]
    ucl[i] = Math.min(1, u);
    rawLcl[i] = lRaw;
    lcl[i] = Math.max(0, Math.min(1, l));

    beyond[i] =
  !points[i]
    ?._excludeFromLimits &&
  (
    pi > ucl[i] ||
    pi < lcl[i]
  );
  }

  return { pbar, pVals, ucl, lcl, rawLcl, beyond };
}

function computePPrime(points, baselineCount = null) {
  const nPts = points.length;

  const baseN =
    baselineCount && baselineCount >= 1
      ? Math.min(baselineCount, nPts)
      : nPts;

  // Baseline data used to estimate p-bar
  const base = points
    .slice(0, baseN)
    .filter(
      p =>
        !p._excludeFromLimits &&
	isFinite(p.y) &&
        isFinite(p.n) &&
        p.n > 0
    );

  const sumD = base.reduce(
    (acc, p) => acc + p.y,
    0
  );

  const sumN = base.reduce(
    (acc, p) => acc + p.n,
    0
  );

  const pbar =
    sumN > 0
      ? sumD / sumN
      : NaN;

  const pVals =
    new Array(nPts).fill(NaN);

  const zValues =
    new Array(nPts).fill(NaN);

  const ucl =
    new Array(nPts).fill(NaN);

  const lcl =
    new Array(nPts).fill(NaN);

  const rawLcl =
    new Array(nPts).fill(NaN);

  const beyond =
    new Array(nPts).fill(false);

  // First calculate proportions and standardised z-values.
  for (let i = 0; i < nPts; i++) {
    const d = points[i].y;
    const ni = points[i].n;

    if (
      !isFinite(d) ||
      !isFinite(ni) ||
      ni <= 0 ||
      !isFinite(pbar)
    ) {
      continue;
    }

    const pi = d / ni;
    pVals[i] = pi;

    const standardSigma =
      Math.sqrt(
        Math.max(
          pbar * (1 - pbar) / ni,
          0
        )
      );

    if (standardSigma > 0) {
      zValues[i] =
        (pi - pbar) /
        standardSigma;
    }
  }

  // Laney Sigma Z:
  // moving ranges of consecutive baseline z-scores.
  const mrValues = [];

  for (let i = 1; i < baseN; i++) {
    const zPrev = zValues[i - 1];
    const zNow = zValues[i];

    if (
  !points[i - 1]
    ?._excludeFromLimits &&
  !points[i]
    ?._excludeFromLimits &&
  Number.isFinite(zPrev) &&
  Number.isFinite(zNow)
) {
      mrValues.push(
        Math.abs(zNow - zPrev)
      );
    }
  }

  const avgMR =
    mrValues.length
      ? mrValues.reduce(
          (a, b) => a + b,
          0
        ) / mrValues.length
      : NaN;

  const sigmaZ =
    Number.isFinite(avgMR)
      ? avgMR / 1.128
      : NaN;

  // Apply the Laney adjustment to each subgroup's
  // ordinary binomial standard deviation.
  for (let i = 0; i < nPts; i++) {
    const ni = points[i].n;
    const pi = pVals[i];

    if (
      !Number.isFinite(pi) ||
      !Number.isFinite(ni) ||
      ni <= 0 ||
      !Number.isFinite(pbar) ||
      !Number.isFinite(sigmaZ)
    ) {
      continue;
    }

    const standardSigma =
      Math.sqrt(
        Math.max(
          pbar * (1 - pbar) / ni,
          0
        )
      );

    const adjustedSigma =
      standardSigma * sigmaZ;

    const u =
      pbar +
      3 * adjustedSigma;

    const lRaw =
      pbar -
      3 * adjustedSigma;

    // P′ limits remain bounded between 0 and 1.
    ucl[i] =
      Math.min(1, u);

    rawLcl[i] =
      lRaw;

    lcl[i] =
      Math.max(
        0,
        Math.min(1, lRaw)
      );

    beyond[i] =
  !points[i]
    ?._excludeFromLimits &&
  (
    pi > ucl[i] ||
    pi < lcl[i]
  );
  }

  return {
    pbar,
    pVals,
    zValues,
    mrValues,
    avgMR,
    sigmaZ,
    ucl,
    lcl,
    rawLcl,
    beyond,
    baselineCountUsed: baseN,
    baselineMRCount: mrValues.length
  };
}

function computePDispersionDiagnostic(
  points,
  baselineCount = null
) {
  const nPts =
    Array.isArray(points)
      ? points.length
      : 0;

  const baseN =
    baselineCount && baselineCount >= 1
      ? Math.min(baselineCount, nPts)
      : nPts;

  const base =
    (points || [])
      .slice(0, baseN)
      .filter(p =>
        !p._excludeFromLimits &&
	Number.isFinite(Number(p.y)) &&
        Number.isFinite(Number(p.n)) &&
        Number(p.n) > 0 &&
        Number(p.y) >= 0 &&
        Number(p.y) <= Number(p.n)
      )
      .map(p => ({
        d: Number(p.y),
        n: Number(p.n)
      }));

  const m = base.length;

  if (m < 4) {
    return {
      ok: false,
      reason:
        "Too few usable subgroups for the dispersion check."
    };
  }

  const sumD =
    base.reduce(
      (acc, p) => acc + p.d,
      0
    );

  const sumN =
    base.reduce(
      (acc, p) => acc + p.n,
      0
    );

  const nBar =
    sumN / m;

  const pBar =
    sumN > 0
      ? sumD / sumN
      : NaN;

  if (
    !Number.isFinite(nBar) ||
    nBar <= 0 ||
    !Number.isFinite(pBar) ||
    pBar <= 0 ||
    pBar >= 1
  ) {
    return {
      ok: false,
      reason:
        "The dispersion check could not be calculated for these data."
    };
  }

  // -------------------------------------------------
  // P-chart diagnostic:
  // adjust each count to the average subgroup size
  // and apply the arcsine-square-root transformation.
  // -------------------------------------------------
  const transformed =
    base
      .map(p => {
        const adjustedCount =
          (p.d / p.n) * nBar;

        const fraction =
          (adjustedCount + 0.375) /
          (nBar + 0.75);

        const safeFraction =
          Math.max(
            0,
            Math.min(1, fraction)
          );

        return Math.asin(
          Math.sqrt(safeFraction)
        );
      })
      .filter(Number.isFinite)
      .sort((a, b) => a - b);

  if (transformed.length < 4) {
    return {
      ok: false,
      reason:
        "The dispersion check could not be calculated for these data."
    };
  }

  // Inverse standard-normal CDF.
  // Uses the existing normalCdf() helper via binary search.
  function inverseNormalCdf(probability) {
    if (
      !Number.isFinite(probability) ||
      probability <= 0 ||
      probability >= 1
    ) {
      return NaN;
    }

    let low = -8;
    let high = 8;

    for (let i = 0; i < 60; i++) {
      const mid =
        (low + high) / 2;

      if (normalCdf(mid) < probability) {
        low = mid;
      } else {
        high = mid;
      }
    }

    return (low + high) / 2;
  }

  // Give tied transformed values the same
  // normal score using their average rank.
  const ranked = [];

  let i = 0;

  while (i < transformed.length) {
    let j = i + 1;

    while (
      j < transformed.length &&
      Math.abs(
        transformed[j] -
        transformed[i]
      ) <= 1e-12
    ) {
      j++;
    }

    // Array indices i ... j-1 correspond
    // to ranks i+1 ... j.
    const meanRank =
      ((i + 1) + j) / 2;

    const probability =
      (meanRank - 0.375) /
      (transformed.length + 0.25);

    const normalScore =
      inverseNormalCdf(probability);

    for (let k = i; k < j; k++) {
      ranked.push({
        x: transformed[k],
        y: normalScore
      });
    }

    i = j;
  }

  function quantile(sortedValues, q) {
    if (!sortedValues.length) {
      return NaN;
    }

    const position =
      (sortedValues.length - 1) * q;

    const lower =
      Math.floor(position);

    const upper =
      Math.ceil(position);

    if (lower === upper) {
      return sortedValues[lower];
    }

    const fraction =
      position - lower;

    return (
      sortedValues[lower] +
      (
        sortedValues[upper] -
        sortedValues[lower]
      ) * fraction
    );
  }

  const q1 =
    quantile(transformed, 0.25);

  const q3 =
    quantile(transformed, 0.75);

  // Published diagnostic uses only the
  // middle 50% for the fitted line.
  const middle =
    ranked.filter(item =>
      item.x >= q1 &&
      item.x <= q3 &&
      Number.isFinite(item.y)
    );

  if (middle.length < 2) {
    return {
      ok: false,
      reason:
        "There were not enough usable values for the dispersion check."
    };
  }

  const meanX =
    middle.reduce(
      (acc, item) => acc + item.x,
      0
    ) / middle.length;

  const meanY =
    middle.reduce(
      (acc, item) => acc + item.y,
      0
    ) / middle.length;

  let numerator = 0;
  let denominator = 0;

  middle.forEach(item => {
    numerator +=
      (item.x - meanX) *
      (item.y - meanY);

    denominator +=
      Math.pow(
        item.x - meanX,
        2
      );
  });

  if (
    !Number.isFinite(denominator) ||
    denominator <= 0
  ) {
    return {
      ok: false,
      reason:
        "There was too little variation to calculate the dispersion check."
    };
  }

  const slope =
    numerator / denominator;

  if (
    !Number.isFinite(slope) ||
    slope <= 0
  ) {
    return {
      ok: false,
      reason:
        "The dispersion check could not be calculated reliably."
    };
  }

  const observedVariation =
    1 / slope;

  const expectedVariation =
    1 /
    Math.sqrt(
      4 * nBar
    );

  const ratio =
    (
      observedVariation /
      expectedVariation
    ) * 100;

  // 95% upper confidence limit for the
  // observed/expected variation ratio.
  const upperLimit =
    Math.exp(
      0.185 +
      (5.62 / m) +
      (
        0.274 /
        (nBar * pBar)
      )
    ) * 100;

  // Published method uses a fixed
  // lower confidence limit of 60%.
  const lowerLimit = 60;

  let status = "expected";

  if (ratio > upperLimit) {
    status = "overdispersed";
  } else if (ratio < lowerLimit) {
    status = "underdispersed";
  }

  return {
    ok: true,
    m,
    nBar,
    pBar,
    observedVariation,
    expectedVariation,
    ratio,
    lowerLimit,
    upperLimit,
    status,
    recommendPPrime:
      status === "overdispersed" ||
      status === "underdispersed"
  };
}

function computeU(points, baselineCount = null) {
  const nPts = points.length;
  const baseN = baselineCount && baselineCount >= 1 ? Math.min(baselineCount, nPts) : nPts;

  const base =
  points
    .slice(0, baseN)
    .filter(
      p =>
        !p._excludeFromLimits &&
        isFinite(p.y) &&
        isFinite(p.n) &&
        p.n > 0
    );
  const sumC = base.reduce((acc, p) => acc + p.y, 0);
  const sumN = base.reduce((acc, p) => acc + p.n, 0);

  const ubar = sumN > 0 ? (sumC / sumN) : NaN;

  const uVals = new Array(nPts).fill(NaN);
  const ucl = new Array(nPts).fill(NaN);
  const lcl = new Array(nPts).fill(NaN);
  const beyond = new Array(nPts).fill(false);

  for (let i = 0; i < nPts; i++) {
    const c = points[i].y;
    const ni = points[i].n;

    if (!isFinite(c) || !isFinite(ni) || ni <= 0 || !isFinite(ubar)) continue;

    const ui = c / ni;
    uVals[i] = ui;

    const sigma = Math.sqrt(Math.max(ubar / ni, 0));
    const u = ubar + 3 * sigma;
    const l = Math.max(0, ubar - 3 * sigma); // ✅ always clamp

    ucl[i] = u;
    lcl[i] = l;

    beyond[i] =
  !points[i]
    ?._excludeFromLimits &&
  (
    ui > ucl[i] ||
    ui < lcl[i]
  );
  }

  return { ubar, uVals, ucl, lcl, beyond };
}

function computeUPrime(points, baselineCount = null) {
  const nPts = points.length;

  const baseN =
    baselineCount && baselineCount >= 1
      ? Math.min(baselineCount, nPts)
      : nPts;

  // Baseline data used to estimate u-bar
  const base = points
    .slice(0, baseN)
    .filter(
      p =>
	!p._excludeFromLimits &&        
	Number.isFinite(Number(p.y)) &&
        Number.isFinite(Number(p.n)) &&
        Number(p.n) > 0
    );

  const sumC = base.reduce(
    (acc, p) => acc + Number(p.y),
    0
  );

  const sumN = base.reduce(
    (acc, p) => acc + Number(p.n),
    0
  );

  const ubar =
    sumN > 0
      ? sumC / sumN
      : NaN;

  const uVals =
    new Array(nPts).fill(NaN);

  const zValues =
    new Array(nPts).fill(NaN);

  const ucl =
    new Array(nPts).fill(NaN);

  const lcl =
    new Array(nPts).fill(NaN);

  const beyond =
    new Array(nPts).fill(false);

  // Convert each subgroup rate to a z-score.
  for (let i = 0; i < nPts; i++) {
    const c = Number(points[i].y);
    const ni = Number(points[i].n);

    if (
      !Number.isFinite(c) ||
      !Number.isFinite(ni) ||
      ni <= 0 ||
      !Number.isFinite(ubar)
    ) {
      continue;
    }

    const ui =
      c / ni;

    uVals[i] =
      ui;

    // Standard U-chart standard deviation
    const standardSigma =
      Math.sqrt(
        Math.max(
          ubar / ni,
          0
        )
      );

    if (standardSigma > 0) {
      zValues[i] =
        (ui - ubar) /
        standardSigma;
    }
  }

  // Laney Sigma Z:
  // moving ranges of consecutive baseline z-scores.
  const mrValues = [];

  for (let i = 1; i < baseN; i++) {
    const zPrev =
      zValues[i - 1];

    const zNow =
      zValues[i];

    if (
  !points[i - 1]
    ?._excludeFromLimits &&
  !points[i]
    ?._excludeFromLimits &&
  Number.isFinite(zPrev) &&
  Number.isFinite(zNow)
) {
      mrValues.push(
        Math.abs(
          zNow - zPrev
        )
      );
    }
  }

  const avgMR =
    mrValues.length
      ? mrValues.reduce(
          (a, b) => a + b,
          0
        ) / mrValues.length
      : NaN;

  const sigmaZ =
    Number.isFinite(avgMR)
      ? avgMR / 1.128
      : NaN;

  // Apply the Laney adjustment to the
  // usual U-chart standard deviation.
  for (let i = 0; i < nPts; i++) {
    const ni =
      Number(points[i].n);

    const ui =
      uVals[i];

    if (
      !Number.isFinite(ui) ||
      !Number.isFinite(ni) ||
      ni <= 0 ||
      !Number.isFinite(ubar) ||
      !Number.isFinite(sigmaZ)
    ) {
      continue;
    }

    const standardSigma =
      Math.sqrt(
        Math.max(
          ubar / ni,
          0
        )
      );

    const adjustedSigma =
      standardSigma *
      sigmaZ;

    ucl[i] =
      ubar +
      3 * adjustedSigma;

    lcl[i] =
      Math.max(
        0,
        ubar -
        3 * adjustedSigma
      );

    beyond[i] =
  !points[i]
    ?._excludeFromLimits &&
  (
    ui > ucl[i] ||
    ui < lcl[i]
  );
  }

  return {
    ubar,
    uVals,
    zValues,
    mrValues,
    avgMR,
    sigmaZ,
    ucl,
    lcl,
    beyond,
    baselineCountUsed: baseN,
    baselineMRCount: mrValues.length
  };
}

function computeUDispersionDiagnostic(
  points,
  baselineCount = null
) {
  const nPts =
    Array.isArray(points)
      ? points.length
      : 0;

  const baseN =
    baselineCount && baselineCount >= 1
      ? Math.min(baselineCount, nPts)
      : nPts;

  const base =
    (points || [])
      .slice(0, baseN)
      .filter(p =>
        !p._excludeFromLimits &&
	Number.isFinite(Number(p.y)) &&
        Number.isFinite(Number(p.n)) &&
        Number(p.y) >= 0 &&
        Number(p.n) > 0
      )
      .map(p => ({
        c: Number(p.y),
        n: Number(p.n)
      }));

  const m = base.length;

  if (m < 4) {
    return {
      ok: false,
      reason:
        "Too few usable subgroups for the dispersion check."
    };
  }

  const sumC =
    base.reduce(
      (acc, p) => acc + p.c,
      0
    );

  const sumN =
    base.reduce(
      (acc, p) => acc + p.n,
      0
    );

  const nBar =
    sumN / m;

  const uBar =
    sumN > 0
      ? sumC / sumN
      : NaN;

  if (
    !Number.isFinite(nBar) ||
    nBar <= 0 ||
    !Number.isFinite(uBar) ||
    uBar <= 0
  ) {
    return {
      ok: false,
      reason:
        "The dispersion check could not be calculated for these data."
    };
  }

  // Adjust each subgroup count to the
  // average subgroup size.
  const transformed =
    base
      .map(p => {
        const adjustedCount =
          (p.c / p.n) * nBar;

        return (
          Math.sqrt(adjustedCount) +
          Math.sqrt(adjustedCount + 1)
        );
      })
      .filter(Number.isFinite)
      .sort((a, b) => a - b);

  if (transformed.length < 4) {
    return {
      ok: false,
      reason:
        "The dispersion check could not be calculated for these data."
    };
  }

  function inverseNormalCdf(probability) {
    if (
      !Number.isFinite(probability) ||
      probability <= 0 ||
      probability >= 1
    ) {
      return NaN;
    }

    let low = -8;
    let high = 8;

    for (let i = 0; i < 60; i++) {
      const mid =
        (low + high) / 2;

      if (normalCdf(mid) < probability) {
        low = mid;
      } else {
        high = mid;
      }
    }

    return (low + high) / 2;
  }

  function quantile(sortedValues, q) {
    if (!sortedValues.length) {
      return NaN;
    }

    const position =
      (sortedValues.length - 1) * q;

    const lower =
      Math.floor(position);

    const upper =
      Math.ceil(position);

    if (lower === upper) {
      return sortedValues[lower];
    }

    const fraction =
      position - lower;

    return (
      sortedValues[lower] +
      (
        sortedValues[upper] -
        sortedValues[lower]
      ) * fraction
    );
  }

  // Give tied transformed values the same
  // normal score using their average rank.
  const ranked = [];

  let i = 0;

  while (i < transformed.length) {
    let j = i + 1;

    while (
      j < transformed.length &&
      Math.abs(
        transformed[j] -
        transformed[i]
      ) <= 1e-12
    ) {
      j++;
    }

    const meanRank =
      ((i + 1) + j) / 2;

    const probability =
      (meanRank - 0.375) /
      (transformed.length + 0.25);

    const normalScore =
      inverseNormalCdf(probability);

    for (let k = i; k < j; k++) {
      ranked.push({
        x: transformed[k],
        y: normalScore
      });
    }

    i = j;
  }

  const q1 =
    quantile(transformed, 0.25);

  const q3 =
    quantile(transformed, 0.75);

  const middle =
    ranked.filter(item =>
      item.x >= q1 &&
      item.x <= q3 &&
      Number.isFinite(item.y)
    );

  if (middle.length < 2) {
    return {
      ok: false,
      reason:
        "There were not enough usable values for the dispersion check."
    };
  }

  const meanX =
    middle.reduce(
      (acc, item) => acc + item.x,
      0
    ) / middle.length;

  const meanY =
    middle.reduce(
      (acc, item) => acc + item.y,
      0
    ) / middle.length;

  let numerator = 0;
  let denominator = 0;

  middle.forEach(item => {
    numerator +=
      (item.x - meanX) *
      (item.y - meanY);

    denominator +=
      Math.pow(
        item.x - meanX,
        2
      );
  });

  if (
    !Number.isFinite(denominator) ||
    denominator <= 0
  ) {
    return {
      ok: false,
      reason:
        "There was too little variation to calculate the dispersion check."
    };
  }

  const slope =
    numerator / denominator;

  if (
    !Number.isFinite(slope) ||
    slope <= 0
  ) {
    return {
      ok: false,
      reason:
        "The dispersion check could not be calculated reliably."
    };
  }

  // For the U diagnostic, expected variation
  // after transformation is 1.
  const observedVariation =
    1 / slope;

  const expectedVariation = 1;

  const ratio =
    (
      observedVariation /
      expectedVariation
    ) * 100;

  const upperLimit =
    Math.exp(
      0.182 +
      (5.75 / m) +
      (
        0.195 /
        (nBar * uBar)
      )
    ) * 100;

  const lowerLimit = 60;

  let status = "expected";

  if (ratio > upperLimit) {
    status = "overdispersed";
  } else if (ratio < lowerLimit) {
    status = "underdispersed";
  }

  return {
    ok: true,
    m,
    nBar,
    uBar,
    observedVariation,
    expectedVariation,
    ratio,
    lowerLimit,
    upperLimit,
    status,
    recommendUPrime:
      status === "overdispersed" ||
      status === "underdispersed"
  };
}

// -----------------------------
// X̄–S chart calculations + drawing
// -----------------------------

function computeXbarS(points, baselineCount = null) {
  // points: array of { x, y, label, _rowIndex }, where y is the raw measurement
  // subgroup id is read later from rawRows via _rowIndex

  // Placeholder — this function just returns computed subgroup summaries
  // Actual grouping done in drawXbarSChart where we know which column is subgroup.
  return null;
}

function groupBySubgroup(points, subgroupCol) {
  const map = new Map();
  for (const p of points) {
    const row = rawRows[p._rowIndex];
    const sgRaw = row ? row[subgroupCol] : null;
    const sg = (sgRaw === null || sgRaw === undefined || String(sgRaw).trim() === "")
      ? "(missing subgroup)"
      : String(sgRaw);

    if (!map.has(sg)) map.set(sg, []);
    map.get(sg).push(p);
  }
  return map;
}

function mean(arr) {
  return arr.reduce((a,b) => a + b, 0) / arr.length;
}

function sampleStdDev(arr) {
  if (arr.length < 2) return NaN;
  const m = mean(arr);
  const v = arr.reduce((acc, x) => acc + (x - m) * (x - m), 0) / (arr.length - 1);
  return Math.sqrt(v);
}

function drawXbarSChart(
  points,
  baselineCount,
  labels
) {
  if (
    !thirdSelect ||
    !thirdSelect.value
  ) {
    showError(
      "X̄–S chart needs a third column: Subgroup ID."
    );

    return;
  }

  const subgroupCol =
    thirdSelect.value;

  const groups =
    groupBySubgroup(
      points,
      subgroupCol
    );

  const subgroupKeys =
    Array.from(
      groups.keys()
    );

  const xbarVals = [];
  const sVals = [];
  const subgroupLabels = [];
  const subgroupSizes = [];

  subgroupKeys.forEach(
    (k) => {
      const arr =
        groups.get(k) || [];

      const vals =
        arr
          .map(
            p => p.y
          )
          .filter(
            v => isFinite(v)
          );

      if (
        vals.length < 2
      ) {
        return;
      }

      xbarVals.push(
        mean(vals)
      );

      sVals.push(
        sampleStdDev(vals)
      );

      subgroupLabels.push(
        String(k)
      );

      subgroupSizes.push(
        vals.length
      );
    }
  );

  if (
    xbarVals.length < 4
  ) {
    showError(
      "X̄–S needs at least 4 subgroups."
    );

    return;
  }

  const m =
    xbarVals.length;

    const excludedSubgroupFlags =
    subgroupLabels.map(
      label =>
        isLimitExcluded(
          getLimitExclusionScope(
            "xbars"
          ),
          makeSubgroupLimitExclusionKey(
            label
          )
        )
    );

  // --------------------------------
  // Define process periods from splits
  // --------------------------------
  let effectiveSplits =
    Array.isArray(splits)
      ? splits.slice()
      : [];

  effectiveSplits =
    effectiveSplits
      .filter(
        i =>
          Number.isInteger(i) &&
          i >= 0 &&
          i < m - 1
      )
      .sort(
        (a, b) => a - b
      );

  const segmentStarts = [0];
  const segmentEnds = [];

  effectiveSplits.forEach(
    idx => {
      segmentEnds.push(
        idx
      );

      segmentStarts.push(
        idx + 1
      );
    }
  );

  segmentEnds.push(
    m - 1
  );

  const clX =
    new Array(m).fill(NaN);

  const uclXArr =
    new Array(m).fill(NaN);

  const lclXArr =
    new Array(m).fill(NaN);

  const clS =
    new Array(m).fill(NaN);

  const uclSArr =
    new Array(m).fill(NaN);

  const lclSArr =
    new Array(m).fill(NaN);

  const segmentStats = [];

  // --------------------------------
  // Calculate limits independently
  // for every process period
  // --------------------------------
  for (
    let segmentIndex = 0;
    segmentIndex <
      segmentStarts.length;
    segmentIndex++
  ) {
    const start =
      segmentStarts[
        segmentIndex
      ];

    const end =
      segmentEnds[
        segmentIndex
      ];

    const segX =
      xbarVals.slice(
        start,
        end + 1
      );

    const segS =
      sVals.slice(
        start,
        end + 1
      );

    const segN =
      subgroupSizes.slice(
        start,
        end + 1
      );

    const segBaselineCountUsed =
      (
        segmentIndex === 0 &&
        baselineCount &&
        baselineCount >= 2
      )
        ? Math.min(
            baselineCount,
            segX.length
          )
        : segX.length;

        const baselineLocalIndices =
      Array.from(
        {
          length:
            segBaselineCountUsed
        },
        (_, i) => i
      )
        .filter(
          i =>
            !excludedSubgroupFlags[
              start + i
            ]
        );

    const baseX =
      baselineLocalIndices
        .map(
          i =>
            segX[i]
        );

    const baseS =
      baselineLocalIndices
        .map(
          i =>
            segS[i]
        );

    const baseN =
      baselineLocalIndices
        .map(
          i =>
            segN[i]
        );

    if (
      baseX.length < 2
    ) {
      showError(
        "X̄–S needs at least 2 included baseline subgroups in each period to estimate limits."
      );

      return;
    }

    const totalBaselineN =
      baseN.reduce(
        (acc, n) =>
          acc + n,
        0
      );

    const weightedBaselineSum =
      baseX.reduce(
        (
          acc,
          xbar,
          i
        ) =>
          acc +
          xbar *
          baseN[i],
        0
      );

    const processMean =
      totalBaselineN > 0
        ? weightedBaselineSum /
          totalBaselineN
        : NaN;

    const sigmaHat =
      estimateXbarSSigmaFromSbar(
        baseS,
        baseN
      );

    if (
      !Number.isFinite(
        processMean
      ) ||
      !Number.isFinite(
        sigmaHat
      )
    ) {
      showError(
        "Could not calculate X̄–S limits for these subgroup sizes."
      );

      return;
    }

        segmentStats.push({
      processMean,
      sigmaHat,

      baselineCountUsed:
        segBaselineCountUsed,

      effectiveBaselineCount:
        baseX.length,

      excludedBaselineCount:
        segBaselineCountUsed -
        baseX.length
    });

    for (
      let i = start;
      i <= end;
      i++
    ) {
      const ni =
        subgroupSizes[i];

      const c4 =
        c4Constant(ni);

      if (
        !Number.isFinite(c4) ||
        c4 <= 0 ||
        c4 >= 1
      ) {
        showError(
          "Could not calculate X̄–S constants for one or more subgroup sizes."
        );

        return;
      }

      const c5 =
        Math.sqrt(
          Math.max(
            1 -
            c4 * c4,
            0
          )
        );

      // X̄ chart
      clX[i] =
        processMean;

      uclXArr[i] =
        processMean +
        (
          3 *
          sigmaHat /
          Math.sqrt(ni)
        );

      lclXArr[i] =
        processMean -
        (
          3 *
          sigmaHat /
          Math.sqrt(ni)
        );

      // S chart
      clS[i] =
        c4 *
        sigmaHat;

      uclSArr[i] =
        (
          c4 +
          3 * c5
        ) *
        sigmaHat;

      lclSArr[i] =
        Math.max(
          0,
          (
            c4 -
            3 * c5
          ) *
          sigmaHat
        );
    }
  }

  const flagOnChart =
    (
      typeof shouldFlagSpecialCauseOnChart ===
      "function"
    )
      ? shouldFlagSpecialCauseOnChart()
      : true;

  const pointColoursX =
    new Array(m).fill(
      SPC_STYLE.seriesBlue
    );

  const pointColoursS =
    new Array(m).fill(
      SPC_STYLE.seriesBlue
    );

  const periodAnalyses = [];

  // --------------------------------
  // Analyse each period independently.
  // Rules restart after every split.
  // --------------------------------
  for (
    let segmentIndex = 0;
    segmentIndex <
      segmentStarts.length;
    segmentIndex++
  ) {
    const start =
      segmentStarts[
        segmentIndex
      ];

    const end =
      segmentEnds[
        segmentIndex
      ];

    const localLabels =
      subgroupLabels.slice(
        start,
        end + 1
      );

    const xAnalysis =
      analyzeAttributeChart({
        chartType:
          "xbars",

	        excludedIndices:
          Array.from(
            {
              length:
                end -
                start +
                1
            },
            (_, j) => j
          )
            .filter(
              j =>
                excludedSubgroupFlags[
                  start + j
                ]
            ),

        labels:
          localLabels,

        values:
          xbarVals.slice(
            start,
            end + 1
          ),

        cl:
          clX.slice(
            start,
            end + 1
          ),

        ucl:
          uclXArr.slice(
            start,
            end + 1
          ),

        lcl:
          lclXArr.slice(
            start,
            end + 1
          )
      });

    const sAnalysis =
      analyzeAttributeChart({
        chartType:
          "xbars",
	
	        excludedIndices:
          Array.from(
            {
              length:
                end -
                start +
                1
            },
            (_, j) => j
          )
            .filter(
              j =>
                excludedSubgroupFlags[
                  start + j
                ]
            ),

        labels:
          localLabels,

        values:
          sVals.slice(
            start,
            end + 1
          ),

        cl:
          clS.slice(
            start,
            end + 1
          ),

        ucl:
          uclSArr.slice(
            start,
            end + 1
          ),

        lcl:
          lclSArr.slice(
            start,
            end + 1
          )
      });

    if (flagOnChart) {
      for (
        let j = 0;
        j <=
          end - start;
        j++
      ) {
        const globalIndex =
          start + j;

        if (
          xAnalysis.flags
            ?.beyond
            ?.[j]
        ) {
          pointColoursX[
            globalIndex
          ] =
            SPC_STYLE.pointBeyond;

        } else if (
          xAnalysis.flags
            ?.special
            ?.[j]
        ) {
          pointColoursX[
            globalIndex
          ] =
            SPC_STYLE.pointSpecial;
        }

        if (
          sAnalysis.flags
            ?.beyond
            ?.[j]
        ) {
          pointColoursS[
            globalIndex
          ] =
            SPC_STYLE.pointBeyond;

        } else if (
          sAnalysis.flags
            ?.special
            ?.[j]
        ) {
          pointColoursS[
            globalIndex
          ] =
            SPC_STYLE.pointSpecial;
        }
      }
    }

    const periodSizes =
      subgroupSizes.slice(
        start,
        end + 1
      );

    const uniqueSizes =
      Array.from(
        new Set(
          periodSizes
        )
      );

    const variableSubgroupSizes =
      uniqueSizes.length > 1;

    const periodXUcl =
      uclXArr
        .slice(
          start,
          end + 1
        )
        .filter(
          Number.isFinite
        );

    const periodXLcl =
      lclXArr
        .slice(
          start,
          end + 1
        )
        .filter(
          Number.isFinite
        );

    const periodSCl =
      clS
        .slice(
          start,
          end + 1
        )
        .filter(
          Number.isFinite
        );

    const periodSUcl =
      uclSArr
        .slice(
          start,
          end + 1
        )
        .filter(
          Number.isFinite
        );

    const periodSLcl =
      lclSArr
        .slice(
          start,
          end + 1
        )
        .filter(
          Number.isFinite
        );

    const stats =
      segmentStats[
        segmentIndex
      ] || {};

    periodAnalyses.push({
      xbar:
        xAnalysis,

      s:
        sAnalysis,

      periodIndex:
        segmentIndex + 1,

      periodCount:
        segmentStarts.length,

      startIndex:
        start,

      endIndex:
        end,

      labelStart:
        subgroupLabels[start],

      labelEnd:
        subgroupLabels[end],

      stats: {
        subgroupSize:
          variableSubgroupSizes
            ? null
            : periodSizes[0],

        subgroupSizes:
          periodSizes.slice(),

        subgroupSizeMin:
          Math.min(
            ...periodSizes
          ),

        subgroupSizeMax:
          Math.max(
            ...periodSizes
          ),

        variableSubgroupSizes,

        processSigma:
          stats.sigmaHat,

                baselineCountUsed:
          stats.baselineCountUsed,

        effectiveBaselineCount:
          stats.effectiveBaselineCount,

        excludedBaselineCount:
          stats.excludedBaselineCount,

        xbarbar:
          stats.processMean,

        sbar:
          variableSubgroupSizes
            ? null
            : clS[start],

        uclX:
          variableSubgroupSizes
            ? null
            : uclXArr[start],

        lclX:
          variableSubgroupSizes
            ? null
            : lclXArr[start],

        uclS:
          variableSubgroupSizes
            ? null
            : uclSArr[start],

        lclS:
          variableSubgroupSizes
            ? null
            : lclSArr[start],

        xUclMin:
          periodXUcl.length
            ? Math.min(
                ...periodXUcl
              )
            : NaN,

        xUclMax:
          periodXUcl.length
            ? Math.max(
                ...periodXUcl
              )
            : NaN,

        xLclMin:
          periodXLcl.length
            ? Math.min(
                ...periodXLcl
              )
            : NaN,

        xLclMax:
          periodXLcl.length
            ? Math.max(
                ...periodXLcl
              )
            : NaN,

        sClMin:
          periodSCl.length
            ? Math.min(
                ...periodSCl
              )
            : NaN,

        sClMax:
          periodSCl.length
            ? Math.max(
                ...periodSCl
              )
            : NaN,

        sUclMin:
          periodSUcl.length
            ? Math.min(
                ...periodSUcl
              )
            : NaN,

        sUclMax:
          periodSUcl.length
            ? Math.max(
                ...periodSUcl
              )
            : NaN,

        sLclMin:
          periodSLcl.length
            ? Math.min(
                ...periodSLcl
              )
            : NaN,

        sLclMax:
          periodSLcl.length
            ? Math.max(
                ...periodSLcl
              )
            : NaN
      }
    });
  }

  // --------------------------------
  // Draw charts
  // --------------------------------
  drawXbarSCombinedChart({
    labels:
      subgroupLabels,

    xbarVals,
    sVals,

    pointColoursX,
    pointColoursS,

    clX,
    uclXArr,
    lclXArr,

    clS,
    uclSArr,
    lclSArr,

    splitAfterIndices:
      effectiveSplits
  });

  attachSubgroupLimitExclusionContext(
    currentChart,
    subgroupLabels
  );

  if (mrChart) {
    attachSubgroupLimitExclusionContext(
      mrChart,
      subgroupLabels
    );
  }

  // Keep the existing helper contract:
  // lastXbarSAnalysis remains the
  // latest process period.
  lastXbarSAnalysis =
    periodAnalyses.length
      ? periodAnalyses[
          periodAnalyses.length - 1
        ]
      : null;

  if (
    typeof renderXbarSSummary ===
    "function"
  ) {
    renderXbarSSummary(
      periodAnalyses,
      subgroupLabels.length
    );
  }
}

function computeTLimits(
  values,
  baselineCount = null
) {
  const vals =
    (values || [])
      .map(Number)
      .filter(
        v =>
          Number.isFinite(v) &&
          v >= 0
      );

  if (!vals.length) {
    return {
      mean: NaN,
      ucl: NaN,
      lcl: 0,
      baselineCountUsed: 0
    };
  }

  const baselineCountUsed =
    (
      baselineCount &&
      baselineCount >= 1
    )
      ? Math.min(
          baselineCount,
          vals.length
        )
      : vals.length;

  const base =
    vals.slice(
      0,
      baselineCountUsed
    );

  const mean =
    base.reduce(
      (a, b) => a + b,
      0
    ) /
    base.length;

  const qHigh =
    0.99865;

  const ucl =
    -mean *
    Math.log(
      1 - qHigh
    );

  return {
    mean,
    ucl,
    lcl: 0,
    baselineCountUsed
  };
}

function computeGLimits(
  values,
  baselineCount = null
) {
  const vals =
    (values || [])
      .map(Number)
      .filter(
        v =>
          Number.isFinite(v) &&
          v >= 1
      );

  if (!vals.length) {
    return {
      mean: NaN,
      p: NaN,
      ucl: NaN,
      lcl: NaN,
      baselineCountUsed: 0
    };
  }

  const baselineCountUsed =
    (
      baselineCount &&
      baselineCount >= 1
    )
      ? Math.min(
          baselineCount,
          vals.length
        )
      : vals.length;

  const base =
    vals.slice(
      0,
      baselineCountUsed
    );

  const mean =
    base.reduce(
      (a, b) => a + b,
      0
    ) /
    base.length;

  const p =
    mean > 0
      ? 1 / mean
      : NaN;

  if (
    !Number.isFinite(p) ||
    p <= 0 ||
    p >= 1
  ) {
    return {
      mean,
      p,
      ucl: NaN,
      lcl: NaN,
      baselineCountUsed
    };
  }

  function geomQuantile(
    q
  ) {
    const k =
      Math.log(
        1 - q
      ) /
      Math.log(
        1 - p
      );

    return Math.max(
      1,
      Math.ceil(k)
    );
  }

  return {
    mean,
    p,

    ucl:
      geomQuantile(
        0.99865
      ),

    lcl:
      geomQuantile(
        0.00135
      ),

    baselineCountUsed
  };
}

// -----------------------------
// T chart: time between events (Exponential limits via percentiles)
// -----------------------------
function drawTChart(points, baselineCount, labels) {
  // Sort by time
  const pts = [...points].sort((a, b) => a.x - b.x);
  if (pts.length < 4) {
    showError("T chart needs at least 4 events.");
    return;
  }

  const deltas = [];
  const tLabels = [];

  for (let i = 1; i < pts.length; i++) {
    const dtMs = pts[i].x - pts[i - 1].x;
    const days = dtMs / (1000 * 60 * 60 * 24);
    if (isFinite(days) && days >= 0) {
      deltas.push(days);
      tLabels.push(pts[i].label ?? `Event ${i + 1}`);
    }
  }

  if (deltas.length < 3) {
    showError("T chart needs at least 4 events.");
    return;
  }

  const n = deltas.length;

  // ---- Segment definition from splits ----
  let effectiveSplits = Array.isArray(splits) ? splits.slice() : [];
  effectiveSplits = effectiveSplits
    .filter(i => Number.isInteger(i) && i >= 0 && i < n - 1)
    .sort((a, b) => a - b);

  const segmentStarts = [0];
  const segmentEnds = [];
  effectiveSplits.forEach(idx => { segmentEnds.push(idx); segmentStarts.push(idx + 1); });
  segmentEnds.push(n - 1);

  const cl = new Array(n).fill(NaN);
  const uclArr = new Array(n).fill(NaN);
  const lclArr = new Array(n).fill(0); // practical convention
  const beyond = new Array(n).fill(false);

  const qHigh = 0.99865;

  for (let s = 0; s < segmentStarts.length; s++) {
    const start = segmentStarts[s];
    const end = segmentEnds[s];
    const seg = deltas.slice(start, end + 1);

    const segBaselineCountUsed =
      (s === 0 && baselineCount && baselineCount >= 1)
        ? Math.min(baselineCount, seg.length)
        : seg.length;

    const base = seg.slice(0, segBaselineCountUsed);
    const tbar = base.reduce((a, b) => a + b, 0) / base.length;

    const ucl = -tbar * Math.log(1 - qHigh);

    for (let i = start; i <= end; i++) {
      cl[i] = tbar;
      uclArr[i] = ucl;
      beyond[i] = isFinite(deltas[i]) && deltas[i] > ucl;
    }
  }

    const flagOnChart =
    (typeof shouldFlagSpecialCauseOnChart === "function")
      ? shouldFlagSpecialCauseOnChart()
      : true;

  const analysisForColour = analyzeRareChart({
    chartType: "t",
    labels: tLabels,
    values: deltas,
    cl,
    ucl: uclArr,
    lcl: lclArr
  });

  const pointColours = deltas.map((v, i) => {
    if (!flagOnChart) return SPC_STYLE.seriesBlue;
    if (analysisForColour.flags?.beyond?.[i]) return SPC_STYLE.pointBeyond;
    if (analysisForColour.flags?.special?.[i]) return SPC_STYLE.pointSpecial;
    return SPC_STYLE.seriesBlue;
  });

  drawSimpleSPCChart({
    labels: tLabels,
    values: deltas,
    pointColours,
    cl,
    ucl: uclArr,
    lcl: lclArr,
    yAxisSuggestedMin: 0,
    yAxisSuggestedMax: Math.max(...deltas, ...uclArr.filter(isFinite)),
    chartTitleFallback: "T chart",
    yAxisLabelFallback: "Time between events (days)",
    showUCL: true,
    showLCL: false
  });

    // Build per-period analyses (XmR-style summary, respects splits/baseline)
  const segmentAnalyses = [];

  for (let s = 0; s < segmentStarts.length; s++) {
    const start = segmentStarts[s];
    const end = segmentEnds[s];

    const segBaselineCountUsed =
      (s === 0 && baselineCount && baselineCount >= 1)
        ? Math.min(baselineCount, (end - start + 1))
        : (end - start + 1);

    const a = analyzeRareChart({
      chartType: "t",
      labels: tLabels.slice(start, end + 1),
      values: deltas.slice(start, end + 1),
      cl: cl.slice(start, end + 1),
      ucl: uclArr.slice(start, end + 1),
      lcl: lclArr.slice(start, end + 1)
    });

    // Add metadata for summary formatting (matches your other “multi” summaries)
    a.periodIndex = s + 1;
    a.periodCount = segmentStarts.length;
    a.startIndex = start;
    a.endIndex = end;
    a.labelStart = tLabels[start];
    a.labelEnd = tLabels[end];
    a.nPoints = (end - start + 1);
    a.baselineCountUsed = segBaselineCountUsed;

    // Add “stats” like other chart summaries expect
    // For T: CL is mean gap, UCL from the segment’s constant uclArr
    const segCL = cl[start];
    const segUCL = uclArr[start];
    a.stats = {
      cl: Number(segCL),
      ucl: Number(segUCL),
      lcl: 0
    };

    a.totalPoints = deltas.length;

    segmentAnalyses.push(a);
  }

  // Keep helper behaviour the same: store last period in lastRareAnalysis
  lastRareAnalysis = segmentAnalyses[segmentAnalyses.length - 1];

  // Render new style summary (multi-period)
  renderRareChartSummary(segmentAnalyses, deltas.length);
}


// -----------------------------
// G chart: opportunities between events (Geometric limits via percentiles)
// -----------------------------
function drawGChart(values, baselineCount, labels) {
  if (!Array.isArray(values) || values.length < 4) {
    showError("G chart needs at least 4 points (each value should be 1 or more).");
    return;
  }

  const gVals = values.map(v => Number(v)).filter(v => isFinite(v) && v >= 1);
  if (gVals.length !== values.length) {
    showError("G chart values must be numbers 1 or greater.");
    return;
  }

  const n = gVals.length;

  // ---- Segment definition from splits ----
  let effectiveSplits = Array.isArray(splits) ? splits.slice() : [];
  effectiveSplits = effectiveSplits
    .filter(i => Number.isInteger(i) && i >= 0 && i < n - 1)
    .sort((a, b) => a - b);

  const segmentStarts = [0];
  const segmentEnds = [];
  effectiveSplits.forEach(idx => { segmentEnds.push(idx); segmentStarts.push(idx + 1); });
  segmentEnds.push(n - 1);

  const cl = new Array(n).fill(NaN);
  const uclArr = new Array(n).fill(NaN);
  const lclArr = new Array(n).fill(1); // practical convention for G
  const beyond = new Array(n).fill(false);

  const qLow = 0.00135;
  const qHigh = 0.99865;

  // Quantile for geometric distribution: k = ln(1-q)/ln(1-p)
  function geomQuantile(q, p) {
    const k = Math.log(1 - q) / Math.log(1 - p);
    return Math.max(1, Math.ceil(k));
  }

  for (let s = 0; s < segmentStarts.length; s++) {
    const start = segmentStarts[s];
    const end = segmentEnds[s];

    const seg = gVals.slice(start, end + 1);

    const segBaselineCountUsed =
      (s === 0 && baselineCount && baselineCount >= 1)
        ? Math.min(baselineCount, seg.length)
        : seg.length;

    const base = seg.slice(0, segBaselineCountUsed);
    const gbar = base.reduce((a, b) => a + b, 0) / base.length;

    // Geometric mean ≈ 1/p
    const p = gbar > 0 ? (1 / gbar) : NaN;
    if (!isFinite(p) || p <= 0 || p >= 1) {
      showError("Could not compute G chart probability from your data (check values are >= 1).");
      return;
    }

    const ucl = geomQuantile(qHigh, p);
    const lcl = geomQuantile(qLow, p);

    for (let i = start; i <= end; i++) {
      cl[i] = gbar;
      uclArr[i] = ucl;
      lclArr[i] = lcl;
      beyond[i] = isFinite(gVals[i]) && (gVals[i] > ucl || gVals[i] < lcl);
    }
  }

  const flagOnChart =
    (typeof shouldFlagSpecialCauseOnChart === "function")
      ? shouldFlagSpecialCauseOnChart()
      : true;

  const analysisForColour = analyzeRareChart({
    chartType: "g",
    labels,
    values: gVals,
    cl,
    ucl: uclArr,
    lcl: lclArr
  });

  const pointColours = gVals.map((v, i) => {
    if (!flagOnChart) return SPC_STYLE.seriesBlue;
    if (analysisForColour.flags?.beyond?.[i]) return SPC_STYLE.pointBeyond;
    if (analysisForColour.flags?.special?.[i]) return SPC_STYLE.pointSpecial;
    return SPC_STYLE.seriesBlue;
  });

  drawSimpleSPCChart({
    labels,
    values: gVals,
    pointColours,
    cl,
    ucl: uclArr,
    lcl: lclArr,
    yAxisSuggestedMin: 1,
    yAxisSuggestedMax: Math.max(...gVals, ...uclArr.filter(isFinite)),
    chartTitleFallback: "G chart",
    yAxisLabelFallback: "Opportunities between events",
    showUCL: true,
    showLCL: true
  });

    // Build per-period analyses (XmR-style summary, respects splits/baseline)
  const segmentAnalyses = [];

  for (let s = 0; s < segmentStarts.length; s++) {
    const start = segmentStarts[s];
    const end = segmentEnds[s];

    const segBaselineCountUsed =
      (s === 0 && baselineCount && baselineCount >= 1)
        ? Math.min(baselineCount, (end - start + 1))
        : (end - start + 1);

    const a = analyzeRareChart({
      chartType: "g",
      labels: labels.slice(start, end + 1),
      values: gVals.slice(start, end + 1),
      cl: cl.slice(start, end + 1),
      ucl: uclArr.slice(start, end + 1),
      lcl: lclArr.slice(start, end + 1)
    });

    // Add metadata for summary formatting
    a.periodIndex = s + 1;
    a.periodCount = segmentStarts.length;
    a.startIndex = start;
    a.endIndex = end;
    a.labelStart = labels[start];
    a.labelEnd = labels[end];
    a.nPoints = (end - start + 1);
    a.baselineCountUsed = segBaselineCountUsed;

    // Add “stats” like other chart summaries expect
    const segCL = cl[start];
    const segUCL = uclArr[start];
    const segLCL = lclArr[start];
    a.stats = {
      cl: Number(segCL),
      ucl: Number(segUCL),
      lcl: Number(segLCL)
    };

    a.totalPoints = gVals.length;

    segmentAnalyses.push(a);
  }

  // Keep helper behaviour the same: store last period in lastRareAnalysis
  lastRareAnalysis = segmentAnalyses[segmentAnalyses.length - 1];

  // Render new style summary (multi-period)
  renderRareChartSummary(segmentAnalyses, gVals.length);
}

function parseOptionalNumber(value) {
  const s = String(value ?? "").trim();
  if (s === "") return undefined;
  const n = Number(s);
  return Number.isFinite(n) ? n : undefined;
}

function getYAxisTickStep() {
  const raw = parseOptionalNumber(yAxisTickStepInput?.value);

  if (!Number.isFinite(raw) || raw <= 0) {
    return undefined;
  }

  const format = yAxisFormatInput?.value || "auto";

  // Percentage data is stored internally as proportions.
  // User enters percentage points: 1 = 1%, 0.5 = 0.5%.
  if (format === "percent") {
    return raw / 100;
  }

  return raw;
}

function getAxisSettings() {
  return {
    x: {
      font: {
        family: (xAxisFontFamilyInput?.value || "").trim(),
        size: parseOptionalNumber(xAxisFontSizeInput?.value),
        style: isPressed(xAxisItalicBtn) ? "italic" : "normal",
        weight: isPressed(xAxisBoldBtn) ? "bold" : "normal"
      }
    },
    y: {
  min: parseOptionalNumber(yAxisMinInput?.value),
  max: parseOptionalNumber(yAxisMaxInput?.value),
  format: (yAxisFormatInput?.value || "auto"),
  stepSize: getYAxisTickStep(),
  font: {
        family: (yAxisFontFamilyInput?.value || "").trim(),
        size: parseOptionalNumber(yAxisFontSizeInput?.value),
        style: isPressed(yAxisItalicBtn) ? "italic" : "normal",
        weight: isPressed(yAxisBoldBtn) ? "bold" : "normal"
      }
    }
  };
}

// Draw an underline beneath the main chart title when requested.
// Chart.js supports bold/italic natively, but not underline.
const chartTitleUnderlinePlugin = {
  id: "chartTitleUnderline",

  afterDraw(chart) {
    if (!chart || !chart.ctx) return;

    const settings = getChartTitleSettings();

    if (!settings.show || !settings.underline) return;

    const titleOptions = chart.options?.plugins?.title;
    const titleBlock = chart.titleBlock;

    if (!titleOptions?.display || !titleBlock) return;

    const text = titleOptions.text;

    // Keep underline support simple and safe for normal one-line titles.
    if (typeof text !== "string" || !text.trim()) return;

    const ctx = chart.ctx;
    const fontOptions = titleOptions.font || {};

    ctx.save();

    // Resolve the same font that Chart.js uses for the title.
    if (typeof Chart !== "undefined" && Chart.helpers?.toFont) {
      const resolvedFont = Chart.helpers.toFont(fontOptions);
      ctx.font = resolvedFont.string;
    } else {
      const size = Number(fontOptions.size) || 16;
      const family = fontOptions.family || "sans-serif";
      const style = fontOptions.style || "normal";
      const weight = fontOptions.weight || "normal";
      ctx.font = `${style} ${weight} ${size}px ${family}`;
    }

    const textWidth = ctx.measureText(text).width;

    const centreX = (titleBlock.left + titleBlock.right) / 2;
    const underlineY = titleBlock.bottom - 3;

    ctx.beginPath();
    ctx.moveTo(centreX - textWidth / 2, underlineY);
    ctx.lineTo(centreX + textWidth / 2, underlineY);

    ctx.lineWidth = 1;
    ctx.strokeStyle = titleOptions.color || "#666";
    ctx.stroke();

    ctx.restore();
  }
};

if (typeof Chart !== "undefined" && Chart.register) {
  Chart.register(chartTitleUnderlinePlugin);
}


function positionHelpTooltip(helpIcon) {
  if (!helpIcon) return;

  const tooltip = helpIcon.querySelector(".help-tooltip-text");
  if (!tooltip) return;

  tooltip.classList.add("is-visible");

  // Temporarily place it so we can measure it
  tooltip.style.left = "0px";
  tooltip.style.top = "0px";

  const iconRect = helpIcon.getBoundingClientRect();
  const tipRect = tooltip.getBoundingClientRect();

  const gap = 8;
  const edgePadding = 10;

  // Centre tooltip horizontally on the ? icon
  let left =
    iconRect.left +
    iconRect.width / 2 -
    tipRect.width / 2;

  // Keep the tooltip inside the browser window
  left = Math.max(
    edgePadding,
    Math.min(left, window.innerWidth - tipRect.width - edgePadding)
  );

  // Prefer above the icon
  let top = iconRect.top - tipRect.height - gap;

  // If there isn't enough room above, put it underneath
  if (top < edgePadding) {
    top = iconRect.bottom + gap;
  }

  tooltip.style.left = `${left}px`;
  tooltip.style.top = `${top}px`;
}

function hideHelpTooltip(helpIcon) {
  const tooltip = helpIcon?.querySelector(".help-tooltip-text");

  if (tooltip) {
    tooltip.classList.remove("is-visible");
  }
}

document.querySelectorAll(".help-tooltip").forEach(helpIcon => {
  helpIcon.addEventListener("mouseenter", () => {
    positionHelpTooltip(helpIcon);
  });

  helpIcon.addEventListener("mouseleave", () => {
    hideHelpTooltip(helpIcon);
  });

  helpIcon.addEventListener("focus", () => {
    positionHelpTooltip(helpIcon);
  });

  helpIcon.addEventListener("blur", () => {
    hideHelpTooltip(helpIcon);
  });
});



function cleanFontOptions(font) {
  const out = {};
  if (font?.family) out.family = font.family;
  if (Number.isFinite(font?.size)) out.size = font.size;
  if (font?.style) out.style = font.style;
  if (font?.weight) out.weight = font.weight;
  return out;
}

function getAutoDecimalPlaces(value) {
  const n = Math.abs(Number(value));

  if (!Number.isFinite(n)) return 2;
  if (n > 100) return 0;
  if (n >= 10) return 1;
  if (n >= 1) return 2;
  return 3;
}

function getConfiguredDecimalPlaces(value) {
  const choice = yAxisDecimalsInput?.value || "auto";

  if (choice === "auto") {
    return getAutoDecimalPlaces(value);
  }

  const dp = Number(choice);
  return Number.isFinite(dp) ? dp : getAutoDecimalPlaces(value);
}

function formatSPCNumber(value, fallback = "—") {
  const n = Number(value);
  if (!Number.isFinite(n)) return fallback;

  return n.toFixed(getConfiguredDecimalPlaces(n));
}

function buildTickFormatter(format) {
  return function(value) {
    const n = Number(value);
    if (!Number.isFinite(n)) return value;

    const dp = getConfiguredDecimalPlaces(n);

    if (!format || format === "auto") return n.toFixed(dp);
    if (format === "integer") return String(Math.round(n));
    if (format === "decimal") return n.toFixed(dp);
    if (format === "percent") return `${(n * 100).toFixed(dp)}%`;

    return n.toFixed(dp);
  };
}


function withoutAxisBounds(settings) {
  if (!settings) return settings;
  const copy = { ...settings };
  delete copy.min;
  delete copy.max;
  return copy;
}

function buildAxisConfig(axisLabel, settings, extra = {}) {
  const cfg = {
    grid: { display: false },
    title: {
      display: !!axisLabel,
      text: axisLabel,
      font: cleanFontOptions(settings?.font)
    },
    ticks: {
      font: cleanFontOptions(settings?.font),
      callback: buildTickFormatter(settings?.format)
    },
    ...extra
  };

// Integer axes should use whole-number tick positions.
// This prevents fractional ticks being rounded into duplicate labels
// such as 1, 1, 1, 2, 2.
// Manual tick interval takes priority.
if (Number.isFinite(settings?.stepSize) && settings.stepSize > 0) {
  cfg.ticks.stepSize = settings.stepSize;
} else if (settings?.format === "integer") {
  // Otherwise let Chart.js choose sensible whole-number spacing.
  cfg.ticks.precision = 0;
}

  if (settings && Number.isFinite(settings.min)) {
    cfg.min = settings.min;
    delete cfg.suggestedMin;
  }

  if (settings && Number.isFinite(settings.max)) {
    cfg.max = settings.max;
    delete cfg.suggestedMax;
  }

  return cfg;
}

function buildCategoryXAxisConfig(axisLabel, settings, labels, extra = {}) {
  const cfg = buildAxisConfig(axisLabel, settings, {
    type: "category",
    ...extra
  });

  cfg.ticks = {
    ...(cfg.ticks || {}),
    callback: function(value, index) {
      if (Array.isArray(labels) && index >= 0 && index < labels.length) {
        return labels[index];
      }
      if (typeof this.getLabelForValue === "function") {
        return this.getLabelForValue(value);
      }
      return value;
    }
  };

  return cfg;
}


function validateAxisSettings() {
  const s = getAxisSettings();

  if (Number.isFinite(s.y.min) && Number.isFinite(s.y.max) && s.y.min > s.y.max) {
    if (typeof showError === "function") {
      showError("Y-axis minimum cannot be greater than Y-axis maximum.");
    } else if (errorMessage) {
      errorMessage.textContent = "Y-axis minimum cannot be greater than Y-axis maximum.";
    }
    return false;
  }

  return true;
}
	

// Get title / axis labels with fallbacks
function getChartLabels(defaultTitle, defaultX, defaultY) {
  if (chartTitleInput && !chartTitleManuallyEdited) {
    chartTitleInput.value = defaultTitle || "";
  }

  if (xAxisLabelInput && !xAxisLabelManuallyEdited) {
    xAxisLabelInput.value = defaultX || "";
  }

  if (yAxisLabelInput && !yAxisLabelManuallyEdited) {
    yAxisLabelInput.value = defaultY || "";
  }

  const title = chartTitleInput ? chartTitleInput.value.trim() : (defaultTitle || "");
  const xLabel = xAxisLabelInput ? xAxisLabelInput.value.trim() : (defaultX || "");
  const yLabel = yAxisLabelInput ? yAxisLabelInput.value.trim() : (defaultY || "");

  return { title, xLabel, yLabel };
}


function populateAnnotationDateOptions(labels) {
  if (!annotationDateInput) return;

  // Clear existing options
  annotationDateInput.innerHTML = "";

  // Placeholder option
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "Select date…";
  annotationDateInput.appendChild(placeholder);

  // Add one option per label (these are your x-axis dates like "2024-06-01")
  labels.forEach((lbl) => {
    const opt = document.createElement("option");
    opt.value = lbl;
    opt.textContent = lbl;
    annotationDateInput.appendChild(opt);
  });

  // Reset selection
  annotationDateInput.value = "";
}


function getAxisType() {
  const radios = document.querySelectorAll("input[name='axisType']");
  for (const r of radios) {
    if (r.checked) return r.value;
  }
  return "date"; // sensible default
}

function setAxisType(type) {
  const radios = document.querySelectorAll("input[name='axisType']");
  for (const r of radios) {
    r.checked = (r.value === type);
  }
}

function getAnnotationsAtDate(xVal) {
  if (!Array.isArray(annotations)) return [];
  return annotations
    .map((a, idx) => ({ ...a, _idx: idx }))
    .filter(a => a.date === xVal);
}

function editAnnotationAtIndex(idx) {
  if (!Array.isArray(annotations)) return false;
  if (!Number.isInteger(idx) || idx < 0 || idx >= annotations.length) return false;

  const current = annotations[idx];
  const nextText = prompt(`Edit annotation for ${current.date}:`, current.label);
  if (nextText === null) return false;

  const trimmed = String(nextText).trim();
  if (!trimmed) {
    alert("Annotation text cannot be blank.");
    return false;
  }

  annotations[idx].label = trimmed;
  return true;
}

function deleteAnnotationAtIndex(idx) {
  if (!Array.isArray(annotations)) return false;
  if (!Number.isInteger(idx) || idx < 0 || idx >= annotations.length) return false;

  annotations.splice(idx, 1);
  return true;
}

function wrapAnnotationText(text, maxCharsPerLine = 28) {
  const words = String(text || "").trim().split(/\s+/).filter(Boolean);
  if (!words.length) return [""];

  const lines = [];
  let current = words[0];

  for (let i = 1; i < words.length; i++) {
    const next = words[i];
    if ((current + " " + next).length <= maxCharsPerLine) {
      current += " " + next;
    } else {
      lines.push(current);
      current = next;
    }
  }

  lines.push(current);
  return lines;
}

function getNearestAnnotationAtDate(xVal) {
  const existing = getAnnotationsAtDate(xVal);
  return existing.length ? existing[0] : null;
}

function chooseAnnotationAtDate(xVal, mode) {
  const existing = getAnnotationsAtDate(xVal);
  if (!existing.length) return null;

  if (existing.length === 1) return existing[0];

  const numbered = existing
    .map((a, i) => `${i + 1}. ${a.label}`)
    .join("\n");

  const answer = prompt(
    `${mode === "edit" ? "Edit which annotation?" : "Delete which annotation?"}\n\n` +
    `Annotations at ${xVal}:\n\n${numbered}\n\n` +
    `Type a number from 1 to ${existing.length}:`,
    "1"
  );

  if (answer === null) return null;

  const n = Number(answer);
  if (!Number.isInteger(n) || n < 1 || n > existing.length) {
    alert("Please enter a valid number.");
    return null;
  }

  return existing[n - 1];
}

function buildAnnotationConfig(labels) {
  if (!annotations || annotations.length === 0) {
    return {};
  }

  const cfg = {};

  const items = annotations
    .map((a, idx) => ({
      ...a,
      _idx: idx,
      xIndex: Array.isArray(labels) ? labels.indexOf(a.date) : -1
    }))
    .filter(a => a.xIndex >= 0)
    .sort((a, b) => a.xIndex - b.xIndex);

  const laneLastEnd = [];

  items.forEach((a) => {
    const wrapped = wrapAnnotationText(a.label, 24);
    const longest = wrapped.reduce((m, l) => Math.max(m, l.length), 0);

    const span = Math.max(1, Math.ceil(longest / 7));

    let lane = 0;
    while (laneLastEnd[lane] !== undefined && a.xIndex <= laneLastEnd[lane]) {
      lane++;
    }
    laneLastEnd[lane] = a.xIndex + span;

    const level = Math.floor(lane / 2);
    const above = lane % 2 === 0;

    // Small offsets only, unless the user has dragged the label
	const defaultYAdjust = above
	  ? -(4 + level * 10)
	  : (4 + level * 10);

	const yAdjust = Number.isFinite(a.yAdjust)
	  ? a.yAdjust
	  : defaultYAdjust;

    cfg["annot" + a._idx] = {
      type: "line",
      xMin: a.date,
      xMax: a.date,
      borderColor: "rgba(120,120,120,0.30)",
      borderWidth: 0.6,
      borderDash: [5,4],
      label: {
        display: true,
        content: wrapped,
        backgroundColor: "rgba(255,255,255,0.96)",
        color: "#000000",
        borderColor: "#888888",
        borderWidth: 0.5,
        padding: 6,
        cornerRadius: 6,
        font: {
          size: 10,
          weight: "bold"
        },

        // Keep labels near the ends of the annotation line
        position: above ? "end" : "start",

        yAdjust: yAdjust,
        textAlign: "left"
      }
    };
  });

  return cfg;
}

const draggableAnnotationLabelsPlugin = {
  id: "draggableAnnotationLabels",

  afterInit(chart) {
       const canvas = chart.canvas;
    if (!canvas) return;

    let drag = null;

    
    function getHitAnnotation(event) {
      // Pareto charts do not support draggable user annotations.
      // Their 80% reference line must remain untouched.
      if (chart.$disableUserAnnotations) {
        return null;
      }

      const rect = canvas.getBoundingClientRect();
      const mouseX = event.clientX - rect.left;


      const labels = chart.data.labels || [];
      const xScale = chart.scales.x;
      if (!xScale) return null;

      for (let i = annotations.length - 1; i >= 0; i--) {
        const a = annotations[i];
        const x = xScale.getPixelForValue(a.date);

        // Big forgiving hit area around the dotted vertical line
        if (Math.abs(mouseX - x) <= 18) {
          return i;
        }
      }

      return null;
    }

    function redraw() {
      chart.options.plugins.annotation.annotations = buildAnnotationConfig(chart.data.labels);
      chart.update("none");
    }

    canvas.addEventListener("pointerdown", (event) => {
      if (!Array.isArray(annotations)) return;

      const hitIndex = getHitAnnotation(event);
      if (hitIndex === null) return;

      event.preventDefault();

      drag = {
        index: hitIndex,
        startY: event.clientY,
        startAdjust: Number.isFinite(annotations[hitIndex].yAdjust)
          ? annotations[hitIndex].yAdjust
          : 0
      };

      canvas.setPointerCapture(event.pointerId);
      canvas.style.cursor = "ns-resize";
    });

    canvas.addEventListener("pointermove", (event) => {
      if (!Array.isArray(annotations)) return;

      if (drag) {
        const deltaY = event.clientY - drag.startY;

        const chartTopLimit = chart.chartArea.top +10;
const chartBottomLimit = chart.chartArea.bottom -10;

const proposedAdjust = drag.startAdjust + deltaY;

const annotationConfig = buildAnnotationConfig(chart.data.labels);
const thisConfig = annotationConfig["annot" + drag.index];
const position = thisConfig?.label?.position || "end";

let proposedY = position === "start"
  ? chart.chartArea.bottom + proposedAdjust
  : chart.chartArea.top + proposedAdjust;

if (proposedY < chartTopLimit) {
  proposedY = chartTopLimit;
}

if (proposedY > chartBottomLimit) {
  proposedY = chartBottomLimit;
}

annotations[drag.index].yAdjust = position === "start"
  ? proposedY - chart.chartArea.bottom
  : proposedY - chart.chartArea.top;

        redraw();
        canvas.style.cursor = "ns-resize";
        return;
      }

      canvas.style.cursor = getHitAnnotation(event) !== null ? "ns-resize" : "";
    });

    canvas.addEventListener("pointerup", () => {
      drag = null;
      canvas.style.cursor = "";
    });

    canvas.addEventListener("pointercancel", () => {
      drag = null;
      canvas.style.cursor = "";
    });

    canvas.addEventListener("pointerleave", () => {
      if (!drag) canvas.style.cursor = "";
    });
  }
};

if (typeof Chart !== "undefined" && Chart.register) {
  Chart.register(draggableAnnotationLabelsPlugin);
}


function updateDataEditorWorkbookUi() {
  if (!dataEditorWorkbookBar || !dataEditorSheetSelect || !dataEditorWorkbookStatus) return;

  const isExcelMode =
    dataEditorSourceMode === "excel" &&
    dataEditorWorkbook &&
    Array.isArray(dataEditorWorkbookSheetNames) &&
    dataEditorWorkbookSheetNames.length > 0;

  dataEditorWorkbookBar.style.display = isExcelMode ? "block" : "none";

  const statusText = document.getElementById("dataEditorWorkbookStatusText");

  if (!isExcelMode) {
    dataEditorSheetSelect.innerHTML = "";
    if (statusText) statusText.innerHTML = "";
    return;
  }

  dataEditorSheetSelect.innerHTML = "";
  dataEditorWorkbookSheetNames.forEach(name => {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    dataEditorSheetSelect.appendChild(opt);
  });

  if (dataEditorCurrentSheetName && dataEditorWorkbookSheetNames.includes(dataEditorCurrentSheetName)) {
    dataEditorSheetSelect.value = dataEditorCurrentSheetName;
  }

  if (statusText) {
    statusText.innerHTML =
      `Workbook loaded. You can switch worksheet here, trim rows/columns in the grid, and then click <strong>Apply data</strong>.`;
  }
}

function worksheetToEditorGrid(worksheet) {
  if (!worksheet) {
    return {
      headers: ["Column1", "Column2"],
      data: [["", ""]]
    };
  }

  const rows2D = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    raw: false,
    defval: "",
    blankrows: false
  });

  const maxCols = rows2D.reduce((m, r) => Math.max(m, Array.isArray(r) ? r.length : 0), 0);
  const safeColCount = Math.max(2, maxCols || 0);

  const headers = Array.from({ length: safeColCount }, (_, i) => `Column${i + 1}`);

  const data = rows2D.length
    ? rows2D.map(row => {
        const arr = Array.isArray(row) ? row.slice() : [];
        while (arr.length < safeColCount) arr.push("");
        return arr.map(v => normalizeWorkbookCellValue(v));
      })
    : [Array.from({ length: safeColCount }, () => "")];

  return { headers, data };
}

function renderDataEditorGrid(headers, data) {
  if (!dataEditorGridEl) return;

  gridHeaders = Array.isArray(headers) && headers.length ? headers.slice() : ["Column1", "Column2"];
  const headersKey = JSON.stringify(gridHeaders);

  const mustRebuild = !dataEditorGrid || headersKey !== lastGridHeadersKey;

  if (mustRebuild) {
    if (dataEditorGrid) {
      try { dataEditorGrid.destroy(); } catch (e) { console.warn("Grid destroy failed:", e); }
      dataEditorGrid = null;
    }

    dataEditorGridEl.innerHTML = "";

    dataEditorGrid = jspreadsheet(dataEditorGridEl, {
      data,
      columns: gridHeaders.map(h => ({ title: h, width: 180 })),
      minDimensions: [Math.max(gridHeaders.length, 10), Math.max(20, data.length + 10)],

      allowInsertRow: true,
      allowDeleteRow: true,
      allowInsertColumn: true,
      allowDeleteColumn: true,

      onpaste: function(instance, pasteData, startCol, startRow) {
        if (!pasteData || typeof pasteData !== "string") return;

        const rows = pasteData.split(/\r?\n/).filter(r => r.length > 0);
        const colCount = rows.reduce((m, r) => Math.max(m, r.split("\t").length), 0);
        const rowCount = rows.length;

        const currentCols = instance.options.columns.length;
        const currentRows = instance.getData().length;

        const neededCols = startCol + colCount;
        const neededRows = startRow + rowCount;

        if (neededCols > currentCols) {
          const addN = neededCols - currentCols;
          instance.insertColumn(addN, currentCols);
          for (let i = currentCols; i < neededCols; i++) {
            instance.setHeader(i, `Column${i + 1}`);
          }
        }

        if (neededRows > currentRows) {
          const addN = neededRows - currentRows;
          instance.insertRow(addN);
        }

        setTimeout(() => {
          if (dataEditorHasHeaders) dataEditorHasHeaders.checked = detectHeadersFromGrid();
          renderHeaderStatus();
        }, 0);
      }
    });

    lastGridHeadersKey = headersKey;
  } else if (dataEditorGrid && typeof dataEditorGrid.setData === "function") {
    dataEditorGrid.setData(data);

    if (dataEditorGrid.options && Array.isArray(dataEditorGrid.options.columns)) {
      gridHeaders.forEach((h, i) => {
        if (dataEditorGrid.options.columns[i]) {
          dataEditorGrid.options.columns[i].title = h;
        }
        if (typeof dataEditorGrid.setHeader === "function") {
          dataEditorGrid.setHeader(i, h);
        }
      });
    }
  }

  if (dataEditorHasHeaders) dataEditorHasHeaders.checked = detectHeadersFromGrid();
  renderHeaderStatus();
}

function loadWorkbookSheetIntoDataEditor(sheetName) {
  if (!dataEditorWorkbook || !sheetName) return;

  const worksheet = dataEditorWorkbook.Sheets[sheetName];
  const { headers, data } = worksheetToEditorGrid(worksheet);

  dataEditorCurrentSheetName = sheetName;
  updateDataEditorWorkbookUi();
  renderDataEditorGrid(headers, data);
}

function openExcelWorkbookInDataEditor(workbook, initialSheetName) {
  if (!dataEditorOverlay || !dataEditorGridEl || !workbook) return;

  dataEditorSourceMode = "excel";
  dataEditorWorkbook = workbook;
  dataEditorWorkbookSheetNames = Array.isArray(workbook.SheetNames) ? workbook.SheetNames.slice() : [];
  dataEditorCurrentSheetName =
    initialSheetName && dataEditorWorkbookSheetNames.includes(initialSheetName)
      ? initialSheetName
      : (dataEditorWorkbookSheetNames[0] || "");

  updateDataEditorWorkbookUi();
  loadWorkbookSheetIntoDataEditor(dataEditorCurrentSheetName);

  dataEditorOverlay.style.display = "flex";
}

function openDataEditor() {
  if (!dataEditorOverlay || !dataEditorGridEl) return;

  dataEditorSourceMode = "manual";
  dataEditorWorkbook = null;
  dataEditorWorkbookSheetNames = [];
  dataEditorCurrentSheetName = "";

  updateDataEditorWorkbookUi();

  const { headers, data } = objectsToSheet(rawRows);
  renderDataEditorGrid(headers, data);

  dataEditorOverlay.style.display = "flex";
}

function closeDataEditor() {
  if (dataEditorOverlay) {
    dataEditorOverlay.style.display = "none";
  }
}

if (openDataEditorButton) {
  openDataEditorButton.addEventListener("click", () => {
    openDataEditor();
  });
}

if (dataEditorCancelButton) {
  dataEditorCancelButton.addEventListener("click", () => {
    hideDataEditorDeleteHelp();

    dataEditorSourceMode = "manual";
    dataEditorWorkbook = null;
    dataEditorWorkbookSheetNames = [];
    dataEditorCurrentSheetName = "";
    updateDataEditorWorkbookUi();

    dataEditorOverlay.style.display = "none";
  });
}

function objectsToSheet(rawRows) {
  if (!rawRows || rawRows.length === 0) return { headers: ["Date", "Value"], data: [] };

  const headers = Object.keys(rawRows[0]);
  const data = rawRows.map(r => headers.map(h => r[h] ?? ""));
  return { headers, data };
}

function sheetToObjects(headers, data) {
  return data
    .filter(row => row.some(cell => String(cell ?? "").trim() !== "")) // drop blank rows
    .map(row => {
      const o = {};
      headers.forEach((h, i) => (o[h] = row[i]));
      return o;
    });
}


function rowsEqualNormalized(a, b) {
  if (!Array.isArray(a) || !Array.isArray(b)) return false;
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    const sa = String(a[i] ?? "").trim().toLowerCase();
    const sb = String(b[i] ?? "").trim().toLowerCase();
    if (sa !== sb) return false;
  }
  return true;
}


function rowDataLikenessScore(rowArr) {
  // Score = fraction of cells that look like a date OR a number
  if (!Array.isArray(rowArr) || rowArr.length === 0) return 0;

  let total = 0;
  let looksLikeData = 0;

  for (const cell of rowArr) {
    const s = String(cell ?? "").trim();
    if (!s) continue;
    total++;

    // numeric?
    const num = toNumericValue(s);
    const isNum = isFinite(num);

    // date?
    const d = parseDateValue(s);
    const isDate = isFinite(d.getTime());

    if (isNum || isDate) looksLikeData++;
  }

  return total === 0 ? 0 : looksLikeData / total;
}

function stripDuplicateHeaderRow(rows, headers) {
  // If first "data row" repeats the headers (common after accidental double header),
  // remove it.
  if (!rows || rows.length === 0) return rows;
  const first = rows[0];
  if (!first) return rows;

  const keys = headers || Object.keys(first);
  if (!keys || keys.length === 0) return rows;

  let matches = 0;
  let checked = 0;
  for (const k of keys) {
    const v = first[k];
    if (v === null || v === undefined) continue;
    checked++;
    if (String(v).trim().toLowerCase() === String(k).trim().toLowerCase()) matches++;
  }

  // If most columns match their own header text, treat it as a duplicate header row
  if (checked > 0 && matches / checked >= 0.7) {
    return rows.slice(1);
  }
  return rows;
}

if (dataEditorApplyButton) {
  dataEditorApplyButton.addEventListener("click", () => {
    let loadedOk = false;

    try {
      if (!dataEditorGrid) {
        showError("Spreadsheet editor not initialised. Try reopening the editor.");
        return;
      }

      const data2D = dataEditorGrid.getData();

      // Drop fully blank rows
      const nonBlank = data2D.filter(row =>
        row.some(cell => String(cell ?? "").trim() !== "")
      );

      if (nonBlank.length === 0) {
        showError("Please enter at least one data row.");
        return;
      }

      // Decide if first row is headers (checkbox)
      const useHeaders = !!(dataEditorHasHeaders && dataEditorHasHeaders.checked);

// Safety: warn if checkbox disagrees with auto-detect
const autoGuess = detectHeadersFromGrid();
if (useHeaders !== autoGuess) {
  const msg = useHeaders
    ? "You have 'First row contains headers' ticked, but the first row looks like DATA.\n\nApply anyway?"
    : "You have 'First row contains headers' unticked, but the first row looks like HEADERS.\n\nApply anyway?";

  if (!confirm(msg)) {
    // Keep modal open; let them correct the checkbox
    renderHeaderStatus();
    return;
  }
}


      let headers;
      let body;

      if (useHeaders) {
        headers = nonBlank[0].map((h, i) => {
          const name = String(h ?? "").trim();
          return name || `Column${i + 1}`;
        });
        body = nonBlank.slice(1); // remove header row from data
      } else {
  // No header row: keep the existing column titles from the grid
  const maxCols = nonBlank.reduce((m, r) => Math.max(m, r.length), 0);

  const cols = (dataEditorGrid && dataEditorGrid.options && Array.isArray(dataEditorGrid.options.columns))
    ? dataEditorGrid.options.columns
    : [];

  headers = Array.from({ length: maxCols }, (_, i) => {
    const t = (cols[i] && cols[i].title) ? String(cols[i].title).trim() : "";
    return t || `Column${i + 1}`;
  });

  body = nonBlank;
}

      const rows = sheetToObjects(headers, body);
if (!rows || rows.length === 0) {
  showError("Paste at least one row of data.");
  return;
}

// Applying data from a newly opened Excel workbook means "new chart".
// Applying data from the normal data editor means "edit current chart".
const isNewExcelImport = dataEditorSourceMode === "excel";

if (isNewExcelImport) {
  if (!prepareForNewImportedChart(rows)) return;
}

if (!loadRows(rows)) return;
loadedOk = true;

            clearError();

      // Reset annotations/splits etc...
      annotations = [];
      if (annotationDateInput) annotationDateInput.value = "";
      if (annotationLabelInput) annotationLabelInput.value = "";
      splits = [];
      if (splitPointSelect) splitPointSelect.innerHTML = "";

      hideDataEditorDeleteHelp();
      dataEditorSourceMode = "manual";
      dataEditorWorkbook = null;
      dataEditorWorkbookSheetNames = [];
      dataEditorCurrentSheetName = "";
      updateDataEditorWorkbookUi();

      try { closeDataEditor(); } catch (uiErr) { console.warn("closeDataEditor failed:", uiErr); }

      const hint = document.getElementById("noDataYetHint");
      if (hint) hint.style.display = "none";

      if (isNewExcelImport) {
  // A genuinely new Excel dataset should pause at chart choice
  // rather than immediately implying that the default Run chart
  // is necessarily the right choice.
  showNewDataChartPrompt();
} else {
  // Editing the existing dataset should retain the current behaviour:
  // redraw the existing chart immediately.
  try {
    if (generateButton) generateButton.click();
  } catch (genErr) {
    console.warn("Auto-generate failed:", genErr);
  }
}

    } catch (e) {
      console.error(e);
      if (!loadedOk) showError("Unexpected error reading spreadsheet data.");
      else clearError();
    }
  });
}

function formatYMD_Local(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return "";
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}


function formatDateOnlyLabel(v) {
  if (v === null || v === undefined) return "";

  // If it's already a Date object
  if (v instanceof Date && !isNaN(v.getTime())) {
    return formatYMD_Local(v);
  }

  const s = String(v).trim();
  if (!s) return "";

  // Common cases where labels include time
  if (s.includes("T")) return s.split("T")[0];          // ISO: 2025-01-01T12:00...
  if (s.includes(",")) return s.split(",")[0].trim();   // Locale: "01/01/2025, 12:00"

  // If it’s "YYYY-MM-DD HH:mm..." or "DD/MM/YYYY HH:mm..."
  const firstToken = s.split(/\s+/)[0];
  if (/^\d{4}-\d{2}-\d{2}$/.test(firstToken)) return firstToken;
  if (/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/.test(firstToken)) return firstToken;

  // Fallback: try parsing and formatting
  const d = new Date(s);
  if (!isNaN(d.getTime())) return formatYMD_Local(d);

  return firstToken;
}


// ---- Summary helpers ----

let lastAttributeAnalysis = null;
let lastXbarSAnalysis = null;
let lastRareAnalysis = null;

function getRuleSettingsSafe() {
  const shift = parseInt(shiftRulePointsInput?.value || "8", 10);
  const trend = parseInt(trendRulePointsInput?.value || "6", 10);

  return {
    shiftLength: Number.isFinite(shift) && shift >= 3 ? shift : 8,
    trendLength: Number.isFinite(trend) && trend >= 3 ? trend : 6,

    enableAdvancedTrend: !!enableAdvancedTrendCheckbox?.checked,
    enableRareRunTrend: !!enableRareRunTrendCheckbox?.checked,

    ruleTwoOfThreeOuterThird: !!ruleTwoOfThreeOuterThirdCheckbox?.checked,
    ruleFourOfFiveOneSigma: !!ruleFourOfFiveOneSigmaCheckbox?.checked
  };
}

/* ============================================================
   RULE POLICY (single source of truth)
   Status values:
   - "on"       : always applied / default on
   - "optional" : available only in advanced mode, off by default
   - "warn"     : available only after explicit warning/confirm
   - "blocked"  : not available for this chart type
   ============================================================ */

const RULE_POLICY = {

  // Pareto is a prioritisation / drill-down display, not a control chart.
  pareto: {
    beyondLimits: "blocked",
    runShift: "blocked",
    trend: "blocked",
    zone23: "blocked",
    zone45: "blocked"
  },

  // Attribute charts: conservative defaults only
  c: {
    beyondLimits: "on",
    runShift: "on",
    trend: "blocked",
    zone23: "blocked",
    zone45: "blocked"
  },
  p: {
    beyondLimits: "on",
    runShift: "on",
    trend: "blocked",
    zone23: "blocked",
    zone45: "blocked"
  },
  pprime: {
 	 beyondLimits: "on",
 	 runShift: "on",
 	 trend: "blocked",
 	 zone23: "blocked",
 	 zone45: "blocked"
     },
  u: {
    beyondLimits: "on",
    runShift: "on",
    trend: "blocked",
    zone23: "blocked",
    zone45: "blocked"
  },

  uprime: {
  beyondLimits: "on",
  runShift: "on",
  trend: "blocked",
  zone23: "blocked",
  zone45: "blocked"
},

  // Rare-event charts: beyond limits by default; run/trend only behind warning
  t: {
    beyondLimits: "on",
    runShift: "warn",
    trend: "warn",
    zone23: "blocked",
    zone45: "blocked"
  },
  g: {
    beyondLimits: "on",
    runShift: "warn",
    trend: "warn",
    zone23: "blocked",
    zone45: "blocked"
  },

  // Individuals charts: conservative defaults, advanced pattern rules optional
  xmr: {
    beyondLimits: "on",
    runShift: "on",
    trend: "optional",
    zone23: "optional",
    zone45: "optional"
  },

  // Subgrouped continuous charts: richest defensible advanced rule set
  xbars: {
    beyondLimits: "on",
    runShift: "on",
    trend: "optional",
    zone23: "optional",
    zone45: "optional"
  },

  // Run chart: no control limits, run on by default, trend optional
  run: {
    beyondLimits: "blocked",
    runShift: "on",
    trend: "optional",
    zone23: "blocked",
    zone45: "blocked"
  }
};

function getRulePolicy(chartType) {
  return RULE_POLICY[chartType] || RULE_POLICY.run;
}

function isAdvancedContinuousChartType(chartType) {
  return chartType === "xmr" || chartType === "xbars" || chartType === "run";
}


function getEffectiveRuleSettingsForChart(chartType) {
  const raw = getRuleSettingsSafe();
  const policy = getRulePolicy(chartType);

  const rareAdvancedEnabled = isRareChartType(chartType) && !!raw.enableRareRunTrend;
  const advancedContinuousEnabled = isAdvancedContinuousChartType(chartType) && !!raw.enableAdvancedTrend;

  let allowRunShift = false;
  if (policy.runShift === "on") {
    allowRunShift = true;
  } else if (policy.runShift === "warn") {
    allowRunShift = rareAdvancedEnabled;
  } else {
    allowRunShift = false;
  }

  let allowTrend = false;
  if (policy.trend === "on") {
    allowTrend = true;
  } else if (policy.trend === "warn") {
    allowTrend = rareAdvancedEnabled;
  } else if (policy.trend === "optional") {
    allowTrend = advancedContinuousEnabled;
  } else {
    allowTrend = false;
  }

  let allowZone23 = false;
  if (policy.zone23 === "on") {
    allowZone23 = true;
  } else if (policy.zone23 === "optional") {
    allowZone23 = advancedContinuousEnabled && !!raw.ruleTwoOfThreeOuterThird;
  } else if (policy.zone23 === "warn") {
    allowZone23 = rareAdvancedEnabled && !!raw.ruleTwoOfThreeOuterThird;
  } else {
    allowZone23 = false;
  }

  let allowZone45 = false;
  if (policy.zone45 === "on") {
    allowZone45 = true;
  } else if (policy.zone45 === "optional") {
    allowZone45 = advancedContinuousEnabled && !!raw.ruleFourOfFiveOneSigma;
  } else if (policy.zone45 === "warn") {
    allowZone45 = rareAdvancedEnabled && !!raw.ruleFourOfFiveOneSigma;
  } else {
    allowZone45 = false;
  }

  return {
    ...raw,
    allowRunShift,
    allowTrend,
    allowZone23,
    allowZone45,
    warnRunShift: policy.runShift === "warn",
    warnTrend: policy.trend === "warn"
  };
}

function findShiftWindow(values, cl, shiftLength) {
  let run = 0;
  let side = 0;

  for (let i = 0; i < values.length; i++) {
    const v = values[i];

    const centre =
      Array.isArray(cl)
        ? cl[i]
        : cl;

    if (
      !isFinite(v) ||
      !isFinite(centre)
    ) {
      run = 0;
      side = 0;
      continue;
    }

    const s =
      v > centre
        ? 1
        : (
            v < centre
              ? -1
              : 0
          );

    if (s === 0) {
      run = 0;
      side = 0;
      continue;
    }

    if (s === side) {
      run += 1;
    } else {
      side = s;
      run = 1;
    }

    if (run >= shiftLength) {
      return {
        start:
          i -
          shiftLength +
          1,
        end: i,
        side
      };
    }
  }

  return null;
}


function findTrendWindow(values, trendLength) {
  let inc = 1;
  let dec = 1;

  for (let i = 1; i < values.length; i++) {
    const a = values[i - 1];
    const b = values[i];

    if (!isFinite(a) || !isFinite(b)) {
      inc = 1;
      dec = 1;
      continue;
    }

    if (b > a) {
      inc += 1;
      dec = 1;
    } else if (b < a) {
      dec += 1;
      inc = 1;
    } else {
      inc = 1;
      dec = 1;
    }

    if (inc >= trendLength) {
      return {
        start: i - trendLength + 1,
        end: i,
        direction: "up"
      };
    }

    if (dec >= trendLength) {
      return {
        start: i - trendLength + 1,
        end: i,
        direction: "down"
      };
    }
  }

  return null;
}

function findShiftSignals(values, cl, shiftLength) {
  let bestRun = 0;
  let currentRun = 0;
  let currentSide = 0; // -1 below, +1 above, 0 none

  for (const v of values) {
    if (!isFinite(v) || !isFinite(cl)) { currentRun = 0; currentSide = 0; continue; }
    const side = v > cl ? 1 : (v < cl ? -1 : 0);
    if (side === 0) { currentRun = 0; currentSide = 0; continue; }

    if (side === currentSide) currentRun += 1;
    else { currentSide = side; currentRun = 1; }

    bestRun = Math.max(bestRun, currentRun);
  }

  // Plain English wording (expert-safe)
  if (bestRun >= shiftLength) {
    return `Shift: ${shiftLength} or more points in a row on the same side of the centre line`;
  }
  return null;
}


function findTrendSignals(values, trendLength) {
  let inc = 1, dec = 1;
  let bestInc = 1, bestDec = 1;

  for (let i = 1; i < values.length; i++) {
    const a = values[i - 1], b = values[i];
    if (!isFinite(a) || !isFinite(b)) { inc = 1; dec = 1; continue; }

    if (b > a) { inc += 1; dec = 1; }
    else if (b < a) { dec += 1; inc = 1; }
    else { inc = 1; dec = 1; }

    bestInc = Math.max(bestInc, inc);
    bestDec = Math.max(bestDec, dec);
  }

  // Plain English wording (expert-safe)
  if (bestInc >= trendLength) return `Trend: ${trendLength} or more points in a row steadily increasing`;
  if (bestDec >= trendLength) return `Trend: ${trendLength} or more points in a row steadily decreasing`;
  return null;
}

function zoneThresholdsFromBand(cl, ucl, lcl) {
  if (!isFinite(cl) || !isFinite(ucl) || !isFinite(lcl)) {
    return {
      oneUp: NaN,
      oneDown: NaN,
      twoUp: NaN,
      twoDown: NaN
    };
  }

  const oneUp = cl + (ucl - cl) / 3;
  const twoUp = cl + 2 * (ucl - cl) / 3;
  const oneDown = cl - (cl - lcl) / 3;
  const twoDown = cl - 2 * (cl - lcl) / 3;

  return { oneUp, oneDown, twoUp, twoDown };
}

function detectKofNRule(values, upperThreshArr, lowerThreshArr, windowSize, minCount) {
  const n = values.length;
  const flags = new Array(n).fill(false);
  const hits = [];

  for (let start = 0; start <= n - windowSize; start++) {
    const end = start + windowSize - 1;

    let aboveIdx = [];
    for (let i = start; i <= end; i++) {
      const v = values[i];
      const thr = upperThreshArr[i];
      if (isFinite(v) && isFinite(thr) && v > thr) {
        aboveIdx.push(i);
      }
    }

    if (aboveIdx.length >= minCount) {
      aboveIdx.forEach(i => { flags[i] = true; });
      hits.push({ start, end, side: "above", indices: aboveIdx.slice() });
    }

    let belowIdx = [];
    for (let i = start; i <= end; i++) {
      const v = values[i];
      const thr = lowerThreshArr[i];
      if (isFinite(v) && isFinite(thr) && v < thr) {
        belowIdx.push(i);
      }
    }

    if (belowIdx.length >= minCount) {
      belowIdx.forEach(i => { flags[i] = true; });
      hits.push({ start, end, side: "below", indices: belowIdx.slice() });
    }
  }

  return { flags, hits };
}

function detectTwoOfThreeOuterThird(values, twoUpArr, twoDownArr) {
  return detectKofNRule(values, twoUpArr, twoDownArr, 3, 2);
}

function detectFourOfFiveOneSigma(values, oneUpArr, oneDownArr) {
  return detectKofNRule(values, oneUpArr, oneDownArr, 5, 4);
}

function analyzeLimits({ labels, values, cl, ucl, lcl }) {
  const out = [];
  for (let i = 0; i < values.length; i++) {
    const v = values[i];
    const u = Array.isArray(ucl) ? ucl[i] : ucl;
    const l = Array.isArray(lcl) ? lcl[i] : lcl;

    if (!isFinite(v)) continue;
    if (isFinite(u) && v > u) out.push({ i, label: labels[i], type: "aboveUCL", value: v, limit: u });
    if (isFinite(l) && v < l) out.push({ i, label: labels[i], type: "belowLCL", value: v, limit: l });
  }
  return out;
}

function analyzeAttributeChart({
  chartType,
  labels,
  values,
  cl,
  ucl,
  lcl,
  excludedIndices = []
}) {
  const rs =
    (
      typeof
        getEffectiveRuleSettingsForChart ===
        "function"
    )
      ? getEffectiveRuleSettingsForChart(
          chartType
        )
      : {
          shiftLength: 8,
          trendLength: 6,
          allowRunShift: true,
          allowTrend: false,
          allowZone23: false,
          allowZone45: false
        };

  const signals = [];

  const n =
    values.length;

  const excludedSet =
    new Set(
      Array.isArray(
        excludedIndices
      )
        ? excludedIndices.filter(
            Number.isInteger
          )
        : []
    );

  // NaN deliberately breaks run/trend/zone
  // sequences at an excluded point.
  const analysisValues =
    values.map(
      (
        value,
        index
      ) =>
        excludedSet.has(
          index
        )
          ? NaN
          : value
    );

  const clArr =
    Array.isArray(cl)
      ? cl
      : new Array(n)
          .fill(cl);

  const uclArr =
    Array.isArray(ucl)
      ? ucl
      : new Array(n)
          .fill(ucl);

  const lclArr =
    Array.isArray(lcl)
      ? lcl
      : new Array(n)
          .fill(lcl);

  const hasFiniteCentreLine =
    clArr.some(
      v =>
        Number.isFinite(v)
    );

  // 1) Beyond limits
  let outOfControl = [];

  if (
    ucl !== undefined &&
    lcl !== undefined
  ) {
    outOfControl =
      analyzeLimits({
        labels,

        values:
          analysisValues,

        cl:
          clArr,

        ucl:
          uclArr,

        lcl:
          lclArr
      });

    const hasAbove =
      outOfControl.some(
        o =>
          o.type ===
          "aboveUCL"
      );

    const hasBelow =
      outOfControl.some(
        o =>
          o.type ===
          "belowLCL"
      );

    if (hasAbove) {
      signals.push(
        "One or more points above the upper limit"
      );
    }

    if (hasBelow) {
      signals.push(
        "One or more points below the lower limit"
      );
    }
  }

  // 2) Shift / run
  let shiftWindow =
    null;

  if (
    rs.allowRunShift &&
    hasFiniteCentreLine
  ) {
    shiftWindow =
      findShiftWindow(
        analysisValues,
        clArr,
        rs.shiftLength
      );

    if (shiftWindow) {
      const sideText =
        shiftWindow.side > 0
          ? "above"
          : "below";

      const aLab =
        labels
          ?.[shiftWindow.start] ??
        `point ${
          shiftWindow.start +
          1
        }`;

      const bLab =
        labels
          ?.[shiftWindow.end] ??
        `point ${
          shiftWindow.end +
          1
        }`;

      signals.push(
        `Shift: ${rs.shiftLength}+ points in a row ${sideText} the centre line (from ${aLab} to ${bLab})`
      );
    }
  }

  // 3) Trend
  let trendWindow =
    null;

  if (rs.allowTrend) {
    trendWindow =
      findTrendWindow(
        analysisValues,
        rs.trendLength
      );

    if (trendWindow) {
      const dirText =
        trendWindow.direction ===
        "up"
          ? "increasing"
          : "decreasing";

      const aLab =
        labels
          ?.[trendWindow.start] ??
        `point ${
          trendWindow.start +
          1
        }`;

      const bLab =
        labels
          ?.[trendWindow.end] ??
        `point ${
          trendWindow.end +
          1
        }`;

      signals.push(
        `Trend: ${rs.trendLength}+ points steadily ${dirText} (from ${aLab} to ${bLab})`
      );
    }
  }

  // 4) Zone rules
  let zone23Flags =
    new Array(n)
      .fill(false);

  let zone45Flags =
    new Array(n)
      .fill(false);

  if (
    (
      rs.allowZone23 ||
      rs.allowZone45
    ) &&
    ucl !== undefined &&
    lcl !== undefined
  ) {
    const oneUp =
      new Array(n)
        .fill(NaN);

    const oneDown =
      new Array(n)
        .fill(NaN);

    const twoUp =
      new Array(n)
        .fill(NaN);

    const twoDown =
      new Array(n)
        .fill(NaN);

    for (
      let i = 0;
      i < n;
      i++
    ) {
      const z =
        zoneThresholdsFromBand(
          clArr[i],
          uclArr[i],
          lclArr[i]
        );

      oneUp[i] =
        z.oneUp;

      oneDown[i] =
        z.oneDown;

      twoUp[i] =
        z.twoUp;

      twoDown[i] =
        z.twoDown;
    }

    if (
      rs.allowZone23
    ) {
      const res23 =
        detectTwoOfThreeOuterThird(
          analysisValues,
          twoUp,
          twoDown
        );

      zone23Flags =
        res23.flags ||
        zone23Flags;

      if (
        res23.hits &&
        res23.hits.length
      ) {
        const h =
          res23.hits[0];

        const aLab =
          labels
            ?.[h.start] ??
          `point ${
            h.start + 1
          }`;

        const bLab =
          labels
            ?.[h.end] ??
          `point ${
            h.end + 1
          }`;

        signals.push(
          `Zone: 2 of 3 points in the outer third (from ${aLab} to ${bLab})`
        );
      }
    }

    if (
      rs.allowZone45
    ) {
      const res45 =
        detectFourOfFiveOneSigma(
          analysisValues,
          oneUp,
          oneDown
        );

      zone45Flags =
        res45.flags ||
        zone45Flags;

      if (
        res45.hits &&
        res45.hits.length
      ) {
        const h =
          res45.hits[0];

        const aLab =
          labels
            ?.[h.start] ??
          `point ${
            h.start + 1
          }`;

        const bLab =
          labels
            ?.[h.end] ??
          `point ${
            h.end + 1
          }`;

        signals.push(
          `Zone: 4 of 5 points beyond 1-sigma (from ${aLab} to ${bLab})`
        );
      }
    }
  }

  const beyondFlags =
    new Array(n)
      .fill(false);

  outOfControl.forEach(
    o => {
      if (
        typeof o.i ===
          "number" &&
        o.i >= 0 &&
        o.i < n
      ) {
        beyondFlags[o.i] =
          true;
      }
    }
  );

  const shiftFlags =
    new Array(n)
      .fill(false);

  if (shiftWindow) {
    for (
      let i =
        shiftWindow.start;
      i <=
        shiftWindow.end;
      i++
    ) {
      shiftFlags[i] =
        true;
    }
  }

  const trendFlags =
    new Array(n)
      .fill(false);

  if (trendWindow) {
    for (
      let i =
        trendWindow.start;
      i <=
        trendWindow.end;
      i++
    ) {
      trendFlags[i] =
        true;
    }
  }

  const specialFlags =
    new Array(n)
      .fill(false);

  for (
    let i = 0;
    i < n;
    i++
  ) {
    specialFlags[i] =
      beyondFlags[i] ||
      shiftFlags[i] ||
      trendFlags[i] ||
      zone23Flags[i] ||
      zone45Flags[i];
  }

  return {
    chartType,

    isStable:
      signals.length === 0,

    signals,
    outOfControl,

    shiftLength:
      rs.shiftLength,

    trendLength:
      rs.trendLength,

    rulePolicy: {
      allowRunShift:
        rs.allowRunShift,

      allowTrend:
        rs.allowTrend,

      allowZone23:
        rs.allowZone23,

      allowZone45:
        rs.allowZone45
    },

    flags: {
      beyond:
        beyondFlags,

      shift:
        shiftFlags,

      trend:
        trendFlags,

      zone23:
        zone23Flags,

      zone45:
        zone45Flags,

      special:
        specialFlags
    },

    firstOutOfControl:
      outOfControl.length
        ? outOfControl[0]
        : null
  };
}

function analyzeRareChart({ chartType, labels, values, cl, ucl, lcl }) {
  // Rare charts use the same engine, but the policy layer decides
  // whether run/trend are actually allowed.
  return analyzeAttributeChart({ chartType, labels, values, cl, ucl, lcl });
}

function renderAttributeMultiSummary(segmentAnalyses, totalPoints) {
  if (!summaryDiv) return;

  const nameMap = {
    c: "C chart",
    p: "P chart",
    pprime: "P′ chart",
    u: "U chart",
 uprime: "U′ chart"
  };

  const chartType =
    segmentAnalyses?.[0]?.chartType;

  const chartName =
    nameMap[chartType] || "Chart";

  const fmt = (v) =>
    Number.isFinite(v)
      ? (
          typeof formatNumber === "function"
            ? formatNumber(v, 3)
            : Number(v).toFixed(3)
        )
      : "—";

  const sym = (t) =>
  t === "c"
    ? "c̄"
    : (
        t === "p" ||
        t === "pprime"
      )
      ? "p̄"
      : (
          t === "u" ||
          t === "uprime"
        )
        ? "ū"
        : "CL";

  let html =
    `<h3>Summary (${chartName})</h3>`;

  html +=
    `<p>Total number of points: <strong>${totalPoints}</strong>. `;

  html +=
    `The chart is divided into <strong>${segmentAnalyses.length}</strong> ` +
    `period${segmentAnalyses.length !== 1 ? "s" : ""} `;

  html +=
    `(based on the baseline and any splits).</p>`;

  if (isUFamilyChartType(chartType)) {
    const multiplier =
      getURateMultiplier();

    html +=
      `<p>Rates are displayed per <strong>${formatURateMultiplier(multiplier)}</strong> ` +
      `${multiplier === 1 ? "opportunity" : "opportunities"}.</p>`;
  }

  segmentAnalyses.forEach((a, idx) => {
    html +=
      `<div class="pdf-avoid-break">`;

    html +=
      segmentAnalyses.length > 1
        ? `<h4>Period ${idx + 1}</h4>`
        : `<h4>Single period</h4>`;

    html += `<ul>`;

    // Coverage
    if (
      a.startIndex != null &&
      a.endIndex != null &&
      a.labelStart &&
      a.labelEnd
    ) {
      const n =
        a.endIndex -
        a.startIndex +
        1;

      html +=
        `<li><strong>Coverage:</strong> ` +
        `<strong>points ${a.startIndex + 1}–${a.endIndex + 1}</strong> ` +
        `(${a.labelStart} to ${a.labelEnd}) – ${n} points.</li>`;

    } else if (a.nPoints != null) {
      html +=
        `<li><strong>Coverage:</strong> ${a.nPoints} points.</li>`;
    }

        // Baseline
    if (
      a.baselineCountUsed !=
      null
    ) {
      const periodStart =
        Number.isInteger(
          a.startIndex
        )
          ? a.startIndex
          : 0;

      const excludedSet =
        new Set(
          currentChart
            ?.$limitExcludedIndices ||
          []
        );

      let excludedInBaseline =
        0;

      for (
        let i = 0;
        i <
          a.baselineCountUsed;
        i++
      ) {
        if (
          excludedSet.has(
            periodStart + i
          )
        ) {
          excludedInBaseline +=
            1;
        }
      }

      const usedInBaseline =
        Math.max(
          0,
          a.baselineCountUsed -
          excludedInBaseline
        );

      html +=
        `<li><strong>Baseline for this period:</strong> ` +
        `first ${a.baselineCountUsed} points define the baseline window; ` +
        `${usedInBaseline} used to calculate centre line and limits` +
        `${
          excludedInBaseline
            ? `, ${excludedInBaseline} excluded`
            : ""
        }.</li>`;
    }

    // Stats
    const st =
      a.stats || null;

    // Centre line + limits
    if (
      st &&
      Number.isFinite(st.cl)
    ) {
      if (a.chartType === "c") {
        html +=
          `<li><strong>Centre line (${sym(a.chartType)}):</strong> ` +
          `${fmt(st.cl)}; ` +
          `<strong>control limits:</strong> ` +
          `LCL = ${fmt(st.lcl)}, ` +
          `UCL = ${fmt(st.ucl)}.</li>`;

      } else if (
  a.chartType === "p" ||
  a.chartType === "pprime" ||
  a.chartType === "u" ||
  a.chartType === "uprime"
) {
        html +=
          `<li><strong>Centre line (${sym(a.chartType)}):</strong> ${fmt(st.cl)}; ` +
          `<strong>control limits:</strong> ` +
          `LCL = ${fmt(st.lclAvg)} ` +
          `(range ${fmt(st.lclMin)}–${fmt(st.lclMax)}), ` +
          `UCL = ${fmt(st.uclAvg)} ` +
          `(range ${fmt(st.uclMin)}–${fmt(st.uclMax)}).</li>`;

      } else {
        html +=
          `<li><strong>Centre line:</strong> ${fmt(st.cl)}</li>`;
      }

    } else {
      const fallbackCL =
        (
          typeof a.cl === "number" &&
          Number.isFinite(a.cl)
        )
          ? a.cl
          : (
              typeof a.centerLine === "number" &&
              Number.isFinite(a.centerLine)
            )
            ? a.centerLine
            : null;

      if (fallbackCL != null) {
        html +=
          `<li><strong>Centre line:</strong> ${fmt(fallbackCL)}</li>`;
      }
    }

  if (
  a.chartType === "u" &&
  st?.dispersionDiagnostic?.ok
) {
  const d =
    st.dispersionDiagnostic;

  const ratioText =
    Number(d.ratio).toFixed(1);

  const lowerText =
    Number(d.lowerLimit).toFixed(1);

  const upperText =
    Number(d.upperLimit).toFixed(1);

  if (d.status === "overdispersed") {
    html +=
      `<li>` +
      `<strong>U′ may be a better fit for these data.</strong> ` +
      `The rates vary more from point to point than a standard U chart expects for the amount of opportunity. ` +
      `This can make a standard U chart produce limits that are too narrow and create extra alerts. ` +
      `A U′ chart adjusts the limits for this extra variation while keeping the same rates and centre line.` +
      `<div style="margin-top:0.45rem; font-size:0.92em;">` +
      `<strong>Technical check:</strong> observed variation = ${ratioText}% of expected; ` +
      `diagnostic upper limit = ${upperText}%.` +
      `</div>` +
      `<div style="margin-top:0.55rem;">` +
      `<button type="button" onclick="switchUToUPrime()">` +
      `Use U′ chart instead` +
      `</button>` +
      `</div>` +
      `</li>`;

  } else if (
    d.status === "underdispersed"
  ) {
    html +=
      `<li>` +
      `<strong>U′ may be a better fit for these data.</strong> ` +
      `The rates vary less from point to point than a standard U chart expects. ` +
      `This can make standard U-chart limits too wide and make genuine changes harder to detect. ` +
      `A U′ chart adjusts the limits to better match the variation actually seen in the data.` +
      `<div style="margin-top:0.45rem; font-size:0.92em;">` +
      `<strong>Technical check:</strong> observed variation = ${ratioText}% of expected; ` +
      `diagnostic lower limit = ${lowerText}%.` +
      `</div>` +
      `<div style="margin-top:0.55rem;">` +
      `<button type="button" onclick="switchUToUPrime()">` +
      `Use U′ chart instead` +
      `</button>` +
      `</div>` +
      `</li>`;

  } else {
    html +=
      `<li>` +
      `<strong>A standard U chart looks appropriate for these data.</strong> ` +
      `The amount of variation is within the range expected for defect-rate data, ` +
      `so there is no clear reason to use a U′ chart instead.` +
      `<div style="margin-top:0.45rem; font-size:0.92em;">` +
      `<strong>Technical check:</strong> observed variation = ${ratioText}% of expected; ` +
      `expected range = ${lowerText}%–${upperText}%.` +
      `</div>` +
      `</li>`;
  }
}

if (
  a.chartType === "uprime" &&
  Number.isFinite(st?.sigmaZ)
) {
  let adjustmentText = "";

  if (st.sigmaZ > 1) {
    adjustmentText =
      "These data vary more than a standard U chart expects, so the U′ chart uses wider control limits.";
  } else if (st.sigmaZ < 1) {
    adjustmentText =
      "These data vary less than a standard U chart expects, so the U′ chart uses narrower control limits.";
  } else {
    adjustmentText =
      "The amount of variation is very close to what a standard U chart expects, so the adjustment is small.";
  }

  html +=
    `<li>` +
    `<strong>Why are these limits different?</strong> ` +
    `${adjustmentText} ` +
    `This helps the control limits better reflect the amount of variation actually seen in the data.` +
    `<div style="margin-top:0.45rem; font-size:0.92em;">` +
    `<strong>Technical detail:</strong> Sigma Z = ${fmt(st.sigmaZ)}.` +
    `</div>` +
    `</li>`;
}

    // P vs P′ dispersion diagnostic
    if (
      a.chartType === "p" &&
      st?.dispersionDiagnostic?.ok
    ) {
      const d =
        st.dispersionDiagnostic;

      const ratioText =
        Number(d.ratio).toFixed(1);

      const lowerText =
        Number(d.lowerLimit).toFixed(1);

      const upperText =
        Number(d.upperLimit).toFixed(1);

      if (d.status === "overdispersed") {
        html +=
  `<li>` +
  `<strong>P′ may be a better fit for these data.</strong> ` +
  `The proportions vary more from point to point than a standard P chart expects for these sample sizes. ` +
  `This can make a standard P chart produce limits that are too narrow and create extra alerts. ` +
  `A P′ chart adjusts the limits for this extra variation while keeping the same proportions and centre line.` +
  `<div style="margin-top:0.45rem; font-size:0.92em;">` +
  `<strong>Technical check:</strong> observed variation = ${ratioText}% of expected; ` +
  `diagnostic upper limit = ${upperText}%.` +
  `</div>` +
  `<div style="margin-top:0.55rem;">` +
  `<button type="button" onclick="switchPToPPrime()">` +
  `Use P′ chart instead` +
  `</button>` +
  `</div>` +
  `</li>`;

      } else if (
        d.status === "underdispersed"
      ) {
        html +=
  `<li>` +
  `<strong>P′ may be a better fit for these data.</strong> ` +
  `The proportions vary less from point to point than a standard P chart expects. ` +
  `This can make standard P-chart limits too wide and make genuine changes harder to detect. ` +
  `A P′ chart adjusts the limits to better match the variation actually seen in the data.` +
  `<div style="margin-top:0.45rem; font-size:0.92em;">` +
  `<strong>Technical check:</strong> observed variation = ${ratioText}% of expected; ` +
  `diagnostic lower limit = ${lowerText}%.` +
  `</div>` +
  `<div style="margin-top:0.55rem;">` +
  `<button type="button" onclick="switchPToPPrime()">` +
  `Use P′ chart instead` +
  `</button>` +
  `</div>` +
  `</li>`;

      } else {
        html +=
  `<li>` +
  `<strong>A standard P chart looks appropriate for these data.</strong> ` +
  `The amount of variation is within the range expected for proportion data of these sample sizes, ` +
  `so there is no clear reason to use a P′ chart instead.` +
  `<div style="margin-top:0.45rem; font-size:0.92em;">` +
  `<strong>Technical check:</strong> observed variation = ${ratioText}% of expected; ` +
  `expected range = ${lowerText}%–${upperText}%.` +
  `</div>` +
  `</li>`;
      }
    }

    // Laney P′ adjustment
    if (
      a.chartType === "pprime" &&
      Number.isFinite(st?.sigmaZ)
    ) {
      let adjustmentText = "";

      if (st.sigmaZ > 1) {
  adjustmentText =
    "These data vary more than a standard P chart expects, so the P′ chart uses wider control limits.";
} else if (st.sigmaZ < 1) {
  adjustmentText =
    "These data vary less than a standard P chart expects, so the P′ chart uses narrower control limits.";
} else {
  adjustmentText =
    "The amount of variation is very close to what a standard P chart expects, so the adjustment is small.";
}

      html +=
        `<li>` +
`<strong>Why are these limits different?</strong> ` +
`${adjustmentText} ` +
`This helps the control limits better reflect the amount of variation actually seen in the data.` +
`<div style="margin-top:0.45rem; font-size:0.92em;">` +
`<strong>Technical detail:</strong> Sigma Z = ${fmt(st.sigmaZ)}.` +
`</div>` +
`</li>`;
    }

    // Signals
    if (
      !a.isStable &&
      Array.isArray(a.signals) &&
      a.signals.length
    ) {
      html +=
        `<li><strong>Signals:</strong> ${a.signals.join("; ")}.</li>`;
    }

    // Interpretation
    const interpretation =
      a.isStable
        ? "Based on the selected rules, no clear special-cause signals were detected in this period. This suggests the pattern is consistent with routine (common-cause) variation."
        : "Based on the selected rules, special-cause signals were detected in this period (a pattern unlikely to be routine variation alone).";

    const caution =
      "These rules are prompts, not absolute answers. Interpret alongside local context (changes in process, staffing, demand, definitions/coding) and consider basic SPC assumptions (e.g. reasonably consistent measurement and opportunity over time).";

    html +=
      `<li><strong>Interpretation:</strong> ${interpretation}</li>`;

    html +=
      `<li><strong>Note:</strong> ${caution}</li>`;

    html += `</ul>`;
    html += `</div>`;
  });

  summaryDiv.innerHTML =
    html;
}



function renderAttributeSummary(a) {
  if (!summaryDiv) return;

  const nameMap = {
  c: "C chart",
  p: "P chart",
  pprime: "P′ chart",
  uprime: "U′ chart",
  u: "U chart"
};
  const chartName = nameMap[a.chartType] || "Chart";

  const stableLine = a.isStable
    ? "No clear signal of change (routine ups and downs)."
    : "A signal of change is present (worth investigating).";

  let html = `<h3>${chartName} summary</h3>`;
  html += `<ul>`;
  html += `<li><strong>What this suggests:</strong> ${stableLine}</li>`;

  if (!a.isStable && Array.isArray(a.signals) && a.signals.length) {
    html += `<li><strong>What I can see:</strong> ${a.signals.join("; ")}.</li>`;
  }

  if (a.firstOutOfControl) {
    const ex = a.firstOutOfControl;
    const exText = ex.type === "aboveUCL"
      ? "above the upper limit"
      : "below the lower limit";
    html += `<li><strong>Example to check:</strong> ${ex.label} is ${exText}.</li>`;
  }

  // Very short “what to do next” guidance (plain English)
  html += a.isStable
    ? `<li><strong>What to do next:</strong> If performance isn’t good enough, focus on changing the process (the system) rather than reacting to individual points.</li>`
    : `<li><strong>What to do next:</strong> Look for a real-world explanation (process change, staffing, demand, definition/coding). If it was a planned change, you may want a new baseline after it settles.</li>`;

  // Gentle chart-type hint (reduces misuse)
  if (a.chartType === "c") {
    html += `<li><strong>Best used when:</strong> Each time period is broadly comparable (similar time window / similar-sized service).</li>`;
  } else if (
  a.chartType === "p" ||
  a.chartType === "pprime"
) {
    html += `<li><strong>Best used when:</strong> You have a number out of a total each time (a proportion or %).</li>`;
  } else if (
  a.chartType === "u" ||
  a.chartType === "uprime"
) {
    html += `<li><strong>Best used when:</strong> You have a rate where the “out of how many” changes (e.g., per 1,000 bed days).</li>`;
  }

  html += `</ul>`;
  summaryDiv.innerHTML = html;
}


function renderRareChartSummary(aOrSegments, totalPointsMaybe) {
  // Backwards-compatible: accept either a single analysis object OR an array of analyses
  const segments = Array.isArray(aOrSegments) ? aOrSegments : [aOrSegments];
  const totalPoints = Number.isFinite(totalPointsMaybe)
    ? totalPointsMaybe
    : (Array.isArray(segments) && segments.length && Number.isFinite(segments[segments.length - 1]?.totalPoints))
      ? segments[segments.length - 1].totalPoints
      : null;

  if (!summaryDiv) return;
  if (!segments.length || !segments[0]) return;

  const chartType = segments[0].chartType;
  const chartName = chartType === "t" ? "T chart" : "G chart";

  const fmt = (v) =>
    Number.isFinite(v)
      ? (typeof formatNumber === "function" ? formatNumber(v, 3) : Number(v).toFixed(3))
      : "—";

  // XmR-style header
  let html = `<h3>Summary (${chartName})</h3>`;

  if (Number.isFinite(totalPoints)) {
    html += `<p>Total number of points: <strong>${totalPoints}</strong>. `;
    html += `The chart is divided into <strong>${segments.length}</strong> period${segments.length !== 1 ? "s" : ""} `;
    html += `(based on the baseline and any splits).</p>`;
  } else {
    html += `<p>The chart is divided into <strong>${segments.length}</strong> period${segments.length !== 1 ? "s" : ""} `;
    html += `(based on the baseline and any splits).</p>`;
  }

  // Rare-event note (keep it plain-English + accurate)
  html += `<p><strong>Note:</strong> Rare-event charts are often skewed, so the control limits may not look symmetrical like an XmR chart.</p>`;

  segments.forEach((a, idx) => {
    html += `<div class="pdf-avoid-break">`;
    html += segments.length > 1 ? `<h4>Period ${idx + 1}</h4>` : `<h4>Single period</h4>`;
    html += `<ul>`;

    // Coverage
    if (a.startIndex != null && a.endIndex != null && a.labelStart && a.labelEnd) {
      const n = (a.endIndex - a.startIndex + 1);
      html += `<li><strong>Coverage:</strong> <strong>points ${a.startIndex + 1}–${a.endIndex + 1}</strong> (${escapeHtml(a.labelStart)} to ${escapeHtml(a.labelEnd)}) – ${n} points.</li>`;
    } else if (a.nPoints != null) {
      html += `<li><strong>Coverage:</strong> ${a.nPoints} points.</li>`;
    }

    // Baseline
    if (a.baselineCountUsed != null) {
      html += `<li><strong>Baseline for this period:</strong> first ${a.baselineCountUsed} points used to calculate centre line and limits.</li>`;
    }

    // Stats: centre line + limits
    if (a.stats && Number.isFinite(a.stats.cl)) {
      if (chartType === "t") {
        html += `<li><strong>Centre line (average gap):</strong> ${fmt(a.stats.cl)}; <strong>upper limit:</strong> UCL = ${fmt(a.stats.ucl)}.</li>`;
      } else {
        html += `<li><strong>Centre line (average opportunities):</strong> ${fmt(a.stats.cl)}; <strong>control limits:</strong> LCL = ${fmt(a.stats.lcl)}, UCL = ${fmt(a.stats.ucl)}.</li>`;
      }
    } else {
      // Fallback if stats missing
      html += `<li><strong>Centre line and limits:</strong> (not available).</li>`;
    }

    // Signals
    if (!a.isStable && Array.isArray(a.signals) && a.signals.length) {
      html += `<li><strong>Signals:</strong> ${a.signals.join("; ")}.</li>`;
    }

    // Interpretation (XmR-style)
    const interpretation = a.isStable
      ? "No clear special-cause signals were detected in this period. The pattern is consistent with routine/common variation."
      : "Special-cause signals were detected in this period (pattern inconsistent with routine variation).";

    html += `<li><strong>Interpretation:</strong> ${interpretation}</li>`;

    // “Better” guidance (rare charts often need this)
    html += `<li><strong>Interpreting “better”:</strong> If the event is something you want to avoid, longer gaps (or more opportunities between events) are usually better. If it’s something you want more often, shorter gaps may be better.</li>`;

    // “What next” guidance (brief, not bossy)
    html += a.isStable
      ? `<li><strong>What to do next:</strong> If performance isn’t good enough, focus on improving the system rather than reacting to individual points.</li>`
      : `<li><strong>What to do next:</strong> Look for a real-world explanation (process, staffing, demand, detection/definition changes). If it was planned, consider a new baseline after things settle.</li>`;

    // Data reminder
    html += (chartType === "t")
      ? `<li><strong>Data reminder:</strong> This chart uses the time between events (e.g., days between incidents).</li>`
      : `<li><strong>Data reminder:</strong> This chart uses opportunities between events (values should be 1 or more).</li>`;

    html += `</ul>`;
    html += `</div>`;
  });

  summaryDiv.innerHTML = html;
}



function updateRunSummary(points, medianIgnored, ruleHitsIgnored, baselineCountUsedIgnored) {
  if (!summaryDiv) return;

  const { shiftLength, trendLength } = getRuleSettings();
  const n = points.length;

  // Pull baseline setting from the UI (so it works per-period too)
  const rawBaseline = baselineInput ? parseInt(baselineInput.value, 10) : NaN;
  const baselineSetting = Number.isFinite(rawBaseline) ? rawBaseline : null;

  // Build segments based on splits (splits are “after index”, 0-based)
  const splitIdxs = Array.isArray(splits)
    ? splits
        .map(v => parseInt(v, 10))
        .filter(v => Number.isInteger(v) && v >= 0 && v <= n - 2)
        .sort((a, b) => a - b)
    : [];

  const boundaries = [-1, ...splitIdxs, n - 1]; // inclusive ends
  const segments = [];
  for (let i = 0; i < boundaries.length - 1; i++) {
    const start = boundaries[i] + 1;
    const end = boundaries[i + 1];
    if (start <= end) segments.push({ start, end });
  }

  function rangeText(start, end) {
  const axisType = getAxisType(); // "date" or "sequence"

  // Always show point indices
  const base = `points ${start + 1}–${end + 1}`;

  // Only add date range if DATE axis is selected
  if (axisType !== "date") return base;

  const a = points[start]?.x;
  const b = points[end]?.x;
  const hasDates = a !== undefined && b !== undefined && a !== null && b !== null;

  return hasDates
    ? `${base} (${formatDateOnlyLabel(a)} to ${formatDateOnlyLabel(b)})`
    : base;
}



  function findTrendRanges(values, len) {
    const out = [];
    if (!values || values.length < len) return out;

    let inc = 1, dec = 1;
    let incStart = 0, decStart = 0;

    for (let i = 1; i < values.length; i++) {
      if (values[i] > values[i - 1]) {
        inc++;
        dec = 1;
        decStart = i;
      } else if (values[i] < values[i - 1]) {
        dec++;
        inc = 1;
        incStart = i;
      } else {
        inc = 1; dec = 1;
        incStart = i; decStart = i;
      }

      if (inc === len) out.push({ start: i - len + 1, end: i, dir: "increasing" });
      if (dec === len) out.push({ start: i - len + 1, end: i, dir: "decreasing" });
    }

    // Merge overlaps
    out.sort((a, b) => a.start - b.start);
    const merged = [];
    for (const r of out) {
      const last = merged[merged.length - 1];
      if (!last || r.start > last.end + 1 || r.dir !== last.dir) {
        merged.push({ ...r });
      } else {
        last.end = Math.max(last.end, r.end);
      }
    }
    return merged;
  }

  function flagsToRanges(flags) {
    const ranges = [];
    let i = 0;
    while (i < flags.length) {
      if (!flags[i]) { i++; continue; }
      let j = i;
      while (j < flags.length && flags[j]) j++;
      ranges.push({ start: i, end: j - 1 });
      i = j;
    }
    return ranges;
  }

  let html = `<h3>Summary (Run chart)</h3>`;
  html += `<p>Total number of points: <strong>${n}</strong>. `;
  html += `The chart is divided into <strong>${segments.length}</strong> period${segments.length !== 1 ? "s" : ""}`;
  html += segments.length > 1 ? ` (based on your splits).` : `.`;
  html += `</p>`;

  segments.forEach((seg, idx) => {
    const segPoints = points.slice(seg.start, seg.end + 1);
    const values = segPoints.map(p => p.y);

    const segLen = values.length;

    // Baseline per period (keep the same user setting, but cap to segment length)
    let baselineCountUsed = segLen;
    if (baselineSetting && baselineSetting >= 2) baselineCountUsed = Math.min(baselineSetting, segLen);

    const baselineValues = values.slice(0, baselineCountUsed);
    const median = computeMedian(baselineValues);

    // Signals for this segment
    const runFlags = detectLongRuns(values, median, shiftLength);
    const runRanges = flagsToRanges(runFlags);

    const trendRanges = findTrendRanges(values, trendLength);

    // Astronomical points (use baseline values as reference if possible)
    const astro = findAstronomicalPoints(values, median, baselineValues, 3.5);
    const astroIdx = astro?.indices || [];

    const signals = [];
    if (runRanges.length) signals.push(`a sustained shift (≥ ${shiftLength} points on one side of the median)`);
    if (trendRanges.length) signals.push(`a sustained trend (≥ ${trendLength} points increasing or decreasing)`);
    if (astroIdx.length) signals.push(`an unusual outlier ("astronomical" point)`);

    const periodLabel =
      segments.length === 1
        ? "Single period"
        : idx === 0
          ? "Period 1"
          : `Period ${idx + 1}`;

    html += `<h4>${periodLabel}</h4>`;
    html += `<ul>`;
    html += `<li>Coverage: <strong>${rangeText(seg.start, seg.end)}</strong> – ${segLen} point${segLen !== 1 ? "s" : ""}.</li>`;

    html += (baselineCountUsed < segLen)
      ? `<li>Baseline for this period: first <strong>${baselineCountUsed}</strong> point${baselineCountUsed !== 1 ? "s" : ""} used to calculate the median.</li>`
      : `<li>Baseline for this period: all points in this period used to calculate the median.</li>`;

    html += `<li>Median (this period): <strong>${Number.isFinite(median) ? median.toFixed(3) : "—"}</strong>.</li>`;

    if (!signals.length) {
      html += `<li><strong>Interpretation:</strong> No clear special-cause signals detected in this period (no sustained shift, trend, or unusual outlier). This pattern is consistent with common variation, but always interpret in context.</li>`;
    } else {
      html += `<li><strong>Interpretation:</strong> This period shows special-cause signals: ${signals.join("; ")}.</li>`;

      // “Where to look” (simple + practical)
      const where = [];

      if (runRanges.length) {
        const r = runRanges[0];
        where.push(`shift around points ${seg.start + r.start + 1}–${seg.start + r.end + 1}`);
      }

      if (trendRanges.length) {
        const t = trendRanges[0];
        where.push(`trend around points ${seg.start + t.start + 1}–${seg.start + t.end + 1} (${t.dir})`);
      }

      if (astroIdx.length) {
        const pts = astroIdx.slice(0, 5).map(i => seg.start + i + 1);
        where.push(`outlier at point${pts.length !== 1 ? "s" : ""} ${pts.join(", ")}`);
      }

      if (where.length) {
        html += `<li><strong>Where to look:</strong> ${where.join("; ")}.</li>`;
      }
    }

    html += `</ul>`;
  });

  if (segments.length > 1) {
    html += `<p><em>Note:</em> Each period is summarised separately because splits suggest the process may have changed over time.</p>`;
  }

  summaryDiv.innerHTML = html;
}



function showStatusMessage(msg) {
  if (typeof chartSummaryEl !== "undefined" && chartSummaryEl) {
    chartSummaryEl.textContent = msg;
  } else {
    alert(msg);
  }
}



// ---- Summary helpers ----

function meanFinite(arr) {
  const xs = (arr || []).filter(v => Number.isFinite(v));
  if (!xs.length) return NaN;
  return xs.reduce((a, b) => a + b, 0) / xs.length;
}

function rangeFinite(arr) {
  const xs = (arr || []).filter(v => Number.isFinite(v));
  if (!xs.length) return { min: NaN, max: NaN };
  return { min: Math.min(...xs), max: Math.max(...xs) };
}



// Multi-period XmR summary (handles baseline + splits) — lay-user interpretation + astronomical points
function updateXmRMultiSummary(segments, totalPoints) {
  if (!summaryDiv) return;

  if (!segments || segments.length === 0) {
    summaryDiv.innerHTML = "";
    if (capabilityDiv) capabilityDiv.innerHTML = "";
    return;
  }

  const target = getTargetValue();
  const direction = targetDirectionInput ? targetDirectionInput.value : "above";

  // Use configured thresholds + effective rule policy for XmR
  const rs = (typeof getEffectiveRuleSettingsForChart === "function")
    ? getEffectiveRuleSettingsForChart("xmr")
    : {
        shiftLength: 8,
        trendLength: 6,
        allowRunShift: true,
        allowTrend: false
      };

  const { shiftLength, trendLength } = rs;

  let html = `<h3>Summary (XmR chart)</h3>`;
  html += `<p>Total number of points: <strong>${totalPoints}</strong>. `;
  html += `The chart is divided into <strong>${segments.length}</strong> period${segments.length > 1 ? "s" : ""} `;
  html += `(based on the baseline and any splits).</p>`;

  // For capability badge (last period only)
  let lastPeriodSignals = [];
  let lastPeriodCapability = null;
  let lastPeriodHasCapability = false;

  segments.forEach((seg, idx) => {
    const { startIndex, endIndex, labelStart, labelEnd, result } = seg;
    const {
  mean,
  ucl,
  lcl,
  sigma,
  avgMR,
  baselineCountUsed,
   effectiveBaselinePointCount,
     manualExcludedPointCount,
     manualAdjacentMRExclusions,
  initialAvgMR,
  mrScreeningEnabled,
  mrScreeningThreshold,
  excludedMRCount
} = result;

    const points = result.points || [];
    const n = points.length;
    const values =
  points.map(
    p =>
      p?._excludeFromLimits
        ? NaN
        : p.y
  );

   const limitEstimatePointCount =
  Number.isFinite(
    effectiveBaselinePointCount
  )
    ? effectiveBaselinePointCount
    : (
        Number.isFinite(
          baselineCountUsed
        )
          ? baselineCountUsed
          : n
      );

const hasEnoughPointsForInterpretation =
  limitEstimatePointCount >= MIN_XMR_LIMIT_POINTS;

    // --- Special-cause detection (simple, lay-focused labels) ---
    // 1) Points beyond limits
    const beyondIdx = [];
    points.forEach((p, i) => {
      if (p.beyondLimits) beyondIdx.push(i);
    });

    // 2) Sustained shift (run on one side of mean)
    let runRanges = [];
    if (typeof findLongRunRanges === "function") {
      runRanges = findLongRunRanges(values, mean, shiftLength) || [];
    } else {
      // fallback: your existing boolean flags
      const runFlags = detectLongRuns(values, mean, shiftLength);
      let any = runFlags.some(Boolean);
      if (any) runRanges = [{ start: 0, end: 0 }]; // placeholder (we won't list ranges in fallback)
    }

    // 3) Trend
    let trendRanges = [];

if (rs.allowTrend) {
  if (typeof findTrendRanges === "function") {
    trendRanges = findTrendRanges(values, trendLength) || [];
  } else {
    const hasTrend = detectTrend(values, trendLength);
    if (hasTrend) trendRanges = [{ start: 0, end: 0 }];
  }
}

    // 4) Astronomical point (robust outlier)
    // Use baseline of this *period* to set the reference for outlier detection where possible.
    let astro = { indices: [], flags: [] };
    if (typeof findAstronomicalPoints === "function") {
      const periodBaselineCount = (baselineCountUsed && baselineCountUsed >= 3) ? baselineCountUsed : n;
      const refValues = values.slice(0, Math.min(periodBaselineCount, values.length));
      astro = findAstronomicalPoints(values, mean, refValues, 3.5) || { indices: [], flags: [] };
    }

    // Build simple signals list
    const signals = [];

// Only interpret SPC signals once enough observations have been used
// to establish the XmR limits for this period.
if (hasEnoughPointsForInterpretation) {
  if (beyondIdx.length > 0) {
    signals.push("one or more points are outside the control limits");
  }

  if (runRanges.length > 0) {
    signals.push(
      "a sustained shift (many points on the same side of the mean)"
    );
  }

  if (rs.allowTrend && trendRanges.length > 0) {
    signals.push(
      "a sustained trend (steady increase or decrease)"
    );
  }

  if (astro.indices && astro.indices.length > 0) {
    signals.push(
      "an unusual outlier (an ‘astronomical’ point)"
    );
  }
}

    // Capability (only if target exists and sigma > 0)
    let capability = null;

if (
  hasEnoughPointsForInterpretation &&
  target !== null &&
  sigma > 0
) {
  capability = computeTargetCapability(
    mean,
    sigma,
    target,
    direction
  );
}

    // Target coverage in this period
    let targetCoverageText = "";
    if (target !== null && n > 0) {
      let hits = 0;
      points.forEach(p => {
        if (direction === "above") {
          if (p.y >= target) hits++;
        } else {
          if (p.y <= target) hits++;
        }
      });
      const prop = hits / n;
      targetCoverageText = `${(prop * 100).toFixed(1)}% of points in this period meet the target (${hits}/${n}).`;
    }

    const periodLabel =
      segments.length === 1
        ? "Single period"
        : idx === 0
          ? "Period 1 (initial segment / baseline)"
          : `Period ${idx + 1}`;

    const base = `points ${startIndex + 1}–${endIndex + 1}`;
const rangeText =
  (getAxisType() === "date" && labelStart !== undefined && labelEnd !== undefined)
    ? `${base} (${formatDateOnlyLabel(labelStart)} to ${formatDateOnlyLabel(labelEnd)})`
    : base;


    html += `<div class="pdf-avoid-break">`;    
    html += `<h4>${periodLabel}</h4>`;
    html += `<ul>`;
    html += `<li>Coverage: <strong>${rangeText}</strong> – ${n} point${n !== 1 ? "s" : ""}.</li>`;

        const baselineWindowCount =
      Number.isFinite(
        baselineCountUsed
      )
        ? baselineCountUsed
        : n;

    const baselineUsedCount =
      Number.isFinite(
        effectiveBaselinePointCount
      )
        ? effectiveBaselinePointCount
        : Math.max(
            0,
            baselineWindowCount -
            (
              manualExcludedPointCount ||
              0
            )
          );

    if (
      baselineWindowCount < n ||
      (
        manualExcludedPointCount ||
        0
      ) > 0
    ) {
      html +=
        `<li>Baseline for this period: first <strong>${baselineWindowCount}</strong> point${baselineWindowCount !== 1 ? "s" : ""} define the baseline window; ` +
        `<strong>${baselineUsedCount}</strong> used to calculate mean and limits` +
        `${
          manualExcludedPointCount
            ? `, ${manualExcludedPointCount} excluded`
            : ""
        }.</li>`;

    } else {
      html +=
        `<li>Baseline for this period: all points in this period used to calculate mean and limits.</li>`;
    }

    html += `<li>Mean: <strong>${mean.toFixed(3)}</strong>; control limits: <strong>LCL = ${lcl.toFixed(3)}</strong>, <strong>UCL = ${ucl.toFixed(3)}</strong>.</li>`;
    html += `<li>Estimated σ (from MR): <strong>${sigma.toFixed(3)}</strong> (average MR = ${avgMR.toFixed(3)}).</li>`;
    if (
  manualAdjacentMRExclusions >
  0
) {
  html +=
    `<li><strong>Excluded point adjustment:</strong> ` +
    `${manualAdjacentMRExclusions} adjacent moving range${
      manualAdjacentMRExclusions === 1
        ? " was"
        : "s were"
    } not used in the variation estimate because ${
      manualAdjacentMRExclusions === 1
        ? "it involves"
        : "they involve"
    } an excluded point. ` +
    `SimpleSPC does not create an artificial moving range across an excluded point.</li>`;
}

// Explain XmR moving-range screening only when it is relevant.
// Keep the main wording simple, with the technical method second.
if (mrScreeningEnabled && excludedMRCount > 0) {
  const changeWord =
    excludedMRCount === 1
      ? "change was"
      : "changes were";

  html += `
    <li>
      <strong>XmR limit calculation:</strong>
      ${excludedMRCount} unusually large point-to-point
      ${changeWord} not used when estimating routine variation.
      All data points remain shown on the chart.
    </li>
  `;

  if (
    Number.isFinite(initialAvgMR) &&
    Number.isFinite(mrScreeningThreshold)
  ) {
    html += `
      <li class="hint">
        Nelson moving-range screening was used:
        initial average MR = ${initialAvgMR.toFixed(3)};
        screening threshold = ${mrScreeningThreshold.toFixed(3)};
        revised average MR = ${avgMR.toFixed(3)}.
      </li>
    `;
  }
} else if (mrScreeningEnabled === false) {
  html += `
    <li>
      <strong>XmR limit calculation:</strong>
      Moving-range screening is turned off, so all baseline
      point-to-point changes were used to estimate routine variation.
    </li>
  `;
}

if (!hasEnoughPointsForInterpretation) {
  html += `
    <li>
      <strong>Provisional limits:</strong>
      only <strong>${limitEstimatePointCount}</strong>
      point${limitEstimatePointCount !== 1 ? "s were" : " was"}
      used to estimate the limits for this period.
      SimpleSPC needs at least <strong>${MIN_XMR_LIMIT_POINTS}</strong>
      points before interpreting special-cause signals or capability.
      The displayed limits are shown for context only.
    </li>
  `;
}

if (target !== null) {
      html += `<li>Target: <strong>${target}</strong> (${direction === "above" ? "at or above is better" : "at or below is better"}). `;
      html += targetCoverageText ? (targetCoverageText + `</li>`) : `Target coverage not calculated for this period.</li>`;
    }

    // Simple, clearly labelled interpretation
    if (!hasEnoughPointsForInterpretation) {
  html += `
    <li>
      <strong>Interpretation:</strong>
      Not yet assessed — more data are needed before this period
      can be interpreted as showing routine or special-cause variation.
    </li>
  `;
} else if (signals.length === 0) {
  html += `<li><strong>Interpretation:</strong> No clear special-cause signals were detected in this period. The pattern is consistent with natural/common variation (still interpret in clinical context).</li>`;
} else {
      html += `<li><strong>Interpretation:</strong> This period shows special-cause signals: ${signals.join("; ")}.</li>`;

      // Optional: very short “where” hints (kept minimal)
      const whereBits = [];

      if (beyondIdx.length > 0) {
        const shown = beyondIdx.slice(0, 3).map(i => (startIndex + i + 1));
        whereBits.push(`outside limits at point${shown.length > 1 ? "s" : ""} ${shown.join(", ")}${beyondIdx.length > 3 ? ", …" : ""}`);
      }

      if (astro.indices && astro.indices.length > 0) {
        const shown = astro.indices.slice(0, 3).map(i => (startIndex + i + 1));
        whereBits.push(`outlier at point${shown.length > 1 ? "s" : ""} ${shown.join(", ")}${astro.indices.length > 3 ? ", …" : ""}`);
      }

      // Only add “where” if we actually have something specific to show
      if (whereBits.length > 0) {
        html += `<li><strong>Where to look:</strong> ${whereBits.join("; ")}.</li>`;
      }
    }

    if (capability && sigma > 0) {
      if (signals.length === 0) {
        html += `<li><strong>Estimated capability (this period):</strong> if the process remains stable, about <strong>${(capability.prob * 100).toFixed(1)}%</strong> of future points are expected to meet the target.</li>`;
      } else {
        html += `<li><strong>Capability:</strong> a target has been set, but because special-cause signals are present in this period, any capability estimate would be unreliable.</li>`;
      }
    }

    html += `</ul>`;
    html += `</div>`;

    // Remember last period for badge + helper (store structured information)
    if (idx === segments.length - 1) {
      lastPeriodSignals = signals;
      lastPeriodCapability = capability;
      lastPeriodHasCapability =
  hasEnoughPointsForInterpretation &&
  sigma > 0 &&
  !!capability;

      const hasTrend = trendRanges.length > 0;
      const hasRunViolation = runRanges.length > 0;
      const hasAstronomical = !!(astro.indices && astro.indices.length > 0);
      const nBeyond = beyondIdx.length;

      lastXmRAnalysis = {
  mean,
  ucl,
  lcl,
  sigma,

  // Moving-range calculation
  avgMR,
  initialAvgMR: result.initialAvgMR,
  mrScreeningEnabled: result.mrScreeningEnabled,
  mrScreeningThreshold: result.mrScreeningThreshold,
  excludedMRCount: result.excludedMRCount,
  excludedMRs: Array.isArray(result.excludedMRs)
    ? result.excludedMRs.map(item => ({ ...item }))
    : [],
  baselineMRCount: result.baselineMRCount,
  effectiveMRCount: result.effectiveMRCount,

  n,
  signals: signals.slice(),
        hasTrend,
        hasRunViolation,
        hasAstronomical,
        nBeyond,
        baselineCountUsed,
        target,
        direction,
                capability,

        hasEnoughPointsForInterpretation,
        limitEstimatePointCount,
        minimumPointsForInterpretation: MIN_XMR_LIMIT_POINTS,

        isStable: hasEnoughPointsForInterpretation
          ? signals.length === 0
          : null,

        // thresholds used (handy for helper explanations)
        shiftLength,
        trendLength,
	periodIndex: idx + 1,
 	 	  periodCount: segments.length,
  		  startIndex,
  		  endIndex,
  		  labelStart,
  	  	  labelEnd
      };
    }
  });

  if (target !== null && segments.length > 1) {
    html += `<p><em>Note:</em> comparing means, limits and target performance between periods can indicate whether the process changed after interventions.</p>`;
  }

  summaryDiv.innerHTML = html;

  // Capability badge – last period only
  if (!capabilityDiv) return;

  if (target === null || !lastPeriodHasCapability) {
    capabilityDiv.innerHTML = "";
    return;
  }

  const hasAnySignals = lastPeriodSignals && lastPeriodSignals.length > 0;

  if (!hasAnySignals && lastPeriodCapability) {
    capabilityDiv.innerHTML = `
      <div style="
        display:inline-block;
        padding:0.6rem 1.2rem;
        background:#fff59d;
        border:1px solid #ccc;
        border-radius:0.25rem;
      ">
        <div style="font-weight:bold; text-align:center;">PROCESS CAPABILITY (last period)</div>
        <div style="font-size:1.4rem; font-weight:bold; text-align:center; margin-top:0.2rem;">
          ${(lastPeriodCapability.prob * 100).toFixed(1)}%
        </div>
        <div style="font-size:0.8rem; margin-top:0.2rem;">
          (Estimated probability of meeting the target in the final period, assuming a stable process and approximate normality.)
        </div>
      </div>
    `;
  } else if (target !== null && hasAnySignals) {
    capabilityDiv.innerHTML = `
      <div style="
        display:inline-block;
        padding:0.6rem 1.2rem;
        background:#ffe0b2;
        border:1px solid #ccc;
        border-radius:0.25rem;
        max-width:32rem;
      ">
        <strong>Special-cause signals detected in the last period:</strong> capability estimates may be unreliable until these signals are understood.
        Focus on understanding and addressing these causes before relying on capability estimates.
      </div>
    `;
  } else {
    capabilityDiv.innerHTML = "";
  }
}

function renderXbarSSummary(
  periodAnalyses,
  totalSubgroups
) {
  if (!summaryDiv) return;

  const periods =
    Array.isArray(
      periodAnalyses
    )
      ? periodAnalyses.filter(
          Boolean
        )
      : periodAnalyses
        ? [periodAnalyses]
        : [];

  if (!periods.length) {
    return;
  }

  const fmt =
    (value) =>
      Number.isFinite(value)
        ? Number(value)
            .toFixed(3)
        : "not available";

  const signalText =
    (
      analysis,
      stableText
    ) => {
      if (
        analysis?.isStable
      ) {
        return stableText;
      }

      const signals =
        Array.isArray(
          analysis?.signals
        )
          ? analysis.signals
              .filter(Boolean)
          : [];

      return signals.length
        ? (
            `signal(s): ` +
            `${signals.join("; ")}.`
          )
        : (
            "one or more special-cause signals were detected."
          );
    };

  let html =
    `<h3>X̄–S summary</h3>`;

  html +=
    `<p>Total number of subgroups: ` +
    `<strong>${totalSubgroups}</strong>. ` +
    `The chart is divided into ` +
    `<strong>${periods.length}</strong> ` +
    `period${periods.length === 1 ? "" : "s"} ` +
    `(based on the baseline and any splits).</p>`;

  periods.forEach(
    (
      period,
      index
    ) => {
      const x =
        period.xbar;

      const s =
        period.s;

      const stats =
        period.stats || {};

      const xStable =
        !!x?.isStable;

      const sStable =
        !!s?.isStable;

      const variableSubgroupSizes =
        !!stats.variableSubgroupSizes;

      const subgroupSizeText =
        (
          variableSubgroupSizes &&
          Number.isFinite(
            stats.subgroupSizeMin
          ) &&
          Number.isFinite(
            stats.subgroupSizeMax
          )
        )
          ? (
              `${stats.subgroupSizeMin}–` +
              `${stats.subgroupSizeMax} ` +
              `(varies by subgroup)`
            )
          : Number.isFinite(
              stats.subgroupSize
            )
            ? `${stats.subgroupSize}`
            : "not stated";

      const coverage =
        (
          period.labelStart !==
            undefined &&
          period.labelEnd !==
            undefined
        )
          ? (
              `${period.labelStart} ` +
              `to ` +
              `${period.labelEnd}`
            )
          : (
              `subgroups ` +
              `${period.startIndex + 1}–` +
              `${period.endIndex + 1}`
            );

      let interpretation =
        "";

      if (
        xStable &&
        sStable
      ) {
        interpretation =
          "Both the subgroup averages (X̄) and within-subgroup variation (S) look stable in this period.";

      } else if (
        !xStable &&
        sStable
      ) {
        interpretation =
          "The subgroup averages (X̄) show special-cause signals, while within-subgroup variation (S) looks stable.";

      } else if (
        xStable &&
        !sStable
      ) {
        interpretation =
          "The subgroup averages (X̄) look stable, while within-subgroup variation (S) shows special-cause signals.";

      } else {
        interpretation =
          "Both the subgroup averages (X̄) and within-subgroup variation (S) show special-cause signals.";
      }

      html +=
        `<div class="pdf-avoid-break">`;

      html +=
        `<h4>Period ${index + 1} of ` +
        `${periods.length}: ` +
        `${coverage}</h4>`;

      html += `<ul>`;

      html +=
        `<li>` +
        `<strong>Subgroup size:</strong> ` +
        `${subgroupSizeText} ` +
        `measurement${stats.subgroupSize === 1 ? "" : "s"} ` +
        `per subgroup.` +
        `</li>`;

            if (
        Number.isFinite(
          stats.baselineCountUsed
        ) &&
        (
          index === 0 ||
          Number(
            stats.excludedBaselineCount
          ) > 0
        )
      ) {
        const usedBaseline =
          Number.isFinite(
            stats.effectiveBaselineCount
          )
            ? stats.effectiveBaselineCount
            : stats.baselineCountUsed;

        const excludedBaseline =
          Number.isFinite(
            stats.excludedBaselineCount
          )
            ? stats.excludedBaselineCount
            : 0;

        html +=
          `<li>` +
          `<strong>Baseline:</strong> ` +
          `the first ${stats.baselineCountUsed} subgroup${stats.baselineCountUsed === 1 ? "" : "s"} define the calculation window; ` +
          `${usedBaseline} used to estimate the limits` +
          `${
            excludedBaseline
              ? `, ${excludedBaseline} excluded`
              : ""
          }.` +
          `</li>`;
      }

      if (
        variableSubgroupSizes
      ) {
        html +=
          `<li>` +
          `<strong>X̄ chart centre line:</strong> ` +
          `${fmt(stats.xbarbar)}; ` +
          `<strong>control limits vary with subgroup size:</strong> ` +
          `LCL range = ` +
          `${fmt(stats.xLclMin)}–${fmt(stats.xLclMax)}, ` +
          `UCL range = ` +
          `${fmt(stats.xUclMin)}–${fmt(stats.xUclMax)}.` +
          `</li>`;

        html +=
          `<li>` +
          `<strong>S chart:</strong> ` +
          `the expected S and its control limits vary with subgroup size. ` +
          `Centre-line range = ` +
          `${fmt(stats.sClMin)}–${fmt(stats.sClMax)}; ` +
          `LCL range = ` +
          `${fmt(stats.sLclMin)}–${fmt(stats.sLclMax)}; ` +
          `UCL range = ` +
          `${fmt(stats.sUclMin)}–${fmt(stats.sUclMax)}.` +
          `</li>`;

        html +=
          `<li>` +
          `<strong>Estimated process standard deviation:</strong> ` +
          `${fmt(stats.processSigma)}.` +
          `</li>`;

      } else {
        html +=
          `<li>` +
          `<strong>X̄ chart centre line:</strong> ` +
          `${fmt(stats.xbarbar)}; limits: ` +
          `LCL = ${fmt(stats.lclX)}, ` +
          `UCL = ${fmt(stats.uclX)}.` +
          `</li>`;

        html +=
          `<li>` +
          `<strong>S chart centre line:</strong> ` +
          `${fmt(stats.sbar)}; limits: ` +
          `LCL = ${fmt(stats.lclS)}, ` +
          `UCL = ${fmt(stats.uclS)}.` +
          `</li>`;
      }

      html +=
        `<li>` +
        `<strong>X̄ chart:</strong> ` +
        `${signalText(
          x,
          "stable (no clear signal of change in subgroup averages)."
        )}` +
        `</li>`;

      html +=
        `<li>` +
        `<strong>S chart:</strong> ` +
        `${signalText(
          s,
          "stable (no clear signal of change in within-subgroup variation)."
        )}` +
        `</li>`;

      html +=
        `<li>` +
        `<strong>Interpretation:</strong> ` +
        `${interpretation}` +
        `</li>`;

      html += `</ul>`;
      html += `</div>`;
    }
  );

  const latest =
    periods[
      periods.length - 1
    ];

  const latestXStable =
    !!latest
      ?.xbar
      ?.isStable;

  const latestSStable =
    !!latest
      ?.s
      ?.isStable;

  let latestText = "";

  if (
    latestXStable &&
    latestSStable
  ) {
    latestText =
      "The latest period looks stable on both parts of the X̄–S chart. " +
      "That means the current process level and within-subgroup variation are behaving broadly consistently under the current rule settings.";

  } else if (
    !latestXStable &&
    latestSStable
  ) {
    latestText =
      "The latest period shows possible change in the subgroup averages, while within-subgroup variation remains stable.";

  } else if (
    latestXStable &&
    !latestSStable
  ) {
    latestText =
      "The latest period shows possible change in within-subgroup variation, while the subgroup averages remain stable.";

  } else {
    latestText =
      "The latest period shows possible special-cause variation in both the subgroup averages and within-subgroup variation.";
  }

  html +=
    `<div class="pdf-avoid-break">` +
    `<p><strong>Latest period:</strong> ` +
    `${latestText}</p>` +
    `</div>`;

  summaryDiv.innerHTML =
    html;
}


// Approximate standard normal CDF Φ(z)
function normalCdf(z) {
  // Abramowitz & Stegun approximation
  const t = 1 / (1 + 0.2316419 * Math.abs(z));
  const d = 0.3989423 * Math.exp(-0.5 * z * z);
  let prob = d * t * (0.3193815 +
    t * (-0.3565638 +
    t * (1.781478 +
    t * (-1.821256 +
    t * 1.330274))));
  if (z > 0) prob = 1 - prob;
  return prob;
}

// mean, sigma from XmR; target number; direction "above"/"below"
function computeTargetCapability(mean, sigma, target, direction) {
  if (!isFinite(mean) || !isFinite(sigma) || sigma <= 0 || !isFinite(target)) {
    return null;
  }
  const z = (target - mean) / sigma;
  let p;
  if (direction === "above") {
    // P(X >= target)
    p = 1 - normalCdf(z);
  } else {
    // P(X <= target)
    p = normalCdf(z);
  }
  return { prob: p, z };
}

function getDateFormatPreference() {
  return (dateFormatPreferenceSelect?.value || "uk").toLowerCase();
}

function isAmbiguousNumericDateToken(s) {
  const m = String(s || "").trim().match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (!m) return false;

  const a = Number(m[1]);
  const b = Number(m[2]);

  return a >= 1 && a <= 12 && b >= 1 && b <= 12;
}

function detectNumericDateStyle(values) {
  let seenUkOnly = false;
  let seenUsOnly = false;
  let ambiguousCount = 0;

  for (const raw of values || []) {
    const s = String(raw ?? "").trim();
    if (!s) continue;

    const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (!m) continue;

    const a = Number(m[1]);
    const b = Number(m[2]);

    if (!(a >= 1 && a <= 31 && b >= 1 && b <= 31)) continue;

    if (a > 12 && b <= 12) seenUkOnly = true;   // e.g. 25/01/2024
    else if (b > 12 && a <= 12) seenUsOnly = true; // e.g. 01/25/2024
    else if (a <= 12 && b <= 12) ambiguousCount++;
  }

  if (seenUkOnly && !seenUsOnly) {
    return { style: "uk", ambiguousCount };
  }
  if (seenUsOnly && !seenUkOnly) {
    return { style: "us", ambiguousCount };
  }
  if (seenUkOnly && seenUsOnly) {
    return { style: "mixed", ambiguousCount };
  }
  return { style: "unknown", ambiguousCount };
}

function updateDateFormatWarning() {
  if (!dateFormatWarning) return;

  dateFormatWarning.style.display = "none";
  dateFormatWarning.textContent = "";

  if (!rawRows || !rawRows.length) return;
  if (!dateSelect || !dateSelect.value) return;

  const col = dateSelect.value;
  const axisType = getCheckedRadioValue("axisType");
  if (axisType !== "date") return;

  const values = rawRows
    .map(r => r?.[col])
    .filter(v => v !== null && v !== undefined && String(v).trim() !== "");

  if (!values.length) return;

  const result = detectNumericDateStyle(values);
  const pref = getDateFormatPreference();

  if (result.style === "mixed") {
    dateFormatWarning.textContent =
      "Warning: this column appears to contain a mixture of UK-style and US-style numeric dates. Please standardise the dates if possible.";
    dateFormatWarning.style.display = "block";
    return;
  }

  if (pref === "uk" && result.style === "us") {
    dateFormatWarning.textContent =
      "Warning: these dates look like US month-first dates (mm/dd/yyyy), but UK day-first is selected. Some rows may be ignored. Try switching to US month-first or Auto-detect.";
    dateFormatWarning.style.display = "block";
    return;
  }

  if (pref === "us" && result.style === "uk") {
    dateFormatWarning.textContent =
      "Warning: these dates look like UK day-first dates (dd/mm/yyyy), but US month-first is selected. Some rows may be ignored. Try switching to UK day-first or Auto-detect.";
    dateFormatWarning.style.display = "block";
    return;
  }

  if (result.ambiguousCount > 0) {
    if (pref === "uk") {
      dateFormatWarning.textContent =
        "This column contains ambiguous numeric dates. They will currently be interpreted as UK day-first dates (dd/mm/yyyy).";
      dateFormatWarning.style.display = "block";
      return;
    }

    if (pref === "us") {
      dateFormatWarning.textContent =
        "This column contains ambiguous numeric dates. They will currently be interpreted as US month-first dates (mm/dd/yyyy).";
      dateFormatWarning.style.display = "block";
      return;
    }

    if (pref === "auto" && result.style === "unknown") {
      dateFormatWarning.textContent =
        "This column contains ambiguous numeric dates and the tool cannot confidently auto-detect the style. It will fall back to UK day-first dates unless you choose another option.";
      dateFormatWarning.style.display = "block";
      return;
    }

    if (pref === "iso-only") {
      dateFormatWarning.textContent =
        "This column contains numeric slash/hyphen dates. In 'ISO / Excel dates only' mode, ambiguous numeric dates may not be interpreted as dates.";
      dateFormatWarning.style.display = "block";
      return;
    }
  }
}

function updateDateControlsState() {
  const axisType = document.querySelector("input[name='axisType']:checked")?.value;

  const isDateMode = axisType === "date";

  if (dateFormatPreferenceSelect) {
    dateFormatPreferenceSelect.disabled = !isDateMode;
  }
}

// Parse dates safely, supporting NHS-style dd/mm/yyyy as well as ISO yyyy-mm-dd
function parseDateValue(xRaw) {
  if (xRaw instanceof Date && !isNaN(xRaw)) {
    return xRaw;
  }

  if (xRaw === null || xRaw === undefined) {
    return new Date(NaN);
  }

  const s = String(xRaw).trim();
  if (!s) return new Date(NaN);

  // --- Excel serial date support ---
  const asNumber = Number(s);
  if (Number.isFinite(asNumber) && asNumber > 20000 && asNumber < 60000) {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    return new Date(excelEpoch.getTime() + asNumber * 86400000);
  }

  // ISO style: 2025-10-02 or 2025-10-02T...
  if (/^\d{4}-\d{2}-\d{2}(?:[T\s].*)?$/.test(s)) {
    const d = new Date(s);
    return isNaN(d.getTime()) ? new Date(NaN) : d;
  }

  // Numeric slash or hyphen dates
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:\s+.*)?$/);
  if (m) {
    let a = Number(m[1]);
    let b = Number(m[2]);
    let y = Number(m[3]);

    if (y < 100) {
      y += (y >= 50 ? 1900 : 2000);
    }

    const pref = getDateFormatPreference();

    let styleToUse = pref;

    if (pref === "auto") {
      if (a > 12 && b <= 12) styleToUse = "uk";
      else if (b > 12 && a <= 12) styleToUse = "us";
      else styleToUse = "uk"; // safe fallback for ambiguous cases
    }

    if (pref === "iso-only") {
      return new Date(NaN);
    }

    let day, month;
    if (styleToUse === "us") {
      month = a;
      day = b;
    } else {
      day = a;
      month = b;
    }

    if (!(month >= 1 && month <= 12 && day >= 1 && day <= 31)) {
      return new Date(NaN);
    }

    const d = new Date(y, month - 1, day);
    if (
      d.getFullYear() !== y ||
      d.getMonth() !== month - 1 ||
      d.getDate() !== day
    ) {
      return new Date(NaN);
    }

    return d;
  }

  // Fallback: native parsing for named-month strings like "02 Jan 2025"
  const d = new Date(s);
  return isNaN(d.getTime()) ? new Date(NaN) : d;
}


function validateCountLikeColumn(values, label, { allowZero = true } = {}) {
  for (let i = 0; i < values.length; i++) {
    const v = Number(values[i]);
    if (!Number.isFinite(v)) return `${label} has a non-numeric value at row ${i + 1}.`;
    if (v < 0) return `${label} has a negative value at row ${i + 1}.`;
    if (!allowZero && v === 0) return `${label} has a zero value at row ${i + 1}, but it must be > 0.`;
    // counts are usually integers — warn (not hard fail)
    if (!isIntegerish(v)) {
      return `${label} has a non-integer value at row ${i + 1}. Counts/denominators are usually whole numbers.`;
    }
  }
  return null;
}



// Parse numeric cells, including percentages like "55.17%"
function toNumericValue(raw) {
  if (raw === null || raw === undefined) return NaN;

  if (typeof raw === "number") return raw;

  const s = String(raw).trim();
  if (!s) return NaN;

  // Handle simple percentages, e.g. "55.17%" or "55.17 %"
  const percentMatch = s.match(/^(-?\d+(?:\.\d+)?)\s*%$/);
  if (percentMatch) {
    return Number(percentMatch[1]); // return 55.17
  }

  const num = Number(s);
  return isFinite(num) ? num : NaN;
}

function formatColumnName(colName, fallbackLabel = "this column") {
  return colName ? `"${colName}"` : fallbackLabel;
}

function getSuggestedAlternativeColumn(currentCol, candidates = []) {
  const filtered = (candidates || []).filter(c => c && c !== currentCol);
  return filtered.length ? filtered[0] : "";
}

function buildAlternativeSuggestionText(alternativeCol, purposeText) {
  if (!alternativeCol) return "";
  return ` Try ${formatColumnName(alternativeCol)} instead${purposeText ? ` as the ${purposeText}` : ""}.`;
}

function getCountLikeColumnCandidates({ requirePositive = false } = {}) {
  return allColumns.filter(col => {
    const p = getProfile(col);
    if (!p || !p.isNumeric || p.looksLikeDate || !p.isMostlyInteger || p.hasNeg) return false;
    if (requirePositive) {
      return Number.isFinite(p.min) && p.min > 0;
    }
    return true;
  });
}

function getContinuousMeasureCandidates(excludeCols = []) {
  return allColumns.filter(col => {
    const p = getProfile(col);
    if (!p || !p.isNumeric || p.looksLikeDate) return false;
    if (excludeCols.includes(col)) return false;
    if (p.looksIndexLike) return false;
    return true;
  });
}

function getRepeatingSubgroupCandidates(excludeCols = []) {
  return allColumns.filter(col => {
    const p = getProfile(col);
    if (!p || p.looksLikeDate) return false;
    if (excludeCols.includes(col)) return false;
    return !!p.repeatsOften;
  });
}

/* ============================================================
   VALIDATION HELPERS (P / U / C charts)
   - "error" => block chart generation
   - "warn"  => ask user whether to continue
   ============================================================ */

function isIntegerish(n) {
  return Number.isFinite(n) && Math.abs(n - Math.round(n)) < 1e-9;
}

function validateNonNegativeNumbers(arr, label, options = {}) {
  const {
    columnName = "",
    alternativeColumn = "",
    purposeText = ""
  } = options;

  for (let i = 0; i < arr.length; i++) {
    const v = Number(arr[i]);
    if (!Number.isFinite(v)) {
      return {
        level: "error",
        message:
          `${label} uses ${formatColumnName(columnName, "the selected column")}, ` +
          `but it contains a non-numeric value at row ${i + 1}.` +
          buildAlternativeSuggestionText(alternativeColumn, purposeText)
      };
    }
    if (v < 0) {
      return {
        level: "error",
        message:
          `${label} uses ${formatColumnName(columnName, "the selected column")}, ` +
          `but it contains a negative value at row ${i + 1}.` +
          buildAlternativeSuggestionText(alternativeColumn, purposeText)
      };
    }
  }
  return null;
}

function warnIfNonInteger(arr, label, options = {}) {
  const {
    columnName = "",
    alternativeColumn = "",
    purposeText = ""
  } = options;

  for (let i = 0; i < arr.length; i++) {
    const v = Number(arr[i]);
    if (Number.isFinite(v) && !isIntegerish(v)) {
      return {
        level: "warn",
        message:
          `${label} uses ${formatColumnName(columnName, "the selected column")}, ` +
          `but row ${i + 1} has a non-integer value (${v}). ` +
          `Counts and denominators are usually whole numbers.` +
          buildAlternativeSuggestionText(alternativeColumn, purposeText) +
          `\n\nGenerate the chart anyway?`
      };
    }
  }
  return null;
}

function validateDenominatorPositive(arr, label, options = {}) {
  const {
    columnName = "",
    alternativeColumn = "",
    purposeText = ""
  } = options;

  for (let i = 0; i < arr.length; i++) {
    const v = Number(arr[i]);
    if (!Number.isFinite(v)) {
      return {
        level: "error",
        message:
          `${label} uses ${formatColumnName(columnName, "the selected column")}, ` +
          `but it contains a non-numeric value at row ${i + 1}.` +
          buildAlternativeSuggestionText(alternativeColumn, purposeText)
      };
    }
    if (v <= 0) {
      return {
        level: "error",
        message:
          `${label} uses ${formatColumnName(columnName, "the selected column")}, ` +
          `but row ${i + 1} has value ${v}. Denominators/opportunities must be greater than 0.` +
          buildAlternativeSuggestionText(alternativeColumn, purposeText)
      };
    }
  }
  return null;
}

function validateNumeratorNotGreaterThanDenom(numArr, denomArr, options = {}) {
  const {
    numeratorColumn = "",
    denominatorColumn = "",
    alternativeNumerator = "",
    alternativeDenominator = ""
  } = options;

  for (let i = 0; i < numArr.length; i++) {
    const num = Number(numArr[i]);
    const den = Number(denomArr[i]);

    if (Number.isFinite(num) && Number.isFinite(den) && num > den) {
      let suggestion = "";
      if (alternativeNumerator || alternativeDenominator) {
        const parts = [];
        if (alternativeNumerator) parts.push(`numerator ${formatColumnName(alternativeNumerator)}`);
        if (alternativeDenominator) parts.push(`denominator ${formatColumnName(alternativeDenominator)}`);
        suggestion = ` Try ${parts.join(" and ")} instead.`;
      }

      return {
        level: "error",
        message:
          `P chart setup uses numerator ${formatColumnName(numeratorColumn, "the selected numerator column")} ` +
          `and denominator ${formatColumnName(denominatorColumn, "the selected denominator column")}, ` +
          `but row ${i + 1} has numerator ${num} greater than denominator ${den}.` +
          suggestion
      };
    }
  }
  return null;
}

let lastGenerateWasManual = false;

let axisTypeManuallyChanged = false;

function validateColumnSelectionSafety({ chartType, dateCol, valueCol, axisType }) {
  if (!dateCol || !valueCol) return null;

  // Hard stop: same column chosen for X and Y
  if (dateCol === valueCol) {
    return {
      level: "error",
      message: "Please choose different columns for the X-axis and the value."
    };
  }

  const xProf = (typeof getProfile === "function") ? getProfile(dateCol) : null;
  const yProf = (typeof getProfile === "function") ? getProfile(valueCol) : null;

  const warnings = [];

  const xLooksLikeSequence =
    !!(xProf &&
       xProf.isNumeric &&
       xProf.isMostlyInteger &&
       xProf.uniqueRatio >= 0.85 &&
       !xProf.repeatsOften &&
       !xProf.looksLikeDate);

  const yLooksLikeIndex =
    !!(yProf &&
       yProf.isNumeric &&
       yProf.isMostlyInteger &&
       yProf.uniqueRatio >= 0.85 &&
       !yProf.repeatsOften &&
       !yProf.looksLikeDate);

  // T chart (event dates mode): value column is ignored, so only check X
  if (chartType === "t" && typeof tChartInputMode !== "undefined" && tChartInputMode === "eventDates") {
    if (xProf && !xProf.looksLikeDate) {
      return {
        level: "warning",
        message:
          "T chart (event dates mode): the selected X-axis column does not look like a date/time column.\n\n" +
          "Generate the chart anyway?"
      };
    }
    return null;
  }

  if (axisType === "date" && xLooksLikeSequence) {
    warnings.push(
      "The selected X-axis column looks more like a numeric sequence than a date/time column. " +
      "Consider switching X-axis type to Sequence / category."
    );
  }

  if (yProf && yProf.looksLikeDate) {
    warnings.push(
      "The selected value column looks like a date/time field rather than a measurement."
    );
  }

  if (yLooksLikeIndex && (xProf?.looksLikeDate || xLooksLikeSequence)) {
    warnings.push(
      "The selected value column looks more like an ID / sequence column than an outcome measure. " +
      "Please check that you have chosen the column you want to chart."
    );
  }

  if ((chartType === "run" || chartType === "xmr") && axisType === "date" && xLooksLikeSequence && yLooksLikeIndex) {
    warnings.push(
      "Both selected columns look like numeric sequences. This can create a plausible-looking but misleading chart."
    );
  }

  if (!warnings.length) return null;

  return {
    level: "warning",
    message: warnings.join("\n\n") + "\n\nGenerate the chart anyway?"
  };
}

function handleValidationResult(result, { manual = true } = {}) {
  if (!result) return true;

  if (result.level === "error") {
    alert(result.message);
    return false;
  }

  // Warnings:
  // - If user clicked Generate, ask (confirm)
  // - If auto-regenerate, show inline message but do NOT interrupt
  if (!manual) {
    if (typeof showChartMessage === "function") {
      showChartMessage(result.message.replace(/\n\nGenerate the chart anyway\?/g, ""));
    } else {
      // fallback
      console.warn(result.message);
    }
    return true;
  }

  return confirm(result.message);
}


// ---- Generate chart button ----
generateButton.addEventListener("click", (evt) => {
  lastGenerateWasManual = !!evt.isTrusted;
  clearError();



  if (summaryDiv) summaryDiv.innerHTML = "";
  if (capabilityDiv) capabilityDiv.innerHTML = "";
  if (typeof validateAxisSettings === "function" && !validateAxisSettings()) return;
  if (!validateBeforeGenerate()) return;

  // A previous SPC helper answer may describe a different
  // chart, dataset, target or analysis setting.
  // Clear it before generating the updated analysis.
  if (spcHelperOutput) {
    spcHelperOutput.innerHTML = "";
  }

  if (aiQuestionInput) {
    aiQuestionInput.value = "";
  }

  try {
    const dateCol = dateSelect.value;
    const valueCol = valueSelect.value;
    const axisType = getAxisType();

        const chartType = getSelectedChartType_NoSideEffects();

    if (
      chartType ===
      "pareto"
    ) {
      if (currentChart) {
        currentChart.destroy();
        currentChart = null;
      }

      if (mrChart) {
        mrChart.destroy();
        mrChart = null;
      }

      if (mrPanel) {
        mrPanel.style.display =
          "none";
      }

      drawParetoChart(
        rawRows,
        dateCol,
        valueCol
      );

      if (
        yAxisBoundsManuallyEdited &&
        currentChart
      ) {
        applyCurrentChartYBoundsToInputs(
          currentChart
        );
      }

      if (
        typeof updateYAxisInputStep ===
        "function"
      ) {
        updateYAxisInputStep();
      }

      if (
        typeof clearDataModelDirty ===
        "function"
      ) {
        clearDataModelDirty();
      }

      return;
    }

    // --- 1) Build points depending on axis type and chart type ---
    let parsedPoints;

    // Special handling: T chart
    if (chartType === "t") {
      // If using event dates, we require date axis and we IGNORE the value column.
      if (tChartInputMode === "eventDates") {
        if (axisType !== "date") {
          showError("T chart (event dates mode) needs Date / time axis.");
          return;
        }

        parsedPoints = rawRows
          .map((row, idx) => {
            const d = parseDateValue(row[dateCol]);
            if (!d || !isFinite(d.getTime())) return null;

            // y is not used for event-date mode; keep a harmless constant
            const labelRaw = row[dateCol];
            const label =
              labelRaw !== undefined && labelRaw !== null && String(labelRaw).trim() !== ""
                ? String(labelRaw)
                : `Event ${idx + 1}`;

            return { x: d, y: 1, label, _rowIndex: idx };
          })
          .filter(Boolean);
      } else {
        // gaps mode: use numeric gaps directly from the value column (sequence axis is fine)
        parsedPoints = rawRows
          .map((row, idx) => {
            const gap = toNumericValue(row[valueCol]);
            if (!isFinite(gap)) return null;

            const rawLabel = row[dateCol];
            const label =
              rawLabel !== undefined && rawLabel !== null && String(rawLabel).trim() !== ""
                ? String(rawLabel)
                : `Point ${idx + 1}`;

            // Keep x as sequence index; draw step will treat y as the gap value
            return { x: idx, y: gap, label, _rowIndex: idx };
          })
          .filter(Boolean);
      }
    } else {
      // Normal behaviour for non-T charts
      if (axisType === "date") {
        parsedPoints = rawRows
          .map((row,idx) => {
            const d = parseDateValue(row[dateCol]);
            const y = toNumericValue(row[valueCol]);
            if (!d || !isFinite(d.getTime()) || !isFinite(y)) return null;
            return { x: d, y, _rowIndex: idx  };
          })
          .filter(Boolean);
      } else {
        // sequence/category axis
        parsedPoints = rawRows
          .map((row, idx) => {
            const y = toNumericValue(row[valueCol]);
            if (!isFinite(y)) return null;

            const rawLabel = row[dateCol];
            const label =
              rawLabel !== undefined &&
              rawLabel !== null &&
              String(rawLabel).trim() !== ""
                ? String(rawLabel)
                : `Point ${idx + 1}`;

            return { x: idx, y, label,  _rowIndex: idx };
          })
          .filter(Boolean);
      }
    }


    // You can lower this if you want charts from fewer points
    if (parsedPoints.length < 3) {
      showError("Not enough valid data points after parsing. Check your column choices.");
      return;
    }

    // --- 2) Create points + labels for the chart ---
    let points, labels;

    if (axisType === "date") {
      points = [...parsedPoints].sort((a, b) => a.x - b.x);
      labels = points.map((p) => formatYMD_Local(p.x));
    } else {
      points = parsedPoints;
      labels = points.map((p) => p.label);
    }

    // --- baseline interpretation ---
    let baselineCount = null;
    if (baselineInput && baselineInput.value.trim() !== "") {
      const n = parseInt(baselineInput.value, 10);
      if (!isNaN(n) && n >= 2) baselineCount = Math.min(n, points.length);
    }

// Guard: chart not implemented yet
if (!IMPLEMENTED_CHARTS.has(chartType)) {
  showChartMessage(`"${chartType.toUpperCase()}" charts are not available yet. Please use Run or XmR for now.`);
  return;
}

// If the 3rd column row is visible, ensure it’s selected sensibly
if (thirdColumnRow && thirdColumnRow.style.display !== "none") {
  const yCol = valueSelect.value;
  const thirdCol = thirdSelect ? thirdSelect.value : "";

  if (!thirdCol) {
    showChartMessage(`Please choose the required third column for this ${getChartTypeDisplayName(chartType)}. For example, use a denominator for a P chart or opportunities for a U chart.`);
    return;
  }
  if (thirdCol === yCol) {
    showChartMessage(`The third column is currently set to ${formatColumnName(thirdCol)} but it should be different from the main value column ${formatColumnName(yCol)}.`);
    return;
  }
}


    // clear existing charts
    if (currentChart) {
      currentChart.destroy();
      currentChart = null;
    }
    if (mrChart) {
      mrChart.destroy();
      mrChart = null;
    }
    if (mrPanel) mrPanel.style.display = "none";

// draw the selected chart
if (chartType === "run") {
  drawRunChart(points, baselineCount, labels);

} else if (chartType === "xmr") {
  drawXmRChart(points, baselineCount, labels);

} else if (chartType === "c") {
  // -----------------------------
  // VALIDATION: C chart (counts)
  // -----------------------------
  const cValues = points.map(p => p.y);
  const cColumn = valueSelect?.value || "";
  const cAlternative = getSuggestedAlternativeColumn(cColumn, getCountLikeColumnCandidates());

 if (!handleValidationResult(
      validateNonNegativeNumbers(cValues, "C chart count", {
        columnName: cColumn,
        alternativeColumn: cAlternative,
        purposeText: "count column"
      }),
      { manual: lastGenerateWasManual }
    )) return;

 if (!handleValidationResult(
      warnIfNonInteger(cValues, "C chart count", {
        columnName: cColumn,
        alternativeColumn: cAlternative,
        purposeText: "count column"
      }),
      { manual: lastGenerateWasManual }
    )) return;

  drawCChart(points, baselineCount, labels);

} else if (
  isPFamilyChartType(chartType) ||
  isUFamilyChartType(chartType)
) {
  const isPFamily =
    isPFamilyChartType(chartType);

  const isUFamily =
    isUFamilyChartType(chartType);

  // P/P′ and U/U′ require a third column
  // (denominator or opportunities).
  if (
    !thirdSelect ||
    !thirdSelect.value
  ) {
    showError(
      `The ${getChartTypeDisplayName(chartType)} needs a third column. ` +
      `Please choose ${
        isPFamily
          ? "a denominator (total)"
          : "an opportunities column"
      } before generating the chart.`
    );

    return;
  }

  const denomCol =
    thirdSelect.value;

  // Build points with denominator/opportunities
  // using the original source-row index.
  const pointsWithNOrdered =
    points.map((p, i) => {
      const row =
        rawRows[p._rowIndex];

      const numerator =
        toNumericValue(
          row[valueSelect.value]
        );

      const denom =
        toNumericValue(
          row[denomCol]
        );

      return {
        x: p.x,
        y: numerator,
        n: denom,
        label: labels[i],
        _rowIndex: p._rowIndex
      };
    });

  const numerArr =
    pointsWithNOrdered.map(
      p => p.y
    );

  const denomArr =
    pointsWithNOrdered.map(
      p => p.n
    );

  const currentNumer =
    valueSelect?.value || "";

  const currentDenom =
    thirdSelect?.value || "";

  const fallback =
    (
      typeof chooseDefaultsForChart ===
      "function"
    )
      ? chooseDefaultsForChart(
          chartType
        )
      : null;

  const fallbackNumer =
    fallback?.yCol || "";

  const fallbackDenom =
    fallback?.thirdCol || "";

  // Numerator/count must be numeric and non-negative.
  if (
    !handleValidationResult(
      validateNonNegativeNumbers(
        numerArr,
        isPFamily
          ? "P/P′ chart numerator (d)"
          : "U/U′ chart numerator (c)",
        {
          columnName:
            currentNumer,

          alternativeColumn:
            fallbackNumer &&
            fallbackNumer !== currentNumer
              ? fallbackNumer
              : "",

          purposeText:
            isPFamily
              ? "numerator column"
              : "count column"
        }
      )
    )
  ) {
    return;
  }

  const canTryFallback =
    fallbackNumer &&
    fallbackDenom &&
    fallbackNumer !== fallbackDenom &&
    (
      fallbackNumer !== currentNumer ||
      fallbackDenom !== currentDenom
    );

  // -------------------------------------------------
  // P / P′ only:
  // numerator cannot be greater than denominator.
  // -------------------------------------------------
  if (isPFamily) {
    const pPairValidation =
      validateNumeratorNotGreaterThanDenom(
        numerArr,
        denomArr,
        {
          numeratorColumn:
            currentNumer,

          denominatorColumn:
            currentDenom,

          alternativeNumerator:
            fallbackNumer &&
            fallbackNumer !== currentNumer
              ? fallbackNumer
              : "",

          alternativeDenominator:
            fallbackDenom &&
            fallbackDenom !== currentDenom
              ? fallbackDenom
              : ""
        }
      );

    if (pPairValidation) {
      if (
        canTryFallback &&
        valueSelect &&
        thirdSelect
      ) {
        valueSelect.value =
          fallbackNumer;

        thirdSelect.value =
          fallbackDenom;

        showChartMessage(
          "I changed the numerator and denominator columns to a more suitable pair for a P/P′ chart."
        );

        if (generateButton) {
          lastGenerateWasManual =
            false;

          generateButton.click();
        }

        return;
      }

      if (
        !handleValidationResult(
          pPairValidation
        )
      ) {
        return;
      }
    }
  }

  // -------------------------------------------------
  // U / U′ only:
  // counts may exceed opportunities, but the
  // opportunities must be positive and the selected
  // columns must be different.
  // -------------------------------------------------
  if (isUFamily) {
    const invalidDenominator =
      denomArr.some(
        d =>
          !Number.isFinite(d) ||
          d <= 0
      );

    const sameColumnChosen =
      currentNumer &&
      currentDenom &&
      currentNumer === currentDenom;

    if (
      (
        invalidDenominator ||
        sameColumnChosen
      ) &&
      canTryFallback &&
      valueSelect &&
      thirdSelect
    ) {
      valueSelect.value =
        fallbackNumer;

      thirdSelect.value =
        fallbackDenom;

      showChartMessage(
        "I changed the count and opportunities columns to a more suitable pair for a U/U′ chart."
      );

      if (generateButton) {
        lastGenerateWasManual =
          false;

        generateButton.click();
      }

      return;
    }
  }

  // Denominator / opportunities must be positive.
  if (
    !handleValidationResult(
      validateDenominatorPositive(
        denomArr,
        isPFamily
          ? "P/P′ chart denominator (n)"
          : "U/U′ chart denominator/opportunities (n)",
        {
          columnName:
            currentDenom,

          alternativeColumn:
            fallbackDenom &&
            fallbackDenom !== currentDenom
              ? fallbackDenom
              : "",

          purposeText:
            isPFamily
              ? "denominator column"
              : "opportunities column"
        }
      )
    )
  ) {
    return;
  }

  // Soft warning for non-integer numerator/count.
  if (
    !handleValidationResult(
      warnIfNonInteger(
        numerArr,
        isPFamily
          ? "P/P′ chart numerator (d)"
          : "U/U′ chart numerator (c)",
        {
          columnName:
            currentNumer,

          alternativeColumn:
            fallbackNumer &&
            fallbackNumer !== currentNumer
              ? fallbackNumer
              : "",

          purposeText:
            isPFamily
              ? "numerator column"
              : "count column"
        }
      ),
      {
        manual:
          lastGenerateWasManual
      }
    )
  ) {
    return;
  }

  // Soft warning for non-integer denominator/opportunities.
  if (
    !handleValidationResult(
      warnIfNonInteger(
        denomArr,
        isPFamily
          ? "P/P′ chart denominator (n)"
          : "U/U′ chart denominator/opportunities (n)",
        {
          columnName:
            currentDenom,

          alternativeColumn:
            fallbackDenom &&
            fallbackDenom !== currentDenom
              ? fallbackDenom
              : "",

          purposeText:
            isPFamily
              ? "denominator column"
              : "opportunities column"
        }
      ),
      {
        manual:
          lastGenerateWasManual
      }
    )
  ) {
    return;
  }

  // Draw the selected attribute chart.
  if (chartType === "p") {
    drawPChart(
      pointsWithNOrdered,
      baselineCount,
      labels
    );

  } else if (
    chartType === "pprime"
  ) {
    drawPPrimeChart(
      pointsWithNOrdered,
      baselineCount,
      labels
    );

  } else if (
    chartType === "u"
  ) {
    drawUChart(
      pointsWithNOrdered,
      baselineCount,
      labels
    );

  } else if (
    chartType === "uprime"
  ) {
    drawUPrimeChart(
      pointsWithNOrdered,
      baselineCount,
      labels
    );
  }

} else if (chartType === "xbars") {
  drawXbarSChart(points, baselineCount, labels);

} else if (chartType === "t") {
  if (tChartInputMode === "eventDates") {
    // T chart needs date axis (uses event dates)
    if (document.querySelector("input[name='axisType']:checked")?.value !== "date") {
      showError("T chart needs Date / time axis (it uses event dates).");
      return;
    }
    drawTChart(points, baselineCount, labels);
    } else {
    // Gaps mode: values are already "time between events"
    const gaps = points.map(p => p.y);
    if (gaps.length < 3) {
      showError("T chart (gaps mode) needs at least 3 valid gap values.");
      return;
    }

    // ---- Segment definition from splits ----
    let effectiveSplits = Array.isArray(splits) ? splits.slice() : [];
    effectiveSplits = effectiveSplits
      .filter(i => Number.isInteger(i) && i >= 0 && i < gaps.length - 1)
      .sort((a, b) => a - b);

    const segmentStarts = [0];
    const segmentEnds = [];
    effectiveSplits.forEach(idx => {
      segmentEnds.push(idx);
      segmentStarts.push(idx + 1);
    });
    segmentEnds.push(gaps.length - 1);

    const clArr = new Array(gaps.length).fill(NaN);
    const uclArr = new Array(gaps.length).fill(NaN);
    const lclArr = new Array(gaps.length).fill(NaN);

    // Build period-specific limits
    for (let s = 0; s < segmentStarts.length; s++) {
      const start = segmentStarts[s];
      const end = segmentEnds[s];

      const segGaps = gaps.slice(start, end + 1);

      const segBaselineCountUsed =
        (s === 0 && baselineCount && baselineCount >= 2)
          ? Math.min(baselineCount, segGaps.length)
          : segGaps.length;

      const base = segGaps.slice(0, segBaselineCountUsed);
      const cl = base.reduce((a, b) => a + b, 0) / base.length;

      const qHigh = 0.99865;
      const ucl = -cl * Math.log(1 - qHigh);
      const lcl = 0;

      for (let i = start; i <= end; i++) {
        clArr[i] = cl;
        uclArr[i] = ucl;
        lclArr[i] = lcl;
      }
    }

    const pointColours = gaps.map((v, i) =>
      (v > uclArr[i] || v < lclArr[i]) ? SPC_STYLE.pointBeyond : SPC_STYLE.seriesBlue
    );

    drawSimpleSPCChart({
      labels,
      values: gaps,
      pointColours,
      cl: clArr,
      ucl: uclArr,
      lcl: lclArr,
      yAxisSuggestedMin: 0,
      yAxisSuggestedMax: Math.max(...gaps, ...uclArr.filter(isFinite)),
      chartTitleFallback: chartTitleInput?.value || "T chart (gaps)",
      yAxisLabelFallback: yAxisLabelInput?.value || "Time between events",
      showUCL: true,
      showLCL: false
    });

    // ---- Build per-period analyses for summary ----
    const segmentAnalyses = [];

    for (let s = 0; s < segmentStarts.length; s++) {
      const start = segmentStarts[s];
      const end = segmentEnds[s];

      const segBaselineCountUsed =
        (s === 0 && baselineCount && baselineCount >= 2)
          ? Math.min(baselineCount, (end - start + 1))
          : (end - start + 1);

      const a = analyzeRareChart({
        chartType: "t",
        labels: labels.slice(start, end + 1),
        values: gaps.slice(start, end + 1),
        cl: clArr.slice(start, end + 1),
        ucl: uclArr.slice(start, end + 1),
        lcl: lclArr.slice(start, end + 1)
      });

      a.periodIndex = s + 1;
      a.periodCount = segmentStarts.length;
      a.startIndex = start;
      a.endIndex = end;
      a.labelStart = labels[start];
      a.labelEnd = labels[end];
      a.nPoints = (end - start + 1);
      a.baselineCountUsed = segBaselineCountUsed;

      a.stats = {
        cl: Number(clArr[start]),
        ucl: Number(uclArr[start]),
        lcl: Number(lclArr[start])
      };

      a.totalPoints = gaps.length;

      segmentAnalyses.push(a);
    }

    lastRareAnalysis = segmentAnalyses[segmentAnalyses.length - 1];
    renderRareChartSummary(segmentAnalyses, gaps.length);
  }

} else if (chartType === "g") {
  // drawGChart expects a numeric array of values (not {x,y} point objects)
  const gValues = points.map(p => p.y);
  drawGChart(gValues, baselineCount, labels);

} else {
  showError(`Chart type "${chartType}" is not implemented yet.`);
  return;
}

    appendLimitExclusionSummary();

    // Keep Y-axis inputs blank when the user has not manually set bounds.
// Blank inputs mean "automatic range", which prevents old auto-bounds
// from becoming fixed limits on the next redraw.
if (yAxisBoundsManuallyEdited) {
  applyCurrentChartYBoundsToInputs(currentChart);
}

updateYAxisInputStep();

    // optional: clear dirty flag after successful draw
    if (typeof clearDataModelDirty === "function") clearDataModelDirty();
  } finally {
    // Always re-render helper UI state (even if chart drawing throws)
    if (typeof renderHelperState === "function") renderHelperState();

    // Hide quick-start once a chart exists (robust even if localStorage is blocked)
    if (currentChart) {
      if (typeof markFirstRunComplete === "function") {
        markFirstRunComplete();
      } else {
        const guide = document.getElementById("firstRunGuide");
        if (guide) guide.style.display = "none";
      }
    }
  }
});

// ---- Chart drawing ----

function drawRunChart(points, baselineCount, labels) {
  if (!chartCanvas) return;

  const n = points.length;

  // ---- Read “rules & interpretation” settings (safe fallbacks) ----
  const { shiftLength, trendLength } =
    (typeof getRuleSettings === "function")
      ? getRuleSettings()
      : { shiftLength: 8, trendLength: 6 };

  const flagOnChart =
    (typeof shouldFlagSpecialCauseOnChart === "function")
      ? shouldFlagSpecialCauseOnChart()
      : true;

  // ---- Keep dropdowns in sync ----
  populateAnnotationDateOptions(labels);
  if (typeof populateSplitOptions === "function") {
    populateSplitOptions(labels);
  }

  // ---- Segment definition from splits (same pattern as XmR) ----
  let effectiveSplits = Array.isArray(splits) ? splits.slice() : [];
  effectiveSplits = effectiveSplits
    .filter(i => Number.isInteger(i) && i >= 0 && i < n - 1)
    .sort((a, b) => a - b);

  const segmentStarts = [0];
  const segmentEnds = [];
  effectiveSplits.forEach(idx => {
    segmentEnds.push(idx);
    segmentStarts.push(idx + 1);
  });
  segmentEnds.push(n - 1);

  const values = points.map(p => p.y);

  // ---- Build piecewise median line + colours ----
  const medianLine = new Array(n).fill(NaN);
  const pointColours = new Array(n).fill(SPC_STYLE.seriesBlue);

  // Collect rule hits (optional – useful if your summary wants it)
  const runRangesAll = [];
  const trendRangesAll = [];

  // Baseline applies only to first segment; later segments use full segment
  for (let s = 0; s < segmentStarts.length; s++) {
    const start = segmentStarts[s];
    const end = segmentEnds[s];
    const segPoints = points.slice(start, end + 1);
    const segValues = segPoints.map(p => p.y);

    // baselineCount logic (only first segment honours baselineCount)
    let segBaselineCountUsed;
    if (s === 0 && baselineCount && baselineCount >= 2) {
      segBaselineCountUsed = Math.min(baselineCount, segPoints.length);
    } else {
      segBaselineCountUsed = segPoints.length;
    }

    const segBaselineValues = segValues.slice(0, segBaselineCountUsed);
    const segMedian = computeMedian(segBaselineValues);

    for (let i = start; i <= end; i++) {
      medianLine[i] = segMedian;
    }

    // Rule detection per segment
    const localRunRanges =
      (typeof findLongRunRanges === "function")
        ? findLongRunRanges(segValues, segMedian, shiftLength)
        : [];

    const localTrendRanges =
      (typeof findTrendRanges === "function")
        ? findTrendRanges(segValues, trendLength)
        : [];

    // Convert local ranges into global indices for summary use
    localRunRanges.forEach(r => runRangesAll.push({
      start: r.start + start,
      end: r.end + start,
      len: r.len,
      side: r.side
    }));

    localTrendRanges.forEach(r => trendRangesAll.push({
      start: r.start + start,
      end: r.end + start,
      len: r.len,
      direction: r.direction
    }));

    // Colour flags (per point)
    const runFlags =
      (typeof flagFromRanges === "function")
        ? flagFromRanges(segValues.length, localRunRanges)
        : new Array(segValues.length).fill(false);

    const trendFlags =
      (typeof flagFromRanges === "function")
        ? flagFromRanges(segValues.length, localTrendRanges)
        : new Array(segValues.length).fill(false);

    for (let i = 0; i < segValues.length; i++) {
      const globalIdx = start + i;
      if (flagOnChart && (runFlags[i] || trendFlags[i])) {
        pointColours[globalIdx] = SPC_STYLE.pointSpecial;
      }
    }
  }

  const { title, xLabel, yLabel } = getChartLabels("Run Chart", "Date", "Value");
  const target = getTargetValue();

  const datasets = [
    {
      label: "Value",
      data: values,
      pointRadius: 4,
      pointBackgroundColor: pointColours,
      borderColor: SPC_STYLE.seriesBlue,
      borderWidth: 2,
      fill: false
    },
    {
      label: "Median",
      data: medianLine,
      borderDash: [6, 4],
      borderWidth: 2,
      borderColor: SPC_STYLE.centreRed,
      pointRadius: 0,
      pointHoverRadius: 0,
      fill: false
    }
  ];

  if (targetEnabled &&  target !== null) {
    datasets.push({
      label: "Target",
      data: values.map(() => target),
      borderDash: [4, 2],
      borderWidth: 2,
      borderColor: SPC_STYLE.targetOrange,
      pointRadius: 0,
      pointHoverRadius: 0,
      fill: false
    });
  }

  const centreLineLabelPlugin =
  buildXmRMeanLabelPlugin();

currentChart = new Chart(chartCanvas, {
  type: "line",
  data: { labels, datasets },

  plugins: [
    centreLineLabelPlugin
  ],

  options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        title: buildChartTitleConfig(title),
        legend: SPC_LEGEND,
        annotation: { annotations: buildAnnotationConfig(labels) }
      },
      elements: { point: { radius: 0, hoverRadius: 0 } },
            scales: (() => {
        const axisSettings = getAxisSettings();
        return {
          x: buildCategoryXAxisConfig(xLabel, axisSettings.x, labels),
          y: buildAxisConfig(yLabel, axisSettings.y)
        };
      })()
    }
  });

    clearDataModelDirty();

  // ----- Build summary inputs safely -----
  // Use a simple "overall baseline" for summary fields (the chart median line is still piecewise)
  const baselineCountUsed =
    (baselineCount && baselineCount >= 2)
      ? Math.min(baselineCount, values.length)
      : values.length;

  const median = (typeof computeMedian === "function")
    ? computeMedian(values.slice(0, baselineCountUsed))
    : null;

  // Use the arrays you actually built earlier in drawRunChart
  const runRanges = runRangesAll;       // <- IMPORTANT: these exist
  const trendRanges = trendRangesAll;   // <- IMPORTANT: these exist

  // If you don't calculate astronomical points here, keep a safe empty structure
  const astro = { indices: [], flags: [] };

  // Update summary
  const ruleHits = { runRanges, trendRanges, astro };

  if (typeof updateRunSummary === "function") {
    updateRunSummary(points, median, ruleHits, baselineCountUsed);
  }

  // Store a structured snapshot for the SPC helper (Run chart)
  lastRunAnalysis = {
    chartType: "run",
    n,
    median,
    baselineCountUsed,
    shiftLength,
    trendLength,
    runRanges,
    trendRanges,
    astro,
    splits: Array.isArray(splits) ? splits.slice() : [],
    hasShift: Array.isArray(runRanges) && runRanges.length > 0,
    hasTrend: Array.isArray(trendRanges) && trendRanges.length > 0,
    hasAstronomical: false,
    isStable: !(runRanges.length || trendRanges.length)
  };

  // If helper is open, refresh chips/intro to match the current chart
  if (spcHelperPanel && spcHelperPanel.classList.contains("visible")) {
    if (typeof renderHelperState === "function") renderHelperState();
  }
}


// -----------------------------
// Pareto chart
// -----------------------------

function computeParetoData(
  rows,
  categoryCol,
  countCol
) {
  if (
    !Array.isArray(rows) ||
    !rows.length
  ) {
    return {
      ok: false,
      error:
        "No data are available for the Pareto chart."
    };
  }

  if (!categoryCol) {
    return {
      ok: false,
      error:
        "Please choose a category column for the Pareto chart."
    };
  }

  if (!countCol) {
    return {
      ok: false,
      error:
        "Please choose how the Pareto frequencies should be counted."
    };
  }

  const countRowsMode =
    countCol ===
    PARETO_COUNT_ROWS;

  const totals =
    new Map();

  let blankCategoryRows = 0;
  let invalidCountRows = 0;
  let sourceRowsUsed = 0;

  for (
    const row of rows
  ) {
    const rawCategory =
      row?.[categoryCol];

    const category =
      rawCategory === null ||
      rawCategory === undefined
        ? ""
        : String(
            rawCategory
          ).trim();

    if (!category) {
      blankCategoryRows += 1;
      continue;
    }

    let amount = 1;

    if (!countRowsMode) {
      amount =
        toNumericValue(
          row?.[countCol]
        );

      if (
        !Number.isFinite(amount) ||
        amount < 0 ||
        !Number.isInteger(amount)
      ) {
        invalidCountRows += 1;
        continue;
      }
    }

    sourceRowsUsed += 1;

    totals.set(
      category,
      (
        totals.get(category) ||
        0
      ) +
      amount
    );
  }

  if (
    invalidCountRows > 0
  ) {
    return {
      ok: false,

      error:
        `${invalidCountRows} row${
          invalidCountRows === 1
            ? ""
            : "s"
        } contain an invalid frequency/count value. ` +
        "A classic Pareto chart needs non-negative whole-number counts. " +
        "Please correct those values or choose '(count each row as 1)'."
    };
  }

  let items =
    Array.from(
      totals.entries()
    )
      .map(
        ([category, count]) => ({
          category,
          count
        })
      )
      .filter(
        item =>
          Number.isFinite(
            item.count
          ) &&
          item.count > 0
      );

  items.sort(
    (a, b) => {
      if (
        b.count !==
        a.count
      ) {
        return (
          b.count -
          a.count
        );
      }

      return String(
        a.category
      ).localeCompare(
        String(
          b.category
        )
      );
    }
  );

  if (
    items.length < 2
  ) {
    return {
      ok: false,

      error:
        "A Pareto chart needs at least 2 categories with occurrences."
    };
  }

  const total =
    items.reduce(
      (sum, item) =>
        sum +
        item.count,
      0
    );

  if (
    !Number.isFinite(total) ||
    total <= 0
  ) {
    return {
      ok: false,
      error:
        "No positive occurrence counts were found for the Pareto chart."
    };
  }

  let running = 0;

  items =
    items.map(
      item => {
        running +=
          item.count;

        return {
          ...item,

          share:
            item.count /
            total,

          cumulativeCount:
            running,

          cumulativePct:
            (
              running /
              total
            ) *
            100
        };
      }
    );

  const cutoffIndex =
    items.findIndex(
      item =>
        item.cumulativePct >=
        80
    );

  const cutoff80Count =
    cutoffIndex >= 0
      ? cutoffIndex + 1
      : items.length;

  const cutoff80Pct =
    items[
      cutoff80Count - 1
    ]?.cumulativePct ??
    100;

  const top3Count =
    items
      .slice(0, 3)
      .reduce(
        (sum, item) =>
          sum +
          item.count,
        0
      );

  return {
    ok: true,

    chartType:
      "pareto",

    items,
    total,

    categoryCount:
      items.length,

    cutoff80Count,
    cutoff80Pct,

    top3Count,

    top3Share:
      top3Count /
      total,

    countRowsMode,

    sourceRowsUsed,
    blankCategoryRows
  };
}

function drawParetoChart(
  rows,
  categoryCol,
  countCol
) {
  if (!chartCanvas) {
    return;
  }

  const result =
    computeParetoData(
      rows,
      categoryCol,
      countCol
    );

  if (!result.ok) {
    showError(
      result.error
    );

    return;
  }

  const labels =
    result.items.map(
      item =>
        item.category
    );

  const counts =
    result.items.map(
      item =>
        item.count
    );

  const cumulative =
    result.items.map(
      item =>
        item.cumulativePct
    );

  if (
    typeof populateAnnotationDateOptions ===
    "function"
  ) {
    populateAnnotationDateOptions(
      labels
    );
  }

  const chartLabels =
    getChartLabels(
      "Pareto chart",
      "Category",
      "Count"
    );

  const axisSettings =
    getAxisSettings();

  currentChart =
    new Chart(
      chartCanvas.getContext(
        "2d"
      ),
      {
        data: {
          labels,

          datasets: [
            {
              type:
                "bar",

              label:
                "Count",

              data:
                counts,

              yAxisID:
                "y",

              backgroundColor:
                SPC_STYLE.seriesBlue,

              borderColor:
                SPC_STYLE.seriesBlue,

              borderWidth:
                1
            },

            {
              type:
                "line",

              label:
                "Cumulative %",

              data:
                cumulative,

              yAxisID:
                "y1",

              borderColor:
                SPC_STYLE.centreRed,

              backgroundColor:
                SPC_STYLE.centreRed,

              borderWidth:
                2,

              pointRadius:
                3,

              pointHoverRadius:
                5,

              fill:
                false,

              tension:
                0
            }
          ]
        },

        options: {
          responsive: true,
          maintainAspectRatio: false,

          interaction: {
            mode: "index",
            intersect: false
          },

          plugins: {
            title:
              buildChartTitleConfig(
                chartLabels.title
              ),

            legend: {
              display: true,
              position:
                "bottom",
              align:
                "center"
            },

            tooltip: {
              callbacks: {
                label(context) {
                  const i =
                    context.dataIndex;

                  const item =
                    result.items[i];

                  if (!item) {
                    return "";
                  }

                  if (
                    context.dataset
                      .yAxisID ===
                    "y1"
                  ) {
                    return (
                      `Cumulative: ` +
                      `${item.cumulativePct.toFixed(1)}%`
                    );
                  }

                  return (
                    `Count: ${item.count} ` +
                    `(${(item.share * 100).toFixed(1)}% of total)`
                  );
                },

                afterBody(items) {
                  const i =
                    items?.[0]
                      ?.dataIndex;

                  const item =
                    result.items[i];

                  if (!item) {
                    return [];
                  }

                  return [
                    `Cumulative through this category: ${item.cumulativePct.toFixed(1)}%`
                  ];
                }
              }
            },

            annotation: {
              annotations: {
                pareto80: {
                  type:
                    "line",

                  yScaleID:
                    "y1",

                  yMin:
                    80,

                  yMax:
                    80,

                  borderColor:
                    "rgba(90,90,90,0.65)",

                  borderWidth:
                    1,

                  borderDash:
                    [5, 5]
                }
              }
            }
          },

          scales: {
            x:
              buildCategoryXAxisConfig(
                chartLabels.xLabel,
                axisSettings.x,
                labels
              ),

            y:
              buildAxisConfig(
                chartLabels.yLabel,
                axisSettings.y,
                {
                  beginAtZero:
                    true
                }
              ),

            y1: {
              type:
                "linear",

              position:
                "right",

              min:
                0,

              max:
                100,

              grid: {
                drawOnChartArea:
                  false
              },

              title: {
                display:
                  true,

                text:
                  "Cumulative %"
              },

              ticks: {
                callback(value) {
                  return (
                    `${value}%`
                  );
                }
              }
            }
          }
        }
      }

    );

  // Pareto uses Chart.js annotations for its 80% guide,
  // but does not support draggable user annotations.
  currentChart.$disableUserAnnotations = true;

  const top =
    result.items[0];


  let html =
    `<h3>Summary (Pareto chart)</h3>`;

  html +=
    `<p>` +
    `Total occurrences: ` +
    `<strong>${result.total}</strong> across ` +
    `<strong>${result.categoryCount}</strong> categories.` +
    `</p>`;

  html +=
    `<ul>`;

  html +=
    `<li>` +
    `<strong>Largest category:</strong> ` +
    `${escapeHtml(top.category)} — ` +
    `${top.count} occurrence${top.count === 1 ? "" : "s"} ` +
    `(${(top.share * 100).toFixed(1)}% of the total).` +
    `</li>`;

  html +=
    `<li>` +
    `<strong>Top three categories:</strong> ` +
    `${result.top3Count} occurrences ` +
    `(${(result.top3Share * 100).toFixed(1)}% of the total).` +
    `</li>`;

  html +=
    `<li>` +
    `<strong>80% guide:</strong> ` +
    `the first ${result.cutoff80Count} categor${
      result.cutoff80Count === 1
        ? "y"
        : "ies"
    } account for ` +
    `<strong>${result.cutoff80Pct.toFixed(1)}%</strong> ` +
    `of the occurrences. ` +
    `The dashed 80% line is a prioritisation guide, not a statistical control limit.` +
    `</li>`;

  if (
    result.total < 30
  ) {
    html +=
      `<li>` +
      `<strong>Caution:</strong> ` +
      `this Pareto chart contains fewer than 30 occurrences. ` +
      `With small numbers, the apparent ranking can be unstable, so interpret it cautiously.` +
      `</li>`;
  }

  if (
    result.blankCategoryRows >
    0
  ) {
    html +=
      `<li>` +
      `<strong>Data check:</strong> ` +
      `${result.blankCategoryRows} row${
        result.blankCategoryRows === 1
          ? " was"
          : "s were"
      } ignored because the category was blank.` +
      `</li>`;
  }

  html +=
    `<li>` +
    `<strong>Interpretation:</strong> ` +
    `the categories on the left are the most frequent and may be useful places to focus further investigation or improvement effort. ` +
    `A category is not automatically a cause of the problem.` +
    `</li>`;

  html +=
    `<li>` +
    `<strong>SPC note:</strong> ` +
    `a Pareto chart is not time ordered and cannot tell you whether variation is common cause or special cause. ` +
    `Where relevant, interpret it alongside an appropriate run or Shewhart chart.` +
    `</li>`;

  html +=
    `</ul>`;

  summaryDiv.innerHTML =
    html;

  lastParetoAnalysis = {
    ...result,

    categoryColumn:
      categoryCol,

    countColumn:
      countCol,

    topCategory:
      top.category,

    topCount:
      top.count,

    topShare:
      top.share
  };
}


// -----------------------------
// Draw C / P / U charts
// -----------------------------

function drawCChart(points, baselineCount, labels) {
  points =
    markLimitExclusionsOnPoints(
      "c",
      points
    );
  if (!chartCanvas) return;

  const n = points.length;
  if (n < 2) {
    showError("C chart needs at least 2 points.");
    return;
  }

  // ---- Segment definition from splits ----
  let effectiveSplits = Array.isArray(splits) ? splits.slice() : [];
  effectiveSplits = effectiveSplits
    .filter(i => Number.isInteger(i) && i >= 0 && i < n - 1)
    .sort((a, b) => a - b);

  const segmentStarts = [0];
  const segmentEnds = [];
  effectiveSplits.forEach(idx => { segmentEnds.push(idx); segmentStarts.push(idx + 1); });
  segmentEnds.push(n - 1);

  const values = points.map(p => p.y);

  const clArr = new Array(n).fill(NaN);
  const uclArr = new Array(n).fill(NaN);
  const lclArr = new Array(n).fill(NaN);
  const beyond = new Array(n).fill(false);

  for (let s = 0; s < segmentStarts.length; s++) {
    const start = segmentStarts[s];
    const end = segmentEnds[s];

    const segPoints = points.slice(start, end + 1);

    // baselineCount only applies to first segment; later segments use their whole segment
    const segBaselineCountUsed =
      (s === 0 && baselineCount && baselineCount >= 1)
        ? Math.min(baselineCount, segPoints.length)
        : segPoints.length;

    const res = computeC(segPoints, segBaselineCountUsed);

    for (let i = start; i <= end; i++) {
      clArr[i] = res.cbar;
      uclArr[i] = res.ucl;
      lclArr[i] = res.lcl;
      beyond[i] =
  !points[i]?._excludeFromLimits &&
  isFinite(values[i]) &&
  (
    values[i] > res.ucl ||
    values[i] < res.lcl
  );
    }
  }

  const pointColours = values.map((v, i) => (beyond[i] ? SPC_STYLE.pointBeyond : SPC_STYLE.seriesBlue));

  drawSimpleSPCChart({
    labels,
    values,
    pointColours,
    cl: clArr,
    ucl: uclArr,
    lcl: lclArr,
    chartTitleFallback: "C chart",
    yAxisLabelFallback: "Count",
    showUCL: true,
    showLCL: true
  });

  attachRowLimitExclusionContext(
    currentChart,
    "c",
    points
  );

    // ---- Multi-period analysis (ALL segments) ----
  const analyses = [];

  for (let s = 0; s < segmentStarts.length; s++) {
    const start = segmentStarts[s];
    const end = segmentEnds[s];

    const segPoints = points.slice(start, end + 1);

    // baselineCount only applies to first segment; later segments use their whole segment
    const segBaselineCountUsed =
      (s === 0 && baselineCount && baselineCount >= 1)
        ? Math.min(baselineCount, segPoints.length)
        : segPoints.length;

    const a = analyzeAttributeChart({
      chartType: "c",
	excludedIndices:
  getMarkedExcludedIndices(
    segPoints
  ),
      labels: labels.slice(start, end + 1),
      values: values.slice(start, end + 1),
      cl: clArr.slice(start, end + 1),
      ucl: uclArr.slice(start, end + 1),
      lcl: lclArr.slice(start, end + 1)
    });

    const segValues = values.slice(start, end + 1);
    const segCL = clArr.slice(start, end + 1);
    const segUCL = uclArr.slice(start, end + 1);
    const segLCL = lclArr.slice(start, end + 1);

    a.stats = {
      cl: segCL.find(v => Number.isFinite(v)),     // c̄ (constant within period)
      ucl: segUCL.find(v => Number.isFinite(v)),   // constant within period
      lcl: segLCL.find(v => Number.isFinite(v))    // constant within period
    };


    // Context (so the summary can mirror XmR style)
    a.periodIndex = s + 1;
    a.periodCount = segmentStarts.length;
    a.startIndex = start;
    a.endIndex = end;
    a.labelStart = labels[start];
    a.labelEnd = labels[end];
    a.baselineCountUsed = segBaselineCountUsed;

    analyses.push(a);
  }

  // Render XmR-style multi-period summary
  renderAttributeMultiSummary(analyses, labels.length);

  // Keep "latest" available for anything else that expects it
  lastAttributeAnalysis = analyses[analyses.length - 1];

}


function drawPChart(pointsWithN, baselineCount, labels) {
  pointsWithN =
    markLimitExclusionsOnPoints(
      "p",
      pointsWithN
    );
  if (!chartCanvas) return;

  const n = pointsWithN.length;
  if (n < 2) {
    showError("P chart needs at least 2 points.");
    return;
  }

  const clampLcl =
    (typeof shouldClampLclAtZero === "function")
      ? shouldClampLclAtZero()
      : true;

  // ---- Segment definition from splits ----
  let effectiveSplits = Array.isArray(splits) ? splits.slice() : [];
  effectiveSplits = effectiveSplits
    .filter(i => Number.isInteger(i) && i >= 0 && i < n - 1)
    .sort((a, b) => a - b);

  const segmentStarts = [0];
  const segmentEnds = [];
  effectiveSplits.forEach(idx => { segmentEnds.push(idx); segmentStarts.push(idx + 1); });
  segmentEnds.push(n - 1);

  const values = new Array(n).fill(NaN);
  const clArr = new Array(n).fill(NaN);
  const uclArr = new Array(n).fill(NaN);
  const lclArr = new Array(n).fill(NaN);
  const beyond = new Array(n).fill(false);

  for (let s = 0; s < segmentStarts.length; s++) {
    const start = segmentStarts[s];
    const end = segmentEnds[s];

    const segPoints = pointsWithN.slice(start, end + 1);

    const segBaselineCountUsed =
      (s === 0 && baselineCount && baselineCount >= 1)
        ? Math.min(baselineCount, segPoints.length)
        : segPoints.length;

    const res = computeP(segPoints, segBaselineCountUsed, clampLcl);

    for (let j = 0; j < segPoints.length; j++) {
      const i = start + j;
      values[i] = res.pVals[j];
      clArr[i] = res.pbar;
      uclArr[i] = res.ucl[j];
      lclArr[i] = res.lcl[j];
      beyond[i] = res.beyond[j];
    }
  }

  const pointColours = values.map((v, i) => (beyond[i] ? SPC_STYLE.pointBeyond : SPC_STYLE.seriesBlue));

  drawSimpleSPCChart({
    labels,
    values,
    pointColours,
    cl: clArr,
    ucl: uclArr,
    lcl: lclArr,
    chartTitleFallback: "P chart",
    yAxisLabelFallback: "Proportion / %",
    showUCL: true,
    showLCL: true
  });

  attachRowLimitExclusionContext(
    currentChart,
    "p",
    pointsWithN
  );

    // ---- Multi-period analysis (ALL segments) ----
  const analyses = [];

  for (let s = 0; s < segmentStarts.length; s++) {
    const start = segmentStarts[s];
    const end = segmentEnds[s];

    const segPoints = pointsWithN.slice(start, end + 1);

    const segBaselineCountUsed =
      (s === 0 && baselineCount && baselineCount >= 1)
        ? Math.min(baselineCount, segPoints.length)
        : segPoints.length;

    const a = analyzeAttributeChart({
      chartType: "p",
	excludedIndices:
  getMarkedExcludedIndices(
    segPoints
  ),
      labels: labels.slice(start, end + 1),
      values: values.slice(start, end + 1),
      cl: clArr.slice(start, end + 1),
      ucl: uclArr.slice(start, end + 1),
      lcl: lclArr.slice(start, end + 1)
    });

const segValues = values.slice(start, end + 1);
const segCL = clArr.slice(start, end + 1);
const segUCL = uclArr.slice(start, end + 1);
const segLCL = lclArr.slice(start, end + 1);

const uRange = rangeFinite(segUCL);
const lRange = rangeFinite(segLCL);

const dispersionDiagnostic =
  computePDispersionDiagnostic(
    segPoints,
    segBaselineCountUsed
  );

a.stats = {
  cl:
    segCL.find(
      v => Number.isFinite(v)
    ),

  uclMin: uRange.min,
  uclMax: uRange.max,
  uclAvg: meanFinite(segUCL),

  lclMin: lRange.min,
  lclMax: lRange.max,
  lclAvg: meanFinite(segLCL),

  dispersionDiagnostic
};



    // Context (XmR-style)
    a.periodIndex = s + 1;
    a.periodCount = segmentStarts.length;
    a.startIndex = start;
    a.endIndex = end;
    a.labelStart = labels[start];
    a.labelEnd = labels[end];
    a.baselineCountUsed = segBaselineCountUsed;

    analyses.push(a);
  }

  // Render XmR-style multi-period summary
  renderAttributeMultiSummary(analyses, labels.length);

  // Keep "latest" available for anything else that expects it
  lastAttributeAnalysis = analyses[analyses.length - 1];

}

function drawPPrimeChart(pointsWithN, baselineCount, labels) {
  pointsWithN =
    markLimitExclusionsOnPoints(
      "pprime",
      pointsWithN
    );
  if (!chartCanvas) return;

  const n = pointsWithN.length;

  if (n < 2) {
    showError("P′ chart needs at least 2 points.");
    return;
  }

  // ---- Segment definition from splits ----
  let effectiveSplits =
    Array.isArray(splits) ? splits.slice() : [];

  effectiveSplits = effectiveSplits
    .filter(
      i =>
        Number.isInteger(i) &&
        i >= 0 &&
        i < n - 1
    )
    .sort((a, b) => a - b);

  const segmentStarts = [0];
  const segmentEnds = [];

  effectiveSplits.forEach(idx => {
    segmentEnds.push(idx);
    segmentStarts.push(idx + 1);
  });

  segmentEnds.push(n - 1);

  const values =
    new Array(n).fill(NaN);

  const clArr =
    new Array(n).fill(NaN);

  const uclArr =
    new Array(n).fill(NaN);

  const lclArr =
    new Array(n).fill(NaN);

  const beyond =
    new Array(n).fill(false);

  const segmentResults = [];

  // ---- Calculate each period ----
  for (let s = 0; s < segmentStarts.length; s++) {
    const start = segmentStarts[s];
    const end = segmentEnds[s];

    const segPoints =
      pointsWithN.slice(start, end + 1);

    const segBaselineCountUsed =
      (
        s === 0 &&
        baselineCount &&
        baselineCount >= 1
      )
        ? Math.min(
            baselineCount,
            segPoints.length
          )
        : segPoints.length;

    const res =
      computePPrime(
        segPoints,
        segBaselineCountUsed
      );

    segmentResults.push(res);

    for (let j = 0; j < segPoints.length; j++) {
      const i = start + j;

      values[i] =
        res.pVals[j];

      clArr[i] =
        res.pbar;

      uclArr[i] =
        res.ucl[j];

      lclArr[i] =
        res.lcl[j];

      beyond[i] =
        res.beyond[j];
    }
  }

  const pointColours =
    values.map(
      (v, i) =>
        beyond[i]
          ? SPC_STYLE.pointBeyond
          : SPC_STYLE.seriesBlue
    );

  drawSimpleSPCChart({
    labels,
    values,
    pointColours,
    cl: clArr,
    ucl: uclArr,
    lcl: lclArr,
    chartTitleFallback: "P′ chart",
    yAxisLabelFallback: "Proportion / %",
    showUCL: true,
    showLCL: true
  });

  attachRowLimitExclusionContext(
    currentChart,
    "pprime",
    pointsWithN
  );

  // ---- Multi-period analysis ----
  const analyses = [];

  for (let s = 0; s < segmentStarts.length; s++) {
    const start = segmentStarts[s];
    const end = segmentEnds[s];

    const segPoints =
      pointsWithN.slice(start, end + 1);

    const segBaselineCountUsed =
      (
        s === 0 &&
        baselineCount &&
        baselineCount >= 1
      )
        ? Math.min(
            baselineCount,
            segPoints.length
          )
        : segPoints.length;

    const res =
      segmentResults[s];

    const a =
      analyzeAttributeChart({
        chartType: "pprime",
	excludedIndices:
  getMarkedExcludedIndices(
    segPoints
  ),
        labels:
          labels.slice(start, end + 1),
        values:
          values.slice(start, end + 1),
        cl:
          clArr.slice(start, end + 1),
        ucl:
          uclArr.slice(start, end + 1),
        lcl:
          lclArr.slice(start, end + 1)
      });

    const segCL =
      clArr.slice(start, end + 1);

    const segUCL =
      uclArr.slice(start, end + 1);

    const segLCL =
      lclArr.slice(start, end + 1);

    const uRange =
      rangeFinite(segUCL);

    const lRange =
      rangeFinite(segLCL);

    a.stats = {
      cl:
        segCL.find(
          v => Number.isFinite(v)
        ),

      uclMin: uRange.min,
      uclMax: uRange.max,
      uclAvg: meanFinite(segUCL),

      lclMin: lRange.min,
      lclMax: lRange.max,
      lclAvg: meanFinite(segLCL),

      sigmaZ: res?.sigmaZ,
      avgMR: res?.avgMR
    };

    a.periodIndex = s + 1;
    a.periodCount =
      segmentStarts.length;

    a.startIndex = start;
    a.endIndex = end;

    a.labelStart =
      labels[start];

    a.labelEnd =
      labels[end];

    a.baselineCountUsed =
      segBaselineCountUsed;

    analyses.push(a);
  }

  renderAttributeMultiSummary(
    analyses,
    labels.length
  );

  lastAttributeAnalysis =
    analyses[
      analyses.length - 1
    ];
}

function drawUChart(pointsWithN, baselineCount, labels) {
  pointsWithN =
    markLimitExclusionsOnPoints(
      "u",
      pointsWithN
    );
  if (!chartCanvas) return;

  const n = pointsWithN.length;
  if (n < 2) {
    showError("U chart needs at least 2 points.");
    return;
  }

  const rateMultiplier = getURateMultiplier();

  const clampLcl =
    (typeof shouldClampLclAtZero === "function")
      ? shouldClampLclAtZero()
      : true;

  // ---- Segment definition from splits ----
  let effectiveSplits = Array.isArray(splits) ? splits.slice() : [];
  effectiveSplits = effectiveSplits
    .filter(i => Number.isInteger(i) && i >= 0 && i < n - 1)
    .sort((a, b) => a - b);

  const segmentStarts = [0];
  const segmentEnds = [];
  effectiveSplits.forEach(idx => { segmentEnds.push(idx); segmentStarts.push(idx + 1); });
  segmentEnds.push(n - 1);

  const values = new Array(n).fill(NaN);
  const clArr = new Array(n).fill(NaN);
  const uclArr = new Array(n).fill(NaN);
  const lclArr = new Array(n).fill(NaN);
  const beyond = new Array(n).fill(false);

  for (let s = 0; s < segmentStarts.length; s++) {
    const start = segmentStarts[s];
    const end = segmentEnds[s];

    const segPoints = pointsWithN.slice(start, end + 1);

    const segBaselineCountUsed =
      (s === 0 && baselineCount && baselineCount >= 1)
        ? Math.min(baselineCount, segPoints.length)
        : segPoints.length;

    const res = computeU(segPoints, segBaselineCountUsed, clampLcl);

    for (let j = 0; j < segPoints.length; j++) {
      const i = start + j;
      // computeU() remains in the natural "per 1" scale.
// Apply the user's multiplier only to displayed values.
values[i] = res.uVals[j] * rateMultiplier;
clArr[i] = res.ubar * rateMultiplier;
uclArr[i] = res.ucl[j] * rateMultiplier;
lclArr[i] = res.lcl[j] * rateMultiplier;

beyond[i] = res.beyond[j];
    }
  }

  const pointColours = values.map((v, i) => (beyond[i] ? SPC_STYLE.pointBeyond : SPC_STYLE.seriesBlue));

  drawSimpleSPCChart({
    labels,
    values,
    pointColours,
    cl: clArr,
    ucl: uclArr,
    lcl: lclArr,
    chartTitleFallback: "U chart",
    yAxisLabelFallback:
  rateMultiplier === 1
    ? "Rate per opportunity"
    : `Rate per ${formatURateMultiplier(rateMultiplier)} opportunities`,
    showUCL: true,
    showLCL: true
  });

  attachRowLimitExclusionContext(
    currentChart,
    "u",
    pointsWithN
  );

    // ---- Multi-period analysis (ALL segments) ----
  const analyses = [];

  for (let s = 0; s < segmentStarts.length; s++) {
    const start = segmentStarts[s];
    const end = segmentEnds[s];

    const segPoints = pointsWithN.slice(start, end + 1);

    const segBaselineCountUsed =
      (s === 0 && baselineCount && baselineCount >= 1)
        ? Math.min(baselineCount, segPoints.length)
        : segPoints.length;

    const a = analyzeAttributeChart({
      chartType: "u",
	excludedIndices:
  getMarkedExcludedIndices(
    segPoints
  ),
      labels: labels.slice(start, end + 1),
      values: values.slice(start, end + 1),
      cl: clArr.slice(start, end + 1),
      ucl: uclArr.slice(start, end + 1),
      lcl: lclArr.slice(start, end + 1)
    });

const segValues = values.slice(start, end + 1);
const segCL = clArr.slice(start, end + 1);
const segUCL = uclArr.slice(start, end + 1);
const segLCL = lclArr.slice(start, end + 1);

const uRange = rangeFinite(segUCL);
const lRange = rangeFinite(segLCL);

const dispersionDiagnostic =
  computeUDispersionDiagnostic(
    segPoints,
    segBaselineCountUsed
  );

a.stats = {
  cl: segCL.find(v => Number.isFinite(v)),     // ū
  uclMin: uRange.min,
  uclMax: uRange.max,
  uclAvg: meanFinite(segUCL),                  // ✅
  lclMin: lRange.min,
  lclMax: lRange.max,
  lclAvg: meanFinite(segLCL),                   // ✅

  dispersionDiagnostic
};



    // Context (XmR-style)
    a.periodIndex = s + 1;
    a.periodCount = segmentStarts.length;
    a.startIndex = start;
    a.endIndex = end;
    a.labelStart = labels[start];
    a.labelEnd = labels[end];
    a.baselineCountUsed = segBaselineCountUsed;

    analyses.push(a);
  }

  // Render XmR-style multi-period summary
  renderAttributeMultiSummary(analyses, labels.length);

  // Keep "latest" available for anything else that expects it
  lastAttributeAnalysis = analyses[analyses.length - 1];

}

function drawUPrimeChart(pointsWithN, baselineCount, labels) {
  pointsWithN =
    markLimitExclusionsOnPoints(
      "uprime",
      pointsWithN
    );
  if (!chartCanvas) return;

  const n =
    pointsWithN.length;

  if (n < 2) {
    showError(
      "U′ chart needs at least 2 points."
    );
    return;
  }

  const rateMultiplier =
    getURateMultiplier();

  let effectiveSplits =
    Array.isArray(splits)
      ? splits.slice()
      : [];

  effectiveSplits =
    effectiveSplits
      .filter(
        i =>
          Number.isInteger(i) &&
          i >= 0 &&
          i < n - 1
      )
      .sort((a, b) => a - b);

  const segmentStarts = [0];
  const segmentEnds = [];

  effectiveSplits.forEach(idx => {
    segmentEnds.push(idx);
    segmentStarts.push(idx + 1);
  });

  segmentEnds.push(n - 1);

  const values =
    new Array(n).fill(NaN);

  const clArr =
    new Array(n).fill(NaN);

  const uclArr =
    new Array(n).fill(NaN);

  const lclArr =
    new Array(n).fill(NaN);

  const beyond =
    new Array(n).fill(false);

  const segmentResults = [];

  for (
    let s = 0;
    s < segmentStarts.length;
    s++
  ) {
    const start =
      segmentStarts[s];

    const end =
      segmentEnds[s];

    const segPoints =
      pointsWithN.slice(
        start,
        end + 1
      );

    const segBaselineCountUsed =
      (
        s === 0 &&
        baselineCount &&
        baselineCount >= 1
      )
        ? Math.min(
            baselineCount,
            segPoints.length
          )
        : segPoints.length;

    const res =
      computeUPrime(
        segPoints,
        segBaselineCountUsed
      );

    segmentResults.push(res);

    for (
      let j = 0;
      j < segPoints.length;
      j++
    ) {
      const i =
        start + j;

      values[i] =
        res.uVals[j] *
        rateMultiplier;

      clArr[i] =
        res.ubar *
        rateMultiplier;

      uclArr[i] =
        res.ucl[j] *
        rateMultiplier;

      lclArr[i] =
        res.lcl[j] *
        rateMultiplier;

      beyond[i] =
        res.beyond[j];
    }
  }

  const pointColours =
    values.map(
      (v, i) =>
        beyond[i]
          ? SPC_STYLE.pointBeyond
          : SPC_STYLE.seriesBlue
    );

  drawSimpleSPCChart({
    labels,
    values,
    pointColours,
    cl: clArr,
    ucl: uclArr,
    lcl: lclArr,

    chartTitleFallback:
      "U′ chart",

    yAxisLabelFallback:
      rateMultiplier === 1
        ? "Rate per opportunity"
        : `Rate per ${formatURateMultiplier(rateMultiplier)} opportunities`,

    showUCL: true,
    showLCL: true
  });

  attachRowLimitExclusionContext(
    currentChart,
    "uprime",
    pointsWithN
  );

  const analyses = [];

  for (
    let s = 0;
    s < segmentStarts.length;
    s++
  ) {
    const start =
      segmentStarts[s];

    const end =
      segmentEnds[s];

    const segPoints =
      pointsWithN.slice(
        start,
        end + 1
      );

    const segBaselineCountUsed =
      (
        s === 0 &&
        baselineCount &&
        baselineCount >= 1
      )
        ? Math.min(
            baselineCount,
            segPoints.length
          )
        : segPoints.length;

    const res =
      segmentResults[s];

    const a =
      analyzeAttributeChart({
        chartType: "uprime",

	excludedIndices:
  getMarkedExcludedIndices(
    segPoints
  ),

        labels:
          labels.slice(
            start,
            end + 1
          ),

        values:
          values.slice(
            start,
            end + 1
          ),

        cl:
          clArr.slice(
            start,
            end + 1
          ),

        ucl:
          uclArr.slice(
            start,
            end + 1
          ),

        lcl:
          lclArr.slice(
            start,
            end + 1
          )
      });

    const segCL =
      clArr.slice(
        start,
        end + 1
      );

    const segUCL =
      uclArr.slice(
        start,
        end + 1
      );

    const segLCL =
      lclArr.slice(
        start,
        end + 1
      );

    const uRange =
      rangeFinite(segUCL);

    const lRange =
      rangeFinite(segLCL);

    a.stats = {
      cl:
        segCL.find(
          v => Number.isFinite(v)
        ),

      uclMin: uRange.min,
      uclMax: uRange.max,
      uclAvg:
        meanFinite(segUCL),

      lclMin: lRange.min,
      lclMax: lRange.max,
      lclAvg:
        meanFinite(segLCL),

      sigmaZ:
        res?.sigmaZ,

      avgMR:
        res?.avgMR
    };

    a.periodIndex =
      s + 1;

    a.periodCount =
      segmentStarts.length;

    a.startIndex =
      start;

    a.endIndex =
      end;

    a.labelStart =
      labels[start];

    a.labelEnd =
      labels[end];

    a.baselineCountUsed =
      segBaselineCountUsed;

    analyses.push(a);
  }

  renderAttributeMultiSummary(
    analyses,
    labels.length
  );

  lastAttributeAnalysis =
    analyses[
      analyses.length - 1
    ];
}


/**
 * Reusable SPC chart renderer for C/P/U/T/G/etc
 * Styled to match Run + XmR charts (title/legend/grid/colours).
 */
function drawSimpleSPCChart({
  labels,
  values,
  pointColours,
  cl,
  ucl,
  lcl,
  yAxisSuggestedMin,
  yAxisSuggestedMax,
  chartTitleFallback,
  yAxisLabelFallback,
  // optional toggles (handy for future)
  showUCL = true,
  showLCL = true
}) {
  if (!chartCanvas) return;

  // Keep dropdowns in sync (same behaviour as Run/XmR)
  if (typeof populateAnnotationDateOptions === "function") {
    populateAnnotationDateOptions(labels);
  }
  if (typeof populateSplitOptions === "function") {
    populateSplitOptions(labels);
  }

  // Destroy existing main chart if present
  if (currentChart) {
    currentChart.destroy();
    currentChart = null;
  }

  const labelSet = getChartLabels(chartTitleFallback, "Date", yAxisLabelFallback);
  const title = labelSet.title;
  const xLabel = labelSet.xLabel;
  const yLabel = labelSet.yLabel;

  const datasets = [
    {
      label: "Value",
      data: values,
      borderColor: SPC_STYLE.seriesBlue,
      borderWidth: 2,
      fill: false,
      pointRadius: 4,
      pointBackgroundColor: pointColours,
      pointBorderColor: pointColours,
      tension: 0.1
    },
    {
      label: "Centre line",
      data: cl,
      borderColor: SPC_STYLE.centreRed,
      borderDash: [6, 4],
      borderWidth: 2,
      pointRadius: 0,
      pointHoverRadius: 0,
      fill: false
    }
  ];

  if (showUCL) {
    datasets.push({
      label: "UCL",
      data: ucl,
      borderColor: SPC_STYLE.limitGreen,
      borderDash: [4, 4],
      borderWidth: 2,
      pointRadius: 0,
      pointHoverRadius: 0,
      fill: false
    });
  }

  if (showLCL) {
    datasets.push({
      label: "LCL",
      data: lcl,
      borderColor: SPC_STYLE.limitGreen,
      borderDash: [4, 4],
      borderWidth: 2,
      pointRadius: 0,
      pointHoverRadius: 0,
      fill: false
    });
  }

  // Optional target line
const target = getTargetValue();
if (targetEnabled && target !== null) {
    datasets.push({
      label: "Target",
      data: values.map(() => target),
      borderColor: SPC_STYLE.targetOrange,
      borderDash: [4, 2],
      borderWidth: 2,
      pointRadius: 0,
      pointHoverRadius: 0,
      fill: false
    });
  }

  const centreLineLabelPlugin =
  buildXmRMeanLabelPlugin();

  currentChart = new Chart(chartCanvas.getContext("2d"), {
  type: "line",
  data: { labels, datasets },

  plugins: [
    centreLineLabelPlugin
  ],

  options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        title: buildChartTitleConfig(title),
        legend: SPC_LEGEND,
        annotation: {
          annotations: (typeof buildAnnotationConfig === "function")
            ? buildAnnotationConfig(labels)
            : {}
        }
      },
      elements: { point: { radius: 0, hoverRadius: 0 } },
            scales: (() => {
        const axisSettings = getAxisSettings();
        return {
          x: buildCategoryXAxisConfig(xLabel, axisSettings.x, labels),
          y: buildAxisConfig(yLabel, axisSettings.y, {
            suggestedMin: isFinite(yAxisSuggestedMin) ? yAxisSuggestedMin : undefined,
            suggestedMax: isFinite(yAxisSuggestedMax) ? yAxisSuggestedMax : undefined
          })
        };
      })()
    }
  });
}

// ===============================
// XmR mean labels
// ===============================

// ===============================
// XmR mean labels
// ===============================

// Display preference.
// Save/load persistence will be added later.
let showXmRMeanLabels = false;

// Horizontal position of each period label.
// Stored as a fraction from 0 = left side of period
// to 1 = right side of period.
let xmrMeanLabelPositions = {};
let xmrMeanLabelVerticalPositions = {};
let xmrHiddenMeanLabels = new Set();

function resetXmRMeanLabelState() {
  showXmRMeanLabels = false;
  xmrMeanLabelPositions = {};
  xmrMeanLabelVerticalPositions = {};
  xmrHiddenMeanLabels = new Set();
}

function buildXmRMeanLabelPlugin(segmentSummaries = null) {

  function getPointerPosition(chart, evt) {
    const canvas = chart.canvas;
    const rect = canvas.getBoundingClientRect();

    return {
      x:
        (evt.clientX - rect.left) *
        (chart.width / rect.width),

      y:
        (evt.clientY - rect.top) *
        (chart.height / rect.height)
    };
  }

  function findMeanLabelHitbox(chart, x, y) {
    const boxes = chart.$xmrMeanLabelHitboxes || [];

    return boxes.find(box =>
      x >= box.x &&
      x <= box.x + box.width &&
      y >= box.y &&
      y <= box.y + box.height
    ) || null;
  }

  return {
    id: "xmrMeanLabels",

    afterInit(chart) {
      const canvas = chart.canvas;
      if (!canvas) return;

      const onPointerDown = (evt) => {
        if (!showXmRMeanLabels) return;
        if (evt.button !== 0) return;

        const pos = getPointerPosition(chart, evt);
        const hit = findMeanLabelHitbox(
          chart,
          pos.x,
          pos.y
        );

        if (!hit) return;

        chart.$xmrMeanLabelDrag = {
          segmentIndex: hit.segmentIndex,
          offsetX: pos.x - hit.x
        };

        if (canvas.setPointerCapture) {
          try {
            canvas.setPointerCapture(evt.pointerId);
          } catch (_) {
            // Safe fallback if pointer capture is unavailable
          }
        }

        evt.preventDefault();
      };

      const onPointerMove = (evt) => {
        if (!showXmRMeanLabels) {
          canvas.style.cursor = "";
          return;
        }

        const pos = getPointerPosition(chart, evt);
        const drag = chart.$xmrMeanLabelDrag;

        // Actively dragging a label
        if (drag) {
          const hit =
            (chart.$xmrMeanLabelHitboxes || [])
              .find(
                box =>
                  box.segmentIndex ===
                  drag.segmentIndex
              );

          if (!hit) return;

          let newX =
            pos.x - drag.offsetX;

          newX = Math.max(
            hit.minX,
            Math.min(newX, hit.maxX)
          );

          const availableWidth =
            hit.maxX - hit.minX;

          xmrMeanLabelPositions[
            drag.segmentIndex
          ] =
            availableWidth > 0
              ? (newX - hit.minX) /
                availableWidth
              : 0;

         // Snap vertically to either side of the mean line.
xmrMeanLabelVerticalPositions[
  drag.segmentIndex
] =
  pos.y < hit.meanY
    ? "above"
    : "below";

          chart.draw();

          evt.preventDefault();
          return;
        }

        // Change cursor when hovering over a label
        const hit = findMeanLabelHitbox(
          chart,
          pos.x,
          pos.y
        );

        canvas.style.cursor =
 	 hit ? "move" : "";
      };

      const onPointerUp = (evt) => {
        chart.$xmrMeanLabelDrag = null;

        if (
          canvas.releasePointerCapture &&
          canvas.hasPointerCapture?.(
            evt.pointerId
          )
        ) {
          try {
            canvas.releasePointerCapture(
              evt.pointerId
            );
          } catch (_) {
            // Safe fallback
          }
        }
      };

      canvas.addEventListener(
        "pointerdown",
        onPointerDown
      );

      canvas.addEventListener(
        "pointermove",
        onPointerMove
      );

      canvas.addEventListener(
        "pointerup",
        onPointerUp
      );

      canvas.addEventListener(
        "pointercancel",
        onPointerUp
      );

      chart.$xmrMeanLabelHandlers = {
        onPointerDown,
        onPointerMove,
        onPointerUp
      };
    },

    afterDatasetsDraw(chart) {
      // Always clear old hitboxes first
      chart.$xmrMeanLabelHitboxes = [];

      if (!showXmRMeanLabels) return;

const ctx = chart.ctx;
const chartArea = chart.chartArea;
const xScale = chart.scales?.x;
const yScale = chart.scales?.y;

if (
  !ctx ||
  !chartArea ||
  !xScale ||
  !yScale
) {
  return;
}

// Find the chart's centre-line dataset.
// Run = Median
// XmR = Mean
// Other SPC charts = Centre line
const centreDataset =
  chart.data?.datasets?.find(ds =>
    ds.label === "Mean" ||
    ds.label === "Median" ||
    ds.label === "Centre line"
  );

if (
  !centreDataset ||
  !Array.isArray(centreDataset.data) ||
  !centreDataset.data.length
) {
  return;
}

const lineColour =
  centreDataset.borderColor || "#d33";

let labelSegments = [];

// XmR already provides explicit segment summaries,
// so continue using those when available.
if (
  Array.isArray(segmentSummaries) &&
  segmentSummaries.length
) {
  labelSegments =
    segmentSummaries.map(seg => ({
      startIndex:
        Number(seg?.startIndex),

      endIndex:
        Number(seg?.endIndex),

      value:
        Number(seg?.result?.mean),

      label:
        "Mean"
    }));

} else {
  // For every other chart, derive the periods from
  // the chart labels and the current split positions.
  const totalPoints =
    chart.data?.labels?.length ||
    centreDataset.data.length;

  let effectiveSplits =
    Array.isArray(splits)
      ? splits.slice()
      : [];

  effectiveSplits =
    effectiveSplits
      .filter(
        i =>
          Number.isInteger(i) &&
          i >= 0 &&
          i < totalPoints - 1
      )
      .sort((a, b) => a - b);

  const segmentStarts = [0];
  const segmentEnds = [];

  effectiveSplits.forEach(idx => {
    segmentEnds.push(idx);
    segmentStarts.push(idx + 1);
  });

  segmentEnds.push(totalPoints - 1);

  const labelName =
    centreDataset.label === "Median"
      ? "Median"
      : centreDataset.label === "Mean"
        ? "Mean"
        : "Centre line";

  labelSegments =
    segmentStarts.map((startIndex, segmentIndex) => {
      const endIndex =
        segmentEnds[segmentIndex];

      let value = NaN;

      // Find the first valid centre-line value
      // within this process period.
      for (
        let i = startIndex;
        i <= endIndex;
        i++
      ) {
        const candidate =
          Number(centreDataset.data[i]);

        if (Number.isFinite(candidate)) {
          value = candidate;
          break;
        }
      }

      return {
        startIndex,
        endIndex,
        value,
        label: labelName
      };
    });
}

ctx.save();

ctx.font =
  "600 12px Arial, sans-serif";

ctx.textBaseline = "middle";

labelSegments.forEach(
  (seg, segmentIndex) => {

    if (xmrHiddenMeanLabels.has(segmentIndex)) {
      return;
    }

    const centreValue =
      Number(seg?.value);

    const startIndex =
      Number(seg?.startIndex);

    const endIndex =
      Number(seg?.endIndex);

    if (
      !Number.isFinite(centreValue) ||
      !Number.isFinite(startIndex) ||
      !Number.isFinite(endIndex)
    ) {
      return;
    }

          const startX =
            xScale.getPixelForValue(
              startIndex
            );

          const endX =
            xScale.getPixelForValue(
              endIndex
            );

          const y =
  yScale.getPixelForValue(
    centreValue
  );

          if (
            !Number.isFinite(startX) ||
            !Number.isFinite(endX) ||
            !Number.isFinite(y)
          ) {
            return;
          }

          const text =
  `${seg.label} = ${centreValue.toFixed(3)}`;

          const paddingX = 5;
          const boxHeight = 18;

          const textWidth =
            ctx.measureText(text).width;

          const boxWidth =
            textWidth +
            (paddingX * 2);

          const periodLeft =
            Math.min(startX, endX);

          const periodRight =
            Math.max(startX, endX);

          // Preferred movement range:
          // keep the whole label inside its period.
          let minX = Math.max(
            chartArea.left + 2,
            periodLeft + 2
          );

          let maxX = Math.min(
            chartArea.right -
              boxWidth -
              2,
            periodRight -
              boxWidth -
              2
          );

          // A very short period may physically
          // be narrower than the label itself.
          // In that case allow movement within
          // the overall chart area instead.
          if (maxX < minX) {
            minX =
              chartArea.left + 2;

            maxX =
              chartArea.right -
              boxWidth -
              2;
          }

          const savedPosition =
            Number(
              xmrMeanLabelPositions[
                segmentIndex
              ]
            );

          let boxX;

          if (
            Number.isFinite(savedPosition)
          ) {
            const fraction =
              Math.max(
                0,
                Math.min(
                  1,
                  savedPosition
                )
              );

            boxX =
              minX +
              fraction *
              (maxX - minX);
          } else {
            // Existing default:
            // near right-hand end of period
            boxX =
              endX -
              boxWidth -
              6;

            boxX = Math.max(
              minX,
              Math.min(
                boxX,
                maxX
              )
            );
          }

          // Keep the label tied vertically to the mean line,
// but allow the user to choose above or below.
let verticalSide =
  xmrMeanLabelVerticalPositions[segmentIndex] === "below"
    ? "below"
    : "above";

let boxY =
  verticalSide === "below"
    ? y + 4
    : y - boxHeight - 4;

// If the chosen side would place the label outside
// the chart, automatically use the other side.
if (
  verticalSide === "above" &&
  boxY < chartArea.top + 2
) {
  verticalSide = "below";
  boxY = y + 4;
}

if (
  verticalSide === "below" &&
  boxY + boxHeight > chartArea.bottom - 2
) {
  verticalSide = "above";
  boxY = y - boxHeight - 4;
}

// Final safety clamp.
boxY = Math.max(
  chartArea.top + 2,
  Math.min(
    boxY,
    chartArea.bottom - boxHeight - 2
  )
);

          ctx.fillStyle =
            "rgba(255,255,255,0.92)";

          ctx.fillRect(
            boxX,
            boxY,
            boxWidth,
            boxHeight
          );

          ctx.strokeStyle =
            lineColour;

          ctx.lineWidth = 1;

          ctx.strokeRect(
            boxX,
            boxY,
            boxWidth,
            boxHeight
          );

          ctx.fillStyle = "#222";

          ctx.fillText(
            text,
            boxX + paddingX,
            boxY +
              (boxHeight / 2)
          );

          // Remember where the label is
          // so the pointer handlers know
          // when the user clicks it.
         chart.$xmrMeanLabelHitboxes.push({
  segmentIndex,
  x: boxX,
  y: boxY,
  width: boxWidth,
  height: boxHeight,
  minX,
  maxX,
  meanY: y
});
        }
      );

      ctx.restore();
    },

    afterDestroy(chart) {
      const canvas = chart.canvas;

      const handlers =
        chart.$xmrMeanLabelHandlers;

      if (canvas && handlers) {
        canvas.removeEventListener(
          "pointerdown",
          handlers.onPointerDown
        );

        canvas.removeEventListener(
          "pointermove",
          handlers.onPointerMove
        );

        canvas.removeEventListener(
          "pointerup",
          handlers.onPointerUp
        );

        canvas.removeEventListener(
          "pointercancel",
          handlers.onPointerUp
        );

        canvas.style.cursor = "";
      }
    }
  };
}

function drawXmRChart(points, baselineCount, labels) {
  points =
    markLimitExclusionsOnPoints(
      "xmr",
      points
    );
  if (!chartCanvas) return;

  const n = points.length;
  if (n < MIN_XMR_LIMIT_POINTS) {
  if (errorMessage) {
    errorMessage.textContent =
      `XmR chart needs at least ${MIN_XMR_LIMIT_POINTS} points.`;
  }
  return;
}

  // ---- Read “rules & interpretation” settings (with safe fallbacks) ----
  const { shiftLength, trendLength } =
    (typeof getRuleSettings === "function")
      ? getRuleSettings()
      : { shiftLength: 8, trendLength: 6 };

  const flagOnChart =
    (typeof shouldFlagSpecialCauseOnChart === "function")
      ? shouldFlagSpecialCauseOnChart()
      : true;

  const clampLcl =
    (typeof shouldClampLclAtZero === "function")
      ? shouldClampLclAtZero()
      : false;
  
  const screenLargeMRs =
  screenLargeMRsCheckbox
    ? screenLargeMRsCheckbox.checked
    : true;

  // ----- Segment definition from splits -----
  let effectiveSplits = Array.isArray(splits) ? splits.slice() : [];
  effectiveSplits = effectiveSplits
    .filter(i => Number.isInteger(i) && i >= 0 && i < n - 1)
    .sort((a, b) => a - b);

  const segmentStarts = [0];
  const segmentEnds = [];
  effectiveSplits.forEach(idx => {
    segmentEnds.push(idx);
    segmentStarts.push(idx + 1);
  });
  segmentEnds.push(n - 1);

  // Compute a "global" XmR as a fallback (no splits)
  // (computeXmR should accept clampLcl as third arg; if not, it’ll just ignore it)
  const globalResult = computeXmR(
  points,
  baselineCount,
  clampLcl,
  screenLargeMRs
);

  // ----- Global arrays for plotting -----
  const values = points.map(p => p.y);

  const meanLine     = new Array(n).fill(NaN);
  const uclLine      = new Array(n).fill(NaN);
  const lclLine      = new Array(n).fill(NaN);
  const oneSigmaUp   = new Array(n).fill(NaN);
  const oneSigmaDown = new Array(n).fill(NaN);
  const twoSigmaUp   = new Array(n).fill(NaN);
  const twoSigmaDown = new Array(n).fill(NaN);

  const pointColours = new Array(n).fill(SPC_STYLE.seriesBlue);

  let anySigma = false;

  // We'll collect per-period results for the summary
  const segmentSummaries = [];

  // Track whether any raw LCL would be below 0 (so we can show the option conditionally)
  let anyRawLclBelowZero = false;

    // ----- Per-segment XmR -----
  for (let s = 0; s < segmentStarts.length; s++) {
    const start = segmentStarts[s];
    const end   = segmentEnds[s];

    const segPoints = points.slice(start, end + 1);

    // Only the first segment uses the user baseline; later segments use all points as baseline.
    const segBaseline = s === 0 ? baselineCount : null;

    const segResult = computeXmR(
  segPoints,
  segBaseline,
  clampLcl,
  screenLargeMRs
);
    const segPts    = segResult.points;

    const mean  = segResult.mean;
    const ucl   = segResult.ucl;
    const lcl   = segResult.lcl;
    const sigma = segResult.sigma;

    const limitEstimatePointCount =
  Number.isFinite(
    segResult
      .effectiveBaselinePointCount
  )
    ? segResult
        .effectiveBaselinePointCount
    : (
        Number.isFinite(
          segResult
            .baselineCountUsed
        )
          ? segResult
              .baselineCountUsed
          : segPoints.length
      );

const hasEnoughPointsForInterpretation =
  limitEstimatePointCount >= MIN_XMR_LIMIT_POINTS;

    // If computeXmR returns rawLcl, use it to decide whether to show the clamp option
    if (typeof segResult.rawLcl === "number" && segResult.rawLcl < 0) {
      anyRawLclBelowZero = true;
    }

    // Rule-aware analysis for this segment
    const segValues = segPts.map(p => p.y);

    const segAnalysis = analyzeAttributeChart({
      chartType: "xmr",
      labels: labels.slice(start, end + 1),
      values: segValues,
      cl: new Array(segValues.length).fill(mean),
      ucl:
  new Array(
    segValues.length
  ).fill(ucl),

lcl:
  new Array(
    segValues.length
  ).fill(lcl),

excludedIndices:
  getMarkedExcludedIndices(
    segPoints
  )
    });

    // Store for multi-period summary
    segmentSummaries.push({
      startIndex: start,
      endIndex: end,
      labelStart: labels[start],
      labelEnd: labels[end],
      result: segResult,
      analysis: segAnalysis,
  limitEstimatePointCount,
  hasEnoughPointsForInterpretation
    });

    // Fill chart arrays for this segment
    for (let i = 0; i < segValues.length; i++) {
      const globalIdx = start + i;

      // Colouring:
      // - beyond limits = red
      // - other special-cause signals = orange
      // - otherwise blue
      if (!hasEnoughPointsForInterpretation || !flagOnChart) {
  // Do not present statistical signals from provisional XmR limits.
  pointColours[globalIdx] = SPC_STYLE.pointNormal;
} else if (segAnalysis.flags?.beyond?.[i]) {
  pointColours[globalIdx] = SPC_STYLE.pointBeyond;
} else if (segAnalysis.flags?.special?.[i]) {
  pointColours[globalIdx] = SPC_STYLE.pointSpecial;
} else {
  pointColours[globalIdx] = SPC_STYLE.pointNormal;
}

      // Centre line & limits
      meanLine[globalIdx] = mean;
      uclLine[globalIdx]  = ucl;
      lclLine[globalIdx]  = lcl;

      // Sigma lines (only if sigma is valid)
      if (sigma && sigma > 0) {
        anySigma = true;
        oneSigmaUp[globalIdx]   = mean + sigma;
        oneSigmaDown[globalIdx] = mean - sigma;
        twoSigmaUp[globalIdx]   = mean + 2 * sigma;
        twoSigmaDown[globalIdx] = mean - 2 * sigma;
      }
    }
  }

  // ---- Show/hide the “Fix LCL at 0” option only when relevant ----
  if (typeof setLclClampVisibility === "function") {
    setLclClampVisibility(anyRawLclBelowZero);
  } else {
    // Fallback if you haven't added the helper yet
    const row = document.getElementById("lclClampRow");
    if (row) row.style.display = anyRawLclBelowZero ? "block" : "none";
  }

  // ----- Build datasets -----
  const datasets = [];

  // Main values
  datasets.push({
    label: "Value",
    data: values,
    borderColor: SPC_STYLE.seriesBlue,
    backgroundColor: SPC_STYLE.seriesBlue,
    pointRadius: 3,
    pointHoverRadius: 4,
    pointBackgroundColor: pointColours,
    pointBorderColor: "#ffffff",
    pointBorderWidth: 1,
    tension: 0,
    yAxisID: "y"
  });

  // Mean + limits
  datasets.push(
    {
      label: "Mean",
      data: meanLine,
      borderColor: SPC_STYLE.pointBeyond,
      borderDash: [6, 4],
      pointRadius: 0
    },
    {
      label: "UCL (3σ)",
      data: uclLine,
      borderColor: SPC_STYLE.limitGreen,
      borderDash: [4, 4],
      pointRadius: 0
    },
    {
      label: "LCL (3σ)",
      data: lclLine,
      borderColor: SPC_STYLE.limitGreen,
      borderDash: [4, 4],
      pointRadius: 0
    }
  );

  // Optional sigma reference lines
  if (anySigma) {
    const sigmaStyle = {
      borderColor: "rgba(0,0,0,0.12)",
      borderWidth: 1,
      borderDash: [2, 2],
      pointRadius: 0
    };

    datasets.push(
      { label: "+1σ", data: oneSigmaUp,   ...sigmaStyle },
      { label: "-1σ", data: oneSigmaDown, ...sigmaStyle },
      { label: "+2σ", data: twoSigmaUp,   ...sigmaStyle },
      { label: "-2σ", data: twoSigmaDown, ...sigmaStyle }
    );
  }

  // Target line (optional)
const target = getTargetValue();
if (targetEnabled && target !== null) {
    datasets.push({
      label: "Target",
      data: values.map(() => target),
      borderColor: SPC_STYLE.targetOrange,
      borderWidth: 2,
      borderDash: [4, 2],
      pointRadius: 0,
      tension: 0
    });
  }

  // Update annotation and split dropdowns
  populateAnnotationDateOptions(labels);
  populateSplitOptions(labels);

  // ----- Create chart -----
if (currentChart) currentChart.destroy();

const { title, xLabel, yLabel } =
  getChartLabels("X Chart", "Date", "Value");

const xmrMeanLabelPlugin =
  buildXmRMeanLabelPlugin(segmentSummaries);

const centreLineLabelPlugin =
  buildXmRMeanLabelPlugin();

currentChart = new Chart(chartCanvas, {
  type: "line",
  data: { labels, datasets },

  plugins: [
    centreLineLabelPlugin
  ],

  options: {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      title: buildChartTitleConfig(title),
      legend: SPC_LEGEND,
      annotation: {
        annotations: buildAnnotationConfig(labels)
      }
    },
    elements: {
      point: { radius: 0, hoverRadius: 0 }
    },
    scales: (() => {
      const axisSettings = getAxisSettings();
      return {
        x: buildCategoryXAxisConfig(xLabel, axisSettings.x, labels),
        y: buildAxisConfig(yLabel, axisSettings.y)
      };
    })()
  }
});

attachRowLimitExclusionContext(
  currentChart,
  "xmr",
  points
);

    // Expose all XmR periods to the helper so it can talk about the whole chart,
  // not just the latest segment.
  window.lastXmRPeriods = segmentSummaries.map((seg, idx) => ({
    periodIndex: idx + 1,
    periodCount: segmentSummaries.length,
    startIndex: seg.startIndex,
    endIndex: seg.endIndex,
    labelStart: seg.labelStart,
    labelEnd: seg.labelEnd,
    mean: seg.result?.mean,
    ucl: seg.result?.ucl,
    lcl: seg.result?.lcl,
    sigma:
  seg.result?.sigma,

effectiveBaselinePointCount:
  seg.result
    ?.effectiveBaselinePointCount,

manualExcludedPointCount:
  seg.result
    ?.manualExcludedPointCount ||
  0,

manualAdjacentMRExclusions:
  seg.result
    ?.manualAdjacentMRExclusions ||
  0,

hasEnoughPointsForInterpretation:
  !!seg.hasEnoughPointsForInterpretation,

limitEstimatePointCount:
  seg.limitEstimatePointCount,

minimumPointsForInterpretation:
  MIN_XMR_LIMIT_POINTS,

isStable:
  seg.hasEnoughPointsForInterpretation &&
  Array.isArray(seg.analysis?.signals)
    ? seg.analysis.signals.length === 0
    : null,

signals:
  seg.hasEnoughPointsForInterpretation &&
  Array.isArray(seg.analysis?.signals)
    ? seg.analysis.signals.slice()
    : []
  }));


  // ----- Summary -----
  if (segmentSummaries.length > 0) {
    updateXmRMultiSummary(segmentSummaries, points.length);
  } else {
    updateXmRMultiSummary(
      [{
        startIndex: 0,
        endIndex: n - 1,
        labelStart: labels[0],
        labelEnd: labels[n - 1],
        result: globalResult
      }],
      points.length
    );
  }

    // ----- Show / hide MR chart depending on checkbox -----
  const showMR = showMRCheckbox ? showMRCheckbox.checked : true;

  if (showMR) {
    // If we have splits, pass the same segment structure used by the summary.
    // If no splits, fall back to a single “whole-chart” segment.
    const mrSegments = (segmentSummaries && segmentSummaries.length)
      ? segmentSummaries
      : [{
          startIndex: 0,
          endIndex: n - 1,
          labelStart: labels[0],
          labelEnd: labels[n - 1],
          result: globalResult
        }];

    // IMPORTANT: pass the full points + full labels + the segments array
    drawMrChart(points, labels, mrSegments);

  } else {
    if (mrChart) {
      mrChart.destroy();
      mrChart = null;
    }
    if (mrPanel) {
      mrPanel.style.display = "none";
    }
  }

}


function formatMrValueForDisplay(v) {
  if (!Number.isFinite(v)) return "—";
  return (typeof formatNumber === "function")
    ? formatNumber(v, 3)
    : Number(v).toFixed(3);
}

function buildMrLocalSeries(segmentPoints, segmentResult) {
  const values = Array(segmentPoints.length).fill(null);
  const reasons = Array(segmentPoints.length).fill(null);

  const screenedPairs = new Set(
    Array.isArray(segmentResult?.excludedMRs)
      ? segmentResult.excludedMRs.map(item => `${item.fromPointIndex}-${item.toPointIndex}`)
      : []
  );

  for (let i = 1; i < segmentPoints.length; i++) {
    const prev = segmentPoints[i - 1];
    const curr = segmentPoints[i];

    const mr = Math.abs(curr.y - prev.y);
    values[i] = mr;

    if (prev?._excludeFromLimits || curr?._excludeFromLimits) {
      reasons[i] = "excluded-point";
    } else if (screenedPairs.has(`${i - 1}-${i}`)) {
      reasons[i] = "screened-large";
    } else {
      reasons[i] = null;
    }
  }

  return { values, reasons };
}

function buildMrOverlayArray(values, reasons, targetReason) {
  return values.map((v, i) => (
    reasons[i] === targetReason && Number.isFinite(v) ? v : null
  ));
}

function buildMrMarkerDataset(
  label,
  data,
  colour,
  pointStyle,
  reasonCode
) {
  const hasAny =
    data.some(
      v => Number.isFinite(v)
    );

  if (!hasAny) {
    return null;
  }

  return {
    label,
    data,

    // Used by the MR-specific legend and tooltip.
    _mrOmissionMarker: true,
    _mrOmissionReason: reasonCode,

    type: "line",
    showLine: false,
    spanGaps: false,

    borderColor: colour,
    backgroundColor: "#ffffff",

    pointStyle,

    pointRadius:
      data.map(
        v =>
          Number.isFinite(v)
            ? 8
            : 0
      ),

    pointHoverRadius:
      data.map(
        v =>
          Number.isFinite(v)
            ? 10
            : 0
      ),

    pointBackgroundColor:
      "#ffffff",

    pointBorderColor:
      colour,

    pointBorderWidth: 3,

    fill: false,
    tension: 0
  };
}

function buildMrLegendConfig() {
  return {
    display: true,
    position: "bottom",
    align: "center",

    labels: {
      usePointStyle: true,
      boxWidth: 55,
      boxHeight: 12,
      padding: 14,

      generateLabels(chart) {
        const defaultLabels =
          Chart.defaults.plugins.legend.labels.generateLabels(
            chart
          );

        return defaultLabels.map(
          label => {
            const dataset =
              chart.data.datasets[
                label.datasetIndex
              ] || {};

            if (
              dataset._mrOmissionMarker
            ) {
              return {
                ...label,

                pointStyle:
                  dataset.pointStyle,

                fillStyle:
                  dataset.pointBackgroundColor ||
                  "#ffffff",

                strokeStyle:
                  dataset.pointBorderColor ||
                  dataset.borderColor,

                lineWidth:
                  dataset.pointBorderWidth ||
                  2,

                lineDash: []
              };
            }

            return {
              ...label,

              pointStyle:
                "line",

              strokeStyle:
                dataset.borderColor ||
                label.strokeStyle,

              fillStyle:
                "transparent",

              lineWidth:
                Math.max(
                  dataset.borderWidth ||
                  2,
                  3
                ),

              lineDash:
                dataset.borderDash ||
                [],

              lineCap:
                "round"
            };
          }
        );
      }
    }
  };
}

function buildMrTooltipOptions() {
  return {
    filter(context) {
      const dataset =
        context.dataset || {};

      /*
       * At an omitted MR we deliberately hide the ordinary
       * "Moving range" tooltip item and show only the explanatory
       * marker dataset.
       */
      if (
        dataset.label ===
        "Moving range"
      ) {
        const reason =
          dataset._mrReasons
            ?.[context.dataIndex];

        return !reason;
      }

      return true;
    },

    callbacks: {
      label(context) {
        const value =
          context.parsed?.y;

        if (
          !Number.isFinite(value)
        ) {
          return null;
        }

        const pretty =
          formatMrValueForDisplay(
            value
          );

        const reason =
          context.dataset
            ?._mrOmissionReason;

        if (
          reason ===
          "excluded-point"
        ) {
          return (
            `Not used in the variation estimate: ${pretty} ` +
            `(adjacent to an excluded point)`
          );
        }

        if (
          reason ===
          "screened-large"
        ) {
          return (
            `Not used in the variation estimate: ${pretty} ` +
            `(screened as an unusually large moving range)`
          );
        }

        const label =
          context.dataset?.label ||
          "";

        if (
          label ===
          "Moving range"
        ) {
          return (
            `Moving range: ${pretty}`
          );
        }

        return (
          `${label}: ${pretty}`
        );
      }
    }
  };
}


function drawMrChart(allPoints, labels, segments) {
  // allPoints: full list of points for the current XmR chart (all periods)
  // labels: full x labels used on the X chart
  // segments: [{ startIndex, endIndex, result }, ...] (same segments used by X chart)

  if (!mrCanvas || !mrPanel) return;
  mrPanel.style.display = "block";

  const strong = mrPanel.querySelector("strong");
  if (strong) strong.textContent = "Moving Range chart:";

  const showAll =
    (typeof getMrDisplayMode === "function") &&
    (getMrDisplayMode() === "all");

  const BLUE = SPC_STYLE.seriesBlue;
  const RED = SPC_STYLE.pointBeyond;
  const GREEN = SPC_STYLE.limitGreen;
  const OMIT_EXCLUDED_COLOUR = "#4b5563";
  const OMIT_SCREENED_COLOUR = "#f59e0b";

  function computeAvgMR(mrArr) {
    const vals = mrArr.filter(v => typeof v === "number" && isFinite(v));
    return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : 0;
  }

  // ----- LAST PERIOD ONLY -----
  if (!showAll) {
    const lastSeg = segments && segments.length ? segments[segments.length - 1] : null;
    if (!lastSeg) return;

    const pts = allPoints.slice(lastSeg.startIndex, lastSeg.endIndex + 1);
    const localMr = buildMrLocalSeries(pts, lastSeg.result);

    const mrValues = localMr.values;
    const mrReasons = localMr.reasons;

    const avgMR =
      (lastSeg.result && typeof lastSeg.result.avgMR === "number")
        ? lastSeg.result.avgMR
        : computeAvgMR(mrValues);

    const uclMR = 3.268 * avgMR;
    const mrLabels = labels.slice(lastSeg.startIndex, lastSeg.endIndex + 1);

    renderMrChart(mrLabels, mrValues, mrReasons, avgMR, uclMR);
    return;
  }

  // ----- ALL PERIODS (WITH SPLITS) -----
  const mrAll = Array(labels.length).fill(null);
  const mrReasonsAll = Array(labels.length).fill(null);

  (segments || []).forEach(seg => {
    const pts = allPoints.slice(seg.startIndex, seg.endIndex + 1);
    const localMr = buildMrLocalSeries(pts, seg.result);

    for (let i = 0; i < localMr.values.length; i++) {
      const globalIndex = seg.startIndex + i;
      mrAll[globalIndex] = localMr.values[i];
      mrReasonsAll[globalIndex] = localMr.reasons[i];
    }
  });

  const omittedBecauseExcluded = buildMrOverlayArray(
    mrAll,
    mrReasonsAll,
    "excluded-point"
  );

  const omittedBecauseScreened = buildMrOverlayArray(
    mrAll,
    mrReasonsAll,
    "screened-large"
  );

  const basePointColours = mrAll.map((v, i) => {
    if (!Number.isFinite(v)) return "rgba(0,0,0,0)";
    if (mrReasonsAll[i] === "excluded-point") return OMIT_EXCLUDED_COLOUR;
    if (mrReasonsAll[i] === "screened-large") return OMIT_SCREENED_COLOUR;
    return BLUE;
  });

  const datasets = [];

  // MR line across all periods
  datasets.push({
    label: "Moving range",
    data: mrAll,
    borderColor: BLUE,
    backgroundColor: BLUE,
    borderWidth: 2,
    pointRadius: mrAll.map(v => Number.isFinite(v) ? 3 : 0),
    pointHoverRadius: mrAll.map(v => Number.isFinite(v) ? 4 : 0),
    pointBackgroundColor: basePointColours,
    pointBorderColor: "#ffffff",
    pointBorderWidth: 1,
    spanGaps: false,
    fill: false,
    tension: 0
  });

  const excludedMarkerDataset =
  buildMrMarkerDataset(
    "Not used: adjacent to excluded point",
    omittedBecauseExcluded,
    OMIT_EXCLUDED_COLOUR,
    "crossRot",
    "excluded-point"
  );
  if (excludedMarkerDataset) datasets.push(excludedMarkerDataset);

  const screenedMarkerDataset =
  buildMrMarkerDataset(
    "Not used: unusually large MR screened",
    omittedBecauseScreened,
    OMIT_SCREENED_COLOUR,
    "triangle",
    "screened-large"
  );
  if (screenedMarkerDataset) datasets.push(screenedMarkerDataset);

  // One pair of lines per period
  (segments || []).forEach((seg, idx) => {
    const pts = allPoints.slice(seg.startIndex, seg.endIndex + 1);
    const localMr = buildMrLocalSeries(pts, seg.result);
    const mrValues = localMr.values;

    const avgMR =
      (seg.result && typeof seg.result.avgMR === "number")
        ? seg.result.avgMR
        : computeAvgMR(mrValues);

    const uclMR = 3.268 * avgMR;

    const mrBarLine = Array(labels.length).fill(null);
    const uclLine = Array(labels.length).fill(null);

    for (let i = seg.startIndex; i <= seg.endIndex; i++) {
      mrBarLine[i] = avgMR;
      uclLine[i] = uclMR;
    }

    // MR undefined at first point of each segment
    mrBarLine[seg.startIndex] = null;
    uclLine[seg.startIndex] = null;

    datasets.push({
      label: `MR average (Period ${idx + 1})`,
      data: mrBarLine,
      borderColor: RED,
      borderDash: [6, 4],
      borderWidth: 2,
      pointRadius: 0,
      fill: false,
      tension: 0
    });

    datasets.push({
      label: `MR UCL (Period ${idx + 1})`,
      data: uclLine,
      borderColor: GREEN,
      borderDash: [4, 4],
      borderWidth: 2,
      pointRadius: 0,
      fill: false,
      tension: 0
    });
  });

  if (mrChart) {
    mrChart.destroy();
    mrChart = null;
  }

  mrChart = new Chart(mrCanvas, {
    type: "line",
    data: { labels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        title: {
          display: true,
          text: "Moving Range (MR)",
          font: { size: 14, weight: "bold" }
        },
        legend:
  buildMrLegendConfig(),

tooltip:
  buildMrTooltipOptions(),
        annotation: {
          annotations: (typeof buildAnnotationConfig === "function")
            ? buildAnnotationConfig(labels)
            : {}
        }
      },
      elements: {
        point: { radius: 0, hoverRadius: 0 }
      },
      scales: (() => {
        const axisSettings = getAxisSettings();
        return {
          x: buildCategoryXAxisConfig("", axisSettings.x, labels),
          y: buildAxisConfig("", withoutAxisBounds(axisSettings.y), { beginAtZero: true })
        };
      })()
    }
  });
}

// Helper used by "last period only" mode — uses your existing MR canvas/chart variables
function renderMrChart(mrLabels, mrValues, mrReasons, avgMR, uclMR) {
  if (!mrCanvas) return;

  if (mrPanel) {
    const strong = mrPanel.querySelector("strong");
    if (strong) strong.textContent = "Moving Range chart:";
  }

  if (mrChart) {
    mrChart.destroy();
    mrChart = null;
  }

  const OMIT_EXCLUDED_COLOUR = "#4b5563";
  const OMIT_SCREENED_COLOUR = "#f59e0b";

  const omittedBecauseExcluded = buildMrOverlayArray(
    mrValues,
    mrReasons,
    "excluded-point"
  );

  const omittedBecauseScreened = buildMrOverlayArray(
    mrValues,
    mrReasons,
    "screened-large"
  );

  const basePointColours = mrValues.map((v, i) => {
    if (!Number.isFinite(v)) return "rgba(0,0,0,0)";
    if (mrReasons[i] === "excluded-point") return OMIT_EXCLUDED_COLOUR;
    if (mrReasons[i] === "screened-large") return OMIT_SCREENED_COLOUR;
    return SPC_STYLE.seriesBlue;
  });

  const datasets = [
    {
      label: "Moving range",
      data: mrValues,
      borderColor: SPC_STYLE.seriesBlue,
      backgroundColor: SPC_STYLE.seriesBlue,
      borderWidth: 2,
      _mrReasons:
  mrReasons.slice(),

pointRadius:
  mrValues.map(
    (v, i) =>
      Number.isFinite(v) &&
      !mrReasons[i]
        ? 3
        : 0
  ),

pointHoverRadius:
  mrValues.map(
    (v, i) =>
      Number.isFinite(v) &&
      !mrReasons[i]
        ? 4
        : 0
  ),

pointBackgroundColor:
  basePointColours,
      pointBorderColor: "#ffffff",
      pointBorderWidth: 1,
      spanGaps: false,
      fill: false,
      tension: 0
    },
    {
      label: "MR average",
      data: mrValues.map(() => avgMR),
      borderColor: SPC_STYLE.pointBeyond,
      borderDash: [6, 4],
      borderWidth: 2,
      pointRadius: 0,
      fill: false
    },
    {
      label: "MR UCL",
      data: mrValues.map(() => uclMR),
      borderColor: SPC_STYLE.limitGreen,
      borderDash: [4, 4],
      borderWidth: 2,
      pointRadius: 0,
      fill: false
    }
  ];

  const excludedMarkerDataset =
  buildMrMarkerDataset(
    "Not used: adjacent to excluded point",
    omittedBecauseExcluded,
    OMIT_EXCLUDED_COLOUR,
    "crossRot",
    "excluded-point"
  );
  if (excludedMarkerDataset) datasets.push(excludedMarkerDataset);

  const screenedMarkerDataset =
  buildMrMarkerDataset(
    "Not used: unusually large MR screened",
    omittedBecauseScreened,
    OMIT_SCREENED_COLOUR,
    "triangle",
    "screened-large"
  );
  if (screenedMarkerDataset) datasets.push(screenedMarkerDataset);

  mrChart = new Chart(mrCanvas, {
    type: "line",
    data: { labels: mrLabels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        title: {
          display: true,
          text: "Moving Range (MR)",
          font: { size: 14, weight: "bold" }
        },
        legend:
  buildMrLegendConfig(),

tooltip:
  buildMrTooltipOptions(),
        annotation: {
          annotations: (typeof buildAnnotationConfig === "function")
            ? buildAnnotationConfig(mrLabels)
            : {}
        }
      },
      elements: {
        point: { radius: 0, hoverRadius: 0 }
      },
      scales: (() => {
        const axisSettings = getAxisSettings();
        return {
          x: buildCategoryXAxisConfig("", axisSettings.x, mrLabels),
          y: buildAxisConfig("", withoutAxisBounds(axisSettings.y), { beginAtZero: true })
        };
      })()
    }
  });
}

// ============================================================
// SimpleSPC statistical regression tests
//
// Development / quality-assurance helper.
// Nothing runs automatically.
//
// To run:
//   runSimpleSPCRegressionTests()
//
// from the browser developer console.
// ============================================================

function runSimpleSPCRegressionTests() {
  const results = [];

  function near(
    actual,
    expected,
    tolerance = 1e-9
  ) {
    return (
      Number.isFinite(actual) &&
      Number.isFinite(expected) &&
      Math.abs(
        actual - expected
      ) <= tolerance
    );
  }

  function addNumeric(
    name,
    actual,
    expected,
    tolerance = 1e-9
  ) {
    results.push({
      test: name,
      pass:
        near(
          actual,
          expected,
          tolerance
        ),

      actual,
      expected,
      tolerance
    });
  }

  function addBoolean(
    name,
    actual,
    expected
  ) {
    results.push({
      test: name,
      pass:
        actual === expected,

      actual,
      expected,
      tolerance: ""
    });
  }

  function addText(
    name,
    actual,
    expected
  ) {
    results.push({
      test: name,
      pass:
        String(actual) ===
        String(expected),

      actual,
      expected,
      tolerance: ""
    });
  }

  // ----------------------------------------------------------
  // RUN CHART
  // ----------------------------------------------------------

  addNumeric(
    "Run: median",
    computeMedian(
      [1, 2, 3, 4]
    ),
    2.5
  );

  const runShift =
    findShiftWindow(
      [
        6, 7, 8, 9,
        10, 11, 12, 13
      ],
      5,
      8
    );

  addBoolean(
    "Run: 8-point shift detected",
    !!runShift,
    true
  );

  addNumeric(
    "Run: shift starts at index 0",
    runShift?.start,
    0
  );

  addNumeric(
    "Run: shift ends at index 7",
    runShift?.end,
    7
  );

  const runTrend =
    findTrendWindow(
      [1, 2, 3, 4, 5, 6],
      6
    );

  addBoolean(
    "Run: 6-point trend detected",
    !!runTrend,
    true
  );

  addText(
    "Run: trend direction",
    runTrend?.direction,
    "up"
  );

  // ----------------------------------------------------------
  // XmR — ordinary calculation
  // Independently checked expected values:
  // mean = 10.4
  // MR = [1,2,1,2]
  // MRbar = 1.5
  // sigma = 1.5 / 1.128
  // ----------------------------------------------------------

  const xmrPoints =
    [10, 11, 9, 10, 12]
      .map(
        (y, i) => ({
          x: i,
          y,
          label:
            `P${i + 1}`
        })
      );

  const xmr =
    computeXmR(
      xmrPoints,
      null,
      false,
      false
    );

  addNumeric(
    "XmR: mean",
    xmr.mean,
    10.4,
    1e-10
  );

  addNumeric(
    "XmR: average MR",
    xmr.avgMR,
    1.5,
    1e-10
  );

  addNumeric(
    "XmR: sigma",
    xmr.sigma,
    1.3297872340425534,
    1e-10
  );

  addNumeric(
    "XmR: UCL",
    xmr.ucl,
    14.38936170212766,
    1e-10
  );

  addNumeric(
    "XmR: LCL",
    xmr.lcl,
    6.410638297872341,
    1e-10
  );

  // ----------------------------------------------------------
  // XmR — MR screening
  //
  // Raw MRs:
  // 1,1,1,1,1,25,1,1,1
  //
  // Initial MRbar = 3.666666...
  // threshold = 12.833333...
  // 25 is screened
  // revised MRbar = 1
  // ----------------------------------------------------------

  const screenPoints =
    [
      0, 1, 2, 3, 4,
      5, 30, 31, 32, 33
    ]
      .map(
        (y, i) => ({
          x: i,
          y,
          label:
            `S${i + 1}`
        })
      );

  const screened =
    computeXmR(
      screenPoints,
      null,
      false,
      true
    );

  addNumeric(
    "XmR screening: initial average MR",
    screened.initialAvgMR,
    3.6666666666666665,
    1e-10
  );

  addNumeric(
    "XmR screening: threshold",
    screened.mrScreeningThreshold,
    12.833333333333332,
    1e-10
  );

  addNumeric(
    "XmR screening: revised average MR",
    screened.avgMR,
    1,
    1e-10
  );

  addNumeric(
    "XmR screening: number screened",
    screened.excludedMRCount,
    1
  );

  // ----------------------------------------------------------
  // XmR — manual exclusion
  //
  // Excluding the middle value removes both adjacent MRs.
  // It must NOT create a bridge MR.
  // ----------------------------------------------------------

  const excludedXmRPoints =
    [10, 11, 9, 10, 12]
      .map(
        (y, i) => ({
          x: i,
          y,
          label:
            `E${i + 1}`,

          _excludeFromLimits:
            i === 2
        })
      );

  const excludedXmR =
    computeXmR(
      excludedXmRPoints,
      null,
      false,
      false
    );

  addNumeric(
    "XmR exclusion: mean",
    excludedXmR.mean,
    10.75,
    1e-10
  );

  addNumeric(
    "XmR exclusion: average MR",
    excludedXmR.avgMR,
    1.5,
    1e-10
  );

  addNumeric(
    "XmR exclusion: adjacent MRs omitted",
    excludedXmR.manualAdjacentMRExclusions,
    2
  );

  addNumeric(
    "XmR exclusion: effective baseline points",
    excludedXmR.effectiveBaselinePointCount,
    4
  );

  // Fixed baseline-window test.
  // Point 5 must NOT be pulled into the baseline
  // to replace the excluded point.
  const fixedWindowXmR =
    [
      { x: 0, y: 10 },
      {
        x: 1,
        y: 100,
        _excludeFromLimits:
          true
      },
      { x: 2, y: 12 },
      { x: 3, y: 14 },
      { x: 4, y: 1000 }
    ];

  const fixedWindowResult =
    computeXmR(
      fixedWindowXmR,
      4,
      false,
      false
    );

  addNumeric(
    "XmR exclusion: baseline boundary remains fixed",
    fixedWindowResult.mean,
    12,
    1e-10
  );

  addNumeric(
    "XmR exclusion: baseline used count",
    fixedWindowResult.effectiveBaselinePointCount,
    3
  );

  // ----------------------------------------------------------
  // C CHART
  // ----------------------------------------------------------

  const cPoints =
    [
      5, 4, 6, 5, 5,
      4, 6, 5, 25, 5,
      4, 6, 5, 5, 4,
      6, 5, 5, 4, 6
    ]
      .map(
        (y, i) => ({
          x: i,
          y
        })
      );

  const cNormal =
    computeC(
      cPoints
    );

  addNumeric(
    "C: centre line",
    cNormal.cbar,
    6,
    1e-10
  );

  addNumeric(
    "C: UCL",
    cNormal.ucl,
    13.348469228349533,
    1e-10
  );

  const cExcluded =
    computeC(
      cPoints.map(
        (p, i) => ({
          ...p,

          _excludeFromLimits:
            i === 8
        })
      )
    );

  addNumeric(
    "C exclusion: centre line",
    cExcluded.cbar,
    5,
    1e-10
  );

  addNumeric(
    "C exclusion: UCL",
    cExcluded.ucl,
    11.70820393249937,
    1e-10
  );

  // ----------------------------------------------------------
  // P CHART
  // ----------------------------------------------------------

  const pPoints = [
    { y: 10, n: 100 },
    { y: 18, n: 120 },
    { y: 6,  n: 80 },
    { y: 15, n: 110 }
  ];

  const pResult =
    computeP(
      pPoints,
      null,
      true
    );

  addNumeric(
    "P: p-bar",
    pResult.pbar,
    0.11951219512195121,
    1e-12
  );

  addNumeric(
    "P: first UCL",
    pResult.ucl[0],
    0.21682926829268293,
    1e-12
  );

  addNumeric(
    "P: first LCL",
    pResult.lcl[0],
    0.022195121951219504,
    1e-12
  );

  // ----------------------------------------------------------
  // P′ CHART
  // ----------------------------------------------------------

  const pPrimeResult =
    computePPrime(
      pPoints
    );

  addNumeric(
    "P′: p-bar",
    pPrimeResult.pbar,
    0.11951219512195121,
    1e-12
  );

  addNumeric(
    "P′: Sigma Z",
    pPrimeResult.sigmaZ,
    1.672599851246506,
    1e-10
  );

  addNumeric(
    "P′: first UCL",
    pPrimeResult.ucl[0],
    0.2822847172310624,
    1e-10
  );

  // Verify that an excluded point also breaks
  // the Laney z-moving-range sequence.
  const pPrimeExcluded =
    computePPrime(
      pPoints.map(
        (p, i) => ({
          ...p,

          _excludeFromLimits:
            i === 1
        })
      )
    );

  addNumeric(
    "P′ exclusion: p-bar",
    pPrimeExcluded.pbar,
    0.10689655172413794,
    1e-12
  );

  addNumeric(
    "P′ exclusion: only one valid z-MR remains",
    pPrimeExcluded.baselineMRCount,
    1
  );

  addNumeric(
    "P′ exclusion: Sigma Z",
    pPrimeExcluded.sigmaZ,
    1.7052837483437422,
    1e-10
  );

  // ----------------------------------------------------------
  // U CHART
  // ----------------------------------------------------------

  const uPoints = [
    { y: 5,  n: 100 },
    { y: 8,  n: 120 },
    { y: 3,  n: 90 },
    { y: 10, n: 110 },
    { y: 4,  n: 105 }
  ];

  const uResult =
    computeU(
      uPoints
    );

  addNumeric(
    "U: u-bar",
    uResult.ubar,
    0.05714285714285714,
    1e-12
  );

  addNumeric(
    "U: first UCL",
    uResult.ucl[0],
    0.12885657370292075,
    1e-10
  );

  addNumeric(
    "U: first LCL",
    uResult.lcl[0],
    0,
    1e-12
  );

  // ----------------------------------------------------------
  // U′ CHART
  // ----------------------------------------------------------

  const uPrimeResult =
    computeUPrime(
      uPoints
    );

  addNumeric(
    "U′: u-bar",
    uPrimeResult.ubar,
    0.05714285714285714,
    1e-12
  );

  addNumeric(
    "U′: Sigma Z",
    uPrimeResult.sigmaZ,
    1.5161730157805888,
    1e-10
  );

  addNumeric(
    "U′: first UCL",
    uPrimeResult.ucl[0],
    0.16587325905256314,
    1e-10
  );

  // ----------------------------------------------------------
  // P / U dispersion diagnostics
  // Independently calculated reference dataset.
  // ----------------------------------------------------------

  const dispersionPoints = [
    { y: 4,  n: 100 },
    { y: 15, n: 110 },
    { y: 3,  n: 95 },
    { y: 17, n: 105 },
    { y: 5,  n: 120 },
    { y: 18, n: 115 },
    { y: 4,  n: 90 },
    { y: 16, n: 100 },
    { y: 3,  n: 125 },
    { y: 19, n: 130 },
    { y: 5,  n: 105 },
    { y: 17, n: 110 },
    { y: 4,  n: 115 },
    { y: 18, n: 120 },
    { y: 3,  n: 100 },
    { y: 16, n: 95 },
    { y: 5,  n: 130 },
    { y: 20, n: 125 },
    { y: 4,  n: 110 },
    { y: 17, n: 105 },
    { y: 3,  n: 90 },
    { y: 18, n: 115 },
    { y: 5,  n: 120 },
    { y: 19, n: 130 }
  ];

  const pDiag =
    computePDispersionDiagnostic(
      dispersionPoints
    );

  addBoolean(
    "P diagnostic: calculation succeeds",
    pDiag.ok,
    true
  );

  addNumeric(
    "P diagnostic: variation ratio %",
    pDiag.ratio,
    535.36162769218,
    0.05
  );

  addNumeric(
    "P diagnostic: upper threshold %",
    pDiag.upperLimit,
    155.9951980036352,
    0.05
  );

  addText(
    "P diagnostic: overdispersion",
    pDiag.status,
    "overdispersed"
  );

  const uDiag =
    computeUDispersionDiagnostic(
      dispersionPoints
    );

  addBoolean(
    "U diagnostic: calculation succeeds",
    uDiag.ok,
    true
  );

  addNumeric(
    "U diagnostic: variation ratio %",
    uDiag.ratio,
    508.5863522156604,
    0.05
  );

  addNumeric(
    "U diagnostic: upper threshold %",
    uDiag.upperLimit,
    155.2276973611499,
    0.05
  );

  addText(
    "U diagnostic: overdispersion",
    uDiag.status,
    "overdispersed"
  );

  // ----------------------------------------------------------
  // X̄–S
  // ----------------------------------------------------------

  addNumeric(
    "Xbar-S: c4 for n=5",
    c4Constant(5),
    0.9399856029866254,
    1e-12
  );

  const equalSigma =
    estimateXbarSSigmaFromSbar(
      [2, 2, 2, 2, 2, 2],
      [5, 5, 5, 5, 5, 5]
    );

  addNumeric(
    "Xbar-S: equal-size sigma estimate",
    equalSigma,
    2.1276921621409737,
    1e-10
  );

  const variableSigma =
    estimateXbarSSigmaFromSbar(
      [2, 2, 2, 2, 2, 2],
      [3, 4, 5, 6, 8, 5]
    );

  addNumeric(
    "Xbar-S: variable-size sigma estimate",
    variableSigma,
    2.1217839603123774,
    1e-10
  );

  addNumeric(
    "Xbar-S: n=3 Xbar LCL",
    100 -
      (
        3 *
        variableSigma /
        Math.sqrt(3)
      ),
    96.32496237805425,
    1e-10
  );

  addNumeric(
    "Xbar-S: n=8 Xbar UCL",
    100 +
      (
        3 *
        variableSigma /
        Math.sqrt(8)
      ),
    102.2504917398246,
    1e-10
  );

  // ----------------------------------------------------------
  // T CHART
  // ----------------------------------------------------------

  const tResult =
    computeTLimits(
      [2, 3, 5, 4]
    );

  addNumeric(
    "T: mean time between events",
    tResult.mean,
    3.5,
    1e-12
  );

  addNumeric(
    "T: UCL",
    tResult.ucl,
    23.126777402861393,
    1e-10
  );

  addNumeric(
    "T: LCL",
    tResult.lcl,
    0,
    1e-12
  );

  // ----------------------------------------------------------
  // G CHART
  // ----------------------------------------------------------

  const gResult =
    computeGLimits(
      [5, 7, 6, 8, 4]
    );

  addNumeric(
    "G: centre line",
    gResult.mean,
    6,
    1e-12
  );

  addNumeric(
    "G: UCL",
    gResult.ucl,
    37
  );

  addNumeric(
    "G: LCL",
    gResult.lcl,
    1
  );

  // ----------------------------------------------------------
  // PARETO
  // ----------------------------------------------------------

  const paretoTest =
    computeParetoData(
      [
        { category: "A" },
        { category: "A" },
        { category: "A" },
        { category: "B" },
        { category: "B" },
        { category: "C" }
      ],
      "category",
      PARETO_COUNT_ROWS
    );

  addBoolean(
    "Pareto: calculation succeeds",
    paretoTest.ok,
    true
  );

  addNumeric(
    "Pareto: total occurrences",
    paretoTest.total,
    6
  );

  addText(
    "Pareto: largest category",
    paretoTest.items?.[0]?.category,
    "A"
  );

  addNumeric(
    "Pareto: largest category count",
    paretoTest.items?.[0]?.count,
    3
  );

  addNumeric(
    "Pareto: categories needed to reach 80%",
    paretoTest.cutoff80Count,
    2
  );

  // ----------------------------------------------------------
  // RESULTS
  // ----------------------------------------------------------

  const passed =
    results.filter(
      r => r.pass
    ).length;

  const failed =
    results.length -
    passed;

  console.group(
    `SimpleSPC regression tests — ${passed}/${results.length} passed`
  );

  console.table(
    results.map(
      r => ({
        Test: r.test,
        Result:
          r.pass
            ? "PASS"
            : "FAIL",

        Actual:
          r.actual,

        Expected:
          r.expected,

        Tolerance:
          r.tolerance
      })
    )
  );

  if (failed === 0) {
    console.log(
      `✅ All ${passed} statistical regression tests passed.`
    );
  } else {
    console.error(
      `❌ ${failed} of ${results.length} statistical regression tests failed.`
    );
  }

  console.groupEnd();

  return {
    passed,
    failed,
    total:
      results.length,

    results
  };
}

window.runSimpleSPCRegressionTests =
  runSimpleSPCRegressionTests;

// ---- AI helper function  -----

function setHelperSectionExpanded(toggleEl, bodyEl, expanded) {
  if (!toggleEl || !bodyEl) return;
  toggleEl.setAttribute("aria-expanded", expanded ? "true" : "false");
  bodyEl.classList.toggle("is-collapsed", !expanded);
}

function toggleHelperSection(toggleEl, bodyEl) {
  if (!toggleEl || !bodyEl) return;
  const isExpanded = toggleEl.getAttribute("aria-expanded") === "true";
  setHelperSectionExpanded(toggleEl, bodyEl, !isExpanded);
}

function updateHelperSectionDefaults(hasChart) {
  // Default behavior:
  // - no chart: General open, My chart collapsed
  // - has chart: General collapsed, My chart open
  setHelperSectionExpanded(spcHelperToggleGeneral, spcHelperGeneralSection, !hasChart);
  setHelperSectionExpanded(spcHelperToggleChart, spcHelperChartSection, !!hasChart);
}

function answerSpcQuestion(question) {
  if (window.SPC_HELPER_LIBRARY && typeof window.SPC_HELPER_LIBRARY.answerQuestion === "function") {
    return window.SPC_HELPER_LIBRARY.answerQuestion(question);
  }

  return "The SPC helper library is not available. Please check that spc-helper-library.js is loaded before spc.js.";
}


function renderHelperState() {
  if (!spcHelperIntro) return;

  // Treat "has chart" as "the currently selected chart has matching analysis".
const helperChartType =
  (typeof getSelectedChartType_NoSideEffects === "function")
    ? (getSelectedChartType_NoSideEffects() || "run")
    : (
        (typeof getSelectedChartType === "function")
          ? (getSelectedChartType() || "run")
          : "run"
      );

const hasChart =
  helperChartType === "run"
    ? !!lastRunAnalysis
    : helperChartType === "xmr"
      ? !!lastXmRAnalysis
      : helperChartType === "xbars"
        ? !!lastXbarSAnalysis
        : ["c", "p", "pprime", "u", "uprime"].includes(helperChartType)
          ? (
              !!lastAttributeAnalysis &&
              lastAttributeAnalysis.chartType === helperChartType
            )
          : ["t", "g"].includes(helperChartType)
            ? (
                !!lastRareAnalysis &&
                lastRareAnalysis.chartType === helperChartType
              )
            : helperChartType === "pareto"
              ? !!lastParetoAnalysis
              : false;

  // 1) Intro text
  if (!hasChart) {
    spcHelperIntro.innerHTML = `
      <div><strong>SPC helper</strong></div>
      <div>Ask a general question before you load any data, or use a suggested prompt below.</div>
    `;
  } else {
    spcHelperIntro.innerHTML = `
      <div><strong>Chart helper</strong></div>
      <div>Use the <strong>My chart</strong> questions for a tailored interpretation.</div>
    `;
  }

  // 2) Suggested questions from external helper library
  const helperQuestions =
    (window.SPC_HELPER_LIBRARY &&
     typeof window.SPC_HELPER_LIBRARY.getSuggestedQuestions === "function")
      ? window.SPC_HELPER_LIBRARY.getSuggestedQuestions(hasChart)
      : { general: [], chart: [] };

  const generalQs = helperQuestions.general || [];
  const chartQs = helperQuestions.chart || [];

  if (spcHelperChipsGeneral) {
    spcHelperChipsGeneral.innerHTML = generalQs
      .map(q => `<button type="button" class="spc-chip" data-q="${escapeHtml(q)}">${escapeHtml(q)}</button>`)
      .join("");
    spcHelperChipsGeneral.classList.remove("is-disabled");
  }

  if (spcHelperChipsChart) {
    spcHelperChipsChart.innerHTML = chartQs
      .map(q => `<button type="button" class="spc-chip" data-q="${escapeHtml(q)}">${escapeHtml(q)}</button>`)
      .join("");

    if (!hasChart) spcHelperChipsChart.classList.add("is-disabled");
    else spcHelperChipsChart.classList.remove("is-disabled");
  }

  updateHelperSectionDefaults(hasChart);
}


// ===============================
// Chart context menu + export tools
// ===============================

const chartContextMenu = document.getElementById("chartContextMenu");

// Track which point index was right-clicked
let contextMenuPointIndex = null;
let contextMenuMeanLabelSegmentIndex = null;

// Helper: hide menu
function hideChartContextMenu() {
  if (!chartContextMenu) return;
  chartContextMenu.style.display = "none";
  contextMenuPointIndex = null;
  contextMenuMeanLabelSegmentIndex = null;
}

// Helper: show menu at cursor, clamped to viewport
function showChartContextMenu(clientX, clientY, pointIndex) {
  if (!chartContextMenu) return;

  contextMenuPointIndex = pointIndex;

  // Which charts support splits?
  const chartType =
    (typeof getSelectedChartType_NoSideEffects === "function")
      ? (getSelectedChartType_NoSideEffects() || "run")
      : ((typeof getSelectedChartType === "function") ? (getSelectedChartType() || "run") : "run");

  const supportsSplits = ["run", "xmr", "c", "p", "pprime", "u", "uprime", "xbars", "t", "g"].includes(chartType);


  // User annotations are not supported on Pareto charts.
  // Hide the annotation submenu, but keep it available
  // for all other chart types.
  const annotationMenuGroup =
    chartContextMenu
      .querySelector('button[data-action="addAnnotation"]')
      ?.closest(".context-submenu-wrap");

  if (annotationMenuGroup) {
    annotationMenuGroup.style.display =
      chartType === "pareto" ? "none" : "";

    // Hide the accompanying separator too.
    const separator = annotationMenuGroup.nextElementSibling;

    if (separator?.classList.contains("menu-sep")) {
      separator.style.display =
        chartType === "pareto" ? "none" : "";
    }
  }



  // ---------- XmR mean labels ----------
const xmrMeanLabelsGroup =
  chartContextMenu.querySelector('[data-role="xmrMeanLabelsGroup"]');

const xmrMeanLabelsBtn =
  chartContextMenu.querySelector(
    'button[data-action="toggleXmRMeanLabels"]'
  );

const supportsCentreLineLabels =
  [
    "run",
    "xmr",
    "c",
    "p",
    "pprime",
    "u",
    "uprime",
    "xbars",
    "t",
    "g"
  ].includes(chartType);

if (xmrMeanLabelsGroup) {
  xmrMeanLabelsGroup.style.display =
    supportsCentreLineLabels
      ? ""
      : "none";
}

if (xmrMeanLabelsBtn) {
  xmrMeanLabelsBtn.textContent =
    showXmRMeanLabels
  ? "Hide centre-line labels"
  : "Show centre-line labels";
}

const hideXmRMeanLabelBtn =
  chartContextMenu.querySelector(
    'button[data-action="hideXmRMeanLabel"]'
  );

const restoreXmRMeanLabelsBtn =
  chartContextMenu.querySelector(
    'button[data-action="restoreXmRMeanLabels"]'
  );

if (hideXmRMeanLabelBtn) {
  const canHideThisLabel =
  supportsCentreLineLabels &&
  showXmRMeanLabels &&
    contextMenuMeanLabelSegmentIndex !== null &&
    !xmrHiddenMeanLabels.has(
      contextMenuMeanLabelSegmentIndex
    );

  hideXmRMeanLabelBtn.style.display =
    canHideThisLabel ? "" : "none";
}

if (restoreXmRMeanLabelsBtn) {
  restoreXmRMeanLabelsBtn.style.display =
    supportsCentreLineLabels &&
xmrHiddenMeanLabels.size > 0
      ? ""
      : "none";
}

  // Current x-label at the clicked point (if any)
  const labels = currentChart?.data?.labels || [];
  const noPoint = (pointIndex === null || pointIndex === undefined);
  const xLabel = (!noPoint && labels && labels[pointIndex] !== undefined) ? labels[pointIndex] : null;
  const annsAtPoint = xLabel ? getAnnotationsAtDate(xLabel) : [];
  const hasAnnotationAtPoint = annsAtPoint.length > 0;
  const hasAnyAnnotations = Array.isArray(annotations) && annotations.length > 0;

  // ---------- Annotation buttons ----------
  const addAnnotationBtn = chartContextMenu.querySelector('button[data-action="addAnnotation"]');
  const editAnnotationBtn = chartContextMenu.querySelector('button[data-action="editAnnotation"]');
  const deleteAnnotationBtn = chartContextMenu.querySelector('button[data-action="deleteAnnotation"]');
  const clearAnnotationsBtn = chartContextMenu.querySelector('button[data-action="clearAnnotations"]');

  if (addAnnotationBtn) {
    if (!addAnnotationBtn.dataset.defaultTitle) {
      addAnnotationBtn.dataset.defaultTitle = addAnnotationBtn.getAttribute("title") || "";
    }

    addAnnotationBtn.disabled = noPoint;

    if (noPoint) {
      addAnnotationBtn.title = "Right-click near a data point to add an annotation.";
    } else {
      addAnnotationBtn.title = addAnnotationBtn.dataset.defaultTitle;
    }
  }

  if (editAnnotationBtn) {
    if (!editAnnotationBtn.dataset.defaultTitle) {
      editAnnotationBtn.dataset.defaultTitle = editAnnotationBtn.getAttribute("title") || "";
    }

    editAnnotationBtn.disabled = noPoint || !hasAnnotationAtPoint;

    if (noPoint) {
      editAnnotationBtn.title = "Right-click near a data point to edit an annotation.";
    } else if (!hasAnnotationAtPoint) {
      editAnnotationBtn.title = "There is no annotation at this point to edit.";
    } else {
      editAnnotationBtn.title = editAnnotationBtn.dataset.defaultTitle;
    }
  }

  if (deleteAnnotationBtn) {
    if (!deleteAnnotationBtn.dataset.defaultTitle) {
      deleteAnnotationBtn.dataset.defaultTitle = deleteAnnotationBtn.getAttribute("title") || "";
    }

    deleteAnnotationBtn.disabled = noPoint || !hasAnnotationAtPoint;

    if (noPoint) {
      deleteAnnotationBtn.title = "Right-click near a data point to delete an annotation.";
    } else if (!hasAnnotationAtPoint) {
      deleteAnnotationBtn.title = "There is no annotation at this point to delete.";
    } else {
      deleteAnnotationBtn.title = deleteAnnotationBtn.dataset.defaultTitle;
    }
  }

  if (clearAnnotationsBtn) {
    if (!clearAnnotationsBtn.dataset.defaultTitle) {
      clearAnnotationsBtn.dataset.defaultTitle = clearAnnotationsBtn.getAttribute("title") || "";
    }

    clearAnnotationsBtn.disabled = !hasAnyAnnotations;

    if (!hasAnyAnnotations) {
      clearAnnotationsBtn.title = "There are no annotations to clear.";
    } else {
      clearAnnotationsBtn.title = clearAnnotationsBtn.dataset.defaultTitle;
    }
  }

  // ---------- Limit-calculation exclusion ----------
  const limitExclusionGroup =
    chartContextMenu.querySelector(
      '[data-role="limitExclusionGroup"]'
    );

  const excludeFromLimitsBtn =
    chartContextMenu.querySelector(
      'button[data-action="excludeFromLimits"]'
    );

  const includeInLimitsBtn =
    chartContextMenu.querySelector(
      'button[data-action="includeInLimits"]'
    );

  const limitState =
    getCurrentLimitExclusionMenuState(
      pointIndex
    );

  if (limitExclusionGroup) {
    limitExclusionGroup.style.display =
      limitState.supported
        ? ""
        : "none";
  }

  if (excludeFromLimitsBtn) {
    excludeFromLimitsBtn.style.display =
      limitState.supported &&
      !limitState.excluded
        ? ""
        : "none";

    excludeFromLimitsBtn.disabled =
      !limitState.eligible;

    excludeFromLimitsBtn.textContent =
      chartType === "xbars"
        ? "Exclude this subgroup from limit calculations…"
        : "Exclude from limit calculations…";

    excludeFromLimitsBtn.title =
      limitState.eligible
        ? "Keep this point visible but do not use it to estimate the centre line or control limits."
        : (
            limitState.reason ||
            "Right-click near a plotted point."
          );
  }

  if (includeInLimitsBtn) {
    includeInLimitsBtn.style.display =
      limitState.supported &&
      limitState.excluded
        ? ""
        : "none";

    includeInLimitsBtn.disabled =
      !limitState.key;

    includeInLimitsBtn.textContent =
      chartType === "xbars"
        ? "Include this subgroup in limit calculations again"
        : "Include in limit calculations again";
  }

   // ---------- Split submenu ----------
  const splitsParentBtn = chartContextMenu.querySelector('[data-role="splitsParent"]');
  const splitsSubmenu = chartContextMenu.querySelector('[data-role="splitsSubmenu"]');
  const splitDynamicItems = chartContextMenu.querySelector('[data-role="splitDynamicItems"]');
  const addSplitBtn = chartContextMenu.querySelector('button[data-action="addSplit"]');

  const hasSplits = Array.isArray(splits) && splits.length > 0;

  if (splitsParentBtn) {
    if (!splitsParentBtn.dataset.defaultTitle) {
      splitsParentBtn.dataset.defaultTitle = splitsParentBtn.getAttribute("title") || "";
    }

    splitsParentBtn.disabled = !supportsSplits;

    if (!supportsSplits) {
      splitsParentBtn.title = "Splits are not available for this chart type.";
    } else {
      splitsParentBtn.title = splitsParentBtn.dataset.defaultTitle;
    }
  }

  if (addSplitBtn) {
    if (!addSplitBtn.dataset.defaultTitle) {
      addSplitBtn.dataset.defaultTitle = addSplitBtn.getAttribute("title") || "";
    }

    addSplitBtn.disabled = !supportsSplits || noPoint;

    if (!supportsSplits) {
      addSplitBtn.title = "Splits are not available for this chart type.";
    } else if (noPoint) {
      addSplitBtn.title = "Right-click near a data point to add a split.";
    } else {
      addSplitBtn.title = addSplitBtn.dataset.defaultTitle;
    }
  }

  if (splitDynamicItems) {
    splitDynamicItems.innerHTML = "";

    if (!supportsSplits) {
      const noSupportBtn = document.createElement("button");
      noSupportBtn.type = "button";
      noSupportBtn.disabled = true;
      noSupportBtn.textContent = "Splits not available for this chart type";
      splitDynamicItems.appendChild(noSupportBtn);
    } else {
      if (hasSplits) {
        const labels = currentChart?.data?.labels || [];

        splits
          .slice()
          .sort((a, b) => a - b)
          .forEach((idx) => {
            const btn = document.createElement("button");
            btn.type = "button";
            btn.dataset.action = "removeSplit";
            btn.dataset.splitIndex = String(idx);

            const label = labels[idx] !== undefined ? labels[idx] : `point ${idx + 1}`;
            btn.textContent = `Clear split after ${label} (point ${idx + 1})`;
            btn.title = `Remove the split after ${label}.`;

            splitDynamicItems.appendChild(btn);
          });

        const sep = document.createElement("div");
        sep.className = "menu-sep";
        splitDynamicItems.appendChild(sep);

        const clearAllBtn = document.createElement("button");
        clearAllBtn.type = "button";
        clearAllBtn.dataset.action = "clearSplits";
        clearAllBtn.textContent = "Clear all splits";
        clearAllBtn.title = "Remove all splits and return to a single set of limits.";
        splitDynamicItems.appendChild(clearAllBtn);
      } else {
        const noSplitsBtn = document.createElement("button");
        noSplitsBtn.type = "button";
        noSplitsBtn.disabled = true;
        noSplitsBtn.textContent = "No splits to clear";
        noSplitsBtn.title = "There are no splits to clear.";
        splitDynamicItems.appendChild(noSplitsBtn);

        const sep = document.createElement("div");
        sep.className = "menu-sep";
        splitDynamicItems.appendChild(sep);

        const clearAllBtn = document.createElement("button");
        clearAllBtn.type = "button";
        clearAllBtn.dataset.action = "clearSplits";
        clearAllBtn.textContent = "Clear all splits";
        clearAllBtn.title = "Remove all splits and return to a single set of limits.";
        clearAllBtn.disabled = true;
        splitDynamicItems.appendChild(clearAllBtn);
      }
    }
  }

  chartContextMenu.style.display = "block";
  chartContextMenu.style.left = "0px";
  chartContextMenu.style.top = "0px";

  // Clamp so it stays on-screen
  const menuRect = chartContextMenu.getBoundingClientRect();
  const pad = 8;
  let x = clientX;
  let y = clientY;

  if (x + menuRect.width + pad > window.innerWidth) x = window.innerWidth - menuRect.width - pad;
  if (y + menuRect.height + pad > window.innerHeight) y = window.innerHeight - menuRect.height - pad;
  if (x < pad) x = pad;
  if (y < pad) y = pad;

  chartContextMenu.style.left = `${x}px`;
  chartContextMenu.style.top = `${y}px`;
}

// Helper: get the nearest chart point index from a mouse event
function getNearestPointIndexFromEvent(evt) {
  if (!currentChart) return null;

  const elements = currentChart.getElementsAtEventForMode(
    evt,
    "nearest",
    { intersect: false }, // <-- important
    true
  );


  if (!elements || elements.length === 0) return null;

  // Chart.js v3+: element has .index
  const idx = elements[0].index;
  return Number.isFinite(idx) ? idx : null;
}

// ---- Split helpers ----
function addSplitAfterIndex(splitAfterIndex) {
  if (!Number.isFinite(splitAfterIndex)) return;

  // can’t split after last point
  const labels = currentChart?.data?.labels || [];
  if (labels.length === 0) return;
  if (splitAfterIndex < 0 || splitAfterIndex >= labels.length - 1) {
    alert("You can’t split after the last point.");
    return;
  }

  // avoid duplicates
  if (!splits.includes(splitAfterIndex)) {
    splits.push(splitAfterIndex);
    splits.sort((a, b) => a - b);
  }

  // keep the dropdown in sync (if present)
  if (labels && labels.length) {
    populateSplitOptions(labels);
  }

  // redraw with new split
  if (generateButton) generateButton.click();
}

function removeSplitAtIndex(splitAfterIndex) {
  if (!Array.isArray(splits) || !splits.length) return false;

  const i = splits.indexOf(splitAfterIndex);
  if (i === -1) return false;

  splits.splice(i, 1);

  const labels = currentChart?.data?.labels || [];
  if (labels && labels.length) {
    populateSplitOptions(labels);
  }

  if (splitPointSelect) splitPointSelect.value = "";

  if (generateButton) generateButton.click();
  return true;
}

// ---- Export helpers ----

// Return the canvases to export (main + MR if shown)
function getExportCanvases() {
  const canvases = [];
  if (chartCanvas) canvases.push(chartCanvas);

  const showMR = showMRCheckbox ? showMRCheckbox.checked : false;
  const mrVisible = mrPanel && mrPanel.style.display !== "none";
  if (showMR && mrVisible && mrChartCanvas) canvases.push(mrChartCanvas);

  return canvases;
}

function renderCapabilityBadgeToCanvas(ctx, x, y, maxWidth) {
  if (!capabilityDiv) return 0;

  const txt = (capabilityDiv.innerText || "").trim();
  if (!txt) return 0;

  const baseFont = "system-ui, -apple-system, Segoe UI, sans-serif";

  function setFont(size, bold) {
    ctx.font = `${bold ? "700" : "400"} ${size}px ${baseFont}`;
  }

  function drawWrapped(text, startX, startY, maxW, size, bold) {
    setFont(size, bold);
    const lineHeight = Math.round(size * 1.35);
    const words = (text || "").split(/\s+/).filter(Boolean);

    let line = "";
    let yy = startY;

    for (const w of words) {
      const test = line ? `${line} ${w}` : w;
      const width = ctx.measureText(test).width;

      if (width <= maxW) {
        line = test;
      } else {
        if (line) {
          ctx.fillText(line, startX, yy);
          yy += lineHeight;
        }
        line = w;
      }
    }

    if (line) {
      ctx.fillText(line, startX, yy);
      yy += lineHeight;
    }

    return yy - startY;
  }

  // Decide a background colour (match your UI roughly)
  const isStableBadge = txt.toLowerCase().includes("process capability");
  const bg = isStableBadge ? "#fff59d" : "#ffe0b2";

  // Box layout
  const pad = 14;
  const boxW = Math.min(maxWidth, 520); // keep it readable, like your on-page badge
  const innerW = boxW - pad * 2;

  // Split into logical parts (header / big number / small note)
  const lines = txt.split("\n").map(s => s.trim()).filter(Boolean);

  const header = lines[0] || "";
  // Try to find the big % line (often the 2nd line)
  const bigLine = (lines.length >= 2 && /%/.test(lines[1])) ? lines[1] : "";
  const rest = lines.slice(bigLine ? 2 : 1).join(" ");

  // --- Measure height with a dry run on current ctx ---
  let h = 0;
  h += pad;

  // Header
  h += drawWrapped(header, 0, 0, innerW, 14, true);
  h += 10; // IMPORTANT: extra gap after bold header (fixes “uneven spacing” look)

  // Big value (if present)
  if (bigLine) {
    h += drawWrapped(bigLine, 0, 0, innerW, 22, true);
    h += 8;
  }

  // Small note
  if (rest) {
    h += drawWrapped(rest, 0, 0, innerW, 12, false);
  }

  h += pad;

  // Draw the box
  ctx.save();
  ctx.fillStyle = bg;
  ctx.strokeStyle = "#cccccc";
  ctx.lineWidth = 1;

  ctx.fillRect(x, y, boxW, h);
  ctx.strokeRect(x, y, boxW, h);

  // Draw text inside
  let yy = y + pad;
  ctx.fillStyle = "#111";

  yy += drawWrapped(header, x + pad, yy, innerW, 14, true);
  yy += 10;

  if (bigLine) {
    yy += drawWrapped(bigLine, x + pad, yy, innerW, 22, true);
    yy += 8;
  }

  if (rest) {
    yy += drawWrapped(rest, x + pad, yy, innerW, 12, false);
  }

  ctx.restore();
  return h;
}

function renderSummaryToCanvas(ctx, x, y, maxWidth) {
  if (!summaryDiv) return 0;

  const root = summaryDiv.cloneNode(true);

  const baseFont = "system-ui, -apple-system, Segoe UI, sans-serif";
  const styles = {
    h3: { size: 18, bold: true, gapTop: 6, gapBottom: 8 },
    h4: { size: 14, bold: true, gapTop: 10, gapBottom: 6 },
    p:  { size: 13, bold: false, gapTop: 6, gapBottom: 6 },
    li: { size: 13, bold: false, gapTop: 2, gapBottom: 2 }
  };

  function setFont(size, bold) {
    ctx.font = `${bold ? "700" : "400"} ${size}px ${baseFont}`;
  }

  function drawWrappedText(text, startX, startY, size, bold, indent = 0, bullet = false) {
    setFont(size, bold);

    const clean = String(text || "")
          .replace(/\r\n|\r/g, "\n")   // normalize newlines
          .replace(/\n{2,}/g, "\n")   // collapse multiple blank lines
          .replace(/[ \t]+/g, " ")    // collapse spaces/tabs (but NOT newlines)
          .trim();

    if (!clean) return 0;

    const words = clean.split(" ");
    const lineHeight = Math.round(size * 1.35);
    const bulletText = bullet ? "• " : "";

    const usableWidth = Math.max(80, maxWidth - indent);
    const drawX = startX + indent;

    let line = "";
    let yy = startY;

    for (const w of words) {
      const test = line ? `${line} ${w}` : w;
      const prefix = (line === "") ? bulletText : "";
      const width = ctx.measureText(prefix + test).width;

      if (width <= usableWidth) {
        line = test;
      } else {
        ctx.fillText(((line === "") ? bulletText : "") + line, drawX, yy);
        yy += lineHeight;
        line = w;
      }
    }

    if (line) {
      ctx.fillText(bulletText + line, drawX, yy);
      yy += lineHeight;
    }

    return yy - startY;
  }

  let cursorY = y;
  ctx.fillStyle = "#111";

  const children = Array.from(root.children);

  for (const node of children) {
    const tag = node.tagName ? node.tagName.toLowerCase() : "";
    const text = node.innerText || "";

    if (tag === "h3" || tag === "h4") {
      const st = styles[tag];
      cursorY += st.gapTop;
      cursorY += drawWrappedText(text, x, cursorY, st.size, st.bold, 0, false);
      cursorY += st.gapBottom;
      continue;
    }

    if (tag === "p") {
      const st = styles.p;
      cursorY += st.gapTop;
      cursorY += drawWrappedText(text, x, cursorY, st.size, st.bold, 0, false);
      cursorY += st.gapBottom;
      continue;
    }

    if (tag === "ul") {
      const items = Array.from(node.querySelectorAll(":scope > li"));
      for (const li of items) {
        const st = styles.li;
        cursorY += st.gapTop;
        cursorY += drawWrappedText(li.innerText || "", x, cursorY, st.size, false, 18, true);
        cursorY += st.gapBottom;
      }
      cursorY += 4;
      continue;
    }

    // fallback
    const st = styles.p;
    cursorY += st.gapTop;
    cursorY += drawWrappedText(text, x, cursorY, st.size, st.bold, 0, false);
    cursorY += st.gapBottom;
  }

  return cursorY - y;
}

async function prepareChartsForExport() {
  if (document.fonts && document.fonts.ready) {
    await document.fonts.ready;
  }

  const mainContainer = chartCanvas?.parentElement;
  const mrContainer = mrChartCanvas?.parentElement;

  const backups = {
    mainWidth: mainContainer?.style.width || "",
    mainHeight: mainContainer?.style.height || "",
    mrWidth: mrContainer?.style.width || "",
    mrHeight: mrContainer?.style.height || ""
  };

  if (mainContainer) {
    mainContainer.style.width = "1200px";
    mainContainer.style.height = "640px";
  }

  if (mrContainer) {
    mrContainer.style.width = "1200px";
    mrContainer.style.height = "360px";
  }

  if (currentChart) {
    currentChart.resize(1200, 640);
    currentChart.update("none");
  }

  if (mrChart) {
    mrChart.resize(1200, 360);
    mrChart.update("none");
  }

  await new Promise(requestAnimationFrame);

  return function restoreChartsAfterExport() {
    if (mainContainer) {
      mainContainer.style.width = backups.mainWidth;
      mainContainer.style.height = backups.mainHeight;
    }

    if (mrContainer) {
      mrContainer.style.width = backups.mrWidth;
      mrContainer.style.height = backups.mrHeight;
    }

    if (currentChart) {
      currentChart.resize();
      currentChart.update("none");
    }

    if (mrChart) {
      mrChart.resize();
      mrChart.update("none");
    }
  };
}

// Build one combined image from multiple canvases (stacked vertically).
// Optionally add summary text under the charts.
function buildCompositeCanvas({ includeSummaryText }) {
  const canvases = getExportCanvases();
  if (!canvases.length) return null;

  const widths = canvases.map(c => c.width);
  const heights = canvases.map(c => c.height);

  const outWidth = Math.max(...widths);
  const chartsHeight = heights.reduce((a, b) => a + b, 0);
  const padding = 16;

  const includeSummary =
    !!includeSummaryText &&
    summaryDiv &&
    (summaryDiv.innerText || "").trim().length > 0;

  const includeCapability =
    !!includeSummaryText &&
    capabilityDiv &&
    (capabilityDiv.innerText || "").trim().length > 0;

  // --- Dry-run summary height ---
  let summaryHeight = 0;
  if (includeSummary) {
    const tmp = document.createElement("canvas");
    tmp.width = outWidth;
    tmp.height = 5000;
    const tctx = tmp.getContext("2d");
    tctx.fillStyle = "#111";
    const used = renderSummaryToCanvas(tctx, padding, padding, outWidth - padding * 2);
    summaryHeight = padding + used + padding + 1; // + separator
  }

  // --- Dry-run capability height ---
  let capabilityHeight = 0;
  if (includeCapability) {
    const tmp = document.createElement("canvas");
    tmp.width = outWidth;
    tmp.height = 2000;
    const tctx = tmp.getContext("2d");
    capabilityHeight = padding + renderCapabilityBadgeToCanvas(tctx, padding, padding, outWidth - padding * 2) + padding;
  }

  const out = document.createElement("canvas");
  const ctx = out.getContext("2d");

  out.width = outWidth;
  out.height = chartsHeight + (includeSummary ? summaryHeight : 0) + (includeCapability ? capabilityHeight : 0);

  // Background
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, out.width, out.height);

  // Charts
  let y = 0;
  canvases.forEach((c) => {
    const x = Math.round((outWidth - c.width) / 2);
    ctx.drawImage(c, x, y);
    y += c.height;
  });

  // Summary
  if (includeSummary) {
    ctx.fillStyle = "#eef2f6";
    ctx.fillRect(0, y, out.width, 1);
    y += padding;
    const used = renderSummaryToCanvas(ctx, padding, y, outWidth - padding * 2);
    y += used + padding;
  }

  // Capability badge (if present)
  if (includeCapability) {
    ctx.fillStyle = "#eef2f6";
    ctx.fillRect(0, y, out.width, 1);
    y += padding;
    const used = renderCapabilityBadgeToCanvas(ctx, padding, y, outWidth - padding * 2);
    y += used + padding;
  }

  return out;
}


async function copyCanvasToClipboard(canvas) {
  if (!canvas) return;

  // Modern clipboard image API
  if (navigator.clipboard && window.ClipboardItem) {
    const blob = await new Promise(resolve => canvas.toBlob(resolve, "image/png"));
    if (!blob) throw new Error("Failed to create image.");
    await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]);
    return;
  }

  // Fallback
  alert("Copy to clipboard is not supported in this browser. Try 'Save chart(s) as…' instead.");
}

function downloadCanvasAsPng(canvas, filename) {
  if (!canvas) return;

  // More reliable across browsers than link.click() on a detached node
  canvas.toBlob((blob) => {
    if (!blob) {
      alert("Sorry — your browser could not export this image.");
      return;
    }

    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;

    // Attach to DOM for Safari / locked-down contexts
    document.body.appendChild(link);
    link.click();
    link.remove();

    // Clean up the object URL
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }, "image/png");
}

// ---- Existing top button: Download chart as PNG ----
// Update to download chart(s) (main + MR if shown), as one image.
if (downloadBtn) {
  downloadBtn.addEventListener("click", () => {
    if (!currentChart) {
      alert("Please generate a chart first.");
      return;
    }
    const composite = buildCompositeCanvas({ includeSummaryText: false });
    downloadCanvasAsPng(composite, "spc-charts.png");
  });
}


function toggleChartWizard(forceOpen) {
  const modal = document.getElementById("chartWizardModal");
  if (!modal) return;

  const isOpen = modal.classList.contains("visible");
  const shouldOpen = typeof forceOpen === "boolean" ? forceOpen : !isOpen;

  modal.classList.toggle("visible", shouldOpen);
  modal.setAttribute("aria-hidden", shouldOpen ? "false" : "true");
  document.body.classList.toggle("modal-open", shouldOpen);

  if (shouldOpen) {
    const closeBtn = modal.querySelector(".modal-close");
    if (closeBtn) closeBtn.focus();
  }
}


/* ============================================================
   CHART SETUP MODAL (shared instructions + per-chart options)
   ============================================================ */

let tChartInputMode = (() => {
  try {
    return localStorage.getItem("spc_tChartInputMode") || "eventDates";
  } catch {
    return "eventDates";
  }
})();





function setAutoShowChartSetupModal(shouldShow) {
  try {
    localStorage.setItem("spc_hideChartSetupModal", shouldShow ? "false" : "true");
  } catch {}
}


function openChartSetupForCurrentType() {
  const chartType = (typeof getSelectedChartType_NoSideEffects === "function")
    ? getSelectedChartType_NoSideEffects()
    : (document.querySelector("input[name='chartType']:checked")?.value || "run");

  renderChartSetupModal(chartType);
  toggleChartSetupModal(true);
}


function toggleRuleExplainerModal(forceOpen) {
  const modal = document.getElementById("ruleExplainerModal");
  if (!modal) return;

  const isOpen = modal.classList.contains("visible");
  const shouldOpen = typeof forceOpen === "boolean" ? forceOpen : !isOpen;

  modal.classList.toggle("visible", shouldOpen);
  modal.setAttribute("aria-hidden", shouldOpen ? "false" : "true");
  document.body.classList.toggle("modal-open", shouldOpen);

  if (shouldOpen) {
    const closeBtn = modal.querySelector(".modal-close");
    if (closeBtn) closeBtn.focus();
  }
}

function getChartTypeDisplayName(chartType) {
  const names = {
    run: "Run chart",
    xmr: "X-MR chart",
    c: "C chart",
    p: "P chart",
    pprime: "P′ chart",
    u: "U chart",
    uprime: "U′ chart",
    xbars: "X̄-S chart",
    t: "T chart",
    g: "G chart",
    pareto: "Pareto chart"
  };
  return names[chartType] || "This chart";
}

function describeRuleStatus(status, labels) {
  if (status === "on") {
    return `<li><strong>${labels.name}:</strong> used by default.</li>`;
  }
  if (status === "optional") {
    return `<li><strong>${labels.name}:</strong> available in advanced settings, off by default.</li>`;
  }
  if (status === "warn") {
    return `<li><strong>${labels.name}:</strong> advanced only, off by default, with a warning before use.</li>`;
  }
  return `<li><strong>${labels.name}:</strong> not offered for this chart type.</li>`;
}

function renderRuleExplainerModal(chartType) {
  const body = document.getElementById("ruleExplainerBody");
  const subtitle = document.getElementById("ruleExplainerSubtitle");
  if (!body || !subtitle) return;

  const policy = getRulePolicy(chartType);
  const chartName = getChartTypeDisplayName(chartType);

  subtitle.textContent =
    `${chartName}: this tool uses a conservative, chart-aware rule policy designed to reduce false alerts.`;

  let whyText = "";
  if (
  chartType === "c" ||
  chartType === "p" ||
  chartType === "pprime" ||
  isUFamilyChartType(chartType)
) {
    whyText =
      "Attribute charts use a conservative rule set here. Zone rules are blocked because they can create misleading alerts, especially when limits vary or counts are low.";
  } else if (chartType === "t" || chartType === "g") {
    whyText =
      "Rare-event charts are naturally irregular. Run and trend rules can create false signals, so they are advanced-only and warning-gated.";
  } else if (chartType === "xmr") {
    whyText =
      "XmR charts are used for individual measurements over time. SimpleSPC keeps the default interpretation deliberately simple, while additional pattern rules and calculation options are available for users who need them.";
  } else if (chartType === "xbars") {
    whyText =
      "X̄-S charts are the chart family where advanced rule sets are most defensible, but the default remains conservative to reduce false alarms.";
  } else if (chartType === "run") {
    whyText =
      "Run charts do not use control limits. The tool keeps interpretation simple and conservative by default.";
  } else if (chartType === "pareto") {
 	 whyText =
    		"A Pareto chart is not a time-ordered control chart, so SPC signal rules do not apply. " +
    		"It ranks categories to help focus investigation and improvement effort.";
  } else
	{
    whyText =
      "This tool keeps the default rule set simple and conservative to reduce false alerts.";
  }

  let xmrCalculationHtml = "";

if (chartType === "xmr") {
  const screeningOn =
    screenLargeMRsCheckbox
      ? screenLargeMRsCheckbox.checked
      : true;

  xmrCalculationHtml = `
    <hr style="margin:0.9rem 0;">

    <div class="hint">
      <strong>How XmR limits are set</strong>
    </div>

    <div class="hint" style="margin-top:0.5rem;">
      XmR charts estimate routine variation from the differences between
      consecutive points. These differences are called
      <strong>moving ranges</strong>.
    </div>

    <div class="hint" style="margin-top:0.5rem;">
      A very large jump between two points can make the estimated variation
      too large and produce control limits that are too wide.
      By default, FreeSPC screens unusually large moving ranges when setting
      XmR limits. The original data points and moving ranges are still shown
      on the chart.
    </div>

    <div class="hint" style="margin-top:0.5rem;">
      <strong>Technical detail:</strong>
      FreeSPC uses the Nelson method. Moving ranges greater than
      3.5 times the initial average moving range are not used when estimating
      routine variation.
    </div>

    <div class="hint" style="margin-top:0.5rem;">
      <strong>Current setting:</strong>
      moving-range screening is
      <strong>${screeningOn ? "ON" : "OFF"}</strong>.
      You can change this under <em>Advanced detection options</em>.
    </div>

    <div class="hint" style="margin-top:0.75rem;">
      <strong>What happens when you add a split?</strong>
      A split tells FreeSPC that you want to treat the data before and after
      that point as different process periods. Each period gets its own mean
      and control limits, and the moving range across the split is not used.
    </div>

    <div class="hint" style="margin-top:0.5rem;">
      FreeSPC needs at least
      <strong>${MIN_XMR_LIMIT_POINTS} points</strong>
      to interpret the XmR limits for a period. If a split creates a shorter
      period, the limits are shown as provisional and FreeSPC does not label
      that period as showing routine or special-cause variation yet.
    </div>
  `;
}

  body.innerHTML = `
  <div class="hint">
    <strong>${chartName}</strong>
  </div>

  <div class="hint" style="margin-top:0.5rem;">
    ${whyText}
  </div>

  ${xmrCalculationHtml}

  <hr style="margin:0.9rem 0;">

  <div class="hint">
    <strong>What this chart checks by default</strong>
  </div>
    <ul style="margin-top:0.5rem;">
      ${describeRuleStatus(policy.beyondLimits, { name: "Beyond limits" })}
      ${describeRuleStatus(policy.runShift, { name: "Run rule" })}
      ${describeRuleStatus(policy.trend, { name: "Trend rule" })}
      ${describeRuleStatus(policy.zone23, { name: "2 of 3 in the outer third" })}
      ${describeRuleStatus(policy.zone45, { name: "4 of 5 beyond 1-sigma" })}
    </ul>

    <hr style="margin:0.9rem 0;">

    <div class="hint">
      <strong>Why some rules are limited</strong>
    </div>
    <div class="hint" style="margin-top:0.5rem;">
      This tool is designed for reliable interpretation in healthcare and public-service settings.
      It prefers fewer, more trustworthy signals over a noisier rule set that may create false alerts.
    </div>

    <div class="hint" style="margin-top:0.5rem;">
      A “signal” can trigger meetings, concern, investigation, or action. To reduce wasted effort and alert fatigue,
      some rules are restricted to chart types where they are more dependable.
    </div>

    <hr style="margin:0.9rem 0;">

    <div class="hint">
      <strong>Plain-English rule definitions</strong>
    </div>

    <div class="hint" style="margin-top:0.5rem;">
      <strong>Beyond limits:</strong> a point above the upper limit or below the lower limit.
    </div>

    <div class="hint" style="margin-top:0.5rem;">
      <strong>Run rule:</strong> a sustained run of points on one side of the centre line.
    </div>

    <div class="hint" style="margin-top:0.5rem;">
      <strong>Trend rule:</strong> a sustained pattern of consecutive increases or decreases.
    </div>

    <div class="hint" style="margin-top:0.5rem;">
      <strong>Zone rules:</strong> extra pattern rules based on how far points sit from the centre line.
      These are the least portable rules, so they are only offered on chart types where they are more defensible.
    </div>
  `;
}

function openRuleExplainerForCurrentChart() {
  const chartType =
    (typeof getSelectedChartType_NoSideEffects === "function")
      ? (getSelectedChartType_NoSideEffects() || "run")
      : ((typeof getSelectedChartType === "function") ? (getSelectedChartType() || "run") : "run");

  renderRuleExplainerModal(chartType);
  toggleRuleExplainerModal(true);
}

// -----------------------------
// Chart chooser wizard (Help me choose)
// -----------------------------

const chartWizardBody = document.getElementById("chartWizardBody");

const chartWizardState = {
  step: 0,
  answers: {},
  recommendation: null
};

function setChartType(chartType) {
  const radio = document.querySelector(`input[name="chartType"][value="${chartType}"]`);
  if (radio) {
    radio.checked = true;
    // If chart type is in the "More chart types" details, open it so the selection is visible
    const details = document.getElementById("moreChartTypesDetails");
    if (details && ["c", "p", "pprime", "u", "uprime", "xbars", "t", "g", "pareto"].includes(chartType)) {
      details.open = true;
    }
    updateUIForChartType(chartType);
  }
}

function hideNewDataChartPrompt() {
  if (newDataChartPrompt) {
    newDataChartPrompt.style.display = "none";
  }
}

function showNewDataChartPrompt() {
  if (!newDataChartPrompt) return;

  newDataChartPrompt.style.display = "block";

  // Move naturally to the next step in the workflow: chart choice.
  const chartTabButton =
    document.querySelector('.tab-btn[data-tab="tab-chart"]');

  if (chartTabButton) {
    chartTabButton.click();
  }
}

function startChartWizard() {
  chartWizardState.step = 0;
  chartWizardState.answers = {};
  chartWizardState.recommendation = null;
  renderChartWizard();
}

function closeChartWizard() {
  toggleChartWizard(false);
}

function wizardBack() {
  // If on the results screen, go back to the previous question screen
  if (chartWizardState.step === 99) {
    chartWizardState.recommendation = null;
    chartWizardState.step = 1;
    renderChartWizard();
    return;
  }

  // Otherwise go back one normal step
  if (chartWizardState.step > 0) {
    chartWizardState.step -= 1;
  }

  renderChartWizard();
}

function wizardNext(answerKey, answerValue) {
  chartWizardState.answers[answerKey] = answerValue;
  chartWizardState.step += 1;
  renderChartWizard();
}

function finishWizard(recommendation) {
  chartWizardState.recommendation = recommendation;
  chartWizardState.step = 99; // results screen
  renderChartWizard();
}

// Core decision logic: minimal questions, plain language
function computeRecommendation(answers) {
  // answers.kind: measurement | count | proportion | rare | unsure
  // answers.measurementShape: single | subgroups | unsure
  // answers.countOpportunity: constant | varies | unsure
  // answers.proportionHasDenom: yes | no | unsure
  // answers.rareType: time | opportunities | unsure

  switch (answers.kind) {
    case "measurement": {
      if (answers.measurementShape === "subgroups") return { chartType: "xbars", label: "X̄–S", reason: "You have multiple measurements per time point (subgroups)." };
      return { chartType: "xmr", label: "XmR", reason: "You have one measurement per time point." };
    }

    case "count": {
  if (
    answers.countOpportunity ===
    "varies"
  ) {
    return {
      chartType: "u",
      label: "U chart",
      reason:
        "You have counts where the amount of opportunity or activity varies, so a U chart is the right starting point. " +
        "After the chart is generated, SimpleSPC will check whether the rates vary more or less than a standard U chart expects. " +
        "If so, it will suggest a Laney U′ chart and explain why."
    };
  }

  return {
    chartType: "c",
    label: "C chart",
    reason:
      "You have counts with a roughly constant amount of opportunity or activity each time."
  };
}

    case "proportion": {
  if (answers.proportionHasDenom === "yes") {
    return {
      chartType: "p",
      label: "P chart",
      reason:
        "You have a numerator and denominator for each time point, so a P chart is the right starting point. " +
        "After the chart is generated, SimpleSPC will check whether the data vary more or less than a standard P chart expects. " +
        "If so, it will suggest a Laney P′ chart and explain why."
    };
  }

  return {
    chartType: "xmr",
    label: "XmR",
    reason:
      "Without the numerator and denominator, a P or P′ chart cannot calculate the expected variation correctly. " +
      "The percentage or proportion can instead be treated as a single measurement over time using an XmR chart."
  };
}


    case "pareto": {
      const alreadySummarised =
        answers.paretoDataFormat === "counts";

      return {
        chartType: "pareto",
        label: "Pareto chart",
        reason:
          "You want to compare categories and identify which " +
          "account for the most occurrences. A Pareto chart " +
          "ranks categories from most to least frequent and " +
          "shows their cumulative percentage. " +
          (alreadySummarised
            ? "Your data already contain counts, so select the " +
              "category column and the numeric frequency/count column. "
            : "If each row represents one occurrence, select the " +
              "category column and choose (count each row as 1). ") +
          "The chart helps prioritise investigation, but does not " +
          "identify root causes or assess statistical process stability."
      };
    }



    case "rare": {
      if (answers.rareType === "opportunities") return { chartType: "g", label: "G", reason: "You have opportunities between rare events." };
      return { chartType: "t", label: "T", reason: "You have time between rare events." };
    }

    case "unsure":
    default:
      return { chartType: "xmr", label: "XmR", reason: "When unsure, XmR is a safe default for a single value over time." };
  }
}

function renderChartWizard() {
  if (!chartWizardBody) return;

  const s = chartWizardState;
  const a = s.answers;

  const optionButton = (text, onClick) =>
    `<button type="button" class="chart-wizard-option" onclick="${onClick}">${text}</button>`;

  const actionButtons = `
    <div class="chart-wizard-actions">
      <button type="button" class="chart-wizard-secondary" onclick="wizardBack()">Back</button>
      <button type="button" onclick="finishWizard(computeRecommendation(chartWizardState.answers))">Skip</button>
    </div>
  `;

  if (s.step === 0) {
    chartWizardBody.innerHTML = `
      <p class="chart-wizard-question">What are you charting?</p>

      <div class="chart-wizard-options">
        ${optionButton("A measurement — e.g. waiting time, score, temperature, length of stay", `wizardNext('kind','measurement')`)}
        ${optionButton("A count per time period — e.g. falls per week, complaints per month", `wizardNext('kind','count')`)}
        ${optionButton("A proportion out of a total — e.g. 5 out of 100, pass rate", `wizardNext('kind','proportion')`)}
	${optionButton("Categories — e.g. reasons for delays, types of complaints", `wizardNext('kind','pareto')`)}
        ${optionButton("Rare events — time or opportunities between events", `wizardNext('kind','rare')`)}
        ${optionButton("Not sure", `wizardNext('kind','unsure')`)}
      </div>
    `;
    return;
  }

if (s.step === 1 && a.kind === "unsure") {
  chartWizardBody.innerHTML = `
    <p class="chart-wizard-question">Which description is closest to your data?</p>

    <p class="hint small-hint">
      The chart type depends on what each value represents. Choose the closest
      description below. If none fits, you can return to the chart list and
      select a chart manually.
    </p>

    <div class="chart-wizard-options">
      ${optionButton(
        "One measurement at each time point — e.g. waiting time, score, length of stay",
        `wizardNext('kind','measurement')`
      )}

      ${optionButton(
        "A number of events in each time period — e.g. falls, complaints, incidents",
        `wizardNext('kind','count')`
      )}

      ${optionButton(
        "A proportion or percentage — e.g. 5 out of 100, compliance rate",
        `wizardNext('kind','proportion')`
      )}

      ${optionButton(
        "Categories of occurrences — e.g. complaint types or reasons for delays",
        `wizardNext('kind','pareto')`
      )}

      ${optionButton(
        "The time or opportunities between rare events",
        `wizardNext('kind','rare')`
      )}
    </div>

    <div class="chart-wizard-actions">
      <button
        type="button"
        class="chart-wizard-secondary"
        onclick="wizardBack()"
      >
        Back
      </button>

      <button
        type="button"
        onclick="closeChartWizard()"
      >
        Choose manually
      </button>
    </div>
  `;

  return;
}


  if (s.step === 1 && a.kind === "pareto") {
    chartWizardBody.innerHTML = `
      <p class="chart-wizard-question">
        How are your category data recorded?
      </p>

      <p class="hint small-hint">
        A Pareto chart compares the frequency of different
        categories. You can use individual records or data
        that already contain counts for each category.
      </p>

      <div class="chart-wizard-options">
        ${optionButton(
          "One row per occurrence — e.g. one row for each complaint",
          `wizardNext('paretoDataFormat','rows')`
        )}

        ${optionButton(
          "Summarised counts — e.g. Falls: 20, Delays: 15",
          `wizardNext('paretoDataFormat','counts')`
        )}

        ${optionButton(
          "Not sure",
          `wizardNext('paretoDataFormat','unsure')`
        )}
      </div>

      ${actionButtons}
    `;
    return;
  }



  if (s.step === 1 && a.kind === "measurement") {
    chartWizardBody.innerHTML = `
      <p class="chart-wizard-question">Do you have one value per time point, or multiple values?</p>

      <div class="chart-wizard-options">
        ${optionButton("One value each time point", `wizardNext('measurementShape','single')`)}
        ${optionButton("Multiple values per time point — subgroups or samples", `wizardNext('measurementShape','subgroups')`)}
        ${optionButton("Not sure", `wizardNext('measurementShape','unsure')`)}
      </div>

      ${actionButtons}
    `;
    return;
  }

  if (s.step === 1 && a.kind === "count") {
    chartWizardBody.innerHTML = `
      <p class="chart-wizard-question">Does the amount of work or opportunity vary?</p>

      <p class="hint small-hint">
  If activity is broadly similar each time, treat it as roughly constant.
  If volume changes a lot, or you have a denominator such as bed-days,
  attendances or inspections, choose varies.
  <br><br>
  SimpleSPC will normally start with a U chart and then check whether
  a Laney U′ chart would better match the amount of variation in your data.
</p>

      <div class="chart-wizard-options">
        ${optionButton("No — roughly similar volume each time", `wizardNext('countOpportunity','constant')`)}
        ${optionButton("Yes — volume varies, or I have a denominator column", `wizardNext('countOpportunity','varies')`)}
        ${optionButton("Not sure", `wizardNext('countOpportunity','unsure')`)}
      </div>

      ${actionButtons}
    `;
    return;
  }

  if (s.step === 1 && a.kind === "proportion") {
  chartWizardBody.innerHTML = `
    <p class="chart-wizard-question">
      Do you have both parts of the proportion?
    </p>

    <p class="hint small-hint">
      For example, if 8 out of 100 patients met the measure,
      the numerator is 8 and the denominator is 100.
      <br><br>
      SimpleSPC will normally start with a P chart and then
      check whether a Laney P′ chart would better match the
      amount of variation in your data.
    </p>

    <div class="chart-wizard-options">
      ${optionButton(
        "Yes — I have numerator and denominator columns",
        `wizardNext('proportionHasDenom','yes')`
      )}

      ${optionButton(
        "No — I only have the percentage or proportion value",
        `wizardNext('proportionHasDenom','no')`
      )}

      ${optionButton(
        "Not sure",
        `wizardNext('proportionHasDenom','unsure')`
      )}
    </div>

    ${actionButtons}
  `;

  return;
}

  if (s.step === 1 && a.kind === "rare") {
    chartWizardBody.innerHTML = `
      <p class="chart-wizard-question">Which best describes your data?</p>

      <p class="hint small-hint">
        Choose this when the event is uncommon and you are looking at the gap between events.
      </p>

      <div class="chart-wizard-options">
        ${optionButton("Time between events — e.g. days between incidents", `wizardNext('rareType','time')`)}
        ${optionButton("Opportunities between events — e.g. procedures between harms", `wizardNext('rareType','opportunities')`)}
        ${optionButton("Not sure", `wizardNext('rareType','unsure')`)}
      </div>

      ${actionButtons}
    `;
    return;
  }

  if (s.step >= 2 && s.step !== 99) {
    finishWizard(computeRecommendation(s.answers));
    return;
  }

  if (s.step === 99 && s.recommendation) {
    const rec = s.recommendation;

    chartWizardBody.innerHTML = `
      <div class="chart-wizard-result">
        <div class="chart-wizard-result-title">Recommended chart: ${rec.label}</div>
        <div>${rec.reason}</div>
      </div>

      <div class="chart-wizard-actions">
        <button type="button" class="chart-wizard-secondary" onclick="wizardBack()">Back</button>
        <button type="button" class="wizard-primary" onclick="useWizardChart('${rec.chartType}')">Use this chart</button>
      </div>

      <p class="hint small-hint" style="margin-top:0.75rem;">
        You can still pick a different chart type manually if you prefer.
      </p>
    `;
    return;
  }
}

// Make wizard functions callable from inline onclick in the HTML strings
window.wizardNext = wizardNext;
window.wizardBack = wizardBack;
window.finishWizard = finishWizard;
window.computeRecommendation = computeRecommendation;
window.setChartType = setChartType;
window.closeChartWizard = closeChartWizard;
window.useWizardChart = useWizardChart;

function useWizardChart(chartType) {
  setChartType(chartType);
  closeChartWizard();

  if (rawRows && rawRows.length && generateButton) {
    lastGenerateWasManual = false;
    generateButton.click();
  }
}

function switchPToPPrime() {
  setChartType("pprime");

  if (
    rawRows &&
    rawRows.length &&
    generateButton
  ) {
    lastGenerateWasManual = false;
    generateButton.click();
  }
}

window.switchPToPPrime =
  switchPToPPrime;

function switchUToUPrime() {
  setChartType("uprime");

  if (
    rawRows &&
    rawRows.length &&
    generateButton
  ) {
    lastGenerateWasManual = false;
    generateButton.click();
  }
}

window.switchUToUPrime =
  switchUToUPrime;

if (newDataHelpChooseBtn) {
  newDataHelpChooseBtn.addEventListener("click", () => {
    hideNewDataChartPrompt();

    toggleChartWizard(true);
    startChartWizard();
  });
}

if (newDataChooseManuallyBtn) {
  newDataChooseManuallyBtn.addEventListener("click", () => {
    hideNewDataChartPrompt();
  });
}

// Hook wizard start into the existing button/modal
if (helpChooseChartBtn) {
  helpChooseChartBtn.addEventListener("click", () => {
    hideNewDataChartPrompt();

    toggleChartWizard(true);
    startChartWizard();
  });
}

if (chartSetupBtn) {
  chartSetupBtn.addEventListener("click", () => {
    openChartSetupForCurrentType();
  });
}

// Optional: close wizard when clicking the backdrop
(function wireWizardBackdropClose() {
  const modal = document.getElementById("chartWizardModal");
  if (!modal) return;
  const backdrop = modal.querySelector(".modal-backdrop");
  if (backdrop) {
    backdrop.addEventListener("click", () => toggleChartWizard(false));
  }
})();

if (ruleExplainerBtn) {
  ruleExplainerBtn.addEventListener("click", () => {
    openRuleExplainerForCurrentChart();
  });
}

// ---- Existing split dropdown button still works ----
function applySplitFromSidebarSelection() {
  if (!splitPointSelect) return false;

  // Which charts support splits (recalculated limits / median)
  const chartType =
    (typeof getSelectedChartType_NoSideEffects === "function")
      ? (getSelectedChartType_NoSideEffects() || "run")
      : ((typeof getSelectedChartType === "function") ? (getSelectedChartType() || "run") : "run");

  const supportsSplits = ["run", "xmr", "c", "p", "pprime", "u", "uprime", "xbars", "t", "g"].includes(chartType);

  if (!supportsSplits) {
    alert("Splits / recalculating limits are not available for this chart type.");
    return false;
  }

  const idx = parseInt(splitPointSelect.value, 10);
  if (!Number.isInteger(idx) || idx < 0) {
    alert("Please select a valid split point.");
    return false;
  }

  // Avoid duplicates
  if (!Array.isArray(splits)) splits = [];
  if (!splits.includes(idx)) {
    splits.push(idx);
    splits.sort((a, b) => a - b);
  }

  // Keep dropdown in sync (use current chart labels if available)
  const labels =
    (currentChart && currentChart.data && Array.isArray(currentChart.data.labels))
      ? currentChart.data.labels
      : null;

  if (labels && typeof populateSplitOptions === "function") {
    populateSplitOptions(labels);
  }

    // If the user has not manually locked the Y-axis bounds,
  // return to automatic scaling before regenerating.
  if (!yAxisBoundsManuallyEdited) {
    if (yAxisMinInput) yAxisMinInput.value = "";
    if (yAxisMaxInput) yAxisMaxInput.value = "";
  }

  // Redraw whichever chart is currently selected
  if (generateButton) generateButton.click();

  return true;
}


if (addSplitButton) {
  addSplitButton.addEventListener("click", () => {
    applySplitFromSidebarSelection();
  });
}


// ---- Right-click on chart: show menu ----
if (chartCanvas) {
  chartCanvas.addEventListener("contextmenu", (evt) => {
    evt.preventDefault();

    const idx = getNearestPointIndexFromEvent(evt);

    contextMenuMeanLabelSegmentIndex = null;

    if (
      currentChart &&
      Array.isArray(currentChart.$xmrMeanLabelHitboxes)
    ) {
      const rect = chartCanvas.getBoundingClientRect();

      const x =
        (evt.clientX - rect.left) *
        (currentChart.width / rect.width);

      const y =
        (evt.clientY - rect.top) *
        (currentChart.height / rect.height);

      const hit =
        currentChart.$xmrMeanLabelHitboxes.find(box =>
          x >= box.x &&
          x <= box.x + box.width &&
          y >= box.y &&
          y <= box.y + box.height
        );

      if (hit) {
        contextMenuMeanLabelSegmentIndex =
          hit.segmentIndex;
      }
    }

    showChartContextMenu(
      evt.clientX,
      evt.clientY,
      idx
    );
  });
}


 // Hide the context menu when clicking outside it,
 // pressing Escape, or scrolling.
 // Clicking inside the menu should not dismiss it.
 document.addEventListener("click", (e) => {
   if (chartContextMenu && chartContextMenu.contains(e.target)) {
     return;
   }
   hideChartContextMenu();
 });

 document.addEventListener("keydown", (e) => {
   if (e.key === "Escape") hideChartContextMenu();
 });

 document.addEventListener("scroll", () => hideChartContextMenu(), true);


// Menu actions (right-click menu)
if (chartContextMenu) {
  chartContextMenu.addEventListener("click", async (e) => {
    const btn = e.target.closest("button[data-action]");
    if (!btn) return;

    const action = btn.getAttribute("data-action");

    // Capture context BEFORE hiding the menu
    const clickedPointIndex =
      contextMenuPointIndex;

    const clickedMeanLabelSegmentIndex =
      contextMenuMeanLabelSegmentIndex;

    hideChartContextMenu();

    if (!currentChart) {
      alert("Please generate a chart first.");
      return;
    }

    try {
      if (action === "toggleXmRMeanLabels") {
        showXmRMeanLabels =
          !showXmRMeanLabels;

        if (generateButton) {
          generateButton.click();
        }

        return;
      }

      if (action === "hideXmRMeanLabel") {
        if (
          clickedMeanLabelSegmentIndex !== null &&
          clickedMeanLabelSegmentIndex !== undefined
        ) {
          xmrHiddenMeanLabels.add(
            clickedMeanLabelSegmentIndex
          );

          if (currentChart) {
            currentChart.draw();
          }
        }

        return;
      }

      if (action === "restoreXmRMeanLabels") {
        xmrHiddenMeanLabels.clear();

        if (currentChart) {
          currentChart.draw();
        }

        return;
      }

	      if (
        action ===
          "excludeFromLimits" ||
        action ===
          "includeInLimits"
      ) {
        if (
          clickedPointIndex ===
            null ||
          clickedPointIndex ===
            undefined
        ) {
          alert(
            "Right-click near a plotted point first."
          );

          return;
        }

        const state =
          getCurrentLimitExclusionMenuState(
            clickedPointIndex
          );

        if (
          !state.supported ||
          !state.key
        ) {
          return;
        }

        // Restoring a point is immediate.
        if (
          action ===
          "includeInLimits"
        ) {
          setLimitExcluded(
            state.scope,
            state.key,
            false
          );

          if (generateButton) {
            lastGenerateWasManual =
              false;

            generateButton.click();
          }

          return;
        }

        if (!state.eligible) {
          alert(
            state.reason ||
            "This point is not part of the current limit-estimation window."
          );

          return;
        }

        const isSubgroup =
          state.chartType ===
          "xbars";

        const ok =
          window.confirm(
            `Exclude this ${
              isSubgroup
                ? "subgroup"
                : "point"
            } from the limit calculations?\n\n` +

            `The ${
              isSubgroup
                ? "subgroup"
                : "point"
            } will stay visible on the chart, but it will not be used to estimate the centre line or control limits.\n\n` +

            "Only do this after you have investigated it and have good reason to believe it reflects a known special cause rather than the routine variation of the process.\n\n" +

            "Do not exclude a point simply because it is unusually high or low. If it represents a lasting change in the process, use a split instead."
          );

        if (!ok) {
          return;
        }

        setLimitExcluded(
          state.scope,
          state.key,
          true
        );

        const xLabel =
          currentChart
            ?.data
            ?.labels
            ?.[clickedPointIndex] ??
          `point ${
            clickedPointIndex + 1
          }`;

        const addNote =
          window.confirm(
            `${
              isSubgroup
                ? "Subgroup"
                : "Point"
            } excluded from the limit calculations. It will remain visible on the chart.\n\n` +

            "Would you like to add an annotation explaining why it was excluded?"
          );

        if (addNote) {
          const text =
            prompt(
              `Annotation for ${xLabel}:`,
              ""
            );

          if (
            text !== null &&
            String(text)
              .trim()
          ) {
            annotations.push({
              date:
                xLabel,

              label:
                String(text)
                  .trim()
            });
          }
        }

        if (generateButton) {
          lastGenerateWasManual =
            false;

          generateButton.click();
        }

        return;
      }


      if (action === "addAnnotation") {
        if (
          clickedPointIndex === null ||
          clickedPointIndex === undefined
        ) {
          alert(
            "Right-click near a data point to add an annotation."
          );
          return;
        }

        const labels =
          currentChart?.data?.labels || [];

        const xLabel =
          labels?.[clickedPointIndex];

        if (!xLabel) {
          alert(
            "Could not determine the selected x-position for annotation."
          );
          return;
        }

        const text = prompt(
          `New annotation for ${xLabel}:`,
          ""
        );

        if (text === null) return;

        const trimmed =
          String(text).trim();

        if (!trimmed) {
          alert(
            "Annotation text cannot be blank."
          );
          return;
        }

        if (annotationDateInput) {
          annotationDateInput.value =
            xLabel;
        }

        if (annotationLabelInput) {
          annotationLabelInput.value =
            trimmed;
        }

        annotations.push({
          date: xLabel,
          label: trimmed
        });

        if (generateButton) {
          generateButton.click();
        }

        return;
      }
      if (action === "editAnnotation") {
        if (clickedPointIndex === null || clickedPointIndex === undefined) {
          alert("Right-click near a data point to edit an annotation.");
          return;
        }

        const labels = currentChart?.data?.labels || [];
        const xLabel = labels?.[clickedPointIndex];

        if (!xLabel) {
          alert("Could not determine the selected x-position for annotation.");
          return;
        }

        const chosen = chooseAnnotationAtDate(xLabel, "edit");
        if (!chosen) {
          alert(`There are no annotations to edit at ${xLabel}.`);
          return;
        }

        const changed = editAnnotationAtIndex(chosen._idx);
        if (changed && generateButton) generateButton.click();
        return;
      }

      if (action === "deleteAnnotation") {
        if (clickedPointIndex === null || clickedPointIndex === undefined) {
          alert("Right-click near a data point to delete an annotation.");
          return;
        }

        const labels = currentChart?.data?.labels || [];
        const xLabel = labels?.[clickedPointIndex];

        if (!xLabel) {
          alert("Could not determine the selected x-position for annotation.");
          return;
        }

        const chosen = chooseAnnotationAtDate(xLabel, "delete");
        if (!chosen) {
          alert(`There are no annotations to delete at ${xLabel}.`);
          return;
        }

        const ok = confirm(`Delete this annotation?\n\n"${chosen.label}"`);
        if (!ok) return;

        const changed = deleteAnnotationAtIndex(chosen._idx);
        if (changed && generateButton) generateButton.click();
        return;
      }

      if (action === "clearAnnotations") {
        if (!annotations || annotations.length === 0) return;

        const ok = confirm("Clear all annotations?");
        if (!ok) return;

        annotations.length = 0;

        if (annotationDateInput) annotationDateInput.value = "";
        if (annotationLabelInput) annotationLabelInput.value = "";

        if (generateButton) generateButton.click();
        return;
      }

      if (action === "addSplit") {
        if (clickedPointIndex === null || clickedPointIndex === undefined) {
          alert("Right-click near a data point to add a split.");
          return;
        }

        // Try sidebar-style apply ONLY if the split dropdown exists.
        // If it fails (e.g. dropdown removed), fall back to direct add.
        let applied = false;

        if (splitPointSelect && typeof applySplitFromSidebarSelection === "function") {
          splitPointSelect.value = String(clickedPointIndex);
          applied = (applySplitFromSidebarSelection() === true);
        }

        if (!applied) {
          // Direct method that does NOT require sidebar UI
          addSplitAfterIndex(clickedPointIndex);
        }

        return;
      }
	
      if (action === "removeSplit") {
        const splitIdx = Number(btn.getAttribute("data-split-index"));
        if (!Number.isInteger(splitIdx)) return;

        removeSplitAtIndex(splitIdx);
        return;
      }

      if (action === "clearSplits") {
        // Clear splits immediately + redraw
        splits = [];
        if (splitPointSelect) splitPointSelect.value = "";

        if (generateButton) generateButton.click();
        return;
      }

      if (action === "copyCharts") {
  let restoreExportLayout = null;

  try {
    restoreExportLayout = await prepareChartsForExport();

    const composite = buildCompositeCanvas({ includeSummaryText: false });
    await copyCanvasToClipboard(composite);
    alert("Chart image copied to clipboard.");
  } finally {
    if (restoreExportLayout) restoreExportLayout();
  }

  return;
}

if (action === "copyChartsAndAnalysis") {
  let restoreExportLayout = null;

  try {
    restoreExportLayout = await prepareChartsForExport();

    const composite = buildCompositeCanvas({ includeSummaryText: true });
    await copyCanvasToClipboard(composite);
    alert("Chart + analysis image copied to clipboard.");
  } finally {
    if (restoreExportLayout) restoreExportLayout();
  }

  return;
}

     if (action === "saveChartsAs") {
  let restoreExportLayout = null;

  try {
    restoreExportLayout = await prepareChartsForExport();

    const composite = buildCompositeCanvas({ includeSummaryText: true });
    downloadCanvasAsPng(composite, "spc-charts.png");
  } finally {
    if (restoreExportLayout) restoreExportLayout();
  }

  return;
}

      if (action === "downloadPdf") {
        exportPdfReport();
        return;
      }

    } catch (err) {
      console.error(err);
      alert("Sorry — that action failed in this browser. Try 'Save chart(s) as…' instead.");
    }
  });
}


function formatSpcHelperAnswerToHtml(text) {
  const raw = String(text ?? "").replace(/\r\n/g, "\n").trim();
  if (!raw) return `<p>${escapeHtml("No answer available.")}</p>`;

  // Escape any HTML to keep this safe
  const escaped = escapeHtml(raw);
  const lines = escaped.split("\n");

  const out = [];
  let paraBuf = [];
  let listBuf = null; // { type: 'ul'|'ol', items: [] }

  function flushParagraph() {
    if (paraBuf.length === 0) return;
    const html = applyBasicInlineFormatting(paraBuf.join("<br>"));
    out.push(`<p>${html}</p>`);
    paraBuf = [];
  }

  function flushList() {
    if (!listBuf || listBuf.items.length === 0) {
      listBuf = null;
      return;
    }
    const tag = listBuf.type;
    const itemsHtml = listBuf.items
      .map(it => `<li>${applyBasicInlineFormatting(it)}</li>`)
      .join("");
    out.push(`<${tag}>${itemsHtml}</${tag}>`);
    listBuf = null;
  }

  function startList(type) {
    // Switch list types cleanly (paragraph -> list, ul -> ol, etc.)
    flushParagraph();
    if (listBuf && listBuf.type !== type) flushList();
    if (!listBuf) listBuf = { type, items: [] };
  }

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trimEnd();
    const t = line.trim();

    // Blank line = hard break between blocks
    if (t === "") {
      flushParagraph();
      flushList();
      continue;
    }

    // Bullet item: "- " or "• "
    if (t.startsWith("- ") || t.startsWith("• ")) {
      startList("ul");
      listBuf.items.push(t.replace(/^(-\s+|•\s+)/, ""));
      continue;
    }

    // Numbered item: "1. " "2. " etc.
    if (/^\d+\.\s+/.test(t)) {
      startList("ol");
      listBuf.items.push(t.replace(/^\d+\.\s+/, ""));
      continue;
    }

    // Normal text line
    // If we were building a list and now have normal text, close the list first.
    if (listBuf) flushList();

    // Add to paragraph buffer
    paraBuf.push(t);
  }

  flushParagraph();
  flushList();

  return out.join("");
}


// Optional: allow very small “markdown-like” formatting (safe because input is escaped)
function applyBasicInlineFormatting(escapedText) {
  // **bold**
  let t = escapedText.replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>");
  // _italic_
  t = t.replace(/_(.+?)_/g, "<em>$1</em>");
  return t;
}

function showHelperAnswer(questionText) {
  if (!spcHelperOutput) return;

  const q = (questionText ?? aiQuestionInput?.value ?? "").trim();
  if (!q) {
    spcHelperOutput.innerHTML = `<p>${escapeHtml("Type a question (or click a suggestion) to get started.")}</p>`;
    return;
  }

  const ans = answerSpcQuestion(q);

  // Use your formatter if present; otherwise fall back safely
  if (typeof formatSpcHelperAnswerToHtml === "function") {
    spcHelperOutput.innerHTML = formatSpcHelperAnswerToHtml(ans);
  } else {
    spcHelperOutput.innerHTML = `<p>${escapeHtml(ans)}</p>`;
  }

  spcHelperOutput.scrollTop = 0;

  // Auto-collapse the chips after the first answer to free space for reading
  if (!spcHelperAutoCollapsedOnce) {
    setSpcHelperSuggestionsCollapsed(true);
    spcHelperAutoCollapsedOnce = true;
  } else {
    // Also collapse on subsequent answers (keeps focus on reading)
    setSpcHelperSuggestionsCollapsed(true);
  }
}


if (aiAskButton && aiQuestionInput) {
  aiAskButton.addEventListener("click", () => {
    showHelperAnswer();
  });

  aiQuestionInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      showHelperAnswer();
    }
  });
}

function handleChipClick(e) {
  const btn = e.target.closest("button[data-q]");
  if (!btn) return;

  const q = btn.getAttribute("data-q") || "";
  if (aiQuestionInput) aiQuestionInput.value = q;

  showHelperAnswer(q);
}

if (spcHelperChipsGeneral) {
  spcHelperChipsGeneral.addEventListener("click", handleChipClick);
}
if (spcHelperChipsChart) {
  spcHelperChipsChart.addEventListener("click", handleChipClick);
}



if (clearSplitsButton) {
  clearSplitsButton.addEventListener("click", () => {
    splits = [];

    if (splitPointSelect) {
      splitPointSelect.value = "";
    }

    if (!yAxisBoundsManuallyEdited) {
      if (yAxisMinInput) yAxisMinInput.value = "";
      if (yAxisMaxInput) yAxisMaxInput.value = "";
    }

    if (getSelectedChartType_NoSideEffects() === "xmr") {
      generateButton.click();
    }
  });
}



// --- SPC helper: collapse/expand suggested chips for better small-screen UX ---
let spcHelperHasBeenOpened = false;
let spcHelperAutoCollapsedOnce = false;

function setSpcHelperSuggestionsCollapsed(collapsed) {
  const suggestions = document.getElementById("spcHelperSuggestions");
  const toggleBtn = document.getElementById("spcHelperToggleSuggestions");
  if (!suggestions || !toggleBtn) return;

  suggestions.style.display = collapsed ? "none" : "";
  toggleBtn.textContent = collapsed ? "Show suggested questions" : "Hide suggested questions";
  toggleBtn.setAttribute("aria-expanded", collapsed ? "false" : "true");
}

// Hook up the toggle button (safe even if elements aren’t present yet)
function attachSpcHelperSuggestionToggle() {
  const toggleBtn = document.getElementById("spcHelperToggleSuggestions");
  if (!toggleBtn || toggleBtn.dataset.bound === "1") return;

  toggleBtn.dataset.bound = "1";
  toggleBtn.addEventListener("click", () => {
    const suggestions = document.getElementById("spcHelperSuggestions");
    if (!suggestions) return;

    const isHidden = suggestions.style.display === "none";
    setSpcHelperSuggestionsCollapsed(!isHidden);
  });
}

function attachSpcHelperSectionToggles() {
  const generalBtn = document.getElementById("spcHelperToggleGeneral");
  const chartBtn = document.getElementById("spcHelperToggleChart");
  const generalBody = document.getElementById("spcHelperGeneralSection");
  const chartBody = document.getElementById("spcHelperChartSection");

  if (generalBtn && generalBody && generalBtn.dataset.bound !== "1") {
    generalBtn.dataset.bound = "1";
    generalBtn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      toggleHelperSection(generalBtn, generalBody);
    });
  }

  if (chartBtn && chartBody && chartBtn.dataset.bound !== "1") {
    chartBtn.dataset.bound = "1";
    chartBtn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      toggleHelperSection(chartBtn, chartBody);
    });
  }
}

function toggleSpcHelper() {
  const panel = document.getElementById("spcHelperPanel");
  if (!panel) return;

  const isVisible = panel.classList.toggle("visible");

  if (isVisible) {
    // Populate chips / intro once
    if (!spcHelperHasBeenOpened) {
      if (typeof renderHelperState === "function") renderHelperState();
      spcHelperHasBeenOpened = true;
    }

        // Ensure toggle buttons work
    attachSpcHelperSuggestionToggle();
    attachSpcHelperSectionToggles();

    // Re-render helper state when opening so defaults reflect the current chart state
    if (typeof renderHelperState === "function") renderHelperState();

    // When opening: show suggestions by default for discoverability
    setSpcHelperSuggestionsCollapsed(false);
  }
}


const resetButton = document.getElementById("resetButton");
const clearDataButton = document.getElementById("clearDataButton");

if (resetButton) {
  resetButton.addEventListener("click", () => {
    const hasExistingWork =
      (Array.isArray(rawRows) && rawRows.length > 0) ||
      !!currentChart ||
      !!dataModelDirty;

    if (hasExistingWork) {
      const ok = window.confirm(
        "Reset this chart?\n\n" +
        "Your loaded data and selected X/value columns will be kept, " +
        "but chart settings such as chart type, baseline, target, splits, " +
        "annotations, labels, rules and axis formatting will be reset."
      );

      if (!ok) return;
    }

    resetChartKeepData();
  });
}

if (clearDataButton) {
  clearDataButton.addEventListener("click", () => {
    const hasExistingWork =
      (Array.isArray(rawRows) && rawRows.length > 0) ||
      !!currentChart ||
      !!dataModelDirty;

    if (hasExistingWork) {
      const ok = window.confirm(
        "Clear all loaded data?\n\n" +
"This will also reset the current chart and its settings. " +
"Save the chart first if you want to reopen it later."
      );

      if (!ok) return;
    }

    resetAll();
  });
}
updateSaveChartButtonState();
updateDateControlsState();

// Allow Escape key to close the SPC helper
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") {
    const helpModal = document.getElementById("helpModal");
    if (helpModal && helpModal.classList.contains("visible")) {
      toggleHelpSection(false);
      return;
    }

    if (spcHelperPanel && spcHelperPanel.classList.contains("visible")) {
      spcHelperPanel.classList.remove("visible");
    }
  }
});

function countValidNumericPoints() {
  if (!rawRows || !rawRows.length) return 0;
  const valueCol = valueSelect?.value;
  if (!valueCol) return 0;

  let n = 0;
  for (const row of rawRows) {
    const y = toNumericValue(row[valueCol]);
    if (isFinite(y)) n++;
  }
  return n;
}

function enforceChartTypeSuitabilityAndRegen() {
  if (!rawRows || !rawRows.length) return;


  const beforeType = getSelectedChartType_NoSideEffects();
  const availability = applyChartTypeAvailability();
  const afterType = getSelectedChartType_NoSideEffects();
  updateMrToggleVisibility();

  if (!availability[afterType]?.enabled) {
    const reason = availability[afterType]?.reason || "This chart type is not available for the current data.";
    showError(reason);
    return;
  }

  if (beforeType !== afterType) {
    const reason = availability[beforeType]?.reason || "The previously selected chart type is not available for the current data.";
    showError(`${reason} Switched to ${getChartTypeDisplayName(afterType)}.`);
  }

  generateButton.click();
}

// ---- Auto-regenerate when chart type, axis type, or selected columns change ----
function wireAutoRedrawControls() {
  // Chart type radios (run / xmr)
   document.querySelectorAll("input[name='chartType']").forEach(radio => {
  radio.addEventListener("change", () => {
  hideNewDataChartPrompt();  
  applyDefaultYBoundsForSelectedColumn();

    if (typeof updateUIForChartType === "function") {
      updateUIForChartType(radio.value);
      updateRuleUIForChartType(radio.value);

    }
    maybeShowChartSetupModal(radio.value);

    if (typeof updateMrToggleVisibility === "function") {
      updateMrToggleVisibility();
    }

    if (rawRows && rawRows.length) {
      if (typeof enforceChartTypeSuitabilityAndRegen === "function") {
        enforceChartTypeSuitabilityAndRegen();
      } else if (generateButton) {
        generateButton.click();
      }
    }
  });
});

  // Axis type radios (date / sequence)
  document.querySelectorAll("input[name='axisType']").forEach(radio => {
    radio.addEventListener("change", () => {
      axisTypeManuallyChanged = true;
      if (rawRows && rawRows.length) {
        if (typeof enforceChartTypeSuitabilityAndRegen === "function") {
          enforceChartTypeSuitabilityAndRegen();
        } else if (generateButton) {
          generateButton.click();
        }
      }
    });
  });

  // NEW: X / Y / third column dropdowns
  const dateSelect  = document.getElementById("dateColumn");
  const valueSelect = document.getElementById("valueColumn");
  const thirdSelect = document.getElementById("thirdColumn");

  const onColumnChange = () => {
    if (!rawRows || !rawRows.length) return;

    if (typeof enforceChartTypeSuitabilityAndRegen === "function") {
      enforceChartTypeSuitabilityAndRegen();
    } else if (generateButton) {
      generateButton.click();
    }
  };

  if (dateSelect)  dateSelect.addEventListener("change", onColumnChange);
  if (valueSelect) valueSelect.addEventListener("change", onColumnChange);
  if (thirdSelect) thirdSelect.addEventListener("change", onColumnChange);

  document.querySelectorAll("input[name='chartType']").forEach(radio => {
    radio.addEventListener("change", () => {
      yAxisBoundsManuallyEdited = false;
      applyDefaultYBoundsForSelectedColumn();

      if (!rawRows || !rawRows.length) return;

      if (typeof enforceChartTypeSuitabilityAndRegen === "function") {
        enforceChartTypeSuitabilityAndRegen();
      } else if (generateButton) {
        generateButton.click();
      }
    });
  });

    

  // Run once on load so MR toggle visibility matches initial selection
  if (typeof updateMrToggleVisibility === "function") {
    updateMrToggleVisibility();
  }
}




// Initialize UI once on load (in case default is run)
const checked = document.querySelector('input[name="chartType"]:checked');
if (checked) updateUIForChartType(checked.value);


// Call after the DOM is available (safe even if script is at bottom, but robust)
if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", wireAutoRedrawControls);
} else {
  wireAutoRedrawControls();
}

if (dataEditorHasHeaders) {
  dataEditorHasHeaders.addEventListener("change", renderHeaderStatus);
}

if (dataEditorDetectHeadersButton) {
  dataEditorDetectHeadersButton.addEventListener("click", () => {
    const guess = detectHeadersFromGrid();
    if (dataEditorHasHeaders) dataEditorHasHeaders.checked = guess;

    // Give immediate, obvious feedback
    if (dataEditorHeaderStatus) {
      dataEditorHeaderStatus.innerHTML = guess
        ? `Auto-detect: first row looks like <strong>headings</strong>.`
        : `Auto-detect: first row looks like <strong>data</strong>.`;
    }
  });
}

function initSidebarTabs() {

const buttons = document.querySelectorAll(".tab-btn");
const panels = document.querySelectorAll(".tab-panel");

buttons.forEach(btn => {

btn.addEventListener("click", () => {

const target = btn.dataset.tab;

buttons.forEach(b => b.classList.remove("active"));
panels.forEach(p => p.classList.remove("active"));

btn.classList.add("active");
document.getElementById(target).classList.add("active");

});

});

}

document.addEventListener("DOMContentLoaded", initSidebarTabs);

async function exportPdfReport() {
  if (!currentChart) {
    alert("Please generate a chart first.");
    return;
  }

  let restoreExportLayout = null;

  try {
    restoreExportLayout = await prepareChartsForExport();

    const composite = buildCompositeCanvas({ includeSummaryText: true });
    if (!composite) {
      alert("Could not build the PDF export image.");
      return;
    }

    const { jsPDF } = window.jspdf || {};
    if (!jsPDF) {
      alert("PDF export library not available.");
      return;
    }

    const pdf = new jsPDF({
      orientation: "landscape",
      unit: "mm",
      format: "a4"
    });

    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();

    const margin = 8;
    const usableWidth = pageWidth - margin * 2;
    const usableHeight = pageHeight - margin * 2;

    const scale = usableWidth / composite.width;
    const pageCanvasHeight = Math.floor(usableHeight / scale);

    let sourceY = 0;
    let pageNumber = 0;

    while (sourceY < composite.height) {
      const sliceHeight = Math.min(pageCanvasHeight, composite.height - sourceY);

      const pageCanvas = document.createElement("canvas");
      pageCanvas.width = composite.width;
      pageCanvas.height = sliceHeight;

      const ctx = pageCanvas.getContext("2d");
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(0, 0, pageCanvas.width, pageCanvas.height);

      ctx.drawImage(
        composite,
        0,
        sourceY,
        composite.width,
        sliceHeight,
        0,
        0,
        composite.width,
        sliceHeight
      );

      if (pageNumber > 0) pdf.addPage();

      const imgData = pageCanvas.toDataURL("image/png");
      const renderedHeight = sliceHeight * scale;

      pdf.addImage(
        imgData,
        "PNG",
        margin,
        margin,
        usableWidth,
        renderedHeight
      );

      sourceY += sliceHeight;
      pageNumber += 1;
    }

    pdf.save("spc-report.pdf");
  } finally {
    if (restoreExportLayout) restoreExportLayout();
  }
}



// Optional: keep this in case you ever add the top button back
if (downloadPdfBtn) {
  downloadPdfBtn.addEventListener("click", exportPdfReport);
}

renderHelperState();
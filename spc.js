// Simple SPC logic + wiring for Run chart and XmR chart + MR chart
// Features:
//  - CSV upload + column selection
//  - Run chart with run rule (>=8 points on one side of median)
//  - XmR chart with mean, UCL, LCL, +/-1σ, +/-2σ
//  - MR chart paired with XmR
//  - Baseline: use first N points for centre line & limits (optional)
//  - Target line (goal value) optional
//  - Summary panel with basic NHS-style interpretation
//  - Download main chart as PNG
//  - Custom chart title and axis labels

let rawRows = [];
let currentChart = null;   // main I / run chart
let mrChart = null;        // moving range chart
let annotations = [];      // { date: 'YYYY-MM-DD', label: 'text' }
let splits = [];   // indices where a new XmR segment starts (split AFTER index)
let lastXmRAnalysis = null;

const fileInput         = document.getElementById("fileInput");
const columnSelectors   = document.getElementById("columnSelectors");
const dateSelect        = document.getElementById("dateColumn");
const valueSelect       = document.getElementById("valueColumn");
const baselineInput     = document.getElementById("baselinePoints");
const chartTitleInput   = document.getElementById("chartTitle");
const xAxisLabelInput   = document.getElementById("xAxisLabel");
const yAxisLabelInput   = document.getElementById("yAxisLabel");
const targetInput       = document.getElementById("targetValue");
const targetDirectionInput = document.getElementById("targetDirection");
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
const aiQuestionInput  = document.getElementById("aiQuestionInput");
const aiAskButton      = document.getElementById("aiAskButton");
const spcHelperAnswer  = document.getElementById("spcHelperAnswer");
const spcHelperPanel     = document.getElementById("spcHelperPanel");



const mrPanel           = document.getElementById("mrPanel");
const mrChartCanvas     = document.getElementById("mrChartCanvas");

function loadRows(rows) {
  if (!rows || rows.length === 0) {
    errorMessage.textContent = "No rows found in the data.";
    return false;
  }

  rawRows = rows;
  const firstRow = rows[0];
  const columns = Object.keys(firstRow);

  if (!columns || columns.length === 0) {
    errorMessage.textContent = "Could not detect any columns in the data.";
    return false;
  }

  dateSelect.innerHTML = "";
  valueSelect.innerHTML = "";

  columns.forEach(col => {
    const opt1 = document.createElement("option");
    opt1.value = col;
    opt1.textContent = col;
    dateSelect.appendChild(opt1);

    const opt2 = document.createElement("option");
    opt2.value = col;
    opt2.textContent = col;
    valueSelect.appendChild(opt2);
  });

  columnSelectors.style.display = "block";
  errorMessage.textContent = "";
  return true;
}


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
    annotations.push({ date: dateVal, label: labelVal });

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

// ---- CSV upload & column selection ----

fileInput.addEventListener("change", () => {
  const file = fileInput.files[0];
  if (!file) return;

  errorMessage.textContent = "";
  if (summaryDiv) summaryDiv.innerHTML = "";
  if (capabilityDiv) capabilityDiv.innerHTML = "";
  annotations = [];
  if (annotationDateInput) annotationDateInput.value = "";
  if (annotationLabelInput) annotationLabelInput.value = "";
  splits = [];
  if (splitPointSelect) splitPointSelect.innerHTML = "";

  Papa.parse(file, {
    header: true,
    dynamicTyping: true,
    skipEmptyLines: true,
    complete: (results) => {
  const rows = results.data;

  // Use the shared loader so CSV and pasted data behave the same
  if (!loadRows(rows)) {
    return;
  }

  // Reset annotations and splits because the data changed
  annotations = [];
  if (annotationDateInput) annotationDateInput.value = "";
  if (annotationLabelInput) annotationLabelInput.value = "";
  splits = [];
  if (splitPointSelect) splitPointSelect.innerHTML = "";
},

    error: (err) => {
      errorMessage.textContent = "Error parsing CSV: " + err.message;
    }
  });
});

// ---- Helpers ----

function getSelectedChartType() {
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



/**
 * Compute XmR statistics and MR values.
 */
function computeXmR(points, baselineCount) {
  const pts = [...points].sort((a, b) => a.x - b.x);

  let baselineCountUsed;
  if (baselineCount && baselineCount >= 2) {
    baselineCountUsed = Math.min(baselineCount, pts.length);
  } else {
    baselineCountUsed = pts.length;
  }

  const baseline = pts.slice(0, baselineCountUsed);

  const mean =
    baseline.reduce((sum, p) => sum + p.y, 0) / baseline.length;

  // moving ranges for baseline (for sigma estimate)
  const baselineMRs = [];
  for (let i = 1; i < baseline.length; i++) {
    baselineMRs.push(Math.abs(baseline[i].y - baseline[i - 1].y));
  }
  const avgMR =
    baselineMRs.length > 0
      ? baselineMRs.reduce((sum, v) => sum + v, 0) / baselineMRs.length
      : 0;

  const sigma = avgMR === 0 ? 0 : avgMR / 1.128; // d2 for n=2

  const ucl = mean + 3 * sigma;
  const lcl = mean - 3 * sigma;

 // MR values for full series (for MR chart)
  const mrValues = [];
  for (let i = 1; i < pts.length; i++) {
    mrValues.push(Math.abs(pts[i].y - pts[i - 1].y));
  }

  const flagged = pts.map(p => ({
    ...p,
    beyondLimits: sigma > 0 && (p.y > ucl || p.y < lcl)
  }));

  return {
    points: flagged,
    mean,
    ucl,
    lcl,
    sigma,
    avgMR,
    baselineCountUsed,
    mrValues
  };
}

// Get title / axis labels with fallbacks
function getChartLabels(defaultTitle, defaultX, defaultY) {
  const title = chartTitleInput && chartTitleInput.value.trim()
    ? chartTitleInput.value.trim()
    : defaultTitle;

  const xLabel = xAxisLabelInput && xAxisLabelInput.value.trim()
    ? xAxisLabelInput.value.trim()
    : defaultX;

  const yLabel = yAxisLabelInput && yAxisLabelInput.value.trim()
    ? yAxisLabelInput.value.trim()
    : defaultY;

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

function getTargetValue() {
  if (!targetInput) return null;
  const v = targetInput.value.trim();
  if (v === "") return null;
  const num = Number(v);
  return isFinite(num) ? num : null;
}

function getAxisType() {
  const radios = document.querySelectorAll("input[name='axisType']");
  for (const r of radios) {
    if (r.checked) return r.value;
  }
  return "date"; // sensible default
}

function buildAnnotationConfig(labels) {
  if (!annotations || annotations.length === 0) {
    return {};
  }

  const cfg = {};
  annotations.forEach((a, idx) => {
    const xVal = a.date; // 'YYYY-MM-DD' from <input type="date">
    if (!labels.includes(xVal)) {
      return; // skip if this date isn't on the x-axis
    }

    cfg["annot" + idx] = {
      type: "line",
      xMin: xVal,
      xMax: xVal,
      borderColor: "#000000",
      borderWidth: 1,
      borderDash: [2, 2],
      label: {
        display: true,
        content: a.label,
        backgroundColor: "rgba(255,255,255,0.9)",
        color: "#000000",
        borderColor: "#000000",
        borderWidth: 0.5,
        font: {
          size: 10,
          weight: "bold"
        },
        position: "end",   // near the top of the line
        yAdjust: -6        // nudge it up a little
        // no rotation – keep it horizontal so it's easy to read
      }
    };
  });

  return cfg;
}

function openDataEditor() {
  if (!dataEditorOverlay || !dataEditorTextarea) return;

  // If we already have data, show it as CSV; otherwise, give a skeleton
  if (rawRows && rawRows.length > 0) {
    try {
      dataEditorTextarea.value = Papa.unparse(rawRows);
    } catch (e) {
      // Fallback to blank if unparse fails for any reason
      dataEditorTextarea.value = "";
    }
  } else {
    dataEditorTextarea.value = "Date,Value\n";
  }

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
    closeDataEditor();
  });
}

// Optional: close when clicking outside the dialog
if (dataEditorOverlay) {
  dataEditorOverlay.addEventListener("click", (e) => {
    if (e.target === dataEditorOverlay) {
      closeDataEditor();
    }
  });
}

if (dataEditorApplyButton) {
  dataEditorApplyButton.addEventListener("click", () => {
    if (!dataEditorTextarea) return;
    const text = dataEditorTextarea.value.trim();
    if (!text) {
      alert("Please paste or type some data first.");
      return;
    }

    try {
      const results = Papa.parse(text, {
        header: true,
        dynamicTyping: true,
        skipEmptyLines: true
      });

      if (results.errors && results.errors.length > 0) {
        console.error(results.errors);
        errorMessage.textContent = "Error parsing pasted data: " + results.errors[0].message;
        return;
      }

      const rows = results.data;
      if (!loadRows(rows)) {
        return;
      }

      // Reset annotations and splits because the data changed
      annotations = [];
      if (annotationDateInput) annotationDateInput.value = "";
      if (annotationLabelInput) annotationLabelInput.value = "";
      splits = [];
      if (splitPointSelect) splitPointSelect.innerHTML = "";

      closeDataEditor();
    } catch (e) {
      console.error(e);
      errorMessage.textContent = "Unexpected error parsing pasted data.";
    }
  });
}

const resetButton = document.getElementById("resetButton");

if (resetButton) {
  resetButton.addEventListener("click", () => {
    
    // 1. Clear stored data
    rawRows = [];
    annotations = [];
    splits = [];

    // 2. Reset file input
    if (fileInput) fileInput.value = "";

    // 3. Reset pasted-data editor
    if (dataEditorTextarea) dataEditorTextarea.value = "";

    // 4. Reset column selectors
    if (dateSelect) dateSelect.innerHTML = "";
    if (valueSelect) valueSelect.innerHTML = "";
    if (columnSelectors) columnSelectors.style.display = "none";

    // 5. Reset baseline, target, axis type, chart type
    if (baselineInput) baselineInput.value = "";
    if (targetInput) targetInput.value = "";
    if (targetDirectionInput) targetDirectionInput.value = "below";

    const dateAxisRadio = document.querySelector("input[name='axisType'][value='date']");
    const runChartRadio = document.querySelector("input[name='chartType'][value='run']");
    if (dateAxisRadio) dateAxisRadio.checked = true;
    if (runChartRadio) runChartRadio.checked = true;

    // 6. Clear summary + capability blocks
    if (summaryDiv) summaryDiv.innerHTML = "";
    if (capabilityDiv) capabilityDiv.innerHTML = "";

    // 7. Clear MR chart panel
    const mrPanel = document.getElementById("mrPanel");
    if (mrPanel) mrPanel.style.display = "none";

    // 8. Destroy existing chart
    if (currentChart) {
      currentChart.destroy();
      currentChart = null;
    }

    // 9. Reset annotation inputs
    if (annotationDateInput) annotationDateInput.value = "";
    if (annotationLabelInput) annotationLabelInput.value = "";

    // 10. Provide feedback
    //alert("The tool has been reset. You can upload or paste new data.");
  });
}

// Convert CSV cell to a numeric value, handling percentages like "55.17%"
function toNumericValue(raw) {
  if (raw === null || raw === undefined) return NaN;

  // If PapaParse has already converted it to a number, just use it
  if (typeof raw === "number") return raw;

  const s = String(raw).trim();
  if (s === "") return NaN;

  // Handle simple percentages, e.g. "55.17%" or "55.17 %"
  const percentMatch = s.match(/^(-?\d+(?:\.\d+)?)\s*%$/);
  if (percentMatch) {
    return Number(percentMatch[1]);  // return the 55.17 part
  }

  // Fall back to normal numeric parsing
  const num = Number(s);
  return isFinite(num) ? num : NaN;
}

// ---- Summary helpers ----

function updateRunSummary(points, median, runFlags, baselineCountUsed) {
  if (!summaryDiv) return;

  const n = points.length;
  const nRunPoints = runFlags.filter(Boolean).length;
  const hasRunViolation = nRunPoints > 0;

  const values = points.map(p => p.y);
  const hasTrend = detectTrend(values, 6);

  let html = `<h3>Summary (Run chart)</h3>`;
  html += `<ul>`;
  html += `<li>Number of points: <strong>${n}</strong></li>`;
  if (baselineCountUsed && baselineCountUsed < n) {
    html += `<li>Baseline: first <strong>${baselineCountUsed}</strong> points used to calculate median.</li>`;
  } else {
    html += `<li>Baseline: all points used to calculate median.</li>`;
  }
  html += `<li>Median: <strong>${median.toFixed(3)}</strong></li>`;

  const signals = [];
  if (hasRunViolation) {
    signals.push("a run of 8 or more points on one side of the median");
  }
  if (hasTrend) {
    signals.push("a trend of 6 or more points all increasing or all decreasing");
  }

  if (signals.length === 0) {
    html += `<li><strong>Special cause:</strong> No rule breaches detected (based on long runs or trends). Variation appears consistent with common-cause only, but always interpret in clinical context.</li>`;
  } else {
    html += `<li><strong>Special cause:</strong> Signals suggesting special-cause variation based on: ${signals.join("; ")}.</li>`;
  }

  html += `</ul>`;

  summaryDiv.innerHTML = html;
}

// ---- Summary helpers ----

// Multi-period XmR summary (handles baseline + splits)
function updateXmRMultiSummary(segments, totalPoints) {
  if (!summaryDiv) return;

  if (!segments || segments.length === 0) {
    summaryDiv.innerHTML = "";
    if (capabilityDiv) capabilityDiv.innerHTML = "";
    return;
  }

  const target = getTargetValue();
  const direction = targetDirectionInput ? targetDirectionInput.value : "above";

  let html = `<h3>Summary (XmR chart)</h3>`;
  html += `<p>Total number of points: <strong>${totalPoints}</strong>. `;
  html += `The chart is divided into <strong>${segments.length}</strong> period${segments.length > 1 ? "s" : ""} `;
  html += `(based on the baseline and any splits).</p>`;

  // We'll also keep track of the last period's signals for the capability badge
  let lastPeriodSignals = [];
  let lastPeriodCapability = null;
  let lastPeriodHasCapability = false;

  segments.forEach((seg, idx) => {
    const { startIndex, endIndex, labelStart, labelEnd, result } = seg;
    const { mean, ucl, lcl, sigma, avgMR, baselineCountUsed } = result;

    const points = result.points;
    const n = points.length;
    const values = points.map(p => p.y);
    const nBeyond = points.filter(p => p.beyondLimits).length;

    const runFlags = detectLongRuns(values, mean, 8);
    const nRunPoints = runFlags.filter(Boolean).length;
    const hasRunViolation = nRunPoints > 0;
    const hasTrend = detectTrend(values, 6);

    const signals = [];
    if (nBeyond > 0) {
      signals.push("one or more points beyond the control limits");
    }
    if (hasRunViolation) {
      signals.push("a run of 8 or more points on one side of the mean");
    }
    if (hasTrend) {
      signals.push("a trend of 6 or more points all increasing or all decreasing");
    }

    let capability = null;
    if (target !== null && sigma > 0) {
      capability = computeTargetCapability(mean, sigma, target, direction);
    }

    // How many points in this period meet the target?
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

    const rangeText =
      labelStart !== undefined && labelEnd !== undefined
        ? `points ${startIndex + 1}–${endIndex + 1} (${labelStart} to ${labelEnd})`
        : `points ${startIndex + 1}–${endIndex + 1}`;

    html += `<h4>${periodLabel}</h4>`;
    html += `<ul>`;
    html += `<li>Coverage: <strong>${rangeText}</strong> – ${n} point${n !== 1 ? "s" : ""}.</li>`;

    if (baselineCountUsed && baselineCountUsed < n) {
      html += `<li>Baseline for this period: first <strong>${baselineCountUsed}</strong> point${baselineCountUsed !== 1 ? "s" : ""} used to calculate mean and limits.</li>`;
    } else {
      html += `<li>Baseline for this period: all points in this period used to calculate mean and limits.</li>`;
    }

    html += `<li>Mean: <strong>${mean.toFixed(3)}</strong>; control limits: <strong>LCL = ${lcl.toFixed(3)}</strong>, <strong>UCL = ${ucl.toFixed(3)}</strong>.</li>`;
    html += `<li>Estimated σ (from MR): <strong>${sigma.toFixed(3)}</strong> (average MR = ${avgMR.toFixed(3)}).</li>`;

    if (target !== null) {
      html += `<li>Target: <strong>${target}</strong> (${direction === "above" ? "at or above is better" : "at or below is better"}). `;
      if (targetCoverageText) {
        html += targetCoverageText + `</li>`;
      } else {
        html += `Target coverage not calculated for this period.</li>`;
      }
    }

    if (signals.length === 0) {
      html += `<li><strong>Special cause:</strong> No rule breaches detected in this period (points within limits, no long runs or clear trend). Pattern is consistent with common-cause variation, but always interpret in clinical context.</li>`;
    } else {
      html += `<li><strong>Special cause:</strong> In this period, signals suggesting special-cause variation were detected based on: ${signals.join("; ")}.</li>`;
    }

    if (capability && sigma > 0) {
      if (signals.length === 0) {
        html += `<li><strong>Estimated capability (this period):</strong> if the process remains stable, about <strong>${(capability.prob * 100).toFixed(1)}%</strong> of future points are expected to meet the target.</li>`;
      } else {
        html += `<li><strong>Capability:</strong> a target has been set, but because special-cause signals are present in this period, any capability estimate would be unreliable.</li>`;
      }
    }

    html += `</ul>`;

    // Remember last period's signals + capability for the badge
if (idx === segments.length - 1) {
  lastPeriodSignals = signals;
  lastPeriodCapability = capability;
  lastPeriodHasCapability = sigma > 0 && !!capability;

  // Also store a structured summary for the SPC helper (last / current period)
  lastXmRAnalysis = {
    mean,
    ucl,
    lcl,
    sigma,
    avgMR,
    n,
    signals: signals.slice(),
    hasTrend,
    hasRunViolation,
    baselineCountUsed,
    target,
    direction,
    capability,
    isStable: signals.length === 0
  };
}
  });

  if (target !== null && segments.length > 1) {
    html += `<p><em>Note:</em> comparing means, limits and target performance between periods gives an indication of whether the process has changed after interventions.</p>`;
  }

  summaryDiv.innerHTML = html;

  // Capability badge – focus on the last period (as a simple headline)
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
        <strong>Process not stable in the last period:</strong> special-cause signals are present.
        Focus on understanding and addressing these causes before relying on capability estimates.
      </div>
    `;
  } else {
    capabilityDiv.innerHTML = "";
  }
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

  // ISO style: 2025-10-02 or 2025-10-02T...
  const isoMatch = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (isoMatch) {
    const y = Number(isoMatch[1]);
    const m = Number(isoMatch[2]);
    const d = Number(isoMatch[3]);
    return new Date(y, m - 1, d);
  }

  // NHS-style day-first: dd/mm/yyyy or dd-mm-yyyy
  const dmMatch = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (dmMatch) {
    let day   = Number(dmMatch[1]);
    let month = Number(dmMatch[2]);
    let year  = Number(dmMatch[3]);
    if (year < 100) year += 2000; // e.g. 25 -> 2025
    return new Date(year, month - 1, day);
  }

  // Fallback: let the browser try
  return new Date(s);
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


// ---- Generate chart button ----

generateButton.addEventListener("click", () => {
  errorMessage.textContent = "";
  if (summaryDiv) summaryDiv.innerHTML = "";
  if (capabilityDiv) capabilityDiv.innerHTML = "";

  if (!rawRows || rawRows.length === 0) {
    errorMessage.textContent = "Please upload a CSV file first.";
    return;
  }

  const dateCol  = dateSelect.value;
  const valueCol = valueSelect.value;

  if (!dateCol || !valueCol) {
    errorMessage.textContent = "Please choose both a date column and a value column.";
    return;
  }

  const axisType = getAxisType();

let parsedPoints;

// --- 1. Build points depending on axis type ---

if (axisType === "date") {
  // Time series: parse dates, sort later
  parsedPoints = rawRows
    .map((row) => {
      const xRaw  = row[dateCol];
      const yRaw  = row[valueCol];

      const d = parseDateValue(xRaw);
      const y = toNumericValue(yRaw);

      if (!isFinite(d.getTime()) || !isFinite(y)) return null;
      return { x: d, y };
    })
    .filter(p => p !== null);
} else {
  // Sequence / category: keep row order, use label text
  parsedPoints = rawRows
    .map((row, idx) => {
      const labelRaw = row[dateCol];   // may be patient ID / category / blank
      const yRaw     = row[valueCol];

      const y = toNumericValue(yRaw);
      if (!isFinite(y)) return null;

      const labelText = (labelRaw !== undefined && labelRaw !== null && String(labelRaw).trim() !== "")
        ? String(labelRaw)
        : `Point ${idx + 1}`;  // fallback if X column is empty

      return {
        x: idx,        // numeric index just for ordering
        y,
        label: labelText
      };
    })
    .filter(p => p !== null);
}

if (parsedPoints.length < 5) {
  errorMessage.textContent = "Not enough valid data points after parsing. Check your column choices.";
  return;
}

// --- 2. Create points + labels for the chart ---

let points;
let labels;

if (axisType === "date") {
  // sort by time
  points = [...parsedPoints].sort((a, b) => a.x - b.x);
  labels = points.map(p => p.x.toISOString().slice(0, 10));
} else {
  // keep sequence order
  points = parsedPoints;
  labels = points.map(p => p.label);
}

  // baseline interpretation
  let baselineCount = null;
  if (baselineInput && baselineInput.value.trim() !== "") {
    const parsed = parseInt(baselineInput.value, 10);
    if (!isNaN(parsed) && parsed >= 2) {
      baselineCount = Math.min(parsed, points.length);
    }
  }

  const chartType = getSelectedChartType();

  // clear existing charts
  if (currentChart) {
    currentChart.destroy();
    currentChart = null;
  }
  if (mrChart) {
    mrChart.destroy();
    mrChart = null;
  }
  if (mrPanel) {
    mrPanel.style.display = "none";
  }

  if (chartType === "run") {
    drawRunChart(points, baselineCount, labels);
} else {
    drawXmRChart(points, baselineCount, labels);
}
});

// ---- Chart drawing ----

function drawRunChart(points, baselineCount, labels) {
  const n = points.length;

  let baselineCountUsed;
  if (baselineCount && baselineCount >= 2) {
    baselineCountUsed = Math.min(baselineCount, n);
  } else {
    baselineCountUsed = n;
  }

  const baselineValues = points.slice(0, baselineCountUsed).map(p => p.y);

  
  const values = points.map(p => p.y);
  const median = computeMedian(baselineValues);
  
  // Keep annotation date dropdown in sync with the current chart dates
  populateAnnotationDateOptions(labels);

  // Detect runs of >= 8 points on same side of median
  const runFlags = detectLongRuns(values, median, 8);

  // Colours: orange for run violations, dark blue otherwise
  const pointColours = values.map((_, i) => (runFlags[i] ? "#ff8c00" : "#003f87"));

  const { title, xLabel, yLabel } = getChartLabels(
    "Run Chart",
    "Date",
    "Value"
  );

  const target = getTargetValue();

  const datasets = [
    {
      // DATA LINE
      label: "Value",
      data: values,
      pointRadius: 4,
      pointBackgroundColor: pointColours,
      borderColor: "#003f87", // dark blue
      borderWidth: 2,
      fill: false
    },
    {
      // MEDIAN
      label: "Median",
      data: values.map(() => median),
      borderDash: [6, 4],
      borderWidth: 2,
      borderColor: "#e41a1c", // red-ish
      pointRadius: 0,
      pointHoverRadius: 0,
      fill: false
    }
  ];

  if (target !== null) {
    datasets.push({
      label: "Target",
      data: values.map(() => target),
      borderDash: [4, 2],
      borderWidth: 2,
      borderColor: "#fdae61", // orange-ish
      pointRadius: 0,
      pointHoverRadius: 0,
      fill: false
    });
  }

  currentChart = new Chart(chartCanvas, {
    type: "line",
    data: {
      labels: labels,
      datasets: datasets
    },
    options: {
      responsive: true,
 	maintainAspectRatio: false,
      plugins: {
        title: {
          display: true,
          text: title,
          font: {
            size: 16,
            weight: "bold"
          }
        },
        legend: {
          display: true,
          position: "bottom",
          align: "center"
        },
        annotation: {
         annotations: buildAnnotationConfig(labels)
       }
      },
      elements: {
        point: {
          radius: 0,
          hoverRadius: 0
        }
      },
      scales: {
        x: {
          grid: { display: false },
          title: {
            display: !!xLabel,
            text: xLabel
          }
        },
        y: {
          grid: { display: false },
          title: {
            display: !!yLabel,
            text: yLabel
          }
        }
      }
    }
  });

  updateRunSummary(points, median, runFlags, baselineCountUsed);
}

function drawXmRChart(points, baselineCount, labels) {
  if (!chartCanvas) return;

  const n = points.length;
  if (n < 12) {
    errorMessage.textContent = "XmR chart needs at least 12 points.";
    return;
  }

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
  const globalResult = computeXmR(points, baselineCount);

  // ----- Global arrays for plotting -----
  const values = points.map(p => p.y);

  const meanLine      = new Array(n).fill(NaN);
  const uclLine       = new Array(n).fill(NaN);
  const lclLine       = new Array(n).fill(NaN);
  const oneSigmaUp    = new Array(n).fill(NaN);
  const oneSigmaDown  = new Array(n).fill(NaN);
  const twoSigmaUp    = new Array(n).fill(NaN);
  const twoSigmaDown  = new Array(n).fill(NaN);
  const pointColours  = new Array(n).fill("#003f87");

  let anySigma = false;

  // We'll collect per-period results for the new summary
  const segmentSummaries = [];

  // ----- Per-segment XmR -----
  for (let s = 0; s < segmentStarts.length; s++) {
    const start = segmentStarts[s];
    const end   = segmentEnds[s];

    const segPoints = points.slice(start, end + 1);

    // Only the first segment uses the user-specified baseline;
    // later segments use all their points as baseline.
    const segBaseline = s === 0 ? baselineCount : null;

    const segResult = computeXmR(segPoints, segBaseline);
    const segPts    = segResult.points;
    const mean      = segResult.mean;
    const ucl       = segResult.ucl;
    const lcl       = segResult.lcl;
    const sigma     = segResult.sigma;

    // Store summary info for this period
    segmentSummaries.push({
      startIndex: start,
      endIndex: end,
      labelStart: labels[start],
      labelEnd: labels[end],
      result: segResult
    });

    for (let i = 0; i < segPts.length; i++) {
      const globalIdx = start + i;

      // Flag special-cause points within this segment
      if (segPts[i].beyondLimits) {
        pointColours[globalIdx] = "#d73027";
      }

      // Centre line & limits
      meanLine[globalIdx] = mean;
      uclLine[globalIdx]  = ucl;
      lclLine[globalIdx]  = lcl;

      // Sigma lines (if we have a valid sigma)
      if (sigma && sigma > 0) {
        anySigma = true;
        oneSigmaUp[globalIdx]   = mean + sigma;
        oneSigmaDown[globalIdx] = mean - sigma;
        twoSigmaUp[globalIdx]   = mean + 2 * sigma;
        twoSigmaDown[globalIdx] = mean - 2 * sigma;
      }
    }
  }

  // ----- Build datasets -----
  const datasets = [];

  // Main values
  datasets.push({
    label: "Value",
    data: values,
    borderColor: "#003f87",
    backgroundColor: "#003f87",
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
      borderColor: "#d73027",
      borderDash: [6, 4],
      pointRadius: 0
    },
    {
      label: "UCL (3σ)",
      data: uclLine,
      borderColor: "#2ca25f",
      borderDash: [4, 4],
      pointRadius: 0
    },
    {
      label: "LCL (3σ)",
      data: lclLine,
      borderColor: "#2ca25f",
      borderDash: [4, 4],
      pointRadius: 0
    }
  );

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

  // ----- Target line (optional) -----
  const target = getTargetValue();
  if (target !== null) {
    datasets.push({
      label: "Target",
      data: values.map(() => target),
      borderColor: "#fdae61",   // NHS-style orange
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
  if (currentChart) {
    currentChart.destroy();
  }

  const title = chartTitleInput.value.trim() || "I-MR Chart";

  currentChart = new Chart(chartCanvas, {
    type: "line",
    data: {
      labels: labels,
      datasets: datasets
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        title: {
          display: true,
          text: title,
          font: {
            size: 16,
            weight: "bold"
          }
        },
        legend: {
          display: true,
          position: "bottom",
          align: "center"
        },
        annotation: {
          annotations: buildAnnotationConfig(labels)
        }
      },
      elements: {
        point: {
          radius: 0,
          hoverRadius: 0
        }
      },
      scales: {
        x: {
          grid: { display: false },
          title: {
            display: !!xAxisLabelInput.value.trim(),
            text: xAxisLabelInput.value.trim()
          }
        },
        y: {
          grid: { display: false },
          title: {
            display: !!yAxisLabelInput.value.trim(),
            text: yAxisLabelInput.value.trim()
          }
        }
      }
    }
  });

  // ----- New: multi-period summary -----
  if (segmentSummaries.length > 0) {
    updateXmRMultiSummary(segmentSummaries, points.length);
  } else {
    // Fallback: treat whole series as a single period
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

  // Use the last period for the MR chart (as a simple, focused view)
  const lastSegmentResult =
    segmentSummaries.length > 0
      ? segmentSummaries[segmentSummaries.length - 1].result
      : globalResult;

  if (showMR && lastSegmentResult) {
    drawMRChart(lastSegmentResult, labels);
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


// MR chart: average MR as centre, UCL = 3.268 * avgMR, LCL = 0
function drawMRChart(result, labels) {
  if (!mrPanel || !mrChartCanvas) return;

  const mrValues = result.mrValues;
  const mrLabels = labels.slice(1);

  if (!mrValues || mrValues.length === 0) {
    mrPanel.style.display = "none";
    return;
  }

  const avgMR = result.avgMR;
  const centre = avgMR;
  const uclMR = avgMR * 3.268; // D4 for n=2
  const lclMR = 0;

  mrPanel.style.display = "block";

  mrChart = new Chart(mrChartCanvas, {
    type: "line",
    data: {
      labels: mrLabels,
      datasets: [
        {
          label: "Moving range",
          data: mrValues,
          pointRadius: 4,
          pointBackgroundColor: "#003f87",
          borderColor: "#003f87",
          borderWidth: 2,
          fill: false
        },
        {
          label: "Average MR",
          data: mrValues.map(() => centre),
          borderDash: [6, 4],
          borderWidth: 2,
          borderColor: "#e41a1c",
          pointRadius: 0,
          pointHoverRadius: 0,
          fill: false
        },
        {
          label: "UCL (MR)",
          data: mrValues.map(() => uclMR),
          borderDash: [4, 4],
          borderWidth: 2,
          borderColor: "#1a9850",
          pointRadius: 0,
          pointHoverRadius: 0,
          fill: false
        },
        {
          label: "LCL (MR)",
          data: mrValues.map(() => lclMR),
          borderDash: [4, 4],
          borderWidth: 2,
          borderColor: "#1a9850",
          pointRadius: 0,
          pointHoverRadius: 0,
          fill: false
        }
      ]
    },
    options: {
      responsive: true,
	maintainAspectRatio: false,
      plugins: {
        title: {
          display: true,
          text: "Moving Range chart",
          font: {
            size: 14,
            weight: "bold"
          }
        },
        legend: {
          display: true,
          position: "bottom",
          align: "center"
        }
      },
      elements: {
        point: {
          radius: 0,
          hoverRadius: 0
        }
      },
      scales: {
        x: {
          grid: { display: false },
          title: {
            display: true,
            text: "Date (second and subsequent points)"
          }
        },
        y: {
          grid: { display: false },
          title: {
            display: true,
            text: "Moving range"
          }
        }
      }
    }
  });
}

// ---- AI helper function  -----


// Helper for matching user questions to keyword patterns
function matchesKeywords(q, keywords) {
  return keywords.some((k) => {
    // If k is an array, treat it as "all these words must appear"
    if (Array.isArray(k)) {
      return k.every((word) => q.includes(word));
    }
    // Otherwise simple substring match
    return q.includes(k);
  });
}


function answerSpcQuestion(question) {
  const q = question.trim().toLowerCase();
  if (!q) {
    return "Please type a question about SPC or your chart (for example: “Is the process stable?” or “What is a moving range chart?”).";
  }

  // 0. General SPC knowledge – works even without a chart
  const generalFaq = [
    {
      // What is a control chart / why use it?
      keywords: [
        "what is a control chart",
        "what's a control chart",
        "control chart",
        "shewhart chart",
        "spc chart",
        "statistical process control chart",
        "why use a control chart",
        "use one in healthcare"
      ],
      answer:
        "A control chart (also called a Shewhart chart or SPC chart) is a line chart with a centre line and upper and lower control limits calculated from your own data. " +
        "It helps you see whether the variation you are seeing is just routine noise (common-cause variation) or whether there is evidence that something in the system has changed (special-cause variation). " +
        "In healthcare this means you can tell the difference between random ups and downs and true improvement or deterioration, so you can respond appropriately and avoid over-reacting to every small change."
    },
    {
      // Run vs control chart
      keywords: [
        "run chart vs control chart",
        "difference between run chart and control chart",
        "run or control chart",
        "when to use a run chart",
        "when to use a control chart"
      ],
      answer:
        "A run chart shows data over time with a median line and uses simple rules (trend, shift, runs) to look for signals of change. It is quick to produce and works well when you are starting out or have limited data. " +
        "A control chart adds statistically-based control limits around a mean and can tell you more clearly whether the process is stable and how much natural variation there is. " +
        "A common approach in healthcare is to start with a run chart and move to a control chart once you have enough data and want a more precise view of stability and capability."
    },
    {
      // How many data points do I need?
      keywords: [
        "how many data points",
        "how many points do i need",
        "minimum data points",
        "minimum points",
        "how much data before i can make a control chart",
        "how much data before i can make an spc chart"
      ],
      answer:
        "Control charts work best when they have enough data to estimate the natural variation reliably. " +
        "A common rule is to have at least 12 data points for any control chart, and 20 or more points when you are using an individuals (XmR or I chart) or X-bar and S chart. " +
        "With fewer points the limits can be misleading, so many teams start with a run chart and convert to a control chart once more data are available."
    },
    {
      // Special-cause rules
      keywords: [
        "special cause rules",
        "spc rules",
        "signal rules",
        "rules for detecting special cause",
        "how do i know if there is a signal",
        "what are the rules for a signal"
      ],
      answer:
        "Typical special-cause rules on a control chart include: a point outside a control limit; a run of 7–8 or more points on the same side of the centre line; " +
        "a trend of 6 or more points steadily rising or falling; and two out of three points near a control limit on the same side. " +
        "Some guides also include 15 points very close to the mean as a signal of something unusual. " +
        "These patterns are very unlikely to occur by chance if the process is stable, so when they appear it is worth looking for a change in the real system."
    },
    {
      // Sigma lines
      keywords: [
        "sigma line",
        "sigma lines",
        "1 sigma",
        "2 sigma",
        "3 sigma",
        "sigma limits",
        "what is sigma"
      ],
      answer:
        "Sigma is a way of describing how much the data vary – it is closely related to the standard deviation. " +
        "Sigma lines are drawn at fixed multiples of this variation above and below the average, for example ±1 sigma, ±2 sigma and ±3 sigma. " +
        "The ±3 sigma lines are the classic control limits (UCL and LCL). If the process is stable, points beyond ±3 sigma are very rare, so they are treated as potential signals that something has changed."
    },
    {
      // Control limits vs spec/target
      keywords: [
        "control limits vs specification",
        "control limits vs spec limits",
        "control limits vs target",
        "difference between control and specification limits",
        "difference between control and target limits"
      ],
      answer:
        "Control limits are calculated from how your process is currently behaving and show the range you would expect from common-cause variation. " +
        "Specification or target limits are external goals or requirements, such as a waiting-time standard or a clinical threshold. " +
        "You should not move control limits to match a target. Instead, use the chart to ask: “given the current process, how often will we meet the target, and what needs to change if that is not good enough?”"
    },
    {
      // Which chart should I use?
      keywords: [
        "which control chart should i use",
        "what chart should i use",
        "choose right chart",
        "how do i pick the right chart",
        "what spc chart should i use"
      ],
      answer:
        "The right chart depends mainly on the type of data and the way it is collected. " +
        "For individual continuous measurements over time (for example, length of stay per patient or daily waiting time) an XmR chart is often appropriate. " +
        "For small subgroups of measurements at each time point you might use X-bar and R or X-bar and S charts. " +
        "For counts or percentages (falls, infections, readmissions, proportion achieving a standard) you usually use attribute charts such as p, np, c or u charts."
    },
    {
      // Moving range / MR chart
      keywords: [
        "moving range",
        "moving-range",
        "moving range chart",
        "mr chart",
        "mr-chart",
        "use the moving range",
        "interpret the moving range",
        ["use", "moving", "range"]
      ],
      answer:
        "On an XmR chart the moving range (MR) chart shows how much each point changes from the one before it. " +
        "The average moving range is used to estimate the underlying variation (sigma), which then determines the control limits on the X chart. " +
        "You can use the moving range chart to spot sudden jumps in the data, measurement issues, or changes in the short-term variation, even when the X chart itself still looks fairly stable."
    },
    {
      // Normal distribution / normality
      keywords: [
        "normal distribution",
        "normality",
        "do control charts assume normal",
        "do my data need to be normal",
        "does spc require normal distribution"
      ],
      answer:
        "Classical explanations of control charts often mention the normal distribution, but in practice Shewhart charts such as XmR are quite robust to non-normal data. " +
        "You do not need perfectly normal data before you can use a control chart. " +
        "If the data are very skewed or have natural limits, be cautious when interpreting probabilities and capability indices, and focus on the presence or absence of clear signals rather than exact percentages."
    },
    {
      // Process unstable / out of control – what to do?
      keywords: [
        "process unstable",
        "unstable process",
        "out of control",
        "chart unstable",
        "what should i do if the chart is out of control",
        "what should i do if unstable"
      ],
      answer:
        "If your chart shows special-cause signals, treat this as a prompt to understand what changed in the system rather than simply adjusting the chart. " +
        "Look for real-world explanations around the time of the signals: new policies, staffing changes, case-mix changes, data issues, or improvement tests. " +
        "Decide whether the change is desirable or not. For desirable changes you may later re-baseline the chart around the new level; for undesirable changes you may plan PDSA cycles to remove or reduce the special cause."
    },
    {
      // Process capability – what is it?
      keywords: [
        "process capability",
        "capability",
        "cp",
        "cpk",
        "meeting the target",
        "how capable is the process"
      ],
      answer:
        "Process capability asks how well a stable process can meet a particular target or specification. " +
        "In simple terms it answers: “if the process continues like this, what proportion of future results will be on the desired side of the target?” " +
        "Capability measures are only meaningful when the process is stable, so you normally sort out special-cause signals first and then assess capability. " +
        "In healthcare the focus is often on an estimated percent meeting a standard rather than formal Cp/Cpk indices."
    },
    {
      // Why annotate charts?
      keywords: [
        "annotate chart",
        "add notes to chart",
        "why annotate",
        "why put comments on the chart"
      ],
      answer:
        "Annotations link what you see on the chart to what was happening in the real world. " +
        "Marking important events such as new guidelines, staffing changes, pathway redesigns or data definition changes helps people understand why the pattern changed and prevents mis-interpretation later. " +
        "Annotated charts are much easier to use in meetings because they tell the story of the improvement work, not just the numbers."
    },
    {
      // Run chart with few data points
      keywords: [
        "run chart few data points",
        "only a few data points",
        "i don't have much data",
        "not much data yet",
        "why use a run chart first"
      ],
      answer:
        "Run charts are recommended when you are starting out and have only a small number of data points, because they are simpler and need fewer observations to begin giving useful feedback. " +
        "You can begin with around 8–10 points on a run chart and apply simple run-chart rules to look for changes. " +
        "As you collect more data and want to estimate control limits and capability, you can then move to a control chart."
    },
    {
      // Quarterly / annual data
      keywords: [
        "quarterly data",
        "annual data",
        "why not use quarterly",
        "why not use annual data",
        "can i make a chart with quarterly data"
      ],
      answer:
        "Quarterly or annual data give very few points over a reasonable time period and can hide important patterns of change. " +
        "They also make it hard to apply run or control-chart rules, which rely on having enough points to detect signals. " +
        "Where possible, collect data at a more frequent interval such as weekly or monthly so that your charts can provide timely and reliable feedback."
    },
    {
      // Target / goal line on chart
      keywords: [
        "target line",
        "goal line",
        "add target to chart",
        "how do i show the target on the chart"
      ],
      answer:
        "On an SPC chart the control limits come from the data, but you can still add a separate horizontal line to show a target or goal. " +
        "This makes it clear whether the current stable process is good enough compared with what you are aiming for. " +
        "If the average is far from the target, the chart is telling you that improvement work should focus on shifting the process, not on tightening the control limits."
    },
    {
      // Common vs special cause – concept
      keywords: [
        "common cause",
        "special cause",
        "common and special cause",
        "difference between common and special cause"
      ],
      answer:
        "Common-cause variation is the ordinary, expected noise in a stable system – the small ups and downs that are always present. " +
        "Special-cause variation comes from specific circumstances or changes that affect the system at particular times, such as a new process, an outbreak, or a data error. " +
        "Control charts help you distinguish between the two so you can decide when to redesign the system and when to investigate particular events."
    },
    {
      // Can a process be in control but still bad?
      keywords: [
        "in control but bad",
        "process in control but poor performance",
        "stable but poor",
        "is a process in control always good"
      ],
      answer:
        "No. A process can be in control (showing only common-cause variation) and still be performing at an unacceptable level, for example when an average waiting time is stable but far above the standard. " +
        "In that case SPC tells you the current system is delivering exactly what it is designed to deliver – poor performance – and that improvement requires changing the system, not just reacting to individual points."
    },
    {
      // When to recalc control limits / rebaseline
      keywords: [
        "recalculate control limits",
        "recalc control limits",
        "re-baseline",
        "rebaseline",
        "when should i recalculate the limits",
        "when to reset the baseline",
        "when to recalc limits"
      ],
      answer:
        "Recalculate control limits when you have evidence that the process has genuinely shifted to a new, stable level – for example, a clear and sustained signal following an intentional change. " +
        "A common approach is to split the chart at the point of change and use the later data to estimate a new mean and limits. " +
        "Avoid constantly recomputing limits in response to every small fluctuation, as this hides real changes and defeats the purpose of SPC."
    }
  ];

  // Try general SPC FAQ answers first
  for (const item of generalFaq) {
    if (matchesKeywords(q, item.keywords)) {
      return item.answer;
    }
  }



  // 1. From here on: chart-specific interpretation (currently XmR only)
  const chartType = getSelectedChartType ? getSelectedChartType() : "xmr";

  if (chartType !== "xmr") {
    return (
      "I can answer general SPC questions for any chart type, but the automated interpretation of this specific chart " +
      "currently focuses on XmR (I-MR) charts. Please switch to an XmR chart if you want a detailed interpretation of stability and capability."
    );
  }

  if (!lastXmRAnalysis) {
    return "I don't have an XmR analysis yet. Please generate an XmR chart first, then ask about stability, signals, limits or capability.";
  }

  const a = lastXmRAnalysis;
  const splits = (a && a.splits) || [];
  const lines = [];

  // 2. Stability / special-cause questions
  if (
    q.includes("stable") ||
    q.includes("stability") ||
    q.includes("in control") ||
    q.includes("out of control") ||
    q.includes("special cause") ||
    q.includes("common cause") ||
    q.includes("signal") ||
    q.includes("signals") ||
    q.includes("run rule") ||
    q.includes("run rules")
  ) {
    if (a.isStable) {
      lines.push(
        "The most recent segment of your XmR chart looks stable: no special-cause rules are triggered. " +
          "The ups and downs you see are consistent with common-cause variation in the current system."
      );
    } else {
      lines.push(
        "The most recent segment of your XmR chart does not look stable. One or more SPC rules are triggered, suggesting special-cause variation."
      );
      if (a.signals && a.signals.length > 0) {
        lines.push("Signals detected: " + a.signals.join("; ") + ".");
      }
      lines.push(
        "It is worth exploring what was happening in the system around the times where signals appear – for example changes in process, staffing, demand or data quality."
      );
    }
  }

  // 3. Capability / target questions
  if (
    q.includes("capability") ||
    q.includes("capable") ||
    q.includes("target") ||
    q.includes("specification") ||
    q.includes("spec limit") ||
    q.includes("meeting the target")
  ) {
    if (a.target == null) {
      lines.push(
        "No target has been set in the tool. To discuss capability, please enter a target value and whether it is better for values to be above or below that target."
      );
    } else if (!a.isStable) {
      lines.push(
        "Because the process does not appear stable, any capability estimate will be unreliable. " +
          "It is usually better to address the special-cause variation first, then reassess capability once the process is behaving more consistently."
      );
    } else if (a.capability && typeof a.capability.prob === "number") {
      const prob = (a.capability.prob * 100).toFixed(1);
      const dirText = a.direction === "above" ? "at or above" : "at or below";
      lines.push(
        `Based on the current stable process and a target of ${a.target} (${dirText} the target), the estimated proportion of future points meeting the target is about ${prob}%.`
      );
      lines.push(
        "This is an estimate based on the current level of variation. If the process changes, the capability will also change."
      );
    } else {
      lines.push(
        "I could not calculate capability from the current chart. Please check that a numeric target and direction have been set and that there is some variation in the data."
      );
    }
  }

  // 4. Mean / limits / sigma questions
  if (
    q.includes("limit") ||
    q.includes("ucl") ||
    q.includes("lcl") ||
    q.includes("sigma") ||
    q.includes("mean") ||
    q.includes("average") ||
    q.includes("centre") ||
    q.includes("center") ||
    q.includes("central line")
  ) {
    if (a.mean != null && a.lcl != null && a.ucl != null && a.sigma != null) {
      lines.push(
        `For the current segment, the estimated mean (centre line) is ${a.mean.toFixed(3)}.`,
      );
      lines.push(
        `The control limits are approximately LCL = ${a.lcl.toFixed(3)} and UCL = ${a.ucl.toFixed(3)}.`
      );
      lines.push(
        `The estimated sigma (the typical amount of variation, based on the moving ranges) is about ${a.sigma.toFixed(3)}. ` +
          "Points outside the control limits, or unusual patterns within them, suggest special-cause variation."
      );
    } else {
      lines.push("I do not have a full set of statistics (mean, limits and sigma) for this segment.");
    }
  }

  // 5. Splits / baseline / phases
  if (
    q.includes("split") ||
    q.includes("baseline") ||
    q.includes("phase") ||
    q.includes("segment") ||
    q.includes("change point") ||
    q.includes("before") && q.includes("after")
  ) {
    if (splits.length === 0) {
      lines.push(
        "No splits have been added, so the chart is treating all points as one baseline. " +
          "If you know there was a deliberate change to the system at a specific time, you can add a split so that the tool estimates separate baselines before and after the change."
      );
    } else {
      lines.push(
        `You have added ${splits.length} split${splits.length > 1 ? "s" : ""}. Each split marks a point where you want the tool to treat the data as a new segment with its own mean and limits.`
      );
      lines.push(
        "Interpret changes within each segment separately. Large shifts between segments can indicate the effect of planned changes or other step-changes in the system."
      );
    }
  }

  // 6. If nothing matched above, give a general summary for this chart
  if (lines.length === 0) {
    if (a.isStable) {
      lines.push(
        "Overall, the current XmR chart segment looks stable: there are no strong signals of special-cause variation. " +
          "The points vary around a consistent average within the control limits."
      );
    } else {
      lines.push(
        "Overall, the current XmR chart segment appears unstable: at least one SPC rule is triggered, suggesting special-cause variation. " +
          "It would be useful to review the timing of the signals against any known changes in the system."
      );
    }

    if (a.target != null && a.capability && typeof a.capability.prob === "number") {
      const prob = (a.capability.prob * 100).toFixed(1);
      const dirText = a.direction === "above" ? "at or above" : "at or below";
      lines.push(
        `With the current target of ${a.target}, the estimated proportion of future points ${dirText} the target is about ${prob}%, assuming the process continues to behave in the same way.`
      );
    }
  }

  // 7. Always end with a gentle safety reminder
  lines.push(
    "Always interpret SPC results alongside clinical or operational context, and involve your quality improvement or analytics team if you are planning major changes based on these findings."
  );

  return lines.join(" ");
}



// ---- Download chart as PNG ----

if (downloadBtn) {
  downloadBtn.addEventListener("click", () => {
    if (!currentChart) {
      alert("Please generate a chart first.");
      return;
    }
    const link = document.createElement("a");
    link.href = currentChart.toBase64Image(); // Chart.js helper
    link.download = "spc-chart.png";
    link.click();
  });
}

if (addSplitButton) {
  addSplitButton.addEventListener("click", () => {
    if (!splitPointSelect) return;

    const value = splitPointSelect.value;
    if (value === "") {
      alert("Please choose a point to split after.");
      return;
    }

    const idx = parseInt(value, 10);
    if (!Number.isInteger(idx)) return;

    if (!splits.includes(idx)) {
      splits.push(idx);
    }

    // Rebuild chart if we're on XmR
    if (getSelectedChartType() === "xmr") {
      generateButton.click();
    }
  });
}

if (aiAskButton && aiQuestionInput && spcHelperAnswer) {
  aiAskButton.addEventListener("click", () => {
    const q = aiQuestionInput.value || "";
    const ans = answerSpcQuestion(q);
    spcHelperAnswer.innerHTML = `<p>${escapeHtml(ans)}</p>`;
  });

  aiQuestionInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      aiAskButton.click();
    }
  });
}


if (clearSplitsButton) {
  clearSplitsButton.addEventListener("click", () => {
    splits = [];

    if (splitPointSelect) {
      splitPointSelect.value = "";
    }

    if (getSelectedChartType() === "xmr") {
      generateButton.click();
    }
  });
}

// -----------------------------
// Help section toggle
// -----------------------------
function toggleHelpSection() {
  const help = document.getElementById("helpSection");
  const helper = document.getElementById("spcHelperPanel");

  const isHidden =
    !help ||
    help.style.display === "none" ||
    help.style.display === "";

  if (isHidden) {
    if (help) {
      help.style.display = "block";
      help.scrollIntoView({ behavior: "smooth" });
    }
    if (helper) {
      helper.classList.add("visible");   // show AI helper
    }
  } else {
    if (help) {
      help.style.display = "none";
    }
    if (helper) {
      helper.classList.remove("visible"); // hide AI helper
    }
  }
}

const helpToggleButton = document.getElementById("helpToggleButton");
if (helpToggleButton) {
  helpToggleButton.addEventListener("click", toggleHelpSection);
}


if (downloadPdfBtn) {
  downloadPdfBtn.addEventListener("click", () => {
    const reportElement = document.getElementById("reportContent");
    if (!reportElement) {
      alert("Report content not found.");
      return;
    }
    if (!currentChart) {
      alert("Please generate a chart first.");
      return;
    }

    // Basic options – you can tweak orientation/format later
    const opt = {
      margin:       10,
      filename:     "spc-report.pdf",
      image:        { type: "jpeg", quality: 0.98 },
      html2canvas:  { scale: 2, scrollY: -window.scrollY },
      jsPDF:        { unit: "mm", format: "a4", orientation: "landscape" }
    };


    html2pdf().set(opt).from(reportElement).save();
  });
}


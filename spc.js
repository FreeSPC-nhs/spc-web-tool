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

function resetAll() {
  // --- Clear stored data ---
  rawRows = [];
  annotations = [];
  splits = [];
  lastXmRAnalysis = null;

  // --- Reset file input ---
  if (fileInput) fileInput.value = "";

  // --- Hide column selectors ---
  if (columnSelectors) columnSelectors.style.display = "none";

  // --- Reset dropdowns ---
  if (dateSelect) dateSelect.innerHTML = "";
  if (valueSelect) valueSelect.innerHTML = "";
  if (splitPointSelect) splitPointSelect.innerHTML = "";

  // --- Reset text inputs ---
  if (baselineInput) baselineInput.value = "";
  if (chartTitleInput) chartTitleInput.value = "";
  if (xAxisLabelInput) xAxisLabelInput.value = "";
  if (yAxisLabelInput) yAxisLabelInput.value = "";
  if (targetInput) targetInput.value = "";
  if (annotationDateInput) annotationDateInput.value = "";
  if (annotationLabelInput) annotationLabelInput.value = "";

  // --- Reset target direction dropdown ---
  if (targetDirectionInput) targetDirectionInput.value = "above";

  // --- Clear any error message ---
  if (errorMessage) errorMessage.textContent = "";

  // --- Clear summary & capability output ---
  if (summaryDiv) summaryDiv.innerHTML = "";
  if (capabilityDiv) capabilityDiv.innerHTML = "";

  // --- Destroy main chart ---
  if (currentChart) {
    currentChart.destroy();
    currentChart = null;
  }

  // --- Destroy MR chart ---
  if (mrChart) {
    mrChart.destroy();
    mrChart = null;
  }

  // --- Hide MR panel ---
  if (mrPanel) mrPanel.style.display = "none";

  // --- Reset AI helper ---
  if (spcHelperAnswer) spcHelperAnswer.textContent = "";
  if (aiQuestionInput) aiQuestionInput.value = "";

  // Optionally hide the whole AI helper panel:
  if (spcHelperPanel) spcHelperPanel.style.display = "none";

  // --- Reset data editor ---
  if (dataEditorTextarea) dataEditorTextarea.value = "";
  if (dataEditorOverlay) dataEditorOverlay.style.display = "none";

  console.log("All elements reset.");
}



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

function answerSpcQuestion(question) {
  const q = question.trim().toLowerCase();
  if (!q) {
    return "Please type a question about SPC or your chart (for example: “Is the process stable?” or “What is a moving range chart?”).";
  }

  // ----- 0. General SPC knowledge -----
  const generalFaq = [
    {
      keywords: [
        "moving range", "mr chart", "m-r chart",
        "use the moving range", "interpret the moving range",
        ["moving", "range"]
      ],
      answer:
        "A moving range (MR) chart shows how much each value changes from one point to the next. " +
        "On an XmR chart, the X chart shows the individual values over time and the MR chart shows the size of the step between consecutive points. " +
        "If the moving ranges are mostly small and within their limits, the short-term variation looks stable. Large spikes in the moving range can indicate a one-off shock or a change in how the process behaves."
    },

    {
      keywords: ["xmr", "xm r", "i-mr", "individuals chart", "individual chart"],
      answer:
        "An XmR chart (also called an I-MR chart) is used when you have one measurement at each time point – for example, length of stay per day or waiting time for one patient. " +
        "The X chart shows the individual values with a centre line and control limits. The MR chart shows the absolute difference between each pair of consecutive points. " +
        "The moving ranges are used to estimate the natural variation (sigma), which then gives the control limits on the X chart."
    },

    {
      keywords: ["less than 12", "fewer than 12", "< 12", "minimum data", "how many points", "how many data points"],
      answer:
        "For an XmR chart we usually recommend at least 12 data points, and ideally 20 or more. " +
        "The control limits are based on the average moving range. With too few points this estimate is unstable and the limits can be misleading. " +
        "Using at least 12 points gives a more reliable picture of natural variation."
    },

    {
      keywords: ["sigma line", "sigma lines", "1 sigma", "2 sigma", "3 sigma", "standard deviation"],
      answer:
        "Sigma describes the typical amount of variation in your data. Sigma lines are placed at ±1, ±2 and ±3 sigma around the average. " +
        "The ±3 sigma lines are the traditional control limits (UCL/LCL). Points outside ±3 sigma suggest special-cause variation."
    },

    {
      keywords: ["ucl", "lcl", "control limit", "control limits"],
      answer:
        "Control limits (UCL and LCL) show the range of values you would expect from common-cause variation if the process is stable. " +
        "They are not targets. Points outside the limits, or unusual patterns within them, may indicate special-cause variation."
    },

    {
      keywords: ["process capability", "capability", "capable", "meeting the target", "specification"],
      answer:
        "Process capability describes how well a stable process can meet a given target. It answers: “If the process continues like this, what proportion of results will meet the target?” " +
        "Capability should only be assessed when the process is stable."
    },

    {
      keywords: ["common cause", "special cause"],
      answer:
        "Common-cause variation is the natural background noise in a stable process. Special-cause variation occurs when something different happens, such as a change in process, staffing, case-mix or data quality. " +
        "SPC helps you distinguish between the two."
    },

    {
      keywords: ["run rule", "run rules", "spc rule", "spc rules", "signal", "signals"],
      answer:
        "SPC rules help you detect unusual patterns that are unlikely to occur in a stable system. Examples include a point outside the control limits, a long run of points on one side of the mean, or a steady trend. " +
        "These patterns are treated as potential signals of special-cause variation."
    },

    {
      keywords: ["xbar", "x-bar", "x̄", "r chart", "subgroup", "subgrouping"],
      answer:
        "X-bar and R charts are used when you have groups (subgroups) of measurements at each time point. " +
        "The X-bar chart plots the subgroup average while the R chart plots the subgroup range. This helps monitor both central tendency and within-group variation."
    },

    {
      keywords: ["p chart", "u chart", "c chart", "np chart", "attribute chart"],
      answer:
        "Attribute charts are used when you are counting things rather than measuring a continuous value – for example falls, infections or the proportion of patients meeting a standard. " +
        "p-charts and np-charts handle proportions/counts with a fixed denominator; u-charts and c-charts handle varying denominators."
    },

    {
      keywords: ["why spc", "why use spc", "why not use a line", "why can’t i just"],
      answer:
        "Simple line charts show trends but cannot distinguish between common-cause noise and real change. SPC adds statistical structure so you can identify when a change is unlikely to be random. " +
        "This helps avoid over-reacting to routine variation."
    }
  ];

  // ----- Match general SPC questions -----
  for (const item of generalFaq) {
    if (item.keywords.some(k =>
      (Array.isArray(k) ? k.every(word => q.includes(word)) : q.includes(k))
    )) {
      return item.answer;
    }
  }

  // ----- 1. Chart-specific interpretation (XmR only) -----
  const chartType = getSelectedChartType ? getSelectedChartType() : "xmr";
  if (chartType !== "xmr") {
    return (
      "I can answer general SPC questions, but detailed automatic chart interpretation currently applies only to XmR charts. " +
      "Please switch to an XmR chart if you want automated interpretation."
    );
  }

  if (!lastXmRAnalysis) {
    return "Please generate an XmR chart first, then ask about stability, signals, limits or capability.";
  }

  const a = lastXmRAnalysis;
  const splits = (a && a.splits) || [];
  const lines = [];

  // ----- 2. Stability / signals -----
  if (
    q.includes("stable") || q.includes("in control") || q.includes("out of control") ||
    q.includes("special cause") || q.includes("signal") || q.includes("run rule")
  ) {
    if (a.isStable) {
      lines.push(
        "This segment of the XmR chart appears stable: no special-cause rules are triggered and the variation looks consistent with common-cause behaviour."
      );
    } else {
      lines.push(
        "This segment of the XmR chart appears unstable: one or more SPC rules are triggered, suggesting special-cause variation."
      );
      if (a.signals?.length) {
        lines.push("Signals detected: " + a.signals.join("; ") + ".");
      }
      lines.push(
        "It may help to review what was happening in the system at the times signals appear – for example changes in demand, staffing, processes or data definitions."
      );
    }
  }

  // ----- 3. Capability / target -----
  if (
    q.includes("capability") || q.includes("target") ||
    q.includes("specification") || q.includes("meeting the target")
  ) {
    if (a.target == null) {
      lines.push("No target is set. Please enter a target value and whether higher or lower values are better.");
    } else if (!a.isStable) {
      lines.push(
        "Because the process is not stable, capability estimates would be unreliable. Stabilise the process first, then reassess capability."
      );
    } else if (a.capability?.prob != null) {
      const prob = (a.capability.prob * 100).toFixed(1);
      const dirText = a.direction === "above" ? "at or above" : "at or below";
      lines.push(
        `Based on the current stable process and a target of ${a.target} (values ${dirText} the target), about ${prob}% of future points are expected to meet the target.`
      );
    } else {
      lines.push(
        "A capability estimate could not be calculated. Check that a numeric target has been set and that there is some variation in the data."
      );
    }
  }

  // ----- 4. Mean / limits / sigma -----
  if (
    q.includes("limit") || q.includes("ucl") || q.includes("lcl") ||
    q.includes("mean") || q.includes("centre") || q.includes("sigma")
  ) {
    if (a.mean != null && a.lcl != null && a.ucl != null && a.sigma != null) {
      lines.push(`The estimated mean (centre line) is ${a.mean.toFixed(3)}.`);
      lines.push(`Control limits: LCL = ${a.lcl.toFixed(3)}, UCL = ${a.ucl.toFixed(3)}.`);
      lines.push(
        `The estimated sigma (typical variation, based on the moving ranges) is about ${a.sigma.toFixed(3)}.`
      );
    } else {
      lines.push("I do not have complete statistics (mean, limits, sigma) for this segment.");
    }
  }

  // ----- 5. Splits / baselines -----
  if (
    q.includes("split") || q.includes("baseline") ||
    q.includes("phase") || q.includes("segment")
  ) {
    if (splits.length === 0) {
      lines.push(
        "No splits have been added, so the whole chart is treated as one baseline. " +
        "If a known system change occurred, adding a split will give the before-and-after segments their own mean and limits."
      );
    } else {
      lines.push(
        `You have added ${splits.length} split${splits.length > 1 ? "s" : ""}. Each marks a point where the chart begins a new baseline segment.`
      );
      lines.push(
        "Interpret each segment separately. Large shifts between segments may reflect intentional changes or system-level step-changes."
      );
    }
  }

  // ----- 6. If no specific question matched -----
  if (lines.length === 0) {
    if (a.isStable) {
      lines.push(
        "Overall, the current XmR chart segment looks stable: the points vary around a consistent average within the control limits."
      );
    } else {
      lines.push(
        "Overall, the current XmR chart segment appears unstable: at least one SPC rule is triggered, suggesting special-cause variation. " +
        "It may help to compare the timing of the signals with operational or clinical events."
      );
    }

    if (a.target != null && a.capability?.prob != null) {
      const prob = (a.capability.prob * 100).toFixed(1);
      const dirText = a.direction === "above" ? "at or above" : "at or below";
      lines.push(
        `Given the current target of ${a.target}, about ${prob}% of future points are expected to fall ${dirText} the target, assuming the process remains stable.`
      );
    }
  }

  // ----- 7. Final reminder -----
  lines.push(
    "Always interpret SPC results alongside clinical or operational context, and involve your quality improvement or analytics team when planning changes."
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

const resetButton = document.getElementById("resetButton");

if (resetButton) {
  resetButton.addEventListener("click", resetAll);
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


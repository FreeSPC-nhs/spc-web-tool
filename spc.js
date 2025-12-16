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
const aiQuestionInput   = document.getElementById("aiQuestionInput");
const aiAskButton       = document.getElementById("aiAskButton");
const spcHelperPanel    = document.getElementById("spcHelperPanel");

const spcHelperIntro    = document.getElementById("spcHelperIntro");
const spcHelperChipsGeneral = document.getElementById("spcHelperChipsGeneral");
const spcHelperChipsChart   = document.getElementById("spcHelperChipsChart");
const spcHelperOutput   = document.getElementById("spcHelperOutput");



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
  if (aiQuestionInput) aiQuestionInput.value = "";
  if (spcHelperOutput) spcHelperOutput.innerHTML = "";
  if (spcHelperPanel) spcHelperPanel.classList.remove("visible"); // keep consistent with toggleHelpSection()
  renderHelperState();

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

renderHelperState();
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
  const q = (question || "").trim().toLowerCase();
  if (!q) {
    return "Please type a question about SPC or your chart (for example: \"Is the process stable?\" or \"What is a moving range chart?\").";
  }

  // Helper to match simple keyword / phrase based FAQs
  function matchFaq(items, text) {
    for (const item of items) {
      if (
        item.keywords.some(k =>
          Array.isArray(k)
            ? k.every(word => text.includes(word))
            : text.includes(k)
        )
      ) {
        return item.answer;
      }
    }
    return null;
  }

  // ----- 0. Conceptual SPC knowledge (no chart needed at all) -----
  const conceptualFaq = [
    {
      keywords: [
        "what is spc",
        "what is an spc",
        "what is a spc",
        "what is an spc chart",
        "what is a spc chart",
        "what are spc charts",
        "what is statistical process control",
        "spc chart",
        "control chart",
        ["statistical", "process", "control"]
      ],
      answer:
        "Statistical Process Control (SPC) is a way of using time-series charts to separate routine \"common-cause\" variation from unusual \"special-cause\" variation. " +
        "An SPC or control chart plots your measure over time, shows a typical average, and adds upper and lower control limits that represent the range you would expect if the system is stable. " +
        "When the pattern of points breaks simple rules (for example a point outside the limits or a long run of points on one side of the average), this is treated as a signal that the system may have changed."
    },
    {
      keywords: [
        "what is a run chart",
        "what is run chart",
        "run chart",
        ["what", "run chart"]
      ],
      answer:
        "A run chart is a simple time-series chart that shows your data in order with a median line. " +
        "It uses basic run and trend rules (such as long runs of points on one side of the median or steady upward or downward trends) to highlight possible special-cause variation even without formal control limits."
    }
  ];

  const conceptualHit = matchFaq(conceptualFaq, q);
  if (conceptualHit) return conceptualHit;

  // ----- 1. General SPC FAQs (can be answered without your specific chart) -----
  const generalFaq = [
    {
      keywords: [
        "moving range", "mr chart", "m-r chart",
        "use the moving range", "interpret the moving range",
        ["moving", "range"]
      ],
      answer:
        "A moving range (MR) chart shows how much each value changes from one point to the next. " +
        "On an XmR chart, the X chart shows the individual values over time and the MR chart shows the absolute difference between consecutive values. " +
        "If the moving ranges are mostly small and within their limits, the short-term variation looks stable. Large spikes in the moving range can indicate a one-off shock or a change in how the process behaves."
    },
    {
      keywords: ["xmr", "xm r", "i-mr", "individuals chart", "individual chart"],
      answer:
        "An XmR chart (also called an Individuals and Moving Range chart, or I-MR) is used when you have one value per time period, such as daily admissions, length of stay per day, or time for a single patient. " +
        "The X chart shows the individual values with a centre line and control limits. The MR chart shows the size of the step between each pair of consecutive points. " +
        "The average moving range is used to estimate the underlying variation (sigma), which then gives the control limits on the X chart."
    },
    {
      keywords: ["control limit", "control limits", "ucl", "lcl"],
      answer:
        "Control limits show the range of values you would expect to see from a stable process just due to routine variation. " +
        "They are not targets and they are not hard performance thresholds. Points outside the limits or unusual patterns inside the limits suggest special-cause variation that may be worth investigating."
    },
    {
      keywords: ["sigma", "standard deviation", "spread of the data", "variation"],
      answer:
        "In SPC, sigma is an estimate of the usual spread of the process. On an XmR chart, sigma is estimated from the average moving range between consecutive points. " +
        "Control limits are typically placed at plus or minus three sigma from the mean. A larger sigma means a wider spread of routine variation."
    },
    {
      keywords: ["common cause", "special cause"],
      answer:
        "Common-cause variation is the natural background noise of a stable system. Special-cause variation is a signal that the system may have changed, for example due to a new policy, a change in demand, or a data issue. " +
        "SPC helps you distinguish common-cause from special-cause variation so that you can avoid over-reacting to noise while still spotting real changes."
    },
    {
      keywords: ["run rule", "run rules", "spc rule", "spc rules", "signal", "signals"],
      answer:
        "SPC rules are simple patterns that are unlikely to occur if the process is stable. Examples include a point outside the control limits, a long run of points on one side of the mean, or a long trend of points steadily increasing or decreasing. " +
        "When one of these patterns appears, it is treated as a potential special-cause signal that may be worth investigating."
    },
    {
      keywords: ["capability", "capable process", "process capability"],
      answer:
        "Capability in this context is about the chance that future points will meet a chosen target, assuming the process stays as it is now. " +
        "If the process is stable, we can estimate the mean and sigma and then work out the percentage of future points likely to fall above or below a target threshold."
    },
    {
      keywords: ["baseline", "phase", "segment", "split the chart"],
      answer:
        "Splitting an SPC chart into phases (baselines) lets you compare the process before and after a known change, such as a new pathway or intervention. " +
        "Each segment gets its own mean and control limits so you can see whether the system has shifted, rather than averaging everything together."
    }
  ];

  const generalHit = matchFaq(generalFaq, q);
  if (generalHit) return generalHit;

  // ----- 2. Chart-specific interpretation (XmR only) -----
  const chartType = (typeof getSelectedChartType === "function")
    ? getSelectedChartType()
    : "xmr";

  if (chartType !== "xmr") {
    return (
      "I can answer general SPC questions for any chart, but the automatic detailed interpretation currently applies only to XmR charts. " +
      "Please switch to an XmR chart if you want automated interpretation of stability, signals, limits or capability."
    );
  }

  if (!lastXmRAnalysis) {
    return (
      "I can only interpret your chart once an XmR chart has been generated. " +
      "Please create an XmR chart first, then ask me about stability, signals, control limits, target performance or capability."
    );
  }

  // ----- Special: “My chart” standard questions -----
  const isMyChartQ =
    q.includes("what is my chart telling") ||
    q.includes("what's my chart telling") ||
    q.includes("what is this chart telling") ||
    q.includes("what decision should i make") ||
    q.includes("what should i do") ||
    (q.includes("decision") && q.includes("make")) ||
    q.includes("what about my target") ||
    (q.includes("my target") && q.includes("what about"));

  if (isMyChartQ) {
    const a = lastXmRAnalysis;

    // If we somehow got here without analysis
    if (!a) {
      return "Please generate an XmR chart first, then ask one of the “My chart” questions.";
    }

    const signals = Array.isArray(a.signals) ? a.signals : [];
    const stable = !!a.isStable;

    // Helpful phrasing for signals list
    const signalsText =
      signals.length === 0
        ? "No special-cause signals detected."
        : `Signals detected: ${signals.join("; ")}.`;

    // Target summary (if present)
    let targetText = "No target is set on this chart.";
    if (a.target != null && a.direction) {
      const dirText = a.direction === "above" ? "at or above" : "at or below";
      targetText = `Target is ${a.target} (${dirText} is better).`;

      if (stable && a.capability && typeof a.capability.prob === "number") {
        targetText += ` If the process stays stable, about ${(a.capability.prob * 100).toFixed(1)}% of future points are expected to meet the target.`;
      } else if (!stable) {
        targetText += " Because special-cause signals are present, any capability estimate is unreliable until the process is stable.";
      }
    }

    // 1) “What is my chart telling me?”
    if (q.includes("what is my chart telling") || q.includes("what is this chart telling") || q.includes("what's my chart telling")) {
      const meanText = (typeof a.mean === "number") ? a.mean.toFixed(2) : "n/a";
      const uclText  = (typeof a.ucl === "number") ? a.ucl.toFixed(2) : "n/a";
      const lclText  = (typeof a.lcl === "number") ? a.lcl.toFixed(2) : "n/a";

      return (
        `Your chart summary: mean ≈ ${meanText}, limits ≈ [${lclText}, ${uclText}]. ` +
        (stable
          ? "The process looks stable (common-cause variation). "
          : "The process does not look stable (special-cause variation). ") +
        signalsText + " " +
        targetText
      );
    }

    // 2) “What decision should I make?”
    if (q.includes("what decision should i make") || q.includes("what should i do") || (q.includes("decision") && q.includes("make"))) {
      if (!stable) {
        return (
          "Decision guidance: don’t react to individual points as if they are “performance”. " +
          "Because special-cause signals are present, treat this as evidence the system may have changed. " +
          "Investigate the timing of the signals (what changed in the process/data), confirm the change is real, and then re-baseline (use a split) once the new system is established. " +
          targetText
        );
      }

      // Stable process
      if (a.target == null) {
        return (
          "Decision guidance: the process looks stable, so most up-and-down movement is routine variation. " +
          "If performance is not good enough, the decision is to change the system (not chase individual points), then use the chart to see whether a real shift occurs. " +
          "If performance is acceptable, the decision is to hold the system steady and continue monitoring."
        );
      }

      // Stable + target set
      if (a.capability && typeof a.capability.prob === "number") {
        const pct = (a.capability.prob * 100);
        if (pct >= 90) {
          return (
            `Decision guidance: the process is stable and is very likely to meet the target (~${pct.toFixed(1)}%). ` +
            "Hold the gains, standardise the current approach, and keep monitoring for any new special-cause signals."
          );
        }
        if (pct >= 50) {
          return (
            `Decision guidance: the process is stable but only sometimes meets the target (~${pct.toFixed(1)}%). ` +
            "If the target matters, you’ll need a system change to shift the mean and/or reduce variation. " +
            "Use improvement cycles and watch for a sustained shift before re-baselining."
          );
        }
        return (
          `Decision guidance: the process is stable but unlikely to meet the target (~${pct.toFixed(1)}%). ` +
          "A system redesign is needed (shift the mean and/or reduce variation). Consider stratifying data, reviewing drivers of variation, and testing changes."
        );
      }

      return (
        "Decision guidance: the process looks stable. With a target set, the key question is whether the mean is on the right side of the target and whether variation frequently crosses it. " +
        "If it does, you’ll likely need a system change to make target achievement more reliable."
      );
    }

    // 3) “What about my target?”
    if (q.includes("what about my target") || (q.includes("my target") && q.includes("what about"))) {
      return targetText;
    }
  }



  const a = lastXmRAnalysis;
  const lines = [];

  // ----- 2a. Stability / signals -----
  if (
    q.includes("stable") || q.includes("stability") ||
    q.includes("in control") || q.includes("out of control") ||
    q.includes("special cause") || q.includes("signal") ||
    q.includes("run rule") || q.includes("rule broken") ||
    q.includes("any signals")
  ) {
    if (a.isStable) {
      lines.push(
        "This segment of the XmR chart appears stable: no SPC rules are triggered and the points fluctuate randomly around the mean within the control limits."
      );
    } else if (Array.isArray(a.signals) && a.signals.length > 0) {
      const count = a.signals.length;
      const labels = a.signals.map(s => s.description || s.type || "signal").join("; ");
      lines.push(
        `This XmR chart shows evidence of special-cause variation. I can see ${count} signal${count > 1 ? "s" : ""}: ${labels}. ` +
        "These patterns are unlikely to arise from common-cause variation alone and suggest that the system may have changed."
      );
    } else {
      lines.push(
        "The chart does not look completely stable, but no specific SPC signals have been recorded. Check for obvious shifts, trends or outlying points."
      );
    }
  }

  // ----- 2b. Mean and control limits -----
  if (
    q.includes("mean") || q.includes("average") ||
    q.includes("ucl") || q.includes("lcl") ||
    q.includes("control limit") || q.includes("limits")
  ) {
    if (typeof a.mean === "number" && typeof a.sigma === "number") {
      const meanText = a.mean.toFixed(2);
      const uclText = (a.ucl != null ? a.ucl.toFixed(2) : "not calculated");
      const lclText = (a.lcl != null ? a.lcl.toFixed(2) : "not calculated");
      lines.push(
        `The current segment has an estimated mean of ${meanText}. The upper control limit (UCL) is ${uclText} and the lower control limit (LCL) is ${lclText}. ` +
        "These are based on the average moving range and represent the range you would expect from common-cause variation in this period."
      );
    } else {
      lines.push(
        "Mean and control limits could not be calculated for this chart. Check that there are enough data points and that the values are numeric."
      );
    }
  }

  // ----- 2c. Short-term variation / sigma -----
  if (
    q.includes("sigma") || q.includes("variation") ||
    q.includes("spread") || q.includes("variability")
  ) {
    if (typeof a.sigma === "number" && typeof a.avgMR === "number") {
      lines.push(
        `The estimated sigma (spread) for this segment is approximately ${a.sigma.toFixed(2)}, based on an average moving range of ${a.avgMR.toFixed(2)}. ` +
        "This captures the usual short-term variation between consecutive points and is used to set the control limits."
      );
    } else {
      lines.push(
        "An estimate of sigma could not be calculated. This usually happens if there are too few points or no variation in the data."
      );
    }
  }

  // ----- 2d. Target / direction / performance relative to target -----
  if (
    q.includes("target") || q.includes("goal") ||
    q.includes("above target") || q.includes("below target") ||
    q.includes("better") || q.includes("worse") ||
    q.includes("improve") || q.includes("improvement")
  ) {
    if (a.target == null || !a.direction) {
      lines.push(
        "A target has not been set for this chart, or the direction of improvement (above or below the target) is not defined. " +
        "Set a target value and specify whether higher or lower is better to get a clearer view of performance."
      );
    } else if (!a.isStable) {
      lines.push(
        "Because the process is not yet stable, performance against the target may change unpredictably. " +
        "Stabilise the process first, then reassess how reliably the target is being met."
      );
    } else if (a.capability && typeof a.capability.prob === "number") {
      const prob = (a.capability.prob * 100).toFixed(1);
      const dirText = a.direction === "above" ? "at or above" : "at or below";
      lines.push(
        `Given the current stable process and a target of ${a.target}, about ${prob}% of future points are expected to fall ${dirText} the target. ` +
        "This assumes that the underlying process does not change."
      );
    } else {
      lines.push(
        "A formal capability estimate against the target could not be calculated, but you can still use the chart to see whether the mean is comfortably on the desired side of the target and how often points cross it."
      );
    }
  }

  // ----- 2e. Splits / baselines (using the global splits array) -----
  if (
    q.includes("split") || q.includes("baseline") ||
    q.includes("phase") || q.includes("segment")
  ) {
    if (!Array.isArray(splits) || splits.length === 0) {
      lines.push(
        "No splits have been added, so the whole series is treated as one baseline. " +
        "If a known system change occurred, you can add a split so that before-and-after periods each get their own mean and control limits."
      );
    } else {
      lines.push(
        `You have added ${splits.length} split${splits.length > 1 ? "s" : ""} to this chart. ` +
        "Each split marks a point where a new baseline begins with its own mean and limits, allowing you to compare periods before and after key changes."
      );
    }
  }

  // ----- 2f. Moving range chart questions (chart-specific) -----
  if (
    q.includes("moving range") || q.includes("mr chart") || q.includes("m-r chart") ||
    q.includes("mr line") || q.includes("mr panel")
  ) {
    lines.push(
      "The moving range (MR) chart under the main XmR chart shows the size of the jump between one point and the next. " +
      "Large spikes in the MR chart indicate abrupt changes between consecutive observations, while a stable band of small ranges suggests consistent short-term behaviour."
    );
  }

  // ----- 3. Fallback if nothing matched in chart-specific logic -----
  if (lines.length === 0) {
    return (
      "I could not match that question to a specific SPC topic. " +
      "Try asking about stability, signals, control limits, sigma (variation), target performance, capability, moving range, or splits/baselines."
    );
  }

  // ----- 4. Final reminder -----
  lines.push(
    "Always interpret SPC charts alongside clinical or operational context, rather than in isolation. " +
    "Use the signals as prompts for discussion, not as automatic proof that a change has worked."
  );

  return lines.join(" ");
}

function renderHelperState() {
  if (!spcHelperIntro) return;

  const hasChart = !!lastXmRAnalysis;

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

  // 2) General chips (always available)
  const generalQs = [
    "What is an SPC chart?",
    "What is a run chart?",
    "What is an XmR chart?",
    "What is common cause vs special cause variation?",
    "How do control limits work?"
  ];

  if (spcHelperChipsGeneral) {
    spcHelperChipsGeneral.innerHTML = generalQs
      .map(q => `<button type="button" class="spc-chip" data-q="${escapeHtml(q)}">${escapeHtml(q)}</button>`)
      .join("");
    spcHelperChipsGeneral.classList.remove("is-disabled");
  }

  // 3) My chart chips (available only when a chart exists)
  const chartQs = [
    "What is my chart telling me?",
    "What decision should I make?",
    "What about my target?"
  ];

  if (spcHelperChipsChart) {
    spcHelperChipsChart.innerHTML = chartQs
      .map(q => `<button type="button" class="spc-chip" data-q="${escapeHtml(q)}">${escapeHtml(q)}</button>`)
      .join("");

    // Optional: disable interaction until a chart exists
    if (!hasChart) spcHelperChipsChart.classList.add("is-disabled");
    else spcHelperChipsChart.classList.remove("is-disabled");
  }
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

function showHelperAnswer(questionText) {
  if (!spcHelperOutput) return;

  const q = (questionText ?? aiQuestionInput?.value ?? "").trim();
  if (!q) {
    spcHelperOutput.innerHTML = `<p>${escapeHtml("Type a question (or click a suggestion) to get started.")}</p>`;
    return;
  }

  const ans = answerSpcQuestion(q);
  spcHelperOutput.innerHTML = `<p>${escapeHtml(ans)}</p>`;
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
  if (!help) return;

  const isHidden = help.style.display === "none" || help.style.display === "";

  if (isHidden) {
    help.style.display = "block";
    help.scrollIntoView({ behavior: "smooth" });
  } else {
    help.style.display = "none";
  }
}

let spcHelperHasBeenOpened = false;

function toggleSpcHelper() {
  const panel = document.getElementById("spcHelperPanel");
  if (!panel) return;

  const isVisible = panel.classList.toggle("visible");

  // Populate chips / intro once, when the helper is first opened
  if (isVisible && !spcHelperHasBeenOpened) {
    if (typeof renderHelperState === "function") renderHelperState();
    spcHelperHasBeenOpened = true;
  }
}

const spcHelperCloseBtn = document.getElementById("spcHelperCloseBtn");

if (spcHelperCloseBtn) {
  spcHelperCloseBtn.addEventListener("click", () => {
    if (spcHelperPanel) {
      spcHelperPanel.classList.remove("visible");
    }
  });
}


const resetButton = document.getElementById("resetButton");

if (resetButton) {
  resetButton.addEventListener("click", resetAll);
}

// Allow Escape key to close the SPC helper
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") {
    if (spcHelperPanel && spcHelperPanel.classList.contains("visible")) {
      spcHelperPanel.classList.remove("visible");
    }
  }
});


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

renderHelperState();

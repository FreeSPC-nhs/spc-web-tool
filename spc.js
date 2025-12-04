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
    alert("The tool has been reset. You can upload or paste new data.");
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

function updateXmRSummary(result, totalPoints) {
  if (!summaryDiv) return;

  const n = totalPoints;
  const { mean, ucl, lcl, sigma, avgMR, baselineCountUsed } = result;
  const nBeyond = result.points.filter(p => p.beyondLimits).length;
  const values = result.points.map(p => p.y);

  // additional rules for NHS-style interpretation
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

  const target = getTargetValue();
  const direction = targetDirectionInput ? targetDirectionInput.value : "above";

  let capability = null;
  if (target !== null && sigma > 0) {
    capability = computeTargetCapability(mean, sigma, target, direction);
  }

  let html = `<h3>Summary (XmR chart)</h3>`;
  html += `<ul>`;
  html += `<li>Number of points: <strong>${n}</strong></li>`;
  if (baselineCountUsed && baselineCountUsed < n) {
    html += `<li>Baseline: first <strong>${baselineCountUsed}</strong> points used to calculate mean and limits.</li>`;
  } else {
    html += `<li>Baseline: all points used to calculate mean and limits.</li>`;
  }
  html += `<li>Mean: <strong>${mean.toFixed(3)}</strong></li>`;
  html += `<li>Estimated σ (from MR): <strong>${sigma.toFixed(3)}</strong> (avg MR = ${avgMR.toFixed(3)})</li>`;
  html += `<li>Control limits: <strong>LCL = ${lcl.toFixed(3)}</strong>, <strong>UCL = ${ucl.toFixed(3)}</strong></li>`;

if (target !== null) {
  html += `<li>Target: <strong>${target}</strong> (${direction === "above" ? "at or above is better" : "at or below is better"})</li>`;
  if (capability && signals.length === 0) {
    html += `<li>Estimated process capability (assuming a stable process and approximate normality): about <strong>${(capability.prob * 100).toFixed(1)}%</strong> of future points are expected to meet the target.</li>`;
  } else if (capability && signals.length > 0) {
    html += `<li>Capability: a target has been set, but special-cause signals are present. Any capability estimate would be unreliable until the process is stable.</li>`;
  } else {
    html += `<li>Capability: cannot be estimated (insufficient variation to estimate σ).</li>`;
  }
} else {
  html += `<li>Target: not specified – capability not calculated.</li>`;
}

  if (signals.length === 0) {
    html += `<li><strong>Special cause:</strong> No rule breaches detected (points within limits, no long runs or clear trend). Pattern is consistent with common-cause variation, but always interpret in context.</li>`;
  } else {
    html += `<li><strong>Special cause:</strong> Signals suggesting special-cause variation based on: ${signals.join("; ")}.</li>`;
  }

  html += `</ul>`;

  summaryDiv.innerHTML = html;

if (capabilityDiv) {
  if (target !== null && capability && signals.length === 0) {
    capabilityDiv.innerHTML = `
      <div style="
        display:inline-block;
        padding:0.6rem 1.2rem;
        background:#fff59d;
        border:1px solid #ccc;
        border-radius:0.25rem;
      ">
        <div style="font-weight:bold; text-align:center;">PROCESS CAPABILITY</div>
        <div style="font-size:1.4rem; font-weight:bold; text-align:center; margin-top:0.2rem;">
          ${(capability.prob * 100).toFixed(1)}%
        </div>
        <div style="font-size:0.8rem; margin-top:0.2rem;">
          (Estimated probability of meeting the target, assuming no special-cause variation.)
        </div>
      </div>
    `;
  } else if (target !== null && signals.length > 0) {
    capabilityDiv.innerHTML = `
      <div style="
        display:inline-block;
        padding:0.6rem 1.2rem;
        background:#ffe0b2;
        border:1px solid #ccc;
        border-radius:0.25rem;
        max-width:32rem;
      ">
        <strong>Process not stable:</strong> special-cause signals are present.
        Focus on understanding and addressing these causes before relying on capability estimates.
      </div>
    `;
  } else {
    capabilityDiv.innerHTML = "";
  }
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

      const d = new Date(xRaw);
      const y = Number(yRaw);

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

      const y = Number(yRaw);
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

function computeLastSegment(points, baselineCount) {
  const n = points.length;
  if (n < 2) return null;

  let effectiveSplits = Array.isArray(splits) ? splits.slice() : [];
  effectiveSplits = effectiveSplits
    .filter(i => Number.isInteger(i) && i >= 0 && i < n - 1)
    .sort((a, b) => a - b);

  // Determine start index of the last segment
  const lastSplitIndex = effectiveSplits.length > 0
    ? effectiveSplits[effectiveSplits.length - 1] + 1
    : 0;

  const segPoints = points.slice(lastSplitIndex);

  // Baseline for the last segment: use all its points
  return {
    result: computeXmR(segPoints, null),
    count: segPoints.length
  };
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

  // We'll still compute a "global" XmR for the summary panel
  const globalResult = computeXmR(points, baselineCount);

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
    tension: 0,
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

  // If splits exist: use only the last segment for summary & capability
let summaryResult;
let summaryCount;

if (splits.length > 0) {
  const seg = computeLastSegment(points, baselineCount);
  summaryResult = seg.result;
  summaryCount = seg.count;
} else {
  summaryResult = globalResult;
  summaryCount = points.length;
}

updateXmRSummary(summaryResult, summaryCount);
drawMRChart(summaryResult, labels);
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
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

const generateButton    = document.getElementById("generateButton");
const errorMessage      = document.getElementById("errorMessage");
const chartCanvas       = document.getElementById("spcChart");
const summaryDiv        = document.getElementById("summary");
const downloadBtn       = document.getElementById("downloadPngButton");

const mrPanel           = document.getElementById("mrPanel");
const mrChartCanvas     = document.getElementById("mrChartCanvas");

// ---- CSV upload & column selection ----

fileInput.addEventListener("change", () => {
  const file = fileInput.files[0];
  if (!file) return;

  errorMessage.textContent = "";
  if (summaryDiv) summaryDiv.innerHTML = "";
  if (capabilityDiv) capabilityDiv.innerHTML = "";

  Papa.parse(file, {
    header: true,
    dynamicTyping: true,
    skipEmptyLines: true,
    complete: (results) => {
      const rows = results.data;
      if (!rows || rows.length === 0) {
        errorMessage.textContent = "No rows found in this CSV.";
        return;
      }

      rawRows = rows;
      const firstRow = rows[0];
      const columns = Object.keys(firstRow);

      // populate the dropdowns
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
  const mrLabels = [];
  for (let i = 1; i < pts.length; i++) {
    mrValues.push(Math.abs(pts[i].y - pts[i - 1].y));
    mrLabels.push(pts[i].x.toISOString().slice(0, 10));
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
    mrValues,
    mrLabels
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

function getTargetValue() {
  if (!targetInput) return null;
  const v = targetInput.value.trim();
  if (v === "") return null;
  const num = Number(v);
  return isFinite(num) ? num : null;
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

  const parsedPoints = rawRows
    .map((row) => {
      const dateRaw  = row[dateCol];
      const valueRaw = row[valueCol];

      const d = new Date(dateRaw);
      const y = Number(valueRaw);

      if (!isFinite(d.getTime()) || !isFinite(y)) return null;
      return { x: d, y };
    })
    .filter(p => p !== null);

  if (parsedPoints.length < 5) {
    errorMessage.textContent = "Not enough valid data points after parsing. Check your column choices.";
    return;
  }

  // sort by date once here
  const points = [...parsedPoints].sort((a, b) => a.x - b.x);

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
    drawRunChart(points, baselineCount);
  } else {
    drawXmRChart(points, baselineCount);
  }
});

// ---- Chart drawing ----

function drawRunChart(points, baselineCount) {
  const n = points.length;

  let baselineCountUsed;
  if (baselineCount && baselineCount >= 2) {
    baselineCountUsed = Math.min(baselineCount, n);
  } else {
    baselineCountUsed = n;
  }

  const baselineValues = points.slice(0, baselineCountUsed).map(p => p.y);

  const labels = points.map(p => p.x.toISOString().slice(0, 10));
  const values = points.map(p => p.y);
  const median = computeMedian(baselineValues);

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

function drawXmRChart(points, baselineCount) {
  const result = computeXmR(points, baselineCount);
  const pts = result.points;

  const labels = pts.map(p => p.x.toISOString().slice(0, 10));
  const values = pts.map(p => p.y);
  const pointColours = pts.map(p => (p.beyondLimits ? "#d73027" : "#003f87")); // red for breaches, dark blue otherwise

  const { mean, sigma, ucl, lcl } = result;

  const { title, xLabel, yLabel } = getChartLabels(
    "I-MR Chart",
    "Date",
    "Value"
  );

  const target = getTargetValue();

  // 1σ and 2σ lines if sigma > 0
  const oneSigmaUp   = sigma > 0 ? mean + 1 * sigma : null;
  const oneSigmaDown = sigma > 0 ? mean - 1 * sigma : null;
  const twoSigmaUp   = sigma > 0 ? mean + 2 * sigma : null;
  const twoSigmaDown = sigma > 0 ? mean - 2 * sigma : null;

  const sigmaLineColor = "rgba(128,128,128,0.35)"; // faint grey

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
      // MEAN
      label: "Mean",
      data: values.map(() => mean),
      borderDash: [6, 4],
      borderWidth: 2,
      borderColor: "#e41a1c", // red
      pointRadius: 0,
      pointHoverRadius: 0,
      fill: false
    },
    {
      // UCL
      label: "UCL (3σ)",
      data: values.map(() => ucl),
      borderDash: [4, 4],
      borderWidth: 2,
      borderColor: "#1a9850", // green
      pointRadius: 0,
      pointHoverRadius: 0,
      fill: false
    },
    {
      // LCL
      label: "LCL (3σ)",
      data: values.map(() => lcl),
      borderDash: [4, 4],
      borderWidth: 2,
      borderColor: "#1a9850", // green
      pointRadius: 0,
      pointHoverRadius: 0,
      fill: false
    }
  ];

  if (sigma > 0) {
    datasets.push(
      {
        label: "+1σ",
        data: values.map(() => oneSigmaUp),
        borderDash: [2, 2],
        borderWidth: 1,
        borderColor: sigmaLineColor,
        pointRadius: 0,
        pointHoverRadius: 0,
        fill: false
      },
      {
        label: "-1σ",
        data: values.map(() => oneSigmaDown),
        borderDash: [2, 2],
        borderWidth: 1,
        borderColor: sigmaLineColor,
        pointRadius: 0,
        pointHoverRadius: 0,
        fill: false
      },
      {
        label: "+2σ",
        data: values.map(() => twoSigmaUp),
        borderDash: [2, 2],
        borderWidth: 1,
        borderColor: sigmaLineColor,
        pointRadius: 0,
        pointHoverRadius: 0,
        fill: false
      },
      {
        label: "-2σ",
        data: values.map(() => twoSigmaDown),
        borderDash: [2, 2],
        borderWidth: 1,
        borderColor: sigmaLineColor,
        pointRadius: 0,
        pointHoverRadius: 0,
        fill: false
      }
    );
  }

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

  updateXmRSummary(result, points.length);
  drawMRChart(result);
}

// MR chart: average MR as centre, UCL = 3.268 * avgMR, LCL = 0
function drawMRChart(result) {
  if (!mrPanel || !mrChartCanvas) return;

  const mrValues = result.mrValues;
  const mrLabels = result.mrLabels;

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

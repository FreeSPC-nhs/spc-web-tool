// Simple SPC logic + wiring for Run chart and XmR chart
// Features:
//  - CSV upload + column selection
//  - Run chart with run rule (>=8 points on one side of median)
//  - XmR chart with mean, UCL, LCL
//  - Baseline: use first N points for centre line & limits (optional)
//  - Summary panel
//  - Download chart as PNG

let rawRows = [];
let currentChart = null;

const fileInput       = document.getElementById("fileInput");
const columnSelectors = document.getElementById("columnSelectors");
const dateSelect      = document.getElementById("dateColumn");
const valueSelect     = document.getElementById("valueColumn");
const baselineInput   = document.getElementById("baselinePoints");
const generateButton  = document.getElementById("generateButton");
const errorMessage    = document.getElementById("errorMessage");
const chartCanvas     = document.getElementById("spcChart");
const summaryDiv      = document.getElementById("summary");
const downloadBtn     = document.getElementById("downloadPngButton");
const chartTitleInput   = document.getElementById("chartTitle");
const xAxisLabelInput   = document.getElementById("xAxisLabel");
const yAxisLabelInput   = document.getElementById("yAxisLabel");

// ---- CSV upload & column selection ----

fileInput.addEventListener("change", () => {
  const file = fileInput.files[0];
  if (!file) return;

  errorMessage.textContent = "";
  if (summaryDiv) summaryDiv.innerHTML = "";

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


function computeMedian(values) {
  const sorted = [...values].sort((a, b) => a - b);
  const n = sorted.length;
  if (n === 0) return NaN;
  if (n % 2 === 1) return sorted[(n - 1) / 2];
  return (sorted[n / 2 - 1] + sorted[n / 2]) / 2;
}

/**
 * Detect runs of >= runLength points on the same side of the centre line.
 * values: array of numbers
 * centre: median/mean
 * returns: array of booleans, true if that point is part of a violating run
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
 * points: array of {x: Date, y: number} sorted by x
 * baselineCount: optional number of points to use for baseline stats (>=2), else use all
 * returns: { points: [{x,y,beyondLimits}], mean, ucl, lcl, sigma, avgMR, baselineCountUsed }
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

  const mrs = [];
  for (let i = 1; i < baseline.length; i++) {
    mrs.push(Math.abs(baseline[i].y - baseline[i - 1].y));
  }
  const avgMR =
    mrs.length > 0
      ? mrs.reduce((sum, v) => sum + v, 0) / mrs.length
      : 0;

  const sigma = avgMR === 0 ? 0 : avgMR / 1.128; // d2 for n=2

  const ucl = mean + 3 * sigma;
  const lcl = mean - 3 * sigma;

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
    baselineCountUsed
  };
}

// ---- Summary helpers ----

function updateRunSummary(points, median, runFlags, baselineCountUsed) {
  if (!summaryDiv) return;

  const n = points.length;
  const nRunPoints = runFlags.filter(Boolean).length;
  const hasRunViolation = nRunPoints > 0;

  let html = `<h3>Summary (Run chart)</h3>`;
  html += `<ul>`;
  html += `<li>Number of points: <strong>${n}</strong></li>`;
  if (baselineCountUsed && baselineCountUsed < n) {
    html += `<li>Baseline: first <strong>${baselineCountUsed}</strong> points used to calculate median.</li>`;
  } else {
    html += `<li>Baseline: all points used to calculate median.</li>`;
  }
  html += `<li>Median: <strong>${median.toFixed(3)}</strong></li>`;
  if (hasRunViolation) {
    html += `<li><strong>Special cause:</strong> Run rule triggered (≥8 consecutive points on one side of median). Points in long runs are highlighted in orange.</li>`;
  } else {
    html += `<li><strong>Special cause:</strong> No long runs (≥8 points) on one side of the median detected.</li>`;
  }
  html += `</ul>`;

  summaryDiv.innerHTML = html;
}

function updateXmRSummary(result, totalPoints) {
  if (!summaryDiv) return;

  const n = totalPoints;
  const { mean, ucl, lcl, sigma, avgMR, baselineCountUsed } = result;
  const nBeyond = result.points.filter(p => p.beyondLimits).length;
  const hasSpecialCause = nBeyond > 0;

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
  if (hasSpecialCause) {
    html += `<li><strong>Special cause:</strong> ${nBeyond} point(s) beyond control limits (shown in red).</li>`;
  } else {
    html += `<li><strong>Special cause:</strong> No points beyond control limits.</li>`;
  }
  html += `</ul>`;

  summaryDiv.innerHTML = html;
}

// ---- Generate chart button ----

generateButton.addEventListener("click", () => {
  errorMessage.textContent = "";
  if (summaryDiv) summaryDiv.innerHTML = "";

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

  if (currentChart) {
    currentChart.destroy();
    currentChart = null;
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

  currentChart = new Chart(chartCanvas, {
    type: "line",
    data: {
      labels: labels,
      datasets: [
        {
          label: "Value",
          data: values,
          pointRadius: 4,
          pointBackgroundColor: pointColours,
          borderColor: "#003f87", // dark blue
          borderWidth: 2,
          fill: false
        },
        {
          label: "Median",
          data: values.map(() => median),
          borderDash: [6, 4],
          borderWidth: 2,
          borderColor: "#e41a1c", // red-ish
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
          text: title,
          font: {
            size: 16,
            weight: "bold"
          }
        },
        legend: {
          display: true
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

  // 1σ and 2σ lines if sigma > 0
  const oneSigmaUp   = sigma > 0 ? mean + 1 * sigma : null;
  const oneSigmaDown = sigma > 0 ? mean - 1 * sigma : null;
  const twoSigmaUp   = sigma > 0 ? mean + 2 * sigma : null;
  const twoSigmaDown = sigma > 0 ? mean - 2 * sigma : null;

  const sigmaLineColor = "rgba(128,128,128,0.35)"; // faint grey

  const datasets = [
    {
      label: "Value",
      data: values,
      pointRadius: 4,
      pointBackgroundColor: pointColours,
      borderColor: "#003f87", // dark blue
      borderWidth: 2,
      fill: false
    },
    {
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

  // Add 1σ & 2σ lines on both sides if sigma is valid
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
          display: true
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

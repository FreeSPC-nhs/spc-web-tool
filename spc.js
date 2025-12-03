// Simple SPC logic + wiring for Run chart and XmR chart

let rawRows = [];
let currentChart = null;

const fileInput       = document.getElementById("fileInput");
const columnSelectors = document.getElementById("columnSelectors");
const dateSelect      = document.getElementById("dateColumn");
const valueSelect     = document.getElementById("valueColumn");
const generateButton  = document.getElementById("generateButton");
const errorMessage    = document.getElementById("errorMessage");
const chartCanvas     = document.getElementById("spcChart");

// ---- CSV upload & column selection ----

fileInput.addEventListener("change", () => {
  const file = fileInput.files[0];
  if (!file) return;

  errorMessage.textContent = "";

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
 * points: array of {x: Date, y: number}
 * returns: { points: [{x,y,beyondLimits}], mean, ucl, lcl }
 */
function computeXmR(points) {
  // sort by date
  points.sort((a, b) => a.x - b.x);

  const mean = points.reduce((sum, p) => sum + p.y, 0) / points.length;

  const mrs = [];
  for (let i = 1; i < points.length; i++) {
    mrs.push(Math.abs(points[i].y - points[i - 1].y));
  }
  const avgMR = mrs.reduce((sum, v) => sum + v, 0) / mrs.length;
  const sigma = avgMR / 1.128; // d2 for n=2

  const ucl = mean + 3 * sigma;
  const lcl = mean - 3 * sigma;

  const flagged = points.map(p => ({
    ...p,
    beyondLimits: p.y > ucl || p.y < lcl
  }));

  return { points: flagged, mean, ucl, lcl };
}

// ---- Generate chart button ----

generateButton.addEventListener("click", () => {
  errorMessage.textContent = "";

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

  const chartType = getSelectedChartType();

  if (currentChart) {
    currentChart.destroy();
    currentChart = null;
  }

  if (chartType === "run") {
    drawRunChart(parsedPoints);
  } else {
    drawXmRChart(parsedPoints);
  }
});

// ---- Chart drawing ----

function drawRunChart(points) {
  points.sort((a, b) => a.x - b.x);

  const labels = points.map(p => p.x.toISOString().slice(0, 10));
  const values = points.map(p => p.y);
  const median = computeMedian(values);

  currentChart = new Chart(chartCanvas, {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "Value",
          data: values,
          pointRadius: 4,
          borderWidth: 2,
          fill: false
        },
        {
          label: "Median",
          data: values.map(() => median),
          borderDash: [6, 4],
          borderWidth: 1,
          fill: false
        }
      ]
    },
    options: {
      responsive: true,
      plugins: {
        title: {
          display: true,
          text: "Run Chart"
        }
      }
    }
  });
}

function drawXmRChart(points) {
  const { points: pts, mean, ucl, lcl } = computeXmR(points);

  pts.sort((a, b) => a.x - b.x);

  const labels       = pts.map(p => p.x.toISOString().slice(0, 10));
  const values       = pts.map(p => p.y);
  const pointColours = pts.map(p => (p.beyondLimits ? "red" : "blue"));

  currentChart = new Chart(chartCanvas, {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "Value",
          data: values,
          pointRadius: 4,
          pointBackgroundColor: pointColours,
          borderWidth: 2,
          fill: false
        },
        {
          label: "Mean",
          data: values.map(() => mean),
          borderDash: [6, 4],
          borderWidth: 1,
          fill: false
        },
        {
          label: "UCL",
          data: values.map(() => ucl),
          borderDash: [4, 4],
          borderWidth: 1,
          fill: false
        },
        {
          label: "LCL",
          data: values.map(() => lcl),
          borderDash: [4, 4],
          borderWidth: 1,
          fill: false
        }
      ]
    },
    options: {
      responsive: true,
      plugins: {
        title: {
          display: true,
          text: "XmR Chart (points beyond limits shown in red)"
        }
      }
    }
  });
}

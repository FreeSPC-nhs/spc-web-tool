// Universal Escape key handler for overlays
document.addEventListener("keydown", (e) => {
  if (e.key !== "Escape") return;

  const dataEditorOverlay = document.getElementById("dataEditorOverlay");
  if (dataEditorOverlay && dataEditorOverlay.style.display !== "none") {
    dataEditorOverlay.style.display = "none";
  }

  const sheetPickerOverlay = document.getElementById("sheetPickerOverlay");
  if (sheetPickerOverlay && sheetPickerOverlay.style.display !== "none") {
    sheetPickerOverlay.style.display = "none";
  }
});

// -----------------------------
// Help section toggle
// -----------------------------
function toggleHelpSection(forceOpen) {
  const modal = document.getElementById("helpModal");
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

(function initHelpModal() {
  const modal = document.getElementById("helpModal");
  if (!modal) return;

  // Click outside (backdrop) closes
  modal.addEventListener("click", (e) => {
    if (e.target && e.target.classList.contains("modal-backdrop")) {
      toggleHelpSection(false);
    }
  });

  // Escape closes
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && modal.classList.contains("visible")) {
      toggleHelpSection(false);
    }
  });
})();

const spcHelperCloseBtn = document.getElementById("spcHelperCloseBtn");

if (spcHelperCloseBtn) {
  spcHelperCloseBtn.addEventListener("click", () => {
    if (spcHelperPanel) {
      spcHelperPanel.classList.remove("visible");
    }
  });
}

function showDataEditorDeleteHelp() {
  if (!dataEditorDeleteHelpBtn || !dataEditorDeleteHelpPopup) return;
  dataEditorDeleteHelpPopup.classList.add("show");
  dataEditorDeleteHelpBtn.setAttribute("aria-expanded", "true");
  dataEditorDeleteHelpPopup.setAttribute("aria-hidden", "false");
}

function hideDataEditorDeleteHelp() {
  if (!dataEditorDeleteHelpBtn || !dataEditorDeleteHelpPopup) return;
  dataEditorDeleteHelpPopup.classList.remove("show");
  dataEditorDeleteHelpBtn.setAttribute("aria-expanded", "false");
  dataEditorDeleteHelpPopup.setAttribute("aria-hidden", "true");
}

function toggleDataEditorDeleteHelp() {
  if (!dataEditorDeleteHelpPopup) return;
  if (dataEditorDeleteHelpPopup.classList.contains("show")) {
    hideDataEditorDeleteHelp();
  } else {
    showDataEditorDeleteHelp();
  }
}

if (dataEditorDeleteHelpBtn && dataEditorDeleteHelpPopup) {
  dataEditorDeleteHelpBtn.addEventListener("mouseenter", showDataEditorDeleteHelp);
  dataEditorDeleteHelpBtn.addEventListener("mouseleave", hideDataEditorDeleteHelp);

  dataEditorDeleteHelpBtn.addEventListener("click", (e) => {
    e.preventDefault();
    e.stopPropagation();
    toggleDataEditorDeleteHelp();
  });

  dataEditorDeleteHelpBtn.addEventListener("focus", showDataEditorDeleteHelp);
  dataEditorDeleteHelpBtn.addEventListener("blur", hideDataEditorDeleteHelp);

  dataEditorDeleteHelpPopup.addEventListener("mouseenter", showDataEditorDeleteHelp);
  dataEditorDeleteHelpPopup.addEventListener("mouseleave", hideDataEditorDeleteHelp);

  document.addEventListener("click", (e) => {
    const clickedInsideHelp =
      dataEditorDeleteHelpBtn.contains(e.target) ||
      dataEditorDeleteHelpPopup.contains(e.target);

    if (!clickedInsideHelp) {
      hideDataEditorDeleteHelp();
    }
  });
}

function toggleChartSetupModal(forceOpen) {
  const modal = document.getElementById("chartSetupModal");
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


/*
 * Batch 36A-2 — Chart setup localisation.
 *
 * The original English modal content is the fallback.
 * Translate cards while they are rendered, before any
 * T-chart radio-button event listeners are attached.
 */

const chartSetupModalHeadingKeys = {
  "Use this when...": "when",
  "You need these columns": "columns",
  "Example layout": "example",
  "Common mistake": "mistake",
  "Common mistakes": "mistakes",
  "Choosing P or P′": "pprime",
  "Choosing U or U′": "uprime",
  "Displaying rates": "rates",
  "Example 1: one row per occurrence": "paretoRows",
  "Example 2: already summarised counts": "paretoCounts",
  "What the chart shows": "paretoShows",
  "Choose your setup": "tMode",
  "Setup guidance": "fallback"
};

const chartSetupModalExampleKeys = {
  "Week": "week",
  "Waiting time": "waitingTime",
  "Date": "date",
  "Length of stay": "lengthStay",
  "Falls": "falls",
  "Patients with harm": "patientsHarm",
  "Patients reviewed": "patientsReviewed",
  "Month": "month",
  "Infections": "infections",
  "Bed days": "bedDays",
  "Jan": "jan",
  "Feb": "feb",
  "Mar": "mar",
  "Day": "day",
  "Reading": "reading",
  "Sample_ID": "sampleId",
  "Mon": "mon",
  "Tue": "tue",
  "Event date": "eventDate",
  "Procedures between harms": "procedures",
  "Category": "category",
  "Count": "count",
  "Medication": "medication",
  "Pressure damage": "pressureDamage"
};

let lastChartSetupRenderedType = "run";

function chartSetupModalTranslate(
  key,
  englishFallback,
  values = {}
) {
  const i18n = window.SimpleSPCi18n;

  if (
    !i18n ||
    typeof i18n.translate !== "function"
  ) {
    return englishFallback;
  }

  const translated = i18n.translate(key, values);

  return translated && translated !== key
    ? translated
    : englishFallback;
}


function renderChartSetupModal(chartType) {
  const body = document.getElementById("chartSetupBody");
  const subtitle = document.getElementById("chartSetupSubtitle");
  const dontShow = document.getElementById("chartSetupDontShow");
  if (!body || !subtitle) return;

  lastChartSetupRenderedType = chartType;

  if (dontShow) {
    dontShow.checked = !shouldAutoShowChartSetupModal();
    dontShow.onchange = () => setAutoShowChartSetupModal(!dontShow.checked);
  }

  
  // Card numbering restarts for each rendering.
  // This does not affect the underlying chart type.
  let setupCardNumber = 0;

  function card(title, innerHtml) {
    setupCardNumber += 1;

    const headingName =
      chartSetupModalHeadingKeys[title];

    const displayedTitle = headingName
      ? chartSetupModalTranslate(
          `setup.heading.${headingName}`,
          title
        )
      : title;

    // Translate the subtitle once the original renderer
    // has assigned the correct chart-specific wording.
    if (setupCardNumber === 1) {
      subtitle.textContent = chartSetupModalTranslate(
        `setup.subtitle.${chartType}`,
        subtitle.textContent
      );
    }

    // Preserve the currently selected T-chart input mode.
    // These values replace placeholders only in the
    // translated T-chart radio-button card.
    const checkedDates =
      typeof tChartInputMode !== "undefined" &&
      tChartInputMode === "eventDates"
        ? "checked"
        : "";

    const checkedGaps =
      typeof tChartInputMode !== "undefined" &&
      tChartInputMode === "gaps"
        ? "checked"
        : "";

    const displayedHtml = chartSetupModalTranslate(
      `setup.${chartType}.card${setupCardNumber}`,
      innerHtml,
      { checkedDates, checkedGaps }
    );

    return `
      <div style="border:1px solid #d8dde0; border-radius:0.5rem; padding:0.75rem; margin:0.75rem 0; background:#fafcfd;">
        <div style="font-weight:700; color:#003087; margin-bottom:0.4rem;">${displayedTitle}</div>
        ${displayedHtml}
      </div>
    `;
  }


    function exampleTable(headers, rows) {

    // Translate illustrative text only.
    // Keep numerical values and example dates unchanged.
    function exampleLabel(value) {
      const original = String(value);

      const key = chartSetupModalExampleKeys[original];

      if (!key) {
        return original;
      }

      return chartSetupModalTranslate(
        `setup.example.${key}`,
        original
      );
    }

    // Translate the example table's column headings.
    const head = headers.map(h => {
      const translated = exampleLabel(h);

      return `<th style="text-align:left; border-bottom:1px solid #d8dde0; padding:0.35rem 0.5rem;">${translated}</th>`;
    }).join("");

    // Translate illustrative labels within example rows.
    // Do not change any example numbers or dates.
    const bodyRows = rows.map(row => {

      const cells = row.map(value => {
        const translated = exampleLabel(value);

        return `<td style="padding:0.35rem 0.5rem; border-bottom:1px solid #eef2f6;">${translated}</td>`;
      }).join("");

      return `<tr>${cells}</tr>`;

    }).join("");

    // Preserve the original table layout and styling.
    return `
      <div style="overflow-x:auto;">
        <table style="border-collapse:collapse; width:100%; font-size:0.9rem; margin-top:0.35rem;">
          <thead><tr>${head}</tr></thead>
          <tbody>${bodyRows}</tbody>
        </table>
      </div>
    `;
  }

  subtitle.textContent = "How to structure your data for this chart.";

  if (chartType === "run") {
    subtitle.textContent = "Run chart: simple view of values over time.";
    body.innerHTML = `
      ${card("Use this when...", `
        <p style="margin:0;">You want a simple chart of a measure over time using a <strong>median</strong>.</p>
      `)}

      ${card("You need these columns", `
        <ul style="margin:0;">
          <li><strong>Date / X-axis column</strong> → time or order</li>
          <li><strong>Value / Y-axis column</strong> → the measure you want to track</li>
        </ul>
      `)}

      ${card("Example layout", exampleTable(
        ["Week", "Waiting time"],
        [["1", "12"], ["2", "10"], ["3", "14"]]
      ))}

      ${card("Common mistake", `
        <p style="margin:0;">Do not choose a row number or ID column as the value. That can create a chart that looks valid but means nothing.</p>
      `)}
    `;
    return;
  }

  if (chartType === "xmr") {
    subtitle.textContent = "XmR chart: individual measurements over time.";
    body.innerHTML = `
      ${card("Use this when...", `
        <p style="margin:0;">You have <strong>one measurement per time point</strong> and want a mean plus control limits.</p>
      `)}

      ${card("You need these columns", `
        <ul style="margin:0;">
          <li><strong>Date / X-axis column</strong> → date, week, month or sequence</li>
          <li><strong>Value / Y-axis column</strong> → the numeric measurement</li>
        </ul>
      `)}

      ${card("Example layout", exampleTable(
        ["Date", "Length of stay"],
        [["2024-01-01", "5.2"], ["2024-01-08", "4.8"], ["2024-01-15", "6.1"]]
      ))}

      ${card("Common mistake", `
        <p style="margin:0;">Use XmR only when there is one value per time point. If each time point has several measurements, use <strong>X̄–S</strong> instead.</p>
      `)}
    `;
    return;
  }

  if (chartType === "c") {
    subtitle.textContent = "C chart: count per time period.";
    body.innerHTML = `
      ${card("Use this when...", `
        <p style="margin:0;">You are plotting a <strong>count</strong> per period and the amount of opportunity is roughly the same each time.</p>
      `)}

      ${card("You need these columns", `
        <ul style="margin:0;">
          <li><strong>Date / X-axis column</strong> → date, week, month or sequence</li>
          <li><strong>Value / Y-axis column</strong> → count of events</li>
        </ul>
      `)}

      ${card("Example layout", exampleTable(
        ["Week", "Falls"],
        [["1", "3"], ["2", "4"], ["3", "2"]]
      ))}

      ${card("Common mistake", `
        <p style="margin:0;">If the denominator changes a lot between points, use a <strong>U chart</strong> instead of a C chart.</p>
      `)}
    `;
    return;
  }

  if (chartType === "p") {
    subtitle.textContent = "P chart: proportion out of a total.";
    body.innerHTML = `
      ${card("Use this when...", `
        <p style="margin:0;">You want to track a <strong>proportion</strong>, such as 5 out of 100 or the percentage meeting a standard.</p>
      `)}

      ${card("You need these columns", `
        <ul style="margin:0;">
          <li><strong>Date / X-axis column</strong> → time or order</li>
          <li><strong>Value / Y-axis column</strong> → numerator</li>
          <li><strong>Third column</strong> → denominator</li>
        </ul>
      `)}

      ${card("Example layout", exampleTable(
        ["Week", "Patients with harm", "Patients reviewed"],
        [["1", "5", "100"], ["2", "7", "110"], ["3", "6", "95"]]
      ))}

      ${card("Common mistake", `
        <p style="margin:0;">The numerator must not be larger than the denominator.</p>
      `)}
    `;
    return;
  }

  if (chartType === "u") {
    subtitle.textContent = "U chart: rate per opportunity.";
    body.innerHTML = `
      ${card("Use this when...", `
        <p style="margin:0;">You want to track a <strong>rate</strong>, where the denominator changes from point to point.</p>
      `)}

      ${card("You need these columns", `
        <ul style="margin:0;">
          <li><strong>Date / X-axis column</strong> → time or order</li>
          <li><strong>Value / Y-axis column</strong> → count of events</li>
          <li><strong>Third column</strong> → opportunities / exposure</li>
        </ul>
      `)}

      ${card("Example layout", exampleTable(
        ["Month", "Infections", "Bed days"],
        [["Jan", "2", "1200"], ["Feb", "3", "1350"], ["Mar", "1", "980"]]
      ))}

      ${card("Common mistake", `
        <p style="margin:0;">Use U when the denominator varies. If the denominator is roughly constant, a <strong>C chart</strong> may be more appropriate.</p>
      `)}
    `;
    return;
  }


  // Laney P-prime chart setup
  if (chartType === "pprime") {
    subtitle.textContent = "Laney P′ chart: adjusted proportion control chart.";

    body.innerHTML = `
      ${card("Use this when...", `
        <p style="margin:0;">
          You are tracking a <strong>proportion</strong> using a numerator
          and denominator, but the observed variation is greater or
          smaller than a standard P chart expects.
          P′ adjusts the width of the control limits to reflect
          this extra or reduced variation.
        </p>
      `)}

      ${card("You need these columns", `
        <ul style="margin:0;">
          <li><strong>Date / X-axis column</strong> → time or order</li>
          <li><strong>Value / Y-axis column</strong> → numerator (number meeting the definition)</li>
          <li><strong>Third column</strong> → denominator (total number assessed)</li>
        </ul>
      `)}

      ${card("Example layout", exampleTable(
        ["Week", "Patients with harm", "Patients reviewed"],
        [["1", "5", "100"], ["2", "7", "110"], ["3", "6", "95"]]
      ))}

      ${card("Choosing P or P′", `
        <p style="margin:0;">
          Normally start with a <strong>P chart</strong>.
          SimpleSPC checks whether the observed variation differs
          from what a standard P chart expects and can recommend
          switching to P′.
          <br><br>
          The adjustment is called <strong>Sigma Z</strong>.
          It changes the control limits, not the plotted proportions.
          The choice to switch charts remains yours.
        </p>
      `)}

      ${card("Common mistake", `
        <p style="margin:0;">
          Enter the actual numerator and denominator counts,
          not a pre-calculated percentage. The numerator cannot
          exceed the denominator.
        </p>
      `)}
    `;
    return;
  }

  // Laney U-prime chart setup
  if (chartType === "uprime") {
    subtitle.textContent = "Laney U′ chart: adjusted rate control chart.";

    body.innerHTML = `
      ${card("Use this when...", `
        <p style="margin:0;">
          You are tracking <strong>events per opportunity</strong>,
          but the observed variation is greater or smaller than
          a standard U chart expects.
          U′ adjusts the control limits to reflect this variation.
        </p>
      `)}

      ${card("You need these columns", `
        <ul style="margin:0;">
          <li><strong>Date / X-axis column</strong> → time or order</li>
          <li><strong>Value / Y-axis column</strong> → count of events</li>
          <li><strong>Third column</strong> → number of opportunities or amount of exposure</li>
        </ul>
      `)}

      ${card("Example layout", exampleTable(
        ["Month", "Infections", "Bed days"],
        [["Jan", "2", "1200"], ["Feb", "3", "1350"], ["Mar", "1", "980"]]
      ))}

      ${card("Choosing U or U′", `
        <p style="margin:0;">
          Normally start with a <strong>U chart</strong>.
          SimpleSPC checks whether the observed variation differs
          from what a standard U chart expects and can recommend
          switching to U′.
          <br><br>
          The adjustment is called <strong>Sigma Z</strong>.
          It changes the control limits, not the underlying event rates.
          The choice to switch charts remains yours.
        </p>
      `)}

      ${card("Displaying rates", `
        <p style="margin:0;">
          Use <strong>Display rate as</strong> in the Chart tab
          to show rates per 1, 100, 1,000 or another available unit.
          This changes the display scale, not the statistical calculation.
        </p>
      `)}

      ${card("Common mistake", `
        <p style="margin:0;">
          Enter the actual event counts and opportunities,
          not a pre-calculated rate.
          Unlike a proportion, a rate can exceed one event
          per opportunity.
        </p>
      `)}
    `;
    return;
  }

  // Pareto chart setup
  if (chartType === "pareto") {
    subtitle.textContent = "Pareto chart: compare and prioritise categories.";

    body.innerHTML = `
      ${card("Use this when...", `
        <p style="margin:0;">
          You want to identify which categories account for
          the most recorded occurrences.
          A Pareto chart ranks categories from most to least frequent,
          with a cumulative percentage line.
        </p>
      `)}

      ${card("You need these columns", `
        <ul style="margin:0;">
          <li><strong>Category column</strong> → the category you want to compare</li>
          <li><strong>Frequency / count column</strong> → choose how occurrences are counted</li>
        </ul>
        <p style="margin-bottom:0;">
          If each row represents one occurrence, select
          <strong>(count each row as 1)</strong>.
          If your data are already summarised, select the numeric
          column containing the counts.
        </p>
      `)}

      ${card("Example 1: one row per occurrence", exampleTable(
        ["Category"],
        [["Medication"], ["Falls"], ["Medication"], ["Pressure damage"], ["Falls"]]
      ))}

      ${card("Example 2: already summarised counts", exampleTable(
        ["Category", "Count"],
        [["Medication", "25"], ["Falls", "18"], ["Pressure damage", "7"]]
      ))}

      ${card("What the chart shows", `
        <p style="margin:0;">
          The blue bars show the number of occurrences in each category.
          The red line shows the cumulative percentage.
          <br><br>
          The dashed 80% line is a guide for prioritisation,
          <strong>not a statistical control limit</strong>.
        </p>
      `)}

      ${card("Common mistakes", `
        <p style="margin:0;">
          Do not assume that the most frequent categories
          are necessarily the underlying causes of the problem.
          They are useful starting points for investigation.
          <br><br>
          When using a numeric count column, counts must be
          non-negative whole numbers. At least two categories
          with positive counts are needed.
        </p>
      `)}
    `;
    return;
  }


  if (chartType === "xbars") {
    subtitle.textContent = "X̄–S chart: grouped measurements.";
    body.innerHTML = `
      ${card("Use this when...", `
        <p style="margin:0;">You have <strong>multiple measurements within each subgroup</strong> and want to understand both subgroup averages and within-group variation.</p>
      `)}

      ${card("You need these columns", `
        <ul style="margin:0;">
          <li><strong>Date / X-axis column</strong> → subgroup label or time label</li>
          <li><strong>Value / Y-axis column</strong> → measurement value</li>
          <li><strong>Third column</strong> → subgroup ID</li>
        </ul>
      `)}

      ${card("Example layout", exampleTable(
        ["Day", "Reading", "Sample_ID"],
        [["Mon", "10.2", "A"], ["Mon", "10.5", "A"], ["Tue", "9.8", "B"], ["Tue", "10.1", "B"]]
      ))}

      ${card("Common mistake", `
        <p style="margin:0;">Do not use X̄–S if each subgroup only has one reading. Use <strong>XmR</strong> instead.</p>
      `)}
    `;
    return;
  }

  if (chartType === "t") {
    subtitle.textContent = "T chart: time between rare events.";

    const checkedDates = tChartInputMode === "eventDates" ? "checked" : "";
    const checkedGaps  = tChartInputMode === "gaps" ? "checked" : "";

    body.innerHTML = `
      ${card("Use this when...", `
        <p style="margin:0;">You want to track the <strong>time between rare events</strong>.</p>
      `)}

      ${card("Choose your setup", `
        <label style="display:block; margin:0.25rem 0;">
          <input type="radio" name="tChartInputMode" value="eventDates" ${checkedDates}>
          <strong>I have event dates</strong> (one row per event)
        </label>
        <div class="hint small-hint" style="margin-top:0.15rem; margin-bottom:0.5rem;">
          Put the event date/time in <em>Date / X-axis column</em>. The <em>Value / Y-axis column</em> is not used.
        </div>

        <label style="display:block; margin:0.25rem 0;">
          <input type="radio" name="tChartInputMode" value="gaps" ${checkedGaps}>
          <strong>I already have the gaps</strong> (numeric time between events)
        </label>
        <div class="hint small-hint" style="margin-top:0.15rem;">
          Put the gap values in <em>Value / Y-axis column</em>.
        </div>
      `)}

      ${card("Example layout", exampleTable(
        ["Event date"],
        [["01/01/2024"], ["12/01/2024"], ["20/01/2024"]]
      ))}

      ${card("Common mistake", `
        <p style="margin:0;">Do not use a T chart for event counts per month. Use a <strong>C</strong>, <strong>P</strong> or <strong>U</strong> chart instead.</p>
      `)}
    `;

    body.querySelectorAll("input[name='tChartInputMode']").forEach(r => {
      r.addEventListener("change", () => {
        tChartInputMode = r.value;
        try { localStorage.setItem("spc_tChartInputMode", tChartInputMode); } catch {}

        if (typeof updateUIForChartType === "function") {
          updateUIForChartType("t");
        }

        if (rawRows && rawRows.length && generateButton) {
          generateButton.click();
        }
      });
    });

    return;
  }

  if (chartType === "g") {
    subtitle.textContent = "G chart: opportunities between rare events.";
    body.innerHTML = `
      ${card("Use this when...", `
        <p style="margin:0;">You want to track the <strong>number of opportunities between rare events</strong>.</p>
      `)}

      ${card("You need these columns", `
        <ul style="margin:0;">
          <li><strong>Date / X-axis column</strong> → optional label / order column</li>
          <li><strong>Value / Y-axis column</strong> → number of opportunities between events</li>
        </ul>
      `)}

      ${card("Example layout", exampleTable(
        ["Week", "Procedures between harms"],
        [["1", "35"], ["2", "48"], ["3", "27"]]
      ))}

      ${card("Common mistake", `
        <p style="margin:0;">Use G when the thing between events is a <strong>count of opportunities</strong>. If it is elapsed time, use a <strong>T chart</strong>.</p>
      `)}
    `;
    return;
  }

  body.innerHTML = `
    ${card("Setup guidance", `
      <p style="margin:0;">Use the column labels shown in <strong>Choose columns</strong>.</p>
    `)}
  `;
}

/*
 * Batch 36A-2 — Live language switching.
 *
 * If the Chart Setup modal is open, refresh its
 * translated instructions in the selected language.
 *
 * Keep the same chart type, T-chart input mode,
 * and saved "don't show again" preference.
 */
document.addEventListener(
  "simplespc:languagechange",
  () => {

    const modal = document.getElementById(
      "chartSetupModal"
    );

    // Nothing to do when the modal is closed.
    if (!modal || !modal.classList.contains("visible")) {
      return;
    }

    // Refresh instructions using the current chart type.
    // This does not regenerate the SPC chart.
    renderChartSetupModal(
      lastChartSetupRenderedType
    );

  }
);

function maybeShowChartSetupModal(chartType) {
  if (!shouldAutoShowChartSetupModal()) return;
  renderChartSetupModal(chartType);
  toggleChartSetupModal(true);
}

function shouldAutoShowChartSetupModal() {
  try {
    return localStorage.getItem("spc_hideChartSetupModal") !== "true";
  } catch {
    return true;
  }
}


/*
 * SimpleSPC English translations
 *
 * This is the reference language for the application.
 * Other languages will use the same translation keys.
 */

window.SimpleSPCi18n.registerLanguage("en", {

  // Application heading
  "app.title": "Simple SPC Web Tool",
// Browser tab and file upload
"app.browserTitle": "Simple SPC Tool",
"data.chooseFile": "Choose file",
"data.noFileChosen": "No file selected",

  "app.subtitle": "Upload data and create SPC charts in your browser. Data never leaves your browser.",

  // Main buttons
  "buttons.generate": "Generate / Recalculate",
  "buttons.reset": "Reset chart",
  "buttons.help": "Help & guidance",
  "buttons.spcHelper": "SPC helper",

  // Main tabs
  "tabs.data": "Data",
  "tabs.chart": "Chart",
  "tabs.labels": "Labels",
  "tabs.axes": "Axes",
  "tabs.baseline": "Baseline",
  
  "tabs.rules": "Rules",

  // Language selector
  "language.label": "Language",

  // Introduction
  "app.intro": "Add your data, choose your chart and settings, then click 'Generate / Recalculate' when needed. You can upload CSV or Excel files, clean data in the editor, and save chart to reopen later. Not sure about SPC? Ask the helper.",

  // Tooltips
  "buttons.resetTitle": "Keep your loaded data and selected columns, but reset chart settings including targets, splits, annotations, rules and formatting.",

  
"tabs.baselineTitle": "Set an initial baseline period or a performance target",

// DATA PANEL
"data.addHeading": "Add data:",
"data.requirementsAria": "Data requirements help",
"data.requirementsTooltip": "The columns you need depend on the chart type. Most SPC charts need a time or sequence column and a measurement. Some also need counts and denominators. Pareto charts need a category column and can optionally use a frequency/count column.",
"data.uploadHint": "Upload a CSV/Excel file or paste data in the editor.",
"data.openEditor": "Open data editor (paste/type data)",
"data.clear": "Clear data & chart",
"data.clearTitle": "Remove the loaded data and reset the chart. Save your work first if you want to reopen it later.",

"data.saveHeading": "Save or open a chart",
"data.saveHint": "Save your chart, data and settings as a JSON file to reopen and edit later.",
"data.saveButton": "Save chart",
"data.openButton": "Open saved chart",

"data.portableHeading": "Save an offline copy",
"data.portableAria": "Portable chart help",
"data.portableTooltip": "Save a self-contained HTML copy of your chart. The file includes your data and a working copy of SimpleSPC. You can open it without an internet connection, make changes and save it again. Portable files contain your data. Store and share them in line with your organisation's information-governance requirements.",
"data.portableButton": "Save portable chart",

// CHART PANEL
"chart.noData": "Load data to enable chart and column selection.",
"chart.newDataHeading": "New data loaded",
"chart.newDataQuestion": "Would you like help choosing the most appropriate chart for these data?",
"chart.helpChoose": "Help me choose",
"chart.chooseManually": "Choose manually",
"chart.typeHeading": "Chart type:",
"chart.run": "Run chart",
"chart.xmr": "XmR chart",
"chart.moreTypes": "More chart types",

"chart.setup": "Chart setup",

// SECONDARY MR / S CHART CONTROLS
"chart.secondary.showMR": "Show MR",
"chart.secondary.showS": "Show S chart",
"chart.secondary.mrDisplay": "MR display",
"chart.secondary.sDisplay": "S chart display",
"chart.secondary.includesSplits": "(includes splits)",
"chart.secondary.current": "Current",
"chart.secondary.all": "All",
"chart.secondary.helpAria": "Secondary chart display help",
"chart.secondary.helpText": "XmR charts can show a Moving Range (MR) chart, which helps explain the variation between consecutive measurements. X-bar and S charts can show an S chart, which displays variation within subgroups. Current shows the latest analysis period only. All shows the full history, including split periods. Use the checkbox to show or hide the secondary chart.",

// COLUMN SELECTION
"chart.columns.heading": "Choose columns:",
"chart.columns.xDefault": "Date / X-axis column",
"chart.columns.yDefault": "Value / Y-axis column",
"chart.columns.thirdDefault": "Denominator / opportunities",
"chart.columns.xmrY": "Measure (used for XmR limits)",
"chart.columns.cY": "Count (c) per time period",
"chart.columns.pY": "Numerator: defectives (d)",
"chart.columns.pThird": "Denominator: total (n)",
"chart.columns.pHint": "P chart plots a proportion: d out of n.",
"chart.columns.pprimeHint": "P′ chart plots a proportion: d out of n, with Laney-adjusted control limits.",
"chart.columns.uY": "Numerator: defects (c)",
"chart.columns.uThird": "Denominator: opportunities (n)",
"chart.columns.uHint": "U chart plots defects per opportunity: c per n.",
"chart.columns.uprimeHint": "U′ chart plots defects per opportunity, with Laney-adjusted control limits.",
"chart.columns.xbarsY": "Measurement value",
"chart.columns.xbarsThird": "Subgroup ID (e.g. day / week / sample)",
"chart.columns.xbarsHint": "X̄–S needs multiple measurements per subgroup. The horizontal axis uses subgroup order automatically.",
"chart.columns.tGapsY": "Time between events (e.g. days)",
"chart.columns.tDatesY": "Value column not used (T chart uses dates)",
"chart.columns.tGapsHint": "T chart plots time between rare events (using your numeric gaps).",
"chart.columns.tDatesHint": "T chart plots time between rare events (calculated from event dates).",
"chart.columns.gY": "Opportunities between events",
"chart.columns.gHint": "G chart plots opportunities between rare events.",
"chart.columns.paretoX": "Category column",

"chart.columns.paretoY": "Frequency / count column",

// ADDITIONAL CHART TYPES
"chart.types.c": "C chart (counts)",
"chart.types.p": "P chart (proportion)",
"chart.types.pprime": "P′ chart (Laney)",
"chart.types.u": "U chart (rate)",
"chart.types.uprime": "U′ chart (Laney)",
"chart.types.xbars": "X̄–S chart (subgroups)",
"chart.types.t": "T chart (time between)",
"chart.types.g": "G chart (opportunities between)",
"chart.types.pareto": "Pareto chart (categories)",

// RATE DISPLAY
"chart.rate.label": "Display rate as:",
"chart.rate.per1": "Per 1",
"chart.rate.per10": "Per 10",
"chart.rate.per100": "Per 100",
"chart.rate.per1000": "Per 1,000",
"chart.rate.per10000": "Per 10,000",
"chart.rate.hint": "Changes the displayed rate only; the underlying U-chart calculation is unchanged.",

// X-AXIS TYPE
"chart.axis.type": "X-axis type",
"chart.axis.helpAria": "X-axis type help",
"chart.axis.help": "Use Date / time for time-series SPC. Use Sequence / category to plot rows in their original order without treating the X-axis values as dates.",
"chart.axis.date": "Date / time",
"chart.axis.sequence": "Sequence / category",

// DATE FORMAT
"chart.dateFormat.label": "Date format for numeric dates",
"chart.dateFormat.helpAria": "Date format help",
"chart.dateFormat.help": "Used for ambiguous numeric dates such as 01/02/2024. ISO dates such as 2024-02-01 and Excel dates are recognised automatically.",
"chart.dateFormat.uk": "UK / day-first (dd/mm/yyyy)",
"chart.dateFormat.us": "US / month-first (mm/dd/yyyy)",
"chart.dateFormat.auto": "Auto-detect",

"chart.dateFormat.iso": "ISO / Excel dates only",

// CHART AVAILABILITY MESSAGES
"chart.availability.noData": "Load data to assess which chart types are available.",
"chart.availability.selectedColumn": "the selected value column",
"chart.availability.xmr": "Less than 12 data points available in \"{column}\" — this is a minimum requirement for a valid XmR chart.",
"chart.availability.c": "No non-negative whole-number count column is available yet — C charts need counts per time period.",
"chart.availability.p": "No suitable numerator/denominator pair is available yet — P charts need two whole-number columns where the numerator is part of the denominator.",
"chart.availability.pprime": "No suitable numerator/denominator pair is available yet — P′ charts need two whole-number columns where the numerator is part of the denominator.",
"chart.availability.u": "No suitable count/opportunities pair is available yet — U charts need a whole-number count column plus a positive opportunities column.",
"chart.availability.uprime": "No suitable count/opportunities pair is available yet — U′ charts need a whole-number count column plus a positive opportunities column.",
"chart.availability.xbars": "X̄–S charts need a numeric measurement column and a subgroup column with repeated subgroup labels.",
"chart.availability.t": "T charts need either event dates or a numeric time-between-events column.",
"chart.availability.g": "G charts need a whole-number column where all values are at least 1, because they plot opportunities between rare events.",

// DATE-FORMAT WARNINGS
"chart.warning.mixedDates": "Warning: this column appears to contain a mixture of UK-style and US-style numeric dates. Please standardise the dates if possible.",
"chart.warning.usDatesUkSelected": "Warning: these dates look like US month-first dates (mm/dd/yyyy), but UK day-first is selected. Some rows may be ignored. Try switching to US month-first or Auto-detect.",
"chart.warning.ukDatesUsSelected": "Warning: these dates look like UK day-first dates (dd/mm/yyyy), but US month-first is selected. Some rows may be ignored. Try switching to UK day-first or Auto-detect.",
"chart.warning.ambiguousUk": "This column contains ambiguous numeric dates. They will currently be interpreted as UK day-first dates (dd/mm/yyyy).",
"chart.warning.ambiguousUs": "This column contains ambiguous numeric dates. They will currently be interpreted as US month-first dates (mm/dd/yyyy).",
"chart.warning.ambiguousAuto": "This column contains ambiguous numeric dates and the tool cannot confidently auto-detect the style. It will fall back to UK day-first dates unless you choose another option.",
"chart.warning.ambiguousIso": "This column contains numeric slash/hyphen dates. In 'ISO / Excel dates only' mode, ambiguous numeric dates may not be interpreted as dates.",

// DISABLED DATE-FORMAT CONTROL

"chart.dateFormat.disabled": "Date-format selection is only available when the X-axis type is Date / time.",

// SHARED FORMATTING LABELS
"common.autoDefault": "Auto / default",
"common.auto": "Auto",
"common.bold": "Bold",
"common.italic": "Italic",
"common.underline": "Underline",

// LABELS PANEL
"labels.heading": "Chart labels",
"labels.title": "Chart title",
"labels.xAxis": "X-axis label",
"labels.yAxis": "Y-axis label",
"labels.titleSettings": "Title settings",
"labels.fontFamily": "Title font family",
"labels.fontSize": "Title font size",
"labels.showTitle": "Show chart title",

// AXES PANEL
"axes.heading": "Axes",
"axes.xSection": "X axis",
"axes.ySection": "Y axis",
"axes.xFontFamily": "X-axis font family",
"axes.xFontSize": "X-axis font size",
"axes.yMin": "Y-axis minimum",
"axes.yMax": "Y-axis maximum",
"axes.yRangeAria": "Y-axis range help",
"axes.yRangeHelp": "Leave the minimum and maximum blank to let the chart choose its range automatically.",
"axes.yFontFamily": "Y-axis font family",
"axes.yFontSize": "Y-axis font size",
"axes.yFormat": "Y-axis format",
"axes.formatAuto": "Auto",
"axes.formatDecimal": "Decimal",
"axes.formatInteger": "Integer",
"axes.formatPercent": "Percent",
"axes.decimals": "Y-axis decimal places",
"axes.decimalsAuto": "Auto",
"axes.dp0": "0 dp",
"axes.dp1": "1 dp",
"axes.dp2": "2 dp",
"axes.dp3": "3 dp",
"axes.tickInterval": "Y-axis tick interval",
"axes.tickAria": "Y-axis tick interval help",
"axes.tickHelp": "Leave blank for automatic spacing. For percentages, enter percentage points: for example, 1 means 1%, and 0.5 means 0.5%.",

// BASELINE PANEL
"baseline.heading": "Baseline data",
"baseline.helpAria": "Baseline data help",
"baseline.help": "A baseline is an initial reference period used to understand how your process was performing before later observations. Enter the number of points to use as your starting reference. SimpleSPC uses these to calculate the initial centre line and, where applicable, control limits. Later observations remain visible for comparison. Leave blank if you do not need a separate initial baseline period. Choose a baseline based on your knowledge of the process, not simply to remove inconvenient signals.",
"baseline.useFirst": "Use first",
"baseline.pointsAsBaseline": "points as baseline",
"baseline.blankHint": "(leave blank if no baseline period)",
"baseline.targetHeading": "Target",
"baseline.targetHelpAria": "How to set a target",
"baseline.targetHelp": "Enter the performance level you want to achieve, then choose whether values at or above, or at or below, that level are good. A target is a performance goal, not a statistical control limit.",
"baseline.targetHint": "Enter your target value, then choose whether higher or lower values are better.",
"baseline.targetValue": "Target value",
"baseline.atOrAbove": "At or above target is good",
"baseline.atOrBelow": "At or below target is good",
"baseline.hideTarget": "Hide target line",

"baseline.showTarget": "Show target line",

// RULES PANEL
"rules.heading": "Detection rules",
"rules.explainer": "Rule explainer",

"rules.shift": "Shift (run) length",
"rules.shiftAria": "Shift rule help",
"rules.shiftHelp": "Number of consecutive points on the same side of the centre line required to flag a shift. Default: 8.",

"rules.trend": "Trend length",
"rules.trendAria": "Trend rule help",
"rules.trendHelp": "Number of consecutive increases or decreases required to flag a trend. Default: 6. This setting is only used when the trend rule is available and enabled.",

"rules.advanced": "Advanced detection options",
"rules.enableTrend": "Enable trend rule",
"rules.enableTrendHelp": "Looks for a sustained sequence of consecutive increases or decreases. The required length is set by the Trend length control above.",
"rules.advancedCaution": "Advanced option — may increase false alerts.",

"rules.rareLabel": "Run & trend rules for T/G charts",
"rules.rareAria": "Rare event rules help",
"rules.rareHelp": "Enables run and trend detection on T and G rare-event charts. These chart types are naturally irregular, so pattern rules should be used cautiously.",
"rules.rareCaution": "Advanced option — rare-event charts can produce false run/trend signals.",

"rules.clamp": "Clamp LCL at zero",
"rules.clampAria": "Clamp LCL help",
"rules.clampHelp": "If the calculated lower control limit is below zero, display the LCL at zero instead. Useful for measures such as counts or times that cannot be negative.",

"rules.zoneIntro": "Optional zone rules (off by default). Use only on chart types where they are supported.",
"rules.zone23": "Apply “2 out of 3 in the outer third” rule",
"rules.zone45": "Apply “4 out of 5 beyond 1-sigma” rule",
"rules.zoneCaution": "Additional pattern rules can increase false alerts on XmR / individuals charts. They are off by default and should only be used when you understand the trade-off.",

"rules.mrHeading": "XmR limit calculation",
"rules.mrAria": "XmR moving range screening help",
"rules.mrHelp": "A very large jump between two consecutive points can make XmR control limits too wide. When this is on, SimpleSPC keeps all data points and moving ranges on the chart, but does not use unusually large moving ranges to estimate routine variation. Uses the Nelson method: moving ranges greater than 3.5 times the initial average moving range are excluded from the variation estimate.",
"rules.mrScreen": "Screen unusually large jumps when setting limits",

"rules.conservative": "This chart uses a conservative rule set designed for reliable interpretation.",
"rules.flag": "Flag special-cause signals on the chart",
"rules.colourBlind": "Colour-blind friendly colours and symbols",

// DYNAMIC RULE CONTROL TOOLTIPS
"rules.disabledTrend": "Trend rule is not offered for this chart type.",
"rules.enableTrendFirst": "Tick 'Enable trend rule' to use this setting.",
"rules.enableRareFirst": "Enable run & trend rules for this chart type to use this setting.",
"rules.rareOnly": "Only used for T and G charts.",
"rules.zoneBlocked": "Zone rules are not available for this chart type because they may create misleading alerts.",

// CONFIRMATIONS WHEN ENABLING RARE-EVENT RULES
"rules.rareConfirmImmediate": "T and G charts are naturally irregular.\n\nRun and trend rules may create false signals on this chart type and are off by default.\n\nOnly enable these rules if you understand this trade-off.",

"rules.rareConfirmOnce": "Advanced option: Run & trend rules on {chart} charts\n\nThese patterns often happen by chance on T/G charts and can increase false alerts.\n\nEnable anyway?",

// RULE EXPLAINER POP-UP
"explainer.title": "SPC rule explainer",
"explainer.close": "Close",
"explainer.closeAria": "Close rule explainer",
"explainer.subtitle": "{chart}: this tool uses a conservative, chart-aware rule policy designed to reduce false alerts.",

// CHART NAMES
"explainer.chart.run": "Run chart",
"explainer.chart.xmr": "X-MR chart",
"explainer.chart.c": "C chart",
"explainer.chart.p": "P chart",
"explainer.chart.pprime": "P′ chart",
"explainer.chart.u": "U chart",
"explainer.chart.uprime": "U′ chart",
"explainer.chart.xbars": "X̄-S chart",
"explainer.chart.t": "T chart",
"explainer.chart.g": "G chart",
"explainer.chart.pareto": "Pareto chart",

// WHY DIFFERENT CHARTS USE DIFFERENT RULES
"explainer.why.attribute": "Attribute charts use a conservative rule set here. Zone rules are blocked because they can create misleading alerts, especially when limits vary or counts are low.",
"explainer.why.rare": "Rare-event charts are naturally irregular. Run and trend rules can create false signals, so they are advanced-only and warning-gated.",
"explainer.why.xmr": "XmR charts are used for individual measurements over time. SimpleSPC keeps the default interpretation deliberately simple, while additional pattern rules and calculation options are available for users who need them.",
"explainer.why.xbars": "X̄-S charts are the chart family where advanced rule sets are most defensible, but the default remains conservative to reduce false alarms.",
"explainer.why.run": "Run charts do not use control limits. The tool keeps interpretation simple and conservative by default.",
"explainer.why.pareto": "A Pareto chart is not a time-ordered control chart, so SPC signal rules do not apply. It ranks categories to help focus investigation and improvement effort.",
"explainer.why.general": "This tool keeps the default rule set simple and conservative to reduce false alerts.",

// RULE POLICY
"explainer.checksHeading": "What this chart checks by default",
"explainer.rule.beyond": "Beyond limits",
"explainer.rule.run": "Run rule",
"explainer.rule.trend": "Trend rule",
"explainer.rule.zone23": "2 of 3 in the outer third",
"explainer.rule.zone45": "4 of 5 beyond 1-sigma",
"explainer.status.on": "used by default.",
"explainer.status.optional": "available in advanced settings, off by default.",
"explainer.status.warn": "advanced only, off by default, with a warning before use.",
"explainer.status.unavailable": "not offered for this chart type.",

// WHY RULES ARE LIMITED
"explainer.limitedHeading": "Why some rules are limited",
"explainer.limited1": "This tool is designed for reliable interpretation in healthcare and public-service settings. It prefers fewer, more trustworthy signals over a noisier rule set that may create false alerts.",
"explainer.limited2": "A “signal” can trigger meetings, concern, investigation, or action. To reduce wasted effort and alert fatigue, some rules are restricted to chart types where they are more dependable.",

// RULE DEFINITIONS
"explainer.definitionsHeading": "Plain-English rule definitions",
"explainer.definition.beyond": "<strong>Beyond limits:</strong> a point above the upper limit or below the lower limit.",
"explainer.definition.run": "<strong>Run rule:</strong> a sustained run of points on one side of the centre line.",
"explainer.definition.trend": "<strong>Trend rule:</strong> a sustained pattern of consecutive increases or decreases.",
"explainer.definition.zone": "<strong>Zone rules:</strong> extra pattern rules based on how far points sit from the centre line. These are the least portable rules, so they are only offered on chart types where they are more defensible.",

// XMR TECHNICAL EXPLANATION
"explainer.xmr.heading": "How XmR limits are set",
"explainer.xmr.intro": "XmR charts estimate routine variation from the differences between consecutive points. These differences are called <strong>moving ranges</strong>.",
"explainer.xmr.largeJump": "A very large jump between two points can make the estimated variation too large and produce control limits that are too wide. By default, FreeSPC screens unusually large moving ranges when setting XmR limits. The original data points and moving ranges are still shown on the chart.",
"explainer.xmr.technical": "<strong>Technical detail:</strong> FreeSPC uses the Nelson method. Moving ranges greater than 3.5 times the initial average moving range are not used when estimating routine variation.",
"explainer.xmr.current": "<strong>Current setting:</strong> moving-range screening is <strong>{state}</strong>. You can change this under <em>Advanced detection options</em>.",
"explainer.xmr.split": "<strong>What happens when you add a split?</strong> A split tells FreeSPC that you want to treat the data before and after that point as different process periods. Each period gets its own mean and control limits, and the moving range across the split is not used.",
"explainer.xmr.minimum": "FreeSPC needs at least <strong>{points} points</strong> to interpret the XmR limits for a period. If a split creates a shorter period, the limits are shown as provisional and FreeSPC does not label that period as showing routine or special-cause variation yet.",
"explainer.on": "ON",
"explainer.off": "OFF",

// RESULTS AREA
"results.chartHeading": "Chart:",
"results.chartHelpAria": "Chart right-click menu help",
"results.chartHelp": "Right-click the chart for additional options. Depending on the chart type, you can add annotations, create or remove splits, exclude or reinclude investigated data points from limit calculations, manage centre-line labels, and copy or export your chart. Some actions require you to right-click directly on a plotted data point.",

// QUICK START GUIDE
"results.quickStart": "Quick start:",
"results.quick1": "Upload a CSV or Excel file, or open the Data editor to paste/type data",
"results.quick2Before": "If needed, clean or trim the data in the Data editor, then click",
"results.quickApply": "Apply data",
"results.quick3": "Select your chart type and confirm the X and Y columns",
"results.quick4Before": "Click",
"results.quick4After": "if changes do not load automatically",
"results.quick5": "Add titles, labels, targets, splits, or annotations as needed",
"results.quick6Before": "Right-click the chart for additional options, and use",
"results.quick6After": "if you want to return later",

// DYNAMIC CONTROLS
"results.hideControls": "Hide controls",
"results.showControls": "Show controls",
"results.needsRecalc": "Changes saved — click Generate / Recalculate",

// SECONDARY CHART HEADINGS
"results.mrHeading": "Moving Range chart:",
"results.sHeading": "S chart:",

"results.mrCanvasTitle": "Moving Range (MR)",

// CHART LEGENDS
"visual.value": "Value",
"visual.centreLine": "Centre line",
"visual.median": "Median",
"visual.ucl": "UCL",
"visual.lcl": "LCL",
"visual.target": "Target",
"visual.xbar": "X̄",
"visual.count": "Count",
"visual.cumulative": "Cumulative %",
"visual.movingRange": "Moving range",
"visual.mrAverage": "MR average",
"visual.mrUcl": "MR UCL",
"visual.mrAveragePeriod": "MR average (Period {period})",
"visual.mrUclPeriod": "MR UCL (Period {period})",
"visual.omittedExcluded": "Not used: adjacent to excluded point",
"visual.omittedScreened": "Not used: unusually large MR screened",

// MR CHART HOVER TOOLTIPS
"visual.mrTooltip": "Moving range: {value}",
"visual.mrExcludedTooltip": "Not used in the variation estimate: {value} (adjacent to an excluded point)",
"visual.mrScreenedTooltip": "Not used in the variation estimate: {value} (screened as an unusually large moving range)",

// ADDITIONAL XMR LEGEND LABELS
"visual.mean": "Mean",
"visual.ucl3sigma": "UCL (3σ)",
"visual.lcl3sigma": "LCL (3σ)",

// PARETO HOVER TOOLTIPS
"visual.pareto.count": "Count: {count} ({share}% of total)",
"visual.pareto.cumulative": "Cumulative: {percent}%",
"visual.pareto.through": "Cumulative through this category: {percent}%",

// AUTOMATIC CHART TITLES
"auto.title.run": "Run Chart",
"auto.title.xmr": "X Chart",
"auto.title.c": "C chart",
"auto.title.p": "P chart",
"auto.title.pprime": "P′ chart",
"auto.title.u": "U chart",
"auto.title.uprime": "U′ chart",
"auto.title.t": "T chart",
"auto.title.tGaps": "T chart (gaps)",
"auto.title.g": "G chart",
"auto.title.pareto": "Pareto chart",
"auto.title.xbar": "X̄ chart",
"auto.title.s": "S chart",

// AUTOMATIC AXIS LABELS
"auto.axis.date": "Date",
"auto.axis.value": "Value",
"auto.axis.category": "Category",
"auto.axis.count": "Count",
"auto.axis.subgroup": "Subgroup",
"auto.axis.proportion": "Proportion / %",
"auto.axis.rate": "Rate per opportunity",
"auto.axis.rateN": "Rate per {n} opportunities",
"auto.axis.opportunities": "Opportunities between events",
"auto.axis.time": "Time between events",
"auto.axis.timeDays": "Time between events (days)",
"auto.axis.cumulative": "Cumulative %",
"auto.axis.s": "S",

// STATISTICAL SUMMARY INTRODUCTIONS
"summary.heading": "Summary",
"summary.xbarsHeading": "X̄–S summary",
"summary.chart.run": "Run chart",
"summary.chart.xmr": "XmR chart",

"summary.totalPoints": "Total number of points:",
"summary.totalSubgroups": "Total number of subgroups:",
"summary.divided": "The chart is divided into",
"summary.period": "period",
"summary.periods": "periods",
"summary.basedBaseline": "(based on the baseline and any splits).",
"summary.basedSplits": "(based on your splits).",

"summary.rateIntro": "Rates are displayed per",
"summary.opportunity": "opportunity",
"summary.opportunities": "opportunities",

"summary.note": "Note:",
"summary.rareNote": "Rare-event charts are often skewed, so the control limits may not look symmetrical like an XmR chart.",

"summary.paretoOccurrences": "Total occurrences:",
"summary.paretoAcross": "across",
"summary.paretoCategories": "categories.",

// PERIOD HEADINGS AND COVERAGE
"summary.periodLabel": "Period",
"summary.singlePeriod": "Single period",
"summary.initialSegment": "(initial segment / baseline)",
"summary.of": "of",
"summary.coverage": "Coverage:",
"summary.point": "point",
"summary.points": "points",
"summary.subgroups": "subgroups",
"summary.to": "to",

// STATISTICAL MEASURES
"summary.mean": "Mean",
"summary.medianPeriod": "Median (this period)",
"summary.centreLine": "Centre line",
"summary.controlLimits": "control limits",
"summary.lowerLimit": "LCL",
"summary.upperLimit": "UCL",
"summary.range": "range",
"summary.estimatedSigma": "Estimated σ (from MR)",
"summary.averageMr": "average MR",
"summary.centreGap": "Centre line (average gap)",
"summary.centreOpportunities": "Centre line (average opportunities)",
"summary.upperLimitLabel": "upper limit",
"summary.centreAndLimits": "Centre line and limits",
"summary.notAvailable": "(not available)",

// X-BAR / S MEASURES
"summary.subgroupSize": "Subgroup size",
"summary.measurement": "measurement",
"summary.measurements": "measurements",
"summary.perSubgroup": "per subgroup",
"summary.variesSubgroup": "(varies by subgroup)",
"summary.notStated": "not stated",
"summary.xbarCentre": "X̄ chart centre line",
"summary.sCentre": "S chart centre line",
"summary.estimatedProcessSd": "Estimated process standard deviation",

// X-BAR / S BASELINE EXPLANATIONS
"summary.xbars.baseline": "Baseline",
"summary.xbars.firstOne": "the first subgroup (1)",
"summary.xbars.firstMany": "the first",
"summary.xbars.subgroupsMany": "subgroups",
"summary.xbars.windowOne": "defines the calculation window;",
"summary.xbars.windowMany": "define the calculation window;",
"summary.xbars.usedEstimate": "used to estimate the limits",
"summary.xbars.excludedOne": "excluded",
"summary.xbars.excludedMany": "excluded",

// X-BAR / S CONTROL-LIMIT DESCRIPTIONS
"summary.xbars.limitsVary": "control limits vary with subgroup size",
"summary.xbars.sLimitsVary": "the expected S and its control limits vary with subgroup size.",
"summary.xbars.centreRange": "Centre-line range",
"summary.xbars.lclRange": "LCL range",
"summary.xbars.uclRange": "UCL range",
"summary.xbars.limits": "limits",

// CHART STATUS
"summary.xbars.xChart": "X̄ chart",
"summary.xbars.sChart": "S chart",
"summary.xbars.interpretation": "Interpretation",
"summary.xbars.latestPeriod": "Latest period",
"summary.xbars.stableX": "stable (no clear signal of change in subgroup averages).",
"summary.xbars.stableS": "stable (no clear signal of change in within-subgroup variation).",
"summary.xbars.signals": "signal(s):",
"summary.xbars.signalDetected": "one or more special-cause signals were detected.",

// COMBINED PERIOD INTERPRETATIONS
"summary.xbars.bothStable": "Both the subgroup averages (X̄) and within-subgroup variation (S) look stable in this period.",
"summary.xbars.xUnstable": "The subgroup averages (X̄) show special-cause signals, while within-subgroup variation (S) looks stable.",
"summary.xbars.sUnstable": "The subgroup averages (X̄) look stable, while within-subgroup variation (S) shows special-cause signals.",
"summary.xbars.bothUnstable": "Both the subgroup averages (X̄) and within-subgroup variation (S) show special-cause signals.",

// LATEST-PERIOD ASSESSMENTS
"summary.xbars.latestBothStable": "The latest period looks stable on both parts of the X̄–S chart. That means the current process level and within-subgroup variation are behaving broadly consistently under the current rule settings.",
"summary.xbars.latestXUnstable": "The latest period shows possible change in the subgroup averages, while within-subgroup variation remains stable.",
"summary.xbars.latestSUnstable": "The latest period shows possible change in within-subgroup variation, while the subgroup averages remain stable.",
"summary.xbars.latestBothUnstable": "The latest period shows possible special-cause variation in both the subgroup averages and within-subgroup variation.",

// RUN CHART — BASELINE
"summary.run.baselineFirst": "Baseline for this period: first",
"summary.run.baselinePoints": "points used to calculate the median.",
"summary.run.baselineAll": "Baseline for this period: all points in this period used to calculate the median.",

// RUN CHART — SIGNAL DESCRIPTIONS
"summary.run.shiftBefore": "a sustained shift (≥",
"summary.run.shiftAfter": "points on one side of the median)",
"summary.run.trendBefore": "a sustained trend (≥",
"summary.run.trendAfter": "points increasing or decreasing)",
"summary.run.outlier": "an unusual outlier (\"astronomical\" point)",

// RUN CHART — INTERPRETATION
"summary.run.noSignals": "No clear special-cause signals detected in this period (no sustained shift, trend, or unusual outlier). This pattern is consistent with common variation, but always interpret in context.",
"summary.run.hasSignals": "This period shows special-cause signals:",

// RUN CHART — WHERE TO INVESTIGATE
"summary.run.where": "Where to look:",
"summary.run.whereShift": "shift around points",
"summary.run.whereTrend": "trend around points",
"summary.run.whereOutlierOne": "outlier at point",
"summary.run.whereOutlierMany": "outlier at points",
"summary.run.increasing": "increasing",
"summary.run.decreasing": "decreasing",

// RUN CHART — SPLIT EXPLANATION
"summary.run.splitNote": "Each period is summarised separately because splits suggest the process may have changed over time.",

// XMR — SIGNAL DESCRIPTIONS
"summary.xmr.signalBeyond": "one or more points are outside the control limits",
"summary.xmr.signalShift": "a sustained shift (many points on the same side of the mean)",
"summary.xmr.signalTrend": "a sustained trend (steady increase or decrease)",
"summary.xmr.signalOutlier": "an unusual outlier (an ‘astronomical’ point)",

// XMR — BASELINE
"summary.xmr.baselineWindowStart": "Baseline for this period: first",
"summary.xmr.baselineWindowOne": "point defines the baseline window;",
"summary.xmr.baselineWindowMany": "points define the baseline window;",
"summary.xmr.baselineUsed": "used to calculate mean and limits",
"summary.xmr.excludedOne": "excluded",
"summary.xmr.excludedMany": "excluded",
"summary.xmr.baselineAll": "Baseline for this period: all points in this period used to calculate mean and limits.",

// XMR — MANUAL EXCLUSIONS
"summary.xmr.exclusionHeading": "Excluded point adjustment",
"summary.xmr.adjacentOne": "adjacent moving range was not used in the variation estimate because it involves an excluded point.",
"summary.xmr.adjacentMany": "adjacent moving ranges were not used in the variation estimate because they involve an excluded point.",
"summary.xmr.noArtificialMR": "SimpleSPC does not create an artificial moving range across an excluded point.",

// XMR — NELSON SCREENING
"summary.xmr.limitCalculation": "XmR limit calculation",
"summary.xmr.screenedOne": "unusually large point-to-point change was not used when estimating routine variation. All data points remain shown on the chart.",
"summary.xmr.screenedMany": "unusually large point-to-point changes were not used when estimating routine variation. All data points remain shown on the chart.",
"summary.xmr.nelsonUsed": "Nelson moving-range screening was used:",
"summary.xmr.initialMR": "initial average MR",
"summary.xmr.threshold": "screening threshold",
"summary.xmr.revisedMR": "revised average MR",
"summary.xmr.screeningOff": "Moving-range screening is turned off, so all baseline point-to-point changes were used to estimate routine variation.",

// XMR — PROVISIONAL LIMITS
"summary.xmr.provisionalHeading": "Provisional limits",
"summary.xmr.provisionalOnly": "only",
"summary.xmr.provisionalOne": "point was used to estimate the limits for this period.",
"summary.xmr.provisionalMany": "points were used to estimate the limits for this period.",
"summary.xmr.minimumIntro": "SimpleSPC needs at least",
"summary.xmr.minimumEnd": "points before interpreting special-cause signals or capability.",
"summary.xmr.provisionalContext": "The displayed limits are shown for context only.",

// XMR — TARGETS
"summary.xmr.target": "Target",
"summary.xmr.aboveBetter": "at or above is better",
"summary.xmr.belowBetter": "at or below is better",
"summary.xmr.targetCoverage": "of points in this period meet the target",
"summary.xmr.coverageUnavailable": "Target coverage not calculated for this period.",

// XMR — INTERPRETATION
"summary.xmr.notAssessed": "Not yet assessed — more data are needed before this period can be interpreted as showing routine or special-cause variation.",
"summary.xmr.noSignals": "No clear special-cause signals were detected in this period. The pattern is consistent with natural/common variation (still interpret in clinical context).",
"summary.xmr.hasSignals": "This period shows special-cause signals:",
"summary.xmr.where": "Where to look",
"summary.xmr.outsideOne": "outside limits at point",
"summary.xmr.outsideMany": "outside limits at points",
"summary.xmr.outlierOne": "outlier at point",
"summary.xmr.outlierMany": "outliers at points",

// XMR — CAPABILITY
"summary.xmr.estimatedCapability": "Estimated capability (this period)",
"summary.xmr.ifStable": "if the process remains stable, about",
"summary.xmr.futureMeet": "of future points are expected to meet the target.",
"summary.xmr.capability": "Capability",
"summary.xmr.capabilityUnreliable": "a target has been set, but because special-cause signals are present in this period, any capability estimate would be unreliable.",
"summary.xmr.comparePeriods": "Comparing means, limits and target performance between periods can indicate whether the process changed after interventions.",

// XMR — CAPABILITY BADGE
"summary.xmr.badgeHeading": "PROCESS CAPABILITY (last period)",
"summary.xmr.badgeNote": "(Estimated probability of meeting the target in the final period, assuming a stable process and approximate normality.)",
"summary.xmr.badgeWarning": "Special-cause signals detected in the last period:",
"summary.xmr.badgeUnreliable": "capability estimates may be unreliable until these signals are understood.",
"summary.xmr.badgeAction": "Focus on understanding and addressing these causes before relying on capability estimates.",


"summary.exclusions.heading": "Excluded from limit calculations:",
"summary.exclusions.point": "point",
"summary.exclusions.points": "points",
"summary.exclusions.subgroup": "subgroup",
"summary.exclusions.subgroups": "subgroups",
"summary.exclusions.pointVisible": "The excluded point remains visible on the chart.",
"summary.exclusions.pointsVisible": "Excluded points remain visible on the chart.",
"summary.exclusions.subgroupVisible": "The excluded subgroup remains visible on the chart.",
"summary.exclusions.subgroupsVisible": "Excluded subgroups remain visible on the chart.",

// ATTRIBUTE CHARTS — BASELINE EXPLANATIONS
"summary.attr.baselineTitle": "Baseline for this period",
"summary.attr.firstOne": "the first point (1) defines the baseline window;",
"summary.attr.firstBefore": "the first",
"summary.attr.firstAfter": "points define the baseline window;",
"summary.attr.used": "used to calculate centre line and limits",
"summary.attr.excludedOne": "excluded",
"summary.attr.excludedMany": "excluded",

// ATTRIBUTE CHARTS — SIGNALS AND INTERPRETATION
"summary.attr.signals": "Signals",
"summary.attr.stable": "Based on the selected rules, no clear special-cause signals were detected in this period. This suggests the pattern is consistent with routine (common-cause) variation.",
"summary.attr.unstable": "Based on the selected rules, special-cause signals were detected in this period (a pattern unlikely to be routine variation alone).",

// ATTRIBUTE CHARTS — STATISTICAL CAUTION
"summary.attr.caution": "These rules are prompts, not absolute answers. Interpret alongside local context (changes in process, staffing, demand, definitions/coding) and consider basic SPC assumptions (e.g. reasonably consistent measurement and opportunity over time).",

// ATTRIBUTE CHARTS — INDIVIDUAL SIGNAL MESSAGES
"summary.attr.signalAbove": "One or more points above the upper limit",
"summary.attr.signalBelow": "One or more points below the lower limit",

"summary.attr.shiftName": "Shift",
"summary.attr.shiftAbove": "points in a row above the centre line (from",
"summary.attr.shiftBelow": "points in a row below the centre line (from",

"summary.attr.trendName": "Trend",
"summary.attr.trendIncreasing": "points steadily increasing (from",
"summary.attr.trendDecreasing": "points steadily decreasing (from",

"summary.attr.zone23": "Zone: 2 of 3 points in the outer third (from",
"summary.attr.zone45": "Zone: 4 of 5 points beyond 1-sigma (from",

  // LANEY P / P-PRIME DIAGNOSTICS
  "summary.laney.p.suggest": "P′ may be a better fit for these data.",
  "summary.laney.p.over": "The proportions vary more from point to point than a standard P chart expects for these sample sizes. This can make a standard P chart produce limits that are too narrow and create extra alerts. A P′ chart adjusts the limits for this extra variation while keeping the same proportions and centre line.",
  "summary.laney.p.under": "The proportions vary less from point to point than a standard P chart expects. This can make standard P-chart limits too wide and make genuine changes harder to detect. A P′ chart adjusts the limits to better match the variation actually seen in the data.",
  "summary.laney.p.normalTitle": "A standard P chart looks appropriate for these data.",
  "summary.laney.p.normal": "The amount of variation is within the range expected for proportion data of these sample sizes, so there is no clear reason to use a P′ chart instead.",

  "summary.laney.p.prime.over": "These data vary more than a standard P chart expects, so the P′ chart uses wider control limits.",
  "summary.laney.p.prime.under": "These data vary less than a standard P chart expects, so the P′ chart uses narrower control limits.",
  "summary.laney.p.prime.normal": "The amount of variation is very close to what a standard P chart expects, so the adjustment is small.",

  // LANEY U / U-PRIME DIAGNOSTICS
  "summary.laney.u.suggest": "U′ may be a better fit for these data.",
  "summary.laney.u.over": "The rates vary more from point to point than a standard U chart expects for the amount of opportunity. This can make a standard U chart produce limits that are too narrow and create extra alerts. A U′ chart adjusts the limits for this extra variation while keeping the same rates and centre line.",
  "summary.laney.u.under": "The rates vary less from point to point than a standard U chart expects. This can make standard U-chart limits too wide and make genuine changes harder to detect. A U′ chart adjusts the limits to better match the variation actually seen in the data.",
  "summary.laney.u.normalTitle": "A standard U chart looks appropriate for these data.",
  "summary.laney.u.normal": "The amount of variation is within the range expected for defect-rate data, so there is no clear reason to use a U′ chart instead.",

  "summary.laney.u.prime.over": "These data vary more than a standard U chart expects, so the U′ chart uses wider control limits.",
  "summary.laney.u.prime.under": "These data vary less than a standard U chart expects, so the U′ chart uses narrower control limits.",
  "summary.laney.u.prime.normal": "The amount of variation is very close to what a standard U chart expects, so the adjustment is small.",

  // SHARED DISPERSION DIAGNOSTICS
  "summary.laney.technicalCheck": "Technical check",
  "summary.laney.observedVariation": "observed variation",
  "summary.laney.ofExpected": "of expected",
  "summary.laney.diagnosticUpper": "diagnostic upper limit",
  "summary.laney.diagnosticLower": "diagnostic lower limit",
  "summary.laney.expectedRange": "expected range",
  "summary.laney.switchP": "Use P′ chart instead",
  "summary.laney.switchU": "Use U′ chart instead",
  "summary.laney.whyDifferent": "Why are these limits different?",
  "summary.laney.adjustmentBenefit": "This helps the control limits better reflect the amount of variation actually seen in the data.",
"summary.laney.technicalDetail": "Technical detail",

// T AND G RARE-EVENT SUMMARIES
"summary.rare.baseline": "Initial points in this period used to calculate the centre line and limits",

"summary.rare.stable": "No clear special-cause signals were detected in this period. The pattern is consistent with routine/common variation.",
"summary.rare.unstable": "Special-cause signals were detected in this period (pattern inconsistent with routine variation).",

"summary.rare.betterHeading": "Interpreting “better”",
"summary.rare.better": "If the event is something you want to avoid, longer gaps (or more opportunities between events) are usually better. If it is something you want more often, shorter gaps may be better.",

"summary.rare.nextHeading": "What to do next",
"summary.rare.nextStable": "If performance isn't good enough, focus on improving the system rather than reacting to individual points.",
"summary.rare.nextUnstable": "Look for a real-world explanation (process, staffing, demand, detection/definition changes). If it was planned, consider a new baseline after things settle.",

"summary.rare.dataHeading": "Data reminder",
"summary.rare.dataT": "This chart uses the time between events (e.g., days between incidents).",
"summary.rare.dataG": "This chart uses opportunities between events (values should be 1 or more).",

// CHART CONTEXT MENU — BUTTONS
"context.annotations": "Annotations ▸",
"context.annotationAdd": "Add annotation…",
"context.annotationEdit": "Edit annotation…",
"context.annotationDelete": "Delete annotation…",
"context.annotationClear": "Clear all annotations",

"context.excludePoint": "Exclude from limit calculations…",
"context.excludeSubgroup": "Exclude this subgroup from limit calculations…",
"context.includePoint": "Include in limit calculations again",
"context.includeSubgroup": "Include this subgroup in limit calculations again",

"context.splits": "Splits ▸",
"context.addSplit": "Add split after this point…",
"context.removeSplit": "Clear split after {label} (point {point})",
"context.clearSplits": "Clear all splits",
"context.noSplits": "No splits to clear",
"context.splitsUnavailable": "Splits not available for this chart type",

"context.centreShow": "Show centre-line labels",
"context.centreHide": "Hide centre-line labels",
"context.centreHideOne": "Hide this centre-line label",
"context.centreRestore": "Restore hidden centre-line labels",

"context.copy": "Copy chart(s)",
"context.copyAnalysis": "Copy chart(s) + analysis",
"context.savePng": "Save chart(s) as PNG image",
"context.pdf": "Download PDF report",

// CHART CONTEXT MENU — TOOLTIPS
"context.tip.annotations": "Add, edit or delete annotations for the selected point.",
"context.tip.annotationAdd": "Add a new annotation at this point.",
"context.tip.annotationEdit": "Edit an existing annotation at this point.",
"context.tip.annotationDelete": "Delete an existing annotation at this point.",
"context.tip.annotationClear": "Remove all annotations from the chart.",

"context.tip.excludeDefault": "Keep this point visible but exclude it from centre-line and control-limit estimation.",
"context.tip.includeDefault": "Put this point back into the centre-line and control-limit calculation.",
"context.tip.excludeEligible": "Keep this point visible but do not use it to estimate the centre line or control limits.",
"context.tip.nearPoint": "Right-click near a plotted point.",
"context.tip.outsideBaseline": "This point is outside the baseline used to estimate the current limits, so excluding it would not change the limit calculation.",
"context.tip.tooFewPoint": "Excluding this point would leave too few included observations to estimate the control limits.",
"context.tip.tooFewSubgroup": "Excluding this subgroup would leave too few included observations to estimate the control limits.",
"context.tip.noConsecutive": "Excluding this point would leave no pair of consecutive included observations for the moving-range variation estimate.",

"context.tip.splits": "Add or remove split points for this chart.",
"context.tip.addSplit": "Splits recalculate the central line and control limits from the chosen point onward. Splitting the control limits should only be done where a sustained shift in performance has been identified, investigated, understood and agreed to be representative of the new normal level of performance for the process.",
"context.tip.splitsUnavailable": "Splits are not available for this chart type.",
"context.tip.splitPoint": "Right-click near a data point to add a split.",
"context.tip.removeSplit": "Remove the split after {label}.",
"context.tip.clearSplits": "Remove all splits and return to a single set of limits.",
"context.tip.noSplits": "There are no splits to clear.",

"context.tip.centreLabels": "Show or hide the calculated mean values for this XmR chart.",
"context.tip.centreHideOne": "Hide only the mean label that you right-clicked.",
"context.tip.centreRestore": "Restore any individual mean labels that have been hidden.",

"context.tip.copy": "Copy the chart (only) to the clipboard.",
"context.tip.copyAnalysis": "Copy the chart and the analysis area to the clipboard",
"context.tip.savePng": "Save a copy of the chart as a PNG image file",
"context.tip.pdf": "Export the chart and analysis to PDF",

// DISABLED ANNOTATION OPTIONS
"context.tip.annotationPointAdd": "Right-click near a data point to add an annotation.",
"context.tip.annotationPointEdit": "Right-click near a data point to edit an annotation.",
"context.tip.annotationMissingEdit": "There is no annotation at this point to edit.",
"context.tip.annotationPointDelete": "Right-click near a data point to delete an annotation.",
"context.tip.annotationMissingDelete": "There is no annotation at this point to delete.",
"context.tip.annotationNone": "There are no annotations to clear.",

// CHART ACTION DIALOGS
"dialog.noChart": "Please generate a chart first.",
"dialog.nearPoint": "Right-click near a plotted point first.",
"dialog.notInWindow": "This point is not part of the current limit-estimation window.",

// EXCLUDING OBSERVATIONS
"dialog.excludePoint": "Exclude this point from the limit calculations?\n\nThe point will stay visible on the chart, but it will not be used to estimate the centre line or control limits.\n\nOnly do this after you have investigated it and have good reason to believe it reflects a known special cause rather than the routine variation of the process.\n\nDo not exclude a point simply because it is unusually high or low. If it represents a lasting change in the process, use a split instead.",
"dialog.excludeSubgroup": "Exclude this subgroup from the limit calculations?\n\nThe subgroup will stay visible on the chart, but it will not be used to estimate the centre line or control limits.\n\nOnly do this after you have investigated it and have good reason to believe it reflects a known special cause rather than the routine variation of the process.\n\nDo not exclude a point simply because it is unusually high or low. If it represents a lasting change in the process, use a split instead.",
"dialog.excludedPointNote": "Point excluded from the limit calculations. It will remain visible on the chart.\n\nWould you like to add an annotation explaining why it was excluded?",
"dialog.excludedSubgroupNote": "Subgroup excluded from the limit calculations. It will remain visible on the chart.\n\nWould you like to add an annotation explaining why it was excluded?",

// ANNOTATIONS
"dialog.annotationFor": "Annotation for {label}:",
"dialog.newAnnotationFor": "New annotation for {label}:",
"dialog.editAnnotationFor": "Edit annotation for {label}:",
"dialog.annotationPositionMissing": "Could not determine the selected x-position for annotation.",
"dialog.annotationBlank": "Annotation text cannot be blank.",
"dialog.editWhich": "Edit which annotation?",
"dialog.deleteWhich": "Delete which annotation?",
"dialog.annotationsAt": "Annotations at {label}:",
"dialog.chooseNumber": "Type a number from 1 to {count}:",
"dialog.invalidNumber": "Please enter a valid number.",
"dialog.noEditAt": "There are no annotations to edit at {label}.",
"dialog.noDeleteAt": "There are no annotations to delete at {label}.",
"dialog.deleteAnnotation": "Delete this annotation?\n\n\"{label}\"",
"dialog.clearAnnotations": "Clear all annotations?",

// SPLIT VALIDATION
"dialog.splitUnavailable": "Splits / recalculating limits are not available for this chart type.",
"dialog.splitInvalid": "Please select a valid split point.",
"dialog.splitLast": "You can’t split after the last point.",

// COPY AND EXPORT FEEDBACK
"dialog.copyDone": "Chart image copied to clipboard.",
"dialog.copyAnalysisDone": "Chart + analysis image copied to clipboard.",
"dialog.actionFailed": "Sorry — that action failed in this browser. Try 'Save chart(s) as…' instead.",
"dialog.pdfImageFailed": "Could not build the PDF export image.",
"dialog.pdfLibraryMissing": "PDF export library not available.",

// DATA EDITOR
"editor.title": "Data editor",
"editor.intro": "Paste data directly into the grid, or use this editor to review Excel worksheets before importing. You can switch worksheets, delete unwanted rows/columns, and then click",
"editor.helpButton": "How do I delete rows/columns?",
"editor.helpAria": "How do I delete rows or columns?",
"editor.helpHeading": "Editing tip",
"editor.helpRow": "Right-click a row number to delete a row",
"editor.helpColumn": "Right-click a column heading to delete a column",
"editor.helpInsert": "You can also insert rows/columns the same way",
"editor.worksheet": "Worksheet:",
"editor.hasHeaders": "First row contains headers",
"editor.detect": "Auto-detect",
"editor.apply": "Apply data",
"editor.cancel": "Cancel",
"editor.chooseWorksheet": "Choose worksheet",
"editor.chooseDescription": "This Excel file contains more than one worksheet. Choose which one to import.",
"editor.loadSheet": "Load sheet",

// DYNAMIC DATA EDITOR MESSAGES
"editor.workbookStatus": "Workbook loaded. You can switch worksheet here, trim rows/columns in the grid, and then click <strong>Apply data</strong>.",
"editor.headerApplyHeadings": "Apply will treat the <strong>first row of the grid</strong> as <strong>headings</strong>.<br><small>Tip: If you loaded a CSV normally, your headings are already the column titles — so usually leave this OFF.</small>",
"editor.headerApplyData": "Apply will treat the <strong>first row of the grid</strong> as <strong>data</strong>.<br><small>Tip: If you loaded a CSV normally, your headings are already the column titles — so usually leave this OFF.</small>",
"editor.detectHeadings": "Auto-detect: first row looks like <strong>headings</strong>.",
"editor.detectData": "Auto-detect: first row looks like <strong>data</strong>.",
"editor.headerMismatchChecked": "You have 'First row contains headers' ticked, but the first row looks like DATA.\n\nApply anyway?",
"editor.headerMismatchUnchecked": "You have 'First row contains headers' unticked, but the first row looks like HEADERS.\n\nApply anyway?",
"editor.notInitialised": "Spreadsheet editor not initialised. Try reopening the editor.",
"editor.enterRow": "Please enter at least one data row.",
"editor.pasteRow": "Paste at least one row of data.",
"editor.readError": "Unexpected error reading spreadsheet data.",

// RESET AND REPLACEMENT
"data.confirmReset": "Reset this chart?\n\nYour loaded data and selected X/value columns will be kept, but chart settings such as chart type, baseline, target, splits, annotations, labels, rules and axis formatting will be reset.",
"data.confirmClear": "Clear all loaded data?\n\nThis will also reset the current chart and its settings. Save the chart first if you want to reopen it later.",
"data.confirmReplace": "Load these data as a new chart?\n\nThis will clear the current data and chart settings. Save the current chart first if you want to reopen it later.",
"data.noRows": "No rows found in the data.",
"data.noColumns": "Could not detect any columns in the data.",
"data.fileType": "Please upload a CSV or Excel file (.csv, .xlsx, .xls).",
"data.csvError": "Error parsing CSV: {message}",
"data.csvNoHeaders": "It looks like your CSV does not include column headings.\n\nClick OK to treat the first row as DATA (I will create Column1, Column2...).\nClick Cancel if the first row IS a header row (then add headings and upload again).",
"data.csvAddHeaders": "Please add a header row (e.g. Date,Value) and upload again.",
"data.excelReadError": "Unexpected error reading the Excel file.",
"data.csvReadError": "Unexpected error reading the CSV file.",

// JSON PROJECTS AND LEGACY SETTINGS
"project.invalidObject": "That file could not be read as an SPC project.",
"project.noSavedData": "This project file does not contain any saved data.",
"project.loadFailed": "The project data could not be loaded.",
"project.settingsQueued": "Settings loaded. Now upload your CSV or Excel data and the tool will apply these settings automatically.",
"project.unsupported": "That file doesn’t look like a supported SPC project or settings file.",
"project.invalidJson": "Could not read that JSON file. Please check it is a valid export from this tool.",
"project.restoreFailed": "The existing data could not be restored after resetting the chart.",

// CHART GENERATION VALIDATION
"validation.selectColumns": "Please choose both an X-axis column and a value column.",

//HELP & GUIDANCE: SECTIONS 1–3

// Help window
"help.subtitle": "A practical guide to loading data, building charts, interpreting SPC signals, and saving your work.",
"help.closeAria": "Close help",

// 1. What this tool does
"help.1.title": "1. What this tool does",
"help.1.intro": "is a free, browser-based tool for creating Statistical Process Control (SPC) charts and Pareto charts.",
"help.1.purpose": "It helps you understand patterns in data, distinguish routine variation from signals of change, and identify opportunities for improvement. All calculations take place locally in your browser. Your data are not uploaded to a server.",
"help.1.scope": "SimpleSPC is designed for learning, quality improvement, service monitoring and practical analysis. It is not intended to replace a regulated statistical package.",
"help.1.available": "Available charts:",

"help.chart.runName": "Run chart",
"help.chart.runDesc": "shows values in time or sequence order, using a median and simple rules to identify non-random patterns.",
"help.chart.xmrName": "XmR chart",
"help.chart.xmrDesc": "shows individual measurements with a mean and control limits, alongside an optional moving-range chart.",
"help.chart.xbarsName": "X̄–S chart",
"help.chart.xbarsDesc": "examines averages and variation within groups of continuous measurements.",
"help.chart.cName": "C chart",
"help.chart.cDesc": "monitors counts of events or defects when the opportunity for events is reasonably constant.",
"help.chart.pName": "P and Laney P′ charts",
"help.chart.pDesc": "monitor proportions, such as the percentage of patients experiencing an event. P′ adjusts control limits when observed variation differs from that expected by a standard P chart.",
"help.chart.uName": "U and Laney U′ charts",
"help.chart.uDesc": "monitor event rates when the number of opportunities varies. U′ adjusts control limits when observed variation differs from that expected by a standard U chart.",
"help.chart.tName": "T chart",
"help.chart.tDesc": "monitors the time between rare events.",
"help.chart.gName": "G chart",
"help.chart.gDesc": "monitors the number of opportunities between rare events.",
"help.chart.paretoName": "Pareto chart",
"help.chart.paretoDesc": "ranks categories from most to least frequent to help prioritise investigation and improvement.",

"help.1.importantTitle": "Important:",
"help.1.important": "Pareto charts are different from time-ordered SPC charts. They help identify which categories account for the most occurrences, but do not establish whether a process is stable or whether a change is statistically significant.",

// 2. Quick start
"help.2.title": "2. Quick start",
"help.2.step1": "Load data using CSV, Excel, or the Data editor.",
"help.2.step2": "Select the chart type and choose the relevant columns.",
"help.2.step3Before": "Click",
"help.2.step3After": "if the chart does not update automatically.",
"help.2.step4": "Add titles, labels, targets, splits, or annotations if needed.",
"help.2.step5": "Review the summary text and chart signals.",
"help.2.step6": "Save the chart if you want to return to it later.",

// 3. Preparing and loading data
"help.3.title": "3. Preparing and loading your data",
"help.3.intro": "The columns you need depend on the chart type. Most SPC charts need a time or sequence column and a measurement or count.",

"help.3.pName": "P, P′, U and U′ charts",
"help.3.pDesc": "also need a denominator or opportunities column.",
"help.3.xbarsName": "X̄–S charts",
"help.3.xbarsDesc": "need a subgroup identifier and individual measurement values.",
"help.3.paretoName": "Pareto charts",
"help.3.paretoDesc": "need a category column, with an optional frequency/count column if your data are already summarised.",

"help.3.ways": "You can now load data in three ways:",
"help.3.csv": "Upload a CSV file",
"help.3.excel": "Upload an Excel workbook",
"help.3.paste": "Paste or edit data directly",
"help.3.pasteAfter": "in the Data editor",

"help.3.excelBefore": "If you upload an Excel workbook, the tool opens the",
"help.3.excelMiddle": "so you can inspect the worksheet before importing it. You can switch between worksheets, trim unwanted rows or columns, and then click",
"help.3.axisAuto": "The tool tries to suggest sensible X and Y columns automatically, but you can change them yourself at any time.",
"help.3.dateBefore": "If your date column uses numeric dates such as",
"help.3.dateMiddle": "use the",
"help.3.dateSetting": "Date format for numeric dates",
"help.3.dateAfter": "setting so the tool interprets them correctly.",
"help.3.default": "Default:",
"help.3.defaultFormat": "UK day-first (dd/mm/yyyy)",

// HELP & GUIDANCE: SECTIONS 4–6

// 4. Using the Data editor
"help.4.title": "4. Using the Data editor",
"help.4.intro": "The Data editor is useful when:",
"help.4.paste": "you want to paste data from Excel",
"help.4.sheets": "your Excel workbook contains multiple sheets",
"help.4.trim": "your worksheet contains extra rows or columns that need trimming",
"help.4.sheetBefore": "Use the",
"help.4.sheetName": "Worksheet",
"help.4.sheetAfter": "selector to switch between sheets in an uploaded Excel file.",
"help.4.headersBefore": "Tick or untick",
"help.4.headersAfter": "depending on whether the top row is headings or data.",
"help.4.detectBefore": "Use",
"help.4.detectAfter": "as a quick suggestion, but always sense-check it.",
"help.4.rowAction": "Right-click a row number",
"help.4.rowAfter": "to delete or insert rows.",
"help.4.columnAction": "Right-click a column heading",
"help.4.columnAfter": "to delete or insert columns.",
"help.4.multiple": "You can select multiple rows or columns before deleting them.",
"help.4.tip": "Tip: the Data editor lets you clean a worksheet inside the tool before applying it to the chart.",

// 5. X-axis type and chart type
"help.5.title": "5. X-axis type and chart type",
"help.5.intro": "Choose the chart type that matches the kind of data you have. The main options are:",
"help.5.run": "a simple starting point for values over time.",
"help.5.xmr": "individual continuous measurements, with control limits.",
"help.5.xbars": "continuous measurements collected in subgroups.",
"help.5.c": "counts where the opportunity for events is reasonably constant.",
"help.5.p": "proportions, using a numerator and denominator. P′ adjusts for greater or smaller variation than expected.",
"help.5.u": "event rates with changing opportunities. U′ adjusts for greater or smaller variation than expected.",
"help.5.tgName": "T and G charts",
"help.5.tg": "time or opportunities between rare events.",
"help.5.pareto": "ranks categories from most to least frequent to help prioritise investigation.",
"help.5.xaxis": "X-axis type:",
"help.5.dateName": "Date / time",
"help.5.date": "use when observations have dates or times and you want to display them chronologically.",
"help.5.sequenceName": "Sequence / category",
"help.5.sequence": "use when observations should be shown in order without treating their labels as dates.",
"help.5.paretoNote": "Pareto charts use categories rather than dates. Categories are ranked by frequency, with a cumulative percentage line. They are not time-ordered SPC charts and do not have statistical control limits.",
"help.5.chooseBefore": "If you are unsure which chart to use, click",
"help.5.chooseMiddle": "You can also use",
"help.5.chooseAfter": "for guidance on the data columns required by the selected chart.",

// 6. Baselines, splits, targets and context menu
"help.6.title": "6. Baselines, splits, targets and the right-click menu",
"help.6.intro": "SimpleSPC offers additional features to help you investigate and explain your data.",
"help.6.baselineHeading": "Set a baseline:",
"help.6.baselineBefore": "Open the",
"help.6.baselineAfter": "tab and enter the number of initial points you want to use as a reference period. Leave this blank if you do not need a separate initial baseline.",
"help.6.targetHeading": "Set a target:",
"help.6.targetBefore": "Open the",
"help.6.targetMiddle": "tab and find",
"help.6.targetName": "Target",
"help.6.targetAfter": "Enter the level of performance you want to achieve, then select whether values at or above, or at or below, that target are good. You can also show or hide the target line. A target is a performance goal, not a statistical control limit.",
"help.6.splitHeading": "Add a split:",
"help.6.splitBefore": "Right-click a plotted point and choose",
"help.6.splitsMenu": "Splits",
"help.6.splitMiddle": "then",
"help.6.addSplitAction": "Add split after this point",
"help.6.splitAfter": "This divides the chart into separate analysis periods where supported.",
"help.6.annotationHeading": "Add an annotation:",
"help.6.annotationBefore": "Right-click a plotted point and use the",
"help.6.annotationsMenu": "Annotations",
"help.6.annotationAfter": "menu to record an event, intervention or explanation.",

"help.6.menuHeading": "Discover the right-click menu",
"help.6.menuIntro": "Right-click directly on a chart to discover additional options. Available actions depend on the chart type and where you click.",
"help.6.menuAnnotations": "Add, edit or delete annotations.",
"help.6.menuSplits": "Add or remove split periods.",
"help.6.menuExclusions": "Exclude or reinclude investigated observations from control-limit calculations.",
"help.6.menuCentre": "Show, hide or restore centre-line labels.",
"help.6.menuExport": "Copy the chart, copy its analysis, or export an image.",

"help.6.excludeHeading": "Important – excluding data points:",
"help.6.excludeAdvice": "Only exclude an observation from limit calculations after investigating and understanding the reason it differs, and deciding it does not represent routine process variation. Excluded observations remain visible and can be reincluded. Consider annotating the reason for exclusion.",
"help.6.subgroups": "On X̄–S charts, exclusion applies to an entire subgroup. Pareto charts do not use control-limit splits or point exclusions.",
"help.6.splitAdvice": "Similarly, only use a split when there is good reason to believe the process has changed. Splits should reflect investigated changes, not simply remove inconvenient signals.",
"help.6.extraBefore": "Look for the small",
"help.6.extraMiddle": "symbols beside controls throughout SimpleSPC. Hover over them, or focus them with the keyboard, for additional guidance. You can also ask the",
"help.6.extraAfter": "about your chart and its interpretation.",

// HELP SECTIONS 7–11

// 7. Interpreting the summary
"help.7.title": "7. Interpreting the summary",
"help.7.intro": "The summary underneath the chart helps explain what the tool has found.",
"help.7.commonName": "Common-cause variation",
"help.7.commonDesc": "means the process looks stable within its current level of performance.",
"help.7.specialName": "Special-cause variation",
"help.7.specialDesc": "means the pattern suggests something may have changed.",
"help.7.stableBefore": "A process can be",
"help.7.stableHighlight": "stable but still poor",
"help.7.stableAfter": "if the level of performance is consistently below what you need.",
"help.7.advice": "Look for sustained changes, not just one good or bad point. Use local knowledge to interpret what was happening at that time.",

// 8. Saving and reopening your work
"help.8.title": "8. Saving and reopening your work",
"help.8.intro": "You can save your work in different ways, depending on how you want to use it later.",
"help.8.saveDesc": "stores the current data and chart settings together in a .json project file.",
"help.8.openDesc": "restores a previously saved project file so you can continue working on it later.",
"help.8.portableBefore": "creates a self-contained",
"help.8.portableAfter": "file containing the SPC tool, your data and your chart settings. The saved file can be opened directly in a web browser and used again without an internet connection.",

"help.8.preservedHeading": "Saved charts can preserve things such as:",
"help.8.data": "the underlying chart data",
"help.8.columns": "selected columns and X-axis type",
"help.8.type": "chart type",
"help.8.labels": "titles, labels and axis settings",
"help.8.rules": "rules and targets",
"help.8.baseline": "baseline settings",
"help.8.splits": "splits and annotations",

"help.8.editableBefore": "A portable HTML file remains editable. After opening it, you can change the data or chart settings and use",
"help.8.editableAfter": "again to create an updated copy.",
"help.8.privacy": "Both JSON project files and portable HTML files contain the data used to create the chart. Store and share saved files in line with your organisation's information-governance, confidentiality and cybersecurity requirements.",

// 9. Copying and downloading charts
"help.9.title": "9. Copying and downloading charts",
"help.9.context": "Right-click the chart to open the menu for exporting the chart as an image.",
"help.9.appearance": "Use the chart options to adjust the chart's appearance before exporting.",
"help.9.clipboard": "You can copy the chart(s) and/or analysis to your clipboard for easy pasting into a report, presentation or document.",
"help.9.secondaryBefore": "Use the",
"help.9.secondaryName": "Show/hide MR/S chart",
"help.9.secondaryMiddle": "option in the",
"help.9.secondaryAfter": "tab to control whether the secondary chart is included in copied or downloaded output.",

// 10. Data privacy
"help.10.title": "10. Data privacy",
"help.10.intro": "The tool runs entirely in your browser:",
"help.10.noServer": "data are not sent to a server",
"help.10.localFiles": "uploaded files are processed locally in your browser session",
"help.10.unsavedBefore": "Unsaved work may be lost when you close or refresh the page. Use",
"help.10.unsavedMiddle": "(JSON) or",
"help.10.unsavedAfter": "(HTML) to keep a copy of your work.",
"help.10.caution": "SimpleSPC is not designed for patient-identifiable or other sensitive personal data. Use anonymised or appropriately aggregated data, and follow your organisation's information-governance, confidentiality and cybersecurity requirements. Remember that saved JSON and portable HTML files contain the data used to create the chart.",

// 11. About the tool
"help.11.title": "11. About the tool",
"help.11.authorBefore": "This tool was developed by Jem Ramazanoglu –",
"help.11.authorAfter": "as a free resource to support Quality Improvement in the NHS.",
"help.11.feedback": "Feedback and suggestions are very welcome.",

// OFFLINE SPC HELPER INTERFACE
"helper.title": "SPC helper",
"helper.close": "Close SPC helper",
"helper.closeTitle": "Close",
"helper.hideQuestions": "Hide suggested questions",
"helper.showQuestions": "Show suggested questions",
"helper.generalHeading": "General SPC questions",
"helper.chartHeading": "My chart questions",
"helper.generalAria": "General SPC suggested questions",
"helper.chartAria": "My chart suggested questions",
"helper.askPlaceholder": "Ask a question…",
"helper.askButton": "Ask",

"helper.introGeneralTitle": "SPC helper",
"helper.introGeneral": "Ask a general question before you load any data, or use a suggested prompt below.",
"helper.introChartTitle": "Chart helper",
"helper.introChart": "Use the My chart questions for a tailored interpretation.",
"helper.emptyQuestion": "Type a question (or click a suggestion) to get started.",

// GENERAL SUGGESTED QUESTIONS
"helper.q.whatSpc": "What is SPC?",
"helper.q.commonSpecial": "What is the difference between common and special cause variation?",
"helper.q.chooseChart": "How do I choose the right chart?",
"helper.q.run": "What is a run chart?",
"helper.q.xmr": "What is an XmR chart?",
"helper.q.stable": "What does stable mean?",
"helper.q.baselineSplit": "What is the difference between a baseline and a split?",
"helper.q.limits": "How do control limits work?",
"helper.q.signalAction": "What should I do if I see a signal?",
"helper.q.stablePoor": "What if my process is stable but performance is poor?",
"helper.q.rulesUnavailable": "Why are some SPC rules unavailable?",
"helper.q.rareRules": "Why are T/G run and trend rules off by default?",
"helper.q.zoneRules": "Why can’t I use zone rules on P, P′, U, U′, or C charts?",

// CHART-SPECIFIC SUGGESTED QUESTIONS
"helper.q.myChart": "What is my chart telling me?",
"helper.q.changed": "Has anything changed?",
"helper.q.processStable": "Is my process stable?",
"helper.q.improvement": "Is this likely to be real improvement?",
"helper.q.target": "What does this mean for our target?",
"helper.q.next": "What should I do next?",
"helper.q.addSplit": "Should I add a split here?",
"helper.q.xmrLimits": "Why are my XmR limits different?",
"helper.q.pPrime": "Why am I using a P′ chart?",
"helper.q.uPrime": "Why am I using a U′ chart?",
"helper.q.sigmaZ": "What does Sigma Z mean?",
"helper.q.exclude": "When should I exclude a point from limit calculations?",

// PARETO SUGGESTED QUESTIONS
"helper.q.pareto80": "What does the 80% line mean?",
"helper.q.paretoCauses": "Are the biggest categories the causes?",

  // GENERAL SPC HELPER ANSWERS

  "helper.faq.what-is-spc":
    "Statistical Process Control (SPC) helps you use data over time to understand whether a process is behaving as usual, or whether something may have changed.\n\n" +
    "In healthcare improvement, that helps teams distinguish routine ups and downs from signals that may reflect a real change in the service, pathway, staffing, demand, coding, or measurement system.",

  "helper.faq.common-vs-special":
    "Common-cause variation is the routine fluctuation built into the current system. Special-cause variation means the pattern suggests something unusual may have happened.\n\n" +
    "In practice:\n" +
    "• common cause = the system behaving as it usually does\n" +
    "• special cause = a prompt to investigate what may have changed",

  "helper.faq.choose-chart":
    "Choose the chart based on the type of data you have:\n\n" +
    "• Run chart: simple data over time, often a good starting point\n" +
    "• XmR chart: continuous data with one value per time point\n" +
    "• C chart: counts with roughly constant opportunity\n" +
    "• P chart: proportions with a numerator and denominator\n" +
    "• P′ chart: proportions where the variation is more or less than a standard P chart expects\n" +
    "• U chart: rates with changing opportunity\n" +
    "• U′ chart: rates where the variation is more or less than a standard U chart expects\n" +
    "• X̄–S chart: subgrouped continuous data\n" +
    "• T / G charts: rare events\n\n" +
    "• Pareto chart: ranks categories by frequency to help focus investigation; it supplements rather than replaces a time-ordered SPC chart\n\n" +
    "For proportion data, SimpleSPC normally starts with a P chart and checks whether P′ better matches the observed variation. For rate data with changing opportunity, it normally starts with a U chart and checks whether U′ better matches the observed variation.",

  "helper.faq.run-chart":
    "A run chart shows data over time with a centre line, usually the median. It is a simple way to see whether the pattern looks broadly stable or whether there may be a shift or trend.\n\n" +
    "It is often a very good starting point for healthcare improvement teams.",

  "helper.faq.xmr-chart":
    "An XmR chart is used for continuous data when you have one observation per time point.\n\n" +
    "It usually has:\n" +
    "• an Individuals (X) chart showing the values over time\n" +
    "• a Moving Range (MR) chart showing how much consecutive points differ\n\n" +
    "It helps separate routine variation from possible special-cause signals.",

  "helper.faq.stable-meaning":
    "A stable process shows only routine common-cause variation. In simple terms, it is behaving predictably within its current level of performance.\n\n" +
    "Stable does not necessarily mean good. A stable process can still be consistently underperforming.",

  "helper.faq.baseline-vs-split":
    "A baseline fixes the calculation using an earlier reference period. A split recalculates the centre line and limits from a chosen point onward.\n\n" +
    "In simple terms:\n" +
    "• baseline = compare later performance against the old system\n" +
    "• split = treat the later period as a new system or new normal",

  "helper.faq.control-limits":
    "Control limits show the range you would expect from routine common-cause variation. They are not the same as targets, standards, or pass/fail thresholds.\n\n" +
    "A point outside the limits suggests something may have changed and is worth investigating. Some chart types also support selected pattern rules, but the tool only offers those where they are reasonably defensible.",

  "helper.faq.what-to-do-signal":
    "A special-cause signal is usually a prompt to investigate what changed around that time.\n\n" +
    "In a healthcare service, that might include a pathway change, staffing change, demand pressure, case-mix change, coding change, data-definition change, equipment issue, or an intentional improvement intervention.\n\n" +
    "The aim is not to blame a person or overreact to one point. The aim is to understand whether something in the system changed, whether that change is real, and whether it is likely to continue.",

  "helper.faq.stable-but-poor":
    "If the process is stable but the level of performance is not good enough, that usually means the system is delivering a predictable but unsatisfactory result.\n\n" +
    "In improvement terms, that often points toward redesigning or improving the system rather than reacting to individual good or bad days.",

  "helper.faq.rule-availability":
    "Some SPC rules are only suitable for certain chart types.\n\n" +
    "This tool uses a conservative, chart-aware rule policy. That means it shows safer default rules first and blocks or hides rules that could create misleading alerts.\n\n" +
    "This is especially important for healthcare and public-service use, where false signals can lead to wasted effort, unnecessary concern, or alert fatigue.",

  "helper.faq.tg-run-trend-off":
    "T and G charts are rare-event charts. They can naturally look irregular even when the process has not really changed.\n\n" +
    "For that reason, this tool keeps run and trend rules off by default for T/G charts. You can enable them in advanced settings, but they should be interpreted cautiously because they may create false signals.",

  "helper.faq.puc-zone-rules-blocked":
    "Zone rules are not available for P, P′, U, U′, or C charts because they can create misleading alerts on these chart types.\n\n" +
    "These charts are based on count, proportion, or rate data. Their limits may vary, and their data may not behave like the continuous data that classic zone rules were designed for.\n\n" +
    "The safer default is to use beyond-limit signals and run rules, rather than adding extra zone-rule alerts.",


  // REMAINING FIXED SPC HELPER FAQ ANSWERS

  // XmR calculations and manual exclusions
  "helper.faq.xmr-moving-range-screening":
    "XmR control limits estimate routine variation from the differences between consecutive points, called moving ranges.\n\n" +
    "A very large point-to-point jump can inflate the average moving range and make the control limits unusually wide. By default, SimpleSPC screens unusually large moving ranges when estimating routine variation. The original observations and moving ranges are still shown on the chart.\n\n" +
    "SimpleSPC uses the Nelson method: moving ranges greater than 3.5 times the initial average moving range are not used in the variation estimate. You can turn this screening off in Advanced detection options if you want the unscreened calculation.",

  "helper.faq.exclude-point-from-limits":
    "Exclude a point from the limit calculations only after you have investigated it and have good reason to believe it reflects a known special cause rather than the routine variation of the process.\n\n" +
    "The point is not deleted. It stays visible on the chart, but SimpleSPC does not use it to estimate the centre line or control limits. Do not exclude a point simply because it is high, low, or outside a control limit. First investigate what happened.\n\n" +
    "If the change represents a lasting new way of working or a new process level, a split is usually more appropriate than excluding the point.\n\n" +
    "For XmR charts, excluding an observation also leaves out the moving ranges immediately before and after that observation when routine variation is estimated. SimpleSPC does not create an artificial moving range across the excluded point. For X̄–S charts, exclusion applies to the whole plotted subgroup.\n\n" +
    "Technical note: some SPC texts call this ghosting or causing out a point. It is a way of estimating the process's routine common-cause variation — sometimes called the voice of the process — without allowing an investigated special cause to distort the limits.",

  // Targets and baselines
  "helper.faq.target":
    "A target is the level of performance you are aiming for. In SPC, it should be interpreted alongside stability.\n\n" +
    "A stable process below target often means the system needs redesign or improvement. An unstable process meeting target intermittently may still not be reliable.",

  "helper.faq.baseline":
    "A baseline is the initial set of points used to calculate the centre line and, where relevant, the limits.\n\n" +
    "It is useful when you want to compare later performance against an earlier period that represents the original system.",

  // Data quality and interpretation
  "helper.faq.data-definition-change":
    "A change in coding, inclusion criteria, denominator rules, or case definition can create a chart signal even if the real service has not changed.\n\n" +
    "That is why SPC charts should always be interpreted alongside local knowledge about data collection and operational context.",

  "helper.faq.annotations":
    "Use annotations to mark events that could plausibly explain a change in the pattern, such as a new pathway, staffing change, policy change, or data-definition change.\n\n" +
    "Keep them short and factual. They support interpretation, but they do not prove causation on their own.",

  "helper.faq.improvement-vs-one-good-point":
    "In SPC, improvement is more convincing when the chart shows a sustained change in the desired direction, not just one good point.\n\n" +
    "A single good result can happen by chance. A sustained shift or other clear signal is more consistent with real change.",

  // Chart centre-line explanations
  "helper.faq.runchart-median":
    "Run charts often use the median because it is simple and robust, especially when you are starting improvement work or have limited data.\n\n" +
    "It gives a practical centre line for judging runs and trends.",

  "helper.faq.xmr-mean":
    "An XmR chart uses the mean together with the moving range to estimate routine variation and calculate control limits.\n\n" +
    "That gives stronger statistical signals than a simple run chart when the data suit that chart type.",

  // Splits, targets and causation
  "helper.faq.should-add-split":
    "Usually only add a split when there is good evidence of a real and sustained process change, and you expect the new pattern to continue.\n\n" +
    "A split should help interpretation, not hide instability or make the chart look better.",

  "helper.faq.above-target-solved":
    "Not necessarily. Crossing the target once is not the same as showing a stable, sustained level of performance.\n\n" +
    "A better question is whether the process now looks predictably at or beyond the target, and whether that change is likely to last.",

  "helper.faq.below-target-stable":
    "This usually means the system is delivering a consistent result, but not a good enough one.\n\n" +
    "That often points toward improvement or redesign work rather than reacting to individual points.",

  "helper.faq.caused-improvement":
    "Be cautious. A chart can show that performance changed, and annotations can help link timing to an intervention, but the chart alone does not prove causation.\n\n" +
    "It is usually better to say the change is consistent with the intervention, supported by local knowledge.",

  // Case mix, denominators and rare events
  "helper.faq.case-mix":
    "A real change in the type or complexity of patients can affect the chart, even if the service process stayed the same.\n\n" +
    "That does not make the chart wrong, but it does affect interpretation.",

  "helper.faq.denominator-changed":
    "If the amount of opportunity changes a lot over time, that affects how the chart should be set up and interpreted.\n\n" +
    "That is one reason P and U charts are useful: they allow for changing totals or opportunities.",

  "helper.faq.rare-event-improvement":
    "For rare-event charts, improvement often means longer times between unwanted events or more opportunities between events.\n\n" +
    "Because rare-event data can look dramatic, it helps to focus on sustained pattern changes rather than one unusually long gap.",

// XmR HELPER: SIGNAL DESCRIPTIONS
"helper.xmr.signal.above": "one or more points above the upper limit",
"helper.xmr.signal.below": "one or more points below the lower limit",
"helper.xmr.signal.run": "a sustained run on one side of the centre line",
"helper.xmr.signal.trend": "a sustained trend",
"helper.xmr.signal.outlier": "an unusual outlier",
"helper.xmr.signal.cluster": "an unusual cluster of points far from the centre line",
"helper.xmr.signal.bias": "a sustained pattern away from the centre line",
"helper.xmr.signal.other": "special-cause signals",
"helper.xmr.signal.joinTwo": "{first} and {last}",
"helper.xmr.signal.joinMany": "{first}, and {last}",

// PERIOD-BY-PERIOD EXPLANATIONS
"helper.xmr.period.label": "Period {number}",
"helper.xmr.period.mean": " around a mean of {mean}",
"helper.xmr.period.count": "Only {count} points contributed to the limit estimate, so ",
"helper.xmr.period.provisional": "{label} (points {start}–{end}) has provisional XmR limits{mean}. {count}At least {minimum} usable points are required before SimpleSPC interprets stability and special-cause signals for this period",
"helper.xmr.period.stable": "{label} (points {start}–{end}) appears stable{mean}",
"helper.xmr.period.signals": "{label} (points {start}–{end}) shows {signals}{mean}",

// MULTIPLE PERIODS
"helper.xmr.multi.earlier": "Across the XmR chart, at least one assessable period shows special-cause signals. This suggests the process may have changed.\n\n",
"helper.xmr.multi.splits": "This XmR chart has multiple periods separated by splits. Each split starts a new calculation of the mean and limits, but the presence of a split does not itself prove a process change.\n\n",
"helper.xmr.multi.count": "There are currently {count} usable points in this period, and ",
"helper.xmr.multi.provisional": "The latest period (points {start}–{end}) does not yet have enough data for a firm stability conclusion. {count}SimpleSPC requires at least {minimum} usable points before interpreting XmR stability.",
"helper.xmr.multi.warning": "\n\nSome early warning features are present, but this period is still provisional. More observations are needed to see whether the pattern persists.",
"helper.xmr.multi.calm": "\n\nNo obvious early warning features appear in this short period. That is encouraging, but does not yet establish stability.",
"helper.xmr.multi.stable": "In the latest period (points {start}–{end}), no clear special-cause signal is detected. With sufficient data, this is consistent with routine variation around the current level.",
"helper.xmr.multi.signals": "In the latest period (points {start}–{end}), possible special-cause variation remains: {signals}.",

// SINGLE PERIOD
"helper.xmr.single.noChart": "Generate an XmR chart first, then ask again.",
"helper.xmr.single.count": "There are currently {count} usable points for limit estimation, and ",
"helper.xmr.single.provisional": "There are not yet enough data to make a firm judgement about stability. {count}SimpleSPC requires at least {minimum} usable points for a formal XmR interpretation.",
"helper.xmr.single.warning": "\n\nThere are features worth watching, but more observations are needed to determine whether they form a sustained signal.",
"helper.xmr.single.calm": "\n\nThe pattern currently looks reasonably consistent, with no obvious early warnings. This remains provisional until enough observations are available.",
"helper.xmr.single.stable": "The current XmR chart has enough data for interpretation and shows no clear special-cause signals. This is consistent with routine variation, although stable performance is not necessarily satisfactory performance.",
"helper.xmr.single.signals": "The current XmR chart suggests that the process may have changed. The pattern is not fully explained by routine variation, so investigate what happened in the process or its measurement.",

  //  HELPER STABILITY ASSESSMENTS

  // Chart names for dynamic helper sentences
  "helper.stability.chart.c": "c chart",
  "helper.stability.chart.p": "p chart",
  "helper.stability.chart.pprime": "p′ chart",
  "helper.stability.chart.u": "u chart",
  "helper.stability.chart.uprime": "u′ chart",
  "helper.stability.chart.t": "t chart",
  "helper.stability.chart.g": "g chart",

  "helper.stability.signalFallback": "one or more SPC signals",

  // Pareto: stability is not assessed
  "helper.stability.pareto":
    "A Pareto chart does not assess process stability because the categories are not plotted in time order.",

  // XmR: period identification
  "helper.stability.xmr.latest":
    "Looking at the latest period (Period {period} of {total}, points {start}–{end})",

  "helper.stability.xmr.single":
    "Looking at the chart (points {start}–{end})",

  // XmR: provisional limits
  "helper.stability.xmr.count":
    "There are currently {count} points in this period, and ",

  "helper.stability.xmr.provisional":
    "{scope}, there are not yet enough data to make a firm judgement about stability. {count}SimpleSPC uses at least {minimum} points before making a formal XmR stability interpretation.",

  "helper.stability.xmr.warning":
    "\n\nThere are already some features in this short period that are worth watching, so I would not yet describe the process as stable. More data will help show whether these are temporary fluctuations or part of a real pattern.",

  "helper.stability.xmr.calm":
    "\n\nSo far, the pattern looks fairly consistent and there are no obvious warning signs in this short period. That is encouraging, but it is still provisional rather than proof of stability.",

  // XmR: assessable periods
  "helper.stability.xmr.stable":
    "{scope}, I cannot see a clear special-cause signal using the current rule settings. This suggests the process is behaving broadly consistently around its current level.",

  "helper.stability.xmr.signals":
    "{scope}, this XmR chart shows possible special-cause variation — specifically {signals}.",

  // X-bar and S stability combinations
  "helper.stability.xbars.bothStable":
    "Looking at the latest X̄–S period, both the subgroup averages and the within-subgroup variation look stable. I cannot see a clear special-cause signal in either part of the chart.",

  "helper.stability.xbars.xChanged":
    "Looking at the latest X̄–S period, the subgroup averages show possible change, but the within-subgroup variation looks stable.",

  "helper.stability.xbars.sChanged":
    "Looking at the latest X̄–S period, the subgroup averages look stable, but the within-subgroup variation shows possible change.",

  "helper.stability.xbars.bothChanged":
    "Looking at the latest X̄–S period, both the subgroup averages and the within-subgroup variation show possible change.",

  // Attribute charts
  "helper.stability.attribute.stable":
    "Looking at the latest {chart} period, I cannot see a clear special-cause signal using the current rule settings.",

  "helper.stability.attribute.signals":
    "Looking at the latest {chart} period, there are signs of possible special-cause variation — specifically {signals}.",

  // Rare-event charts
  "helper.stability.rare.stable":
    "Looking at the latest {chart} period, I cannot see a clear special-cause signal using the current rare-event rule settings.",

  "helper.stability.rare.signals":
    "Looking at the latest {chart} period, there are signs of possible special-cause variation — specifically {signals}.",

  "helper.stability.noChart":
    "Generate the currently selected chart first and I can give you a chart-specific interpretation.",

  // Run charts, including phase splits
  "helper.stability.run.noChart":
    "Generate a Run chart first so I can interpret it.",

  "helper.stability.run.splitIntro":
    "This Run chart has been divided into {count} periods using splits. Each period has its own median, and SimpleSPC checks for shifts and trends separately within each period.\n\n",

  "helper.stability.run.splitStable":
    "No shift or trend has been detected within the individual periods using the current rule settings. However, this does not mean that the overall process has remained unchanged.\n\n",

  "helper.stability.run.splitSignals":
    "One or more periods contain a shift or trend signal that may warrant investigation.\n\n",

  "helper.stability.run.splitAdvice":
    "Compare the medians reported for the different periods. A difference in their levels may indicate an important change in performance, even if there are no further signals within each period.\n\nA split is a boundary chosen for analysis; it does not by itself prove a statistically significant change. Consider why the split was introduced, what happened in the process around that time, and whether the new level is sustained.",

  "helper.stability.run.stable":
    "Looking at the current Run chart, no shift or trend has been detected using the selected run-chart rules.",

  "helper.stability.run.signals":
    "The current Run chart shows one or more shift or trend signals, suggesting that the pattern may have changed. Investigate what happened in the process around that time.",

  // DYNAMIC HELPER SIGNALS
  "helper.signal.shiftMean": "a sustained shift (many points on the same side of the mean)",
  "helper.signal.shiftDetailed": "Shift: {count}+ points in a row {direction} the centre line (from {start} to {end})",
  "helper.signal.trendDetailed": "Trend: {count}+ points steadily {direction} (from {start} to {end})",
  "helper.signal.sideAbove": "above",
  "helper.signal.sideBelow": "below",
  "helper.signal.directionIncreasing": "increasing",
  "helper.signal.directionDecreasing": "decreasing",


  //  DYNAMIC TARGET AND IMPROVEMENT ANSWERS

  // Additional chart names
  "helper.stability.chart.run": "run chart",
  "helper.stability.chart.xmr": "XmR chart",
  "helper.stability.chart.xbars": "X̄–S chart",
  "helper.stability.chart.pareto": "Pareto chart",

  // XmR: target not configured
  "helper.target.xmr.none":
    "There is no target currently set on this XmR chart. If you add one, SimpleSPC can help interpret performance against the target alongside process stability.",

  // XmR: comparison of mean and target
  "helper.target.xmr.meanFavourable":
    "The current period mean ({mean}) is on the favourable side of the target ({target}).",
  "helper.target.xmr.meanUnfavourable":
    "The current period mean ({mean}) is not yet on the favourable side of the target ({target}).",

  // XmR: provisional limits
  "helper.target.xmr.provisional":
    "The latest XmR period is still provisional, so it is too early to make a firm judgement about whether the process can reliably meet the target.",
  "helper.target.xmr.count":
    "There are currently {count} points in this period, and ",
  "helper.target.xmr.minimum":
    "{count}SimpleSPC uses at least {minimum} points before making a formal stability or capability interpretation.",
  "helper.target.xmr.earlyFavourable":
    "The early results are encouraging, but they do not yet show that this level of performance is predictable or sustainable.",
  "helper.target.xmr.earlyUnfavourable":
    "The early pattern may still be settling. More data will show whether the process level is genuinely different and whether further improvement is needed to reach the target reliably.",

  // XmR: stable and special-cause periods
  "helper.target.xmr.stable":
    "The latest XmR period has enough data for interpretation and currently looks stable.",
  "helper.target.xmr.capability":
    "Based on the current stable process, SimpleSPC estimates that about {percentage}% of future results would meet the target, assuming the process continues to behave in the same way.",
  "helper.target.xmr.stableNoCapability":
    "Because the process appears stable, the target comparison is more meaningful: the main question is whether the current system is predictably good enough.",
  "helper.target.xmr.unstable":
    "The latest XmR period has enough data for interpretation, but it shows special-cause variation. That means target performance should be interpreted cautiously because the process is not yet behaving predictably.",
  "helper.target.xmr.unstableAction":
    "The priority is to understand what is changing in the system, demand, staffing, pathway, coding, or measurement before relying on a capability estimate.",

  // Other charts: target assessment
  "helper.target.general.none":
    "There is no target currently set on this {chart}. If you add one, SimpleSPC can help you interpret the target alongside the pattern of variation.",
  "helper.target.general.higher":
    "higher values are treated as better",
  "helper.target.general.lower":
    "lower values are treated as better",
  "helper.target.general.intro":
    "The target is {target}, and {direction}.",
  "helper.target.general.centreFavourable":
    "The current centre line ({centre}) is on the favourable side of the target.",
  "helper.target.general.centreUnfavourable":
    "The current centre line ({centre}) is not yet on the favourable side of the target.",
  "helper.target.general.stable":
    "The latest period does not show a clear special-cause signal, so the comparison with the target is more meaningful. Remember that being stable does not automatically mean performance is good enough.",
  "helper.target.general.unstable":
    "The latest period shows possible special-cause variation, so target performance should be interpreted cautiously. The priority is to understand what may have changed before assuming the current level will continue.",
  "helper.target.general.unknown":
    "Interpret the target alongside the pattern of variation rather than judging individual points on their own.",

  // Improvement: XmR
  "helper.improvement.xmr.provisional":
    "There are not yet enough data in the latest XmR period to make a firm judgement about stability or improvement. The early pattern may be encouraging, but more points are needed before calling it a sustained change.",
  "helper.improvement.xmr.stable":
    "The current XmR period does not show a clear special-cause signal, so the chart alone does not provide strong evidence of a new improvement. Real improvement is more convincing when there is a sustained change in the desired direction and the new pattern is maintained.",
  "helper.improvement.xmr.unstable":
    "The XmR chart shows evidence that the process may have changed. That is not automatically the same as improvement. Check whether the change is in the desired direction, whether it is sustained, and whether local knowledge supports a plausible improvement explanation. The chart can show evidence of change, but it cannot prove that an intervention caused it.",

  // Improvement: other charts
  "helper.improvement.general.stable":
    "The current {chart} does not show a clear special-cause signal, so the chart alone does not provide strong evidence of a new improvement. Improvement is more convincing when there is a sustained change in the desired direction, not just one favourable point.",
  "helper.improvement.general.unstable":
    "The current {chart} shows evidence that the process may have changed. That could be improvement, deterioration, or another change in the system. To call it improvement, check that the change is in the desired direction, is sustained, and makes sense in the local context. The chart alone does not prove causation.",
  "helper.improvement.general.noChart":
    "Generate the currently selected chart first and I can help assess whether the pattern is consistent with real improvement.",

  // XmR CONTROL LIMITS AND NELSON SCREENING

  "helper.xmrLimits.noChart":
    "XmR limits are estimated from the moving ranges between consecutive points. Generate an XmR chart and I can explain the calculation used for your current data.",

  "helper.xmrLimits.intro":
    "XmR control limits depend on the estimated routine variation. SimpleSPC estimates this from the moving ranges between consecutive observations.",

  "helper.xmrLimits.screeningOn":
    "\n\nFor your current chart, moving-range screening is ON.",

  "helper.xmrLimits.screenedOne":
    " SimpleSPC found {count} unusually large point-to-point change and did not use it when estimating routine variation. The original data points and moving ranges remain on the chart.",

  "helper.xmrLimits.screenedMany":
    " SimpleSPC found {count} unusually large point-to-point changes and did not use them when estimating routine variation. The original data points and moving ranges remain on the chart.",

  "helper.xmrLimits.whyScreen":
    "\n\nThis prevents an unusually large jump from making the estimated variation — and therefore the XmR control limits — artificially wide.",

  "helper.xmrLimits.noneScreened":
    " No moving ranges in the current period exceeded the screening threshold, so screening has not changed the variation estimate for this period.",

  "helper.xmrLimits.statistics":
    "\n\nFor the current period:\n" +
    "• initial average MR = {initial}\n" +
    "• Nelson screening threshold = {threshold}\n" +
    "• average MR used for the limits = {average}",

  "helper.xmrLimits.screeningOff":
    "\n\nFor your current chart, moving-range screening is OFF. All baseline moving ranges are therefore used to estimate routine variation. If one point-to-point jump is unusually large, this can increase the average moving range and produce wider control limits.",

  "helper.xmrLimits.offAverage":
    "\n\nThe average MR currently being used for the limits is {average}.",

  "helper.xmrLimits.unknown":
    "\n\nI cannot determine the moving-range screening setting from the current chart snapshot.",

  "helper.xmrLimits.splits":
    "\n\nThis chart also contains a split. Each period is treated as a separate process period with its own mean and limits, and the moving range across the split is not used. This can also make the limits different from an unsplit version of the same chart.",

  "helper.xmrLimits.provisionalIntro":
    "\n\nThe latest period is still provisional. ",

  "helper.xmrLimits.provisionalCount":
    "Only {count} points are currently being used to estimate its limits. ",

  "helper.xmrLimits.provisionalMinimum":
    "SimpleSPC waits until at least {minimum} points are available before making a formal stability interpretation.",

// LANEY P-PRIME, U-PRIME AND SIGMA Z

// P′ explanations
"helper.laney.pprime.noChart":
  "A P′ chart is a Laney-adjusted version of a P chart. It is used for proportion data when the amount of variation is noticeably different from what a standard P chart expects.",

"helper.laney.pprime.intro":
  "This P′ chart uses the same numerator, denominator, observed proportions and centre line as a standard P chart. The difference is that it adjusts the width of the control limits so they better match the amount of variation actually seen in the data.",

"helper.laney.pprime.more":
  "\n\nYour data vary more than a standard P chart expects, so P′ uses wider control limits. This helps avoid treating ordinary variation as a possible process change.",

"helper.laney.pprime.less":
  "\n\nYour data vary less than a standard P chart expects, so P′ uses narrower control limits. This can make genuinely unusual variation easier to detect.",

"helper.laney.pprime.equal":
  "\n\nThe amount of variation is very close to what a standard P chart expects, so the P′ adjustment is very small.",

"helper.laney.pprime.factor":
  "\n\nThe Sigma Z adjustment factor for this chart is {sigmaZ}. A value above 1 widens the standard P-chart limits; a value below 1 narrows them.",

// U′ explanations
"helper.laney.uprime.noChart":
  "A U′ chart is a Laney-adjusted version of a U chart. It is used for rate data when the amount of variation is noticeably different from what a standard U chart expects.",

"helper.laney.uprime.intro":
  "This U′ chart uses the same event counts, opportunities, observed rates and centre line as a standard U chart. The difference is that it adjusts the width of the control limits so they better match the amount of variation actually seen in the data.",

"helper.laney.uprime.more":
  "\n\nYour data vary more than a standard U chart expects, so U′ uses wider control limits. This helps avoid treating ordinary variation as a possible process change.",

"helper.laney.uprime.less":
  "\n\nYour data vary less than a standard U chart expects, so U′ uses narrower control limits. This can make genuinely unusual variation easier to detect.",

"helper.laney.uprime.equal":
  "\n\nThe amount of variation is very close to what a standard U chart expects, so the U′ adjustment is very small.",

"helper.laney.uprime.factor":
  "\n\nThe Sigma Z adjustment factor for this chart is {sigmaZ}. A value above 1 widens the standard U-chart limits; a value below 1 narrows them.",

// Chart names for technical explanations
"helper.laney.chart.p": "P chart",
"helper.laney.chart.u": "U chart",
"helper.laney.chart.pprime": "P′ chart",
"helper.laney.chart.uprime": "U′ chart",

// Sigma Z explanations
"helper.laney.sigma.general":
  "Sigma Z is used on Laney P′ and U′ charts. It is an adjustment factor that tells SimpleSPC whether the data are varying more or less than a standard P or U chart would normally expect.",

"helper.laney.sigma.intro":
  "Sigma Z is the adjustment factor used by a Laney chart. It tells SimpleSPC whether your data are varying more or less than a standard {standard} expects.",

"helper.laney.sigma.guidance":
  "\n\nA Sigma Z of about 1 means the amount of variation is close to what the standard chart expects. A value above 1 means there is more variation than expected, so the control limits are made wider. A value below 1 means there is less variation than expected, so the control limits are made narrower.\n\n",

"helper.laney.sigma.unchanged":
  "Sigma Z does not change your original observations or the centre line. It only adjusts the width of the control limits on the {prime}.",

"helper.laney.sigma.value":
  "\n\nFor your current chart, Sigma Z = {sigmaZ}.",

"helper.laney.sigma.more":
  " This means the data are showing more variation than the standard {standard} expects, so the {prime} uses wider limits.",

"helper.laney.sigma.less":
  " This means the data are showing less variation than the standard {standard} expects, so the {prime} uses narrower limits.",

"helper.laney.sigma.equal":
  " This is exactly 1, so the Laney adjustment does not change the width of the standard {standard} limits.",

"helper.laney.sigma.technical":
  "\n\nTechnical detail: SimpleSPC first allows for the amount of variation the standard chart expects at each point. It then looks at the point-to-point changes in those standardised values. Sigma Z summarises that extra or reduced variation into a single adjustment factor.",

  // COLUMN GUIDANCE AND NEXT STEPS

  "chart.columnWarning.c":
    "<strong>Check your selected columns.</strong> For a C chart, the value column should be a count of events/defects per time period. Do not use a denominator column here.",

  "chart.columnWarning.p":
    "<strong>Check your selected columns.</strong> For a P chart, the value column should be the numerator, e.g. number of cases/events, and the third column should be the denominator, e.g. total patients/episodes.",

  "chart.columnWarning.pprime":
    "<strong>Check your selected columns.</strong> For a P′ chart, the value column should be the numerator, e.g. number of cases/events, and the third column should be the denominator, e.g. total patients/episodes. P′ uses the same data structure as a P chart but adjusts the control limits using the Laney method.",

  "chart.columnWarning.u":
    "<strong>Check your selected columns.</strong> For a U chart, the value column should be the numerator, e.g. number of defects/events, and the third column should be the opportunities/denominator, e.g. bed days, attendances, or procedures.",

  "chart.columnWarning.uprime":
    "<strong>Check your selected columns.</strong> For a U′ chart, the value column should be the numerator, e.g. number of defects/events, and the third column should be the opportunities/denominator, e.g. bed days, attendances, procedures or inspections. U′ uses the same data as a U chart but adjusts the control limits using the Laney method.",

  "chart.columnWarning.xbars":
    "<strong>Check your selected columns.</strong> For an X̄–S chart, the value column should contain the individual measurements, and the third column should identify the subgroup, e.g. day, week, batch, or sample.",

  "chart.columnWarning.tGaps":
    "<strong>T chart: using numeric time gaps.</strong><br>Select the column containing the time between consecutive events (e.g. days) as the <strong>Value / Y-axis column</strong>. The X-axis column provides the observation labels.<br><br><strong>Want to use event dates instead?</strong> Click <strong>Chart setup</strong> above and select <strong>I have event dates</strong>.",

  "chart.columnWarning.tDates":
    "<strong>T chart: using event dates.</strong><br>Select the dates or times when events occurred in the <strong>Date / X-axis column</strong>. SimpleSPC calculates the intervals automatically, so the Value / Y-axis column is disabled.<br><br><strong>Already have the time gaps calculated?</strong> Click <strong>Chart setup</strong> above and select <strong>I already have the gaps</strong>.",

  "chart.columnWarning.g":
    "<strong>Check your selected columns.</strong> For a G chart, the value column should usually represent the number of opportunities, cases, or observations between rare events. Check this is not a rate or percentage.",

  "chart.columnWarning.pareto":
    "<strong>Choose how you want to group the occurrences.</strong> Select the category you want to compare. If every row is one occurrence, choose <strong>(count each row as 1)</strong> for the frequency column. If your data are already summarised, choose the numeric count column instead. Pareto categories help focus investigation but should not automatically be assumed to be causes.",

  "helper.decision.general":
    "From a healthcare improvement point of view, the next step is usually:\n" +
    "• if the process looks stable but performance is not good enough, think about redesigning the system\n" +
    "• if the chart shows a signal, investigate what changed in the service, pathway, demand, staffing, coding, or measurement\n" +
    "• avoid reacting to individual points as if each one proves success or failure",

  "helper.decision.paretoNext":
    "\n\nA sensible next step is to investigate the leading categories in more detail, using local knowledge and, where useful, further stratification. Do not assume the category itself is the underlying cause.",

  //PARETO HELPER INTERPRETATIONS

  "helper.pareto.noChart":
    "A Pareto chart ranks categories from most to least frequent. It can help you decide where further investigation or improvement effort may be most useful.",

  "helper.pareto.firstCategory":
    "the first category",

  "helper.pareto.summary":
    "This Pareto chart contains {total} occurrences across {categories} categories. The categories are ordered from most frequent to least frequent.",

  "helper.pareto.topOne":
    "The largest category is {category}, with {count} occurrence{share}.",

  "helper.pareto.topMany":
    "The largest category is {category}, with {count} occurrences{share}.",

  "helper.pareto.topShare":
    " ({percentage}% of the total)",

  "helper.pareto.cutoffOne":
    "The first {count} category accounts for about {percentage}% of the occurrences. That can help you decide where to investigate first, but the 80% line is a prioritisation guide rather than a statistical limit.",

  "helper.pareto.cutoffMany":
    "The first {count} categories account for about {percentage}% of the occurrences. That can help you decide where to investigate first, but the 80% line is a prioritisation guide rather than a statistical limit.",

  "helper.pareto.smallSample":
    "There are fewer than 30 occurrences in this chart, so the ranking may change quite easily as more data are collected.",

  "helper.pareto.causeCaution":
    "Do not assume that the largest categories are necessarily the causes of the problem. They are ways of grouping the data and should normally prompt further investigation.",

  "helper.pareto.notTimeOrdered":
    "A Pareto chart is not time ordered, so it cannot tell you whether a process is stable or whether a change is statistically significant. Use a run or Shewhart chart alongside it when that question matters.",

  // THE 80% REFERENCE LINE

  "helper.pareto.eighty.explain":
    "The dashed 80% line is a guide to help prioritise further investigation. It is not a statistical control limit. The red cumulative line shows how the percentage of occurrences builds up as categories are added from left to right.",

  "helper.pareto.eighty.countOne":
    "In your chart, the first {count} category accounts for {percentage}% of all occurrences.",

  "helper.pareto.eighty.countMany":
    "In your chart, the first {count} categories account for {percentage}% of all occurrences.",

  "helper.pareto.eighty.notRule":
    "The 80% line is a useful reference, not a rule that 80% of problems must come from 20% of categories.",

  // ARE THE BIGGEST CATEGORIES THE CAUSES?

  "helper.pareto.causes":
    "Not necessarily. A Pareto chart shows which categories contain the most recorded occurrences, but it does not establish what caused them.\n\n" +
    "The largest categories are often useful starting points for investigation. Look more closely at what happens within those categories, consider how the data were collected, and use local knowledge to identify possible causes.\n\n" +
    "A frequent category is not automatically a root cause, and a less frequent category may still deserve priority if its consequences are particularly serious.",

  // HAS ANYTHING CHANGED?

  "helper.pareto.changeIntro":
    "A Pareto chart cannot tell whether something has changed over time because it deliberately removes the time order and ranks categories by frequency.\n\n",

  // PARETO STATISTICAL SUMMARY

  "summary.pareto.largestLabel": "Largest category:",
  "summary.pareto.topThreeLabel": "Top three categories:",
  "summary.pareto.eightyHeading": "80% guide:",
  "summary.pareto.cautionLabel": "Caution:",
  "summary.pareto.dataCheckLabel": "Data check:",
  "summary.pareto.interpretationLabel": "Interpretation:",
  "summary.pareto.spcNoteLabel": "SPC note:",

  "summary.pareto.occurrenceOne": "occurrence",
  "summary.pareto.occurrenceMany": "occurrences",
  "summary.pareto.ofTotal": "of the total",

  "summary.pareto.cutoffOneStart": "the first category",
  "summary.pareto.cutoffOneEnd": "accounts for",
  "summary.pareto.cutoffManyStart": "the first",
  "summary.pareto.cutoffManyEnd": "categories account for",
  "summary.pareto.ofOccurrences": "of the occurrences.",

  "summary.pareto.eightyCaution":
    "The dashed 80% line is a prioritisation guide, not a statistical control limit.",

  "summary.pareto.smallSample":
    "this Pareto chart contains fewer than 30 occurrences. With small numbers, the apparent ranking can be unstable, so interpret it cautiously.",

  "summary.pareto.blankOne":
    "row was ignored because the category was blank.",

  "summary.pareto.blankMany":
    "rows were ignored because the category was blank.",

  "summary.pareto.interpretation":
    "the categories on the left are the most frequent and may be useful places to focus further investigation or improvement effort. A category is not automatically a cause of the problem.",

  "summary.pareto.spcNote":
    "a Pareto chart is not time ordered and cannot tell you whether variation is common cause or special cause. Where relevant, interpret it alongside an appropriate run or Shewhart chart.",

  // CHANGE ASSESSMENTS AND SPC RULE POLICY

  // X-bar and S: four existing change outcomes
  "helper.change.xbars.bothStable":
    "I cannot see clear evidence that the process has changed in the latest X̄–S period. Neither the subgroup averages nor the within-subgroup variation shows a special-cause signal.",

  "helper.change.xbars.xOnly":
    "There are signs that the process level may have changed: the X̄ chart shows a signal, while the S chart remains stable.",

  "helper.change.xbars.sOnly":
    "There are signs that the process variation may have changed: the S chart shows a signal, while the X̄ chart remains stable.",

  "helper.change.xbars.bothSignals":
    "There are signs that the process may have changed: both the X̄ and S charts show special-cause signals.",

  // Attribute-chart change assessments
  "helper.change.attribute.stable":
    "I cannot see clear evidence that the process has changed in the latest {chart} period. No special-cause signal is detected using the current rule settings.",

  "helper.change.attribute.signals":
    "There are signs that the process may have changed in the latest {chart} period — specifically {signals}.",

  // Rare-event change assessments
  "helper.change.rare.stable":
    "I cannot see clear evidence that the process has changed in the latest {chart} period. No special-cause signal is detected using the current rule settings.",

  "helper.change.rare.signals":
    "There are signs that the process may have changed in the latest {chart} period — specifically {signals}.",

  "helper.change.investigate":
    "That is a prompt to investigate what changed around that time.",

  "helper.change.noChart":
    "Generate the currently selected chart first and I can look for evidence of change.",

  // Rules applied to each chart family
  "helper.rules.attribute":
    "For this chart, SimpleSPC uses beyond-limit signals and the run rule. Zone rules are blocked because they may create misleading alerts on attribute charts.",

  "helper.rules.rare":
    "For this chart, beyond-limit signals are used by default. Run and trend rules are advanced options only, because rare-event charts can naturally look irregular.",

  "helper.rules.xmr":
    "For this XmR chart, beyond-limit signals and the run rule are used by default. Trend and zone-style rules are advanced options because adding more rules can increase false alerts.",

  "helper.rules.xbars":
    "For this X̄–S chart, beyond-limit signals and the run rule are used by default. Trend and zone rules are available as advanced options.",

  "helper.rules.run":
    "For this run chart, SimpleSPC uses run-chart rules such as shifts and, where enabled, trends. Control-limit rules do not apply to run charts.",

  "helper.rules.pareto":
    "SPC signal rules are not used on a Pareto chart. A Pareto chart ranks categories rather than showing process behaviour over time.",

  "helper.rules.default":
    "SimpleSPC uses a conservative, chart-aware rule set. Not every rule is available for every chart type.",

  // SPC HELPER FALLBACK MESSAGES

  "helper.response.empty":
    "Please type a question about SPC or your chart.",

  "helper.response.unsure":
    "I’m not fully sure which answer fits best. Try one of these:",

  "helper.response.libraryMissing":
    "The SPC helper library is not available. Please check that spc-helper-library.js is loaded before spc.js.",

  // XmR EDUCATIONAL QUESTIONS

  "helper.q.xmrShortSignals":
    "Why do XmR signals disappear with a short baseline?",

  "helper.q.xmrMinimum":
    "Why does XmR need 12 usable points?",

  "helper.q.xmrExclusionDisabled":
    "Why is excluding this point from the limits disabled?",

  "helper.q.xmrNelsonOutside":
    "Why doesn't Nelson screening flag changes after the baseline?",

  // No current XmR chart
  "helper.xmrEducation.noChart":
    "Generate an XmR chart to see information about your current period. The general explanation is below.\n\n",

  // Short baseline and provisional limits
  "helper.xmrEducation.short.intro":
    "SimpleSPC deliberately does not give formal special-cause interpretations when the XmR control limits are provisional. The limits may still be displayed, but an apparent signal is not reported as a confirmed SPC signal until there are enough usable observations in the limit-estimation period.",

  "helper.xmrEducation.short.provisional":
    "Your latest XmR period currently has {count} usable points for estimating the limits; at least {minimum} are required for interpretation. This explains why a signal may disappear when you shorten the baseline, even if later plotted values look unusual.",

  "helper.xmrEducation.short.provisionalUnknown":
    "The latest XmR period does not yet meet the minimum of {minimum} usable points. Its limits and stability interpretation therefore remain provisional.",

  "helper.xmrEducation.short.assessable":
    "Your latest period has {count} usable points and meets the {minimum}-point minimum, so the selected SPC rules can now be interpreted. Not finding a signal does not necessarily mean performance is satisfactory.",

  // Why the threshold is twelve
  "helper.xmrEducation.minimum":
    "SimpleSPC uses a conservative minimum of {minimum} usable points for interpreting XmR limits, special-cause signals and capability. This is a safeguard against treating limits estimated from very little information as sufficiently reliable. It is a rule adopted by this tool, not a guarantee that any process with 12 points is stable.\n\n" +
    "The distinction matters: having {minimum} points somewhere on the chart is not enough. They must contribute to the relevant period's limit estimate. If the initial baseline contains only 10 points, or a split starts a short new period, the limits may remain provisional. Excluded observations are not automatically replaced by later points.",

  "helper.xmrEducation.minimum.current":
    "\n\nFor your current period, {count} usable points contribute to the limit estimate; the minimum is {minimum}.",

  // Disabled manual point exclusion
  "helper.xmrEducation.exclusion":
    "Excluding a point from control-limit calculations is only useful if that point contributes to the limits being estimated. If you fixed the initial baseline, a later point outside that baseline does not contribute to those limits, so excluding it would have no effect. A later period created by a split has its own limit-estimation window.\n\n" +
    "SimpleSPC may also disable a new exclusion if it would leave too few included observations, or — for an XmR chart — no consecutive pair of included observations for estimating moving-range variation.\n\n" +
    "An excluded observation remains visible. Its adjacent moving ranges are omitted from the variation estimate; the tool does not join the observations across the gap to create an artificial moving range.\n\n" +
    "Only exclude an observation after investigating a known special cause. If you believe the process has moved to a lasting new level, consider whether a justified phase split is more appropriate. The right-click menu provides the specific reason when an exclusion is unavailable; this helper answer does not identify which point you clicked.",

  // Nelson screening versus special-cause detection
  "helper.xmrEducation.nelson":
    "Nelson moving-range screening is a method for estimating XmR control limits, not a rule for flagging changes anywhere on the chart.\n\n" +
    "SimpleSPC examines consecutive moving ranges within the period used to estimate the limits. When screening is enabled, moving ranges strictly greater than 3.5 times the initial average moving range are omitted from the variation estimate. The observations and original moving ranges remain visible.\n\n" +
    "A jump after a fixed baseline is not part of that baseline's variation estimate, so Nelson screening will not remove or flag it. Later observations are assessed separately using the selected SPC signal rules, provided the relevant period has enough usable data for interpretation. A split starts a separate period with its own limit estimate.\n\n" +
    "Nelson screening also differs from manually excluding a point: screening omits qualifying moving ranges from the variation estimate; a manual exclusion omits an investigated observation and its adjacent moving ranges.",

  "helper.xmrEducation.nelson.on":
    "\n\nOn your current XmR chart, moving-range screening is enabled.",

  "helper.xmrEducation.nelson.off":
    "\n\nOn your current XmR chart, moving-range screening is disabled.",

  // CHART SELECTION WIZARD

  "wizard.title": "Help me choose a chart",
  "wizard.subtitle": "Answer a couple of questions and the tool will pick a chart type and show the extra columns (if needed).",
  "wizard.close": "Close",
  "wizard.closeAria": "Close chart wizard",

  "wizard.action.back": "Back",
  "wizard.action.skip": "Skip",
  "wizard.action.manual": "Choose manually",
  "wizard.action.use": "Use this chart",

  "wizard.start.question": "What are you charting?",
  "wizard.start.option0": "A measurement — e.g. waiting time, score, temperature, length of stay",
  "wizard.start.option1": "A count per time period — e.g. falls per week, complaints per month",
  "wizard.start.option2": "A proportion out of a total — e.g. 5 out of 100, pass rate",
  "wizard.start.option3": "Categories — e.g. reasons for delays, types of complaints",
  "wizard.start.option4": "Rare events — time or opportunities between events",
  "wizard.start.option5": "Not sure",

  "wizard.unsure.question": "Which description is closest to your data?",
  "wizard.unsure.hint": "The chart type depends on what each value represents. Choose the closest description below. If none fits, you can return to the chart list and select a chart manually.",
  "wizard.unsure.option0": "One measurement at each time point — e.g. waiting time, score, length of stay",
  "wizard.unsure.option1": "A number of events in each time period — e.g. falls, complaints, incidents",
  "wizard.unsure.option2": "A proportion or percentage — e.g. 5 out of 100, compliance rate",
  "wizard.unsure.option3": "Categories of occurrences — e.g. complaint types or reasons for delays",
  "wizard.unsure.option4": "The time or opportunities between rare events",

  "wizard.measurement.question": "Do you have one value per time point, or multiple values?",
  "wizard.measurement.option0": "One value each time point",
  "wizard.measurement.option1": "Multiple values per time point — subgroups or samples",
  "wizard.measurement.option2": "Not sure",

  "wizard.count.question": "Does the amount of work or opportunity vary?",
  "wizard.count.hint": "If activity is broadly similar each time, treat it as roughly constant. If volume changes a lot, or you have a denominator such as bed-days, attendances or inspections, choose varies.<br><br>SimpleSPC will normally start with a U chart and then check whether a Laney U′ chart would better match the amount of variation in your data.",
  "wizard.count.option0": "No — roughly similar volume each time",
  "wizard.count.option1": "Yes — volume varies, or I have a denominator column",
  "wizard.count.option2": "Not sure",

  "wizard.proportion.question": "Do you have both parts of the proportion?",
  "wizard.proportion.hint": "For example, if 8 out of 100 patients met the measure, the numerator is 8 and the denominator is 100.<br><br>SimpleSPC will normally start with a P chart and then check whether a Laney P′ chart would better match the amount of variation in your data.",
  "wizard.proportion.option0": "Yes — I have numerator and denominator columns",
  "wizard.proportion.option1": "No — I only have the percentage or proportion value",
  "wizard.proportion.option2": "Not sure",

  "wizard.pareto.question": "How are your category data recorded?",
  "wizard.pareto.hint": "A Pareto chart compares the frequency of different categories. You can use individual records or data that already contain counts for each category.",
  "wizard.pareto.option0": "One row per occurrence — e.g. one row for each complaint",
  "wizard.pareto.option1": "Summarised counts — e.g. Falls: 20, Delays: 15",
  "wizard.pareto.option2": "Not sure",

  "wizard.rare.question": "Which best describes your data?",
  "wizard.rare.hint": "Choose this when the event is uncommon and you are looking at the gap between events.",
  "wizard.rare.option0": "Time between events — e.g. days between incidents",
  "wizard.rare.option1": "Opportunities between events — e.g. procedures between harms",
  "wizard.rare.option2": "Not sure",

  "wizard.result.heading": "Recommended chart:",
  "wizard.result.hint": "You can still pick a different chart type manually if you prefer.",

  "wizard.chart.xmr": "XmR",
  "wizard.chart.xbars": "X̄–S",
  "wizard.chart.c": "C chart",
  "wizard.chart.p": "P chart",
  "wizard.chart.u": "U chart",
  "wizard.chart.t": "T",
  "wizard.chart.g": "G",
  "wizard.chart.pareto": "Pareto chart",

  "wizard.reason.measurementSingle": "You have one measurement per time point.",
  "wizard.reason.measurementSubgroups": "You have multiple measurements per time point (subgroups).",
  "wizard.reason.countConstant": "You have counts with a roughly constant amount of opportunity or activity each time.",
  "wizard.reason.countVaries": "You have counts where the amount of opportunity or activity varies, so a U chart is the right starting point. After the chart is generated, SimpleSPC will check whether the rates vary more or less than a standard U chart expects. If so, it will suggest a Laney U′ chart and explain why.",
  "wizard.reason.proportionYes": "You have a numerator and denominator for each time point, so a P chart is the right starting point. After the chart is generated, SimpleSPC will check whether the data vary more or less than a standard P chart expects. If so, it will suggest a Laney P′ chart and explain why.",
  "wizard.reason.proportionNo": "Without the numerator and denominator, a P or P′ chart cannot calculate the expected variation correctly. The percentage or proportion can instead be treated as a single measurement over time using an XmR chart.",
  "wizard.reason.paretoCounts": "You want to compare categories and identify which account for the most occurrences. A Pareto chart ranks categories from most to least frequent and shows their cumulative percentage. Your data already contain counts, so select the category column and the numeric frequency/count column. The chart helps prioritise investigation, but does not identify root causes or assess statistical process stability.",
  "wizard.reason.paretoRows": "You want to compare categories and identify which account for the most occurrences. A Pareto chart ranks categories from most to least frequent and shows their cumulative percentage. If each row represents one occurrence, select the category column and choose (count each row as 1). The chart helps prioritise investigation, but does not identify root causes or assess statistical process stability.",
  "wizard.reason.rareTime": "You have time between rare events.",
  "wizard.reason.rareOpportunities": "You have opportunities between rare events.",
  "wizard.reason.unsure": "When unsure, XmR is a safe default for a single value over time.",

  //CHART SETUP MODAL

  "setup.title": "Chart setup",
  "setup.close": "Close",
  "setup.closeAria": "Close chart setup",
  "setup.dontShow": "Don’t automatically show this again",

  // Chart-specific subtitles
  "setup.subtitle.run": "Run chart: simple view of values over time.",
  "setup.subtitle.xmr": "XmR chart: individual measurements over time.",
  "setup.subtitle.c": "C chart: count per time period.",
  "setup.subtitle.p": "P chart: proportion out of a total.",
  "setup.subtitle.pprime": "Laney P′ chart: adjusted proportion control chart.",
  "setup.subtitle.u": "U chart: rate per opportunity.",
  "setup.subtitle.uprime": "Laney U′ chart: adjusted rate control chart.",
  "setup.subtitle.xbars": "X̄–S chart: grouped measurements.",
  "setup.subtitle.t": "T chart: time between rare events.",
  "setup.subtitle.g": "G chart: opportunities between rare events.",
  "setup.subtitle.pareto": "Pareto chart: compare and prioritise categories.",

  // Reusable card headings
  "setup.heading.when": "Use this when...",
  "setup.heading.columns": "You need these columns",
  "setup.heading.example": "Example layout",
  "setup.heading.mistake": "Common mistake",
  "setup.heading.mistakes": "Common mistakes",
  "setup.heading.pprime": "Choosing P or P′",
  "setup.heading.uprime": "Choosing U or U′",
  "setup.heading.rates": "Displaying rates",
  "setup.heading.paretoRows": "Example 1: one row per occurrence",
  "setup.heading.paretoCounts": "Example 2: already summarised counts",
  "setup.heading.paretoShows": "What the chart shows",
  "setup.heading.tMode": "Choose your setup",
  "setup.heading.fallback": "Setup guidance",

  // Example-table headings and illustrative labels
  "setup.example.week": "Week",
  "setup.example.waitingTime": "Waiting time",
  "setup.example.date": "Date",
  "setup.example.lengthStay": "Length of stay",
  "setup.example.falls": "Falls",
  "setup.example.patientsHarm": "Patients with harm",
  "setup.example.patientsReviewed": "Patients reviewed",
  "setup.example.month": "Month",
  "setup.example.infections": "Infections",
  "setup.example.bedDays": "Bed days",
  "setup.example.jan": "Jan",
  "setup.example.feb": "Feb",
  "setup.example.mar": "Mar",
  "setup.example.day": "Day",
  "setup.example.reading": "Reading",
  "setup.example.sampleId": "Sample_ID",
  "setup.example.mon": "Mon",
  "setup.example.tue": "Tue",
  "setup.example.eventDate": "Event date",
  "setup.example.procedures": "Procedures between harms",
  "setup.example.category": "Category",
  "setup.example.count": "Count",
  "setup.example.medication": "Medication",
  "setup.example.pressureDamage": "Pressure damage",

//  PORTABLE SAVE MESSAGES

"portable.saveConfirm":
  "Save an editable offline copy of this SPC chart?\n\n" +
  "The HTML file will contain the underlying chart data as well as " +
  "the FreeSPC application. Store and share the saved file in line " +
  "with your organisation's information-governance requirements.",

"portable.creating":
  "Creating portable file…",

// Additional analysis block
  "summary.xmr.detail.above": "Above",
  "summary.xmr.detail.below": "Below",
  "summary.xmr.detail.increasing": "Increasing",
  "summary.xmr.detail.decreasing": "Decreasing",
  "summary.xmr.detail.disabled": "Not assessed (rule disabled)",
  "summary.xmr.detail.distinct": "Distinct flagged observations",
  "summary.xmr.detail.overlap": "A point can meet more than one rule.",
  "summary.xmr.detail.beyond": "Outside control limits",
  "summary.xmr.detail.zone23": "2 of 3 beyond 2σ",
  "summary.xmr.detail.zone45": "4 of 5 beyond 1σ",
  "summary.xmr.detail.windows": "qualifying windows",
  "summary.xmr.detail.stretches": "grouped stretches",
  "summary.xmr.detail.outliers": "Additional unusual observations",


  // Additional English translations: formatting
  "common.formatLetter.bold": "B",
  "common.formatLetter.italic": "I",
  "common.formatLetter.underline": "U",
  "chart.columns.tDatesTooltip": "Not used for T charts based on event dates.",

  // Additional English translations: error messages
  "errors.annotationRequired": "Please enter a date and a label for the annotation.",
  "errors.noData": "No data loaded yet. Upload a CSV or use the data editor first.",
  "errors.tMinEvents": "T chart needs at least 4 events.",
  "errors.gMinPoints": "G chart needs at least 4 points (each value should be 1 or more).",
  "errors.gPositive": "G chart values must be numbers 1 or greater.",
  "errors.gProbability": "Could not compute G chart probability from your data (check values are >= 1).",
  "errors.axisBounds": "Y-axis minimum cannot be greater than Y-axis maximum.",
  "errors.tDatesAxis": "T chart (event dates mode) needs Date / time axis.",
  "errors.tooFewValidPoints": "Not enough valid data points after parsing. Check your column choices.",
  "errors.tDateModeAxis": "T chart needs Date / time axis (it uses event dates).",
  "errors.tGapsMin": "T chart (gaps mode) needs at least 3 valid gap values.",
  "errors.chartNotImplemented": "Chart type \"{chartType}\" is not implemented yet.",
  "errors.cMinPoints": "C chart needs at least 2 points.",
  "errors.pMinPoints": "P chart needs at least 2 points.",
  "errors.pPrimeMinPoints": "P′ chart needs at least 2 points.",
  "errors.uMinPoints": "U chart needs at least 2 points.",
  "errors.clipboardUnsupported": "Your browser does not support copying to the clipboard. Please use Save chart(s) as… instead.",
  "errors.exportImageFailed": "Your browser could not export this image.",
  "errors.imageCreationFailed": "Could not create image.",
  "errors.projectColumnsMissing": "Project settings were applied, but some saved columns could not be found. Please reselect: {columns}",
  "errors.portableUnrecognised": "This file does not contain a recognised portable FreeSPC project.",
  "errors.portableRead": "Could not read the saved SPC chart. The file may be damaged or incompatible. Technical detail: {message}",
  "errors.portableLoaderMissing": "Could not load the saved SPC chart because the project loader was unavailable.",
  "errors.portableLoad": "Could not load the saved SPC chart. Technical detail: {message}",
  "errors.portableStart": "Could not start the saved FreeSPC file. Technical detail: {message}",

  // Additional English translations: chart setup guidance
  "setup.run.card1": "<p style=\"margin:0;\">You want a simple chart of a measurement over time using the <strong>median</strong>.</p>",
  "setup.run.card2": "<ul style=\"margin:0;\"><li><strong>Date / X-axis column</strong> → time or sequence</li><li><strong>Value / Y-axis column</strong> → the measure to follow</li></ul>",
  "setup.run.card4": "<p style=\"margin:0;\">Do not choose a row number or ID column as your measurement. The chart could look valid but be meaningless.</p>",
  "setup.xmr.card1": "<p style=\"margin:0;\">You have <strong>one measurement per time point</strong> and want a mean with control limits.</p>",
  "setup.xmr.card2": "<ul style=\"margin:0;\"><li><strong>Date / X-axis column</strong> → date, week, month or sequence</li><li><strong>Value / Y-axis column</strong> → numerical measurement</li></ul>",
  "setup.xmr.card4": "<p style=\"margin:0;\">Use XmR when you have one value per time point. If you have multiple measurements in each subgroup, use <strong>X̄–S</strong> instead.</p>",
  "setup.c.card1": "<p style=\"margin:0;\">You want to chart a <strong>count</strong> per period, with roughly the same opportunity for events each time.</p>",
  "setup.c.card2": "<ul style=\"margin:0;\"><li><strong>Date / X-axis column</strong> → date, week, month or sequence</li><li><strong>Value / Y-axis column</strong> → event count</li></ul>",
  "setup.c.card4": "<p style=\"margin:0;\">If the denominator changes substantially between observations, use a <strong>U chart</strong> rather than C.</p>",
  "setup.p.card1": "<p style=\"margin:0;\">You want to monitor a <strong>proportion</strong>, such as 5 out of 100 or the percentage meeting a standard.</p>",
  "setup.p.card2": "<ul style=\"margin:0;\"><li><strong>Date / X-axis column</strong> → time or sequence</li><li><strong>Value / Y-axis column</strong> → numerator</li><li><strong>Third column</strong> → denominator</li></ul>",
  "setup.p.card4": "<p style=\"margin:0;\">The numerator cannot be greater than the denominator.</p>",
  "setup.u.card1": "<p style=\"margin:0;\">You want to monitor a <strong>rate</strong> when the denominator changes between observations.</p>",
  "setup.u.card2": "<ul style=\"margin:0;\"><li><strong>Date / X-axis column</strong> → time or sequence</li><li><strong>Value / Y-axis column</strong> → number of events</li><li><strong>Third column</strong> → opportunities / exposure</li></ul>",
  "setup.u.card4": "<p style=\"margin:0;\">Use U when the denominator varies. If it is approximately constant, a <strong>C chart</strong> may be more appropriate.</p>",
  "setup.xbars.card1": "<p style=\"margin:0;\">You have <strong>multiple measurements in each subgroup</strong> and want to understand subgroup means and within-subgroup variation.</p>",
  "setup.xbars.card2": "<ul style=\"margin:0;\"><li><strong>Date / X-axis column</strong> → subgroup or time label</li><li><strong>Value / Y-axis column</strong> → measured value</li><li><strong>Third column</strong> → subgroup ID</li></ul>",
  "setup.xbars.card4": "<p style=\"margin:0;\">Do not use X̄–S if each subgroup contains only one measurement. Use <strong>XmR</strong> instead.",
  "setup.pprime.card1": "<p style=\"margin:0;\">You are monitoring a <strong>proportion</strong> using a numerator and denominator, but the observed variation is greater or less than a standard P chart expects. P′ adjusts the width of the control limits to reflect this variation.</p>",
  "setup.pprime.card2": "<ul style=\"margin:0;\"><li><strong>Date / X-axis column</strong> → time or sequence</li><li><strong>Value / Y-axis column</strong> → numerator (cases meeting the definition)</li><li><strong>Third column</strong> → denominator (total cases assessed)</li></ul>",
  "setup.pprime.card4": "<p style=\"margin:0;\">Normally start with a <strong>P chart</strong>. SimpleSPC checks whether observed variation differs from expectations and may suggest switching to P′.<br><br>The adjustment is called <strong>Sigma Z</strong>: it changes the control limits, not the plotted proportions. You decide whether to switch chart type.</p>",
  "setup.pprime.card5": "<p style=\"margin:0;\">Enter the actual numerator and denominator counts, not a pre-calculated percentage. The numerator cannot exceed the denominator.</p>",
  "setup.uprime.card1": "<p style=\"margin:0;\">You are monitoring <strong>events per opportunity</strong>, but observed variation is greater or less than a standard U chart expects. U′ adjusts the control limits to match this variation.</p>",
  "setup.uprime.card2": "<ul style=\"margin:0;\"><li><strong>Date / X-axis column</strong> → time or sequence</li><li><strong>Value / Y-axis column</strong> → event count</li><li><strong>Third column</strong> → opportunities or exposure</li></ul>",
  "setup.uprime.card4": "<p style=\"margin:0;\">Normally start with a <strong>U chart</strong>. SimpleSPC checks whether observed variation differs from expectations and may suggest switching to U′.<br><br>The adjustment is called <strong>Sigma Z</strong>: it changes control limits, not the underlying event rates. You decide whether to switch chart type.</p>",
  "setup.uprime.card5": "<p style=\"margin:0;\">In the Chart tab, use <strong>Display rate as</strong> to show rates per 1, 100, 1,000 or another available unit. This changes the display scale, not the statistical calculation.</p>",
  "setup.uprime.card6": "<p style=\"margin:0;\">Enter actual event and opportunity counts, not a pre-calculated rate. Unlike a proportion, a rate can exceed one event per opportunity.</p>",
  "setup.t.card1": "<p style=\"margin:0;\">You want to monitor the <strong>time between rare events</strong>.</p>",
  "setup.t.card2": "<label style=\"display:block; margin:0.25rem 0;\"><input type=\"radio\" name=\"tChartInputMode\" value=\"eventDates\" {checkedDates}><strong>I have event dates</strong> (one row per event)</label><div class=\"hint small-hint\" style=\"margin-top:0.15rem; margin-bottom:0.5rem;\">Enter the event date/time in the <em>Date / X-axis column</em>. The <em>Value / Y-axis column</em> is not used.</div><label style=\"display:block; margin:0.25rem 0;\"><input type=\"radio\" name=\"tChartInputMode\" value=\"gaps\" {checkedGaps}><strong>I already have the gaps</strong> (numeric time between events)</label><div class=\"hint small-hint\" style=\"margin-top:0.15rem;\">Enter numeric gaps in the <em>Value / Y-axis column</em>.</div>",
  "setup.t.card4": "<p style=\"margin:0;\">Do not use a T chart for monthly event counts. Use a <strong>C</strong>, <strong>P</strong> or <strong>U</strong> chart instead.</p>",
  "setup.g.card1": "<p style=\"margin:0;\">You want to monitor the <strong>number of opportunities between rare events</strong>.</p>",
  "setup.g.card2": "<ul style=\"margin:0;\"><li><strong>Date / X-axis column</strong> → optional label or order</li><li><strong>Value / Y-axis column</strong> → opportunities between events</li></ul>",
  "setup.g.card4": "<p style=\"margin:0;\">Use G when you count <strong>opportunities</strong> between events. If you are counting elapsed time, use a <strong>T chart</strong>.</p>",
  "setup.pareto.card1": "<p style=\"margin:0;\">You want to identify which categories account for the most recorded occurrences. A Pareto chart ranks categories from most to least frequent, with a cumulative percentage line.</p>",
  "setup.pareto.card2": "<ul style=\"margin:0;\"><li><strong>Category column</strong> → category to compare</li><li><strong>Frequency / count column</strong> → how to count occurrences</li></ul><p style=\"margin-bottom:0;\">If each row is one occurrence, select <strong>(count each row as 1)</strong>. If your data are already summarised, select the numeric count column.</p>",
  "setup.pareto.card5": "<p style=\"margin:0;\">Blue bars show the number of occurrences per category. The red line shows cumulative percentage.<br><br>The dashed 80% line helps prioritisation; it is <strong>not a statistical control limit</strong>.</p>",
  "setup.pareto.card6": "<p style=\"margin:0;\">Do not assume the most frequent categories are necessarily the root causes; they are starting points for investigation.<br><br>If using a numeric count column, values must be non-negative integers. At least two categories with positive counts are needed.</p>",

  // November 2026 — What's new
  "whatsNew.button": "What's new?",
  "whatsNew.bannerTitle": "SimpleSPC has been updated!",
  "whatsNew.bannerText": "SimpleSPC now supports more languages and offers clearer chart explanations. Your familiar features are still here.",
  "whatsNew.view": "See what's new",
  "whatsNew.dismiss": "Dismiss",
  "whatsNew.title": "What's new in SimpleSPC?",
  "whatsNew.release": "November 2026 update · v2026.11",
  "whatsNew.intro": "We've made a few improvements to help you use and understand SPC charts.",
  "whatsNew.languagesTitle": "More languages",
  "whatsNew.languagesText": "Choose from six languages, with translated guidance and chart explanations.",
  "whatsNew.signalsTitle": "Clearer chart explanations",
  "whatsNew.signalsText": "Find out which observations and rules are behind special-cause signals.",
  "whatsNew.refinementsTitle": "Improvements and fixes",
  "whatsNew.refinementsText": "Refinements to tooltips, X̄–S charts, number formats and exporting.",
  "whatsNew.reassure": "You can keep using SimpleSPC as before. You don't need to do anything differently.",
  "whatsNew.history": "Previous releases on GitHub",
  "whatsNew.closeAria": "Close What's new",

//Button tooltips
  "buttons.helpTitle": "Open instructions and practical guidance for using SimpleSPC.",
  "buttons.spcHelperTitle": "Ask questions about choosing, creating and interpreting SPC charts.",
  "whatsNew.buttonTitle": "See what's changed in the latest SimpleSPC update."




});

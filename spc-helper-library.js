// ============================================================
// SPC helper library
// Designed for healthcare service improvement teams
// Plain English, but SPC-defensible
// ============================================================

(function () {
 function norm(text) {
  return String(text || "")
    .toLowerCase()

    // Preserve the mathematical prime symbol used in P′ / U′
    // by translating it into the searchable word "prime".
    .replace(/′/g, " prime ")

    // Normalise ordinary apostrophes separately.
    .replace(/[’']/g, "'")

    .replace(/[^a-z0-9%+\-.\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

  function includesAll(text, parts) {
    return parts.every(p => text.includes(norm(p)));
  }

  function includesAny(text, parts) {
    return parts.some(p => text.includes(norm(p)));
  }

  function finiteNumber(x) {
    return typeof x === "number" && Number.isFinite(x);
  }

  function titleCaseChartType(chartType) {
  switch (chartType) {
    case "run": return "Run chart";
    case "xmr": return "XmR chart";
    case "xbars": return "X̄–S chart";
    case "c": return "C chart";
    case "p": return "P chart";
    case "pprime": return "P′ chart";
    case "u": return "U chart";
    case "uprime": return "U′ chart";
    case "t": return "T chart";
    case "g": return "G chart";
    case "pareto": return "Pareto chart";
    default: return "chart";
  }
}

  function signalTextList(signals) {
    if (!Array.isArray(signals) || signals.length === 0) return [];

    return signals.map(s => {
      const t = String(s || "").toLowerCase();

      if (t.includes("above ucl")) return "one or more points above the upper limit";
      if (t.includes("below lcl")) return "one or more points below the lower limit";
      if (t.includes("run")) return "a sustained run on one side of the centre line";
      if (t.includes("trend")) return "a sustained trend";
      if (t.includes("astronomical")) return "an unusual outlier";
      if (t.includes("two of three")) return "an unusual clustering of points far from the centre line";
      if (t.includes("four of five")) return "a sustained bias away from the centre line";

      return s;
    });
  }

  function joinNice(items) {
    const arr = (items || []).filter(Boolean);
    if (arr.length === 0) return "";
    if (arr.length === 1) return arr[0];
    if (arr.length === 2) return `${arr[0]} and ${arr[1]}`;
    return `${arr.slice(0, -1).join(", ")}, and ${arr[arr.length - 1]}`;
  }

  function getChartTypeSafe() {
    try {
      if (typeof getSelectedChartType_NoSideEffects === "function") {
        return getSelectedChartType_NoSideEffects() || "run";
      }
      if (typeof getSelectedChartType === "function") {
        return getSelectedChartType() || "run";
      }
    } catch (e) {}
    return "run";
  }

  function getCurrentChartContext() {
  const chartType = getChartTypeSafe();

  const runCandidate =
    typeof lastRunAnalysis !== "undefined"
      ? lastRunAnalysis
      : null;

  const xmrCandidate =
    typeof lastXmRAnalysis !== "undefined"
      ? lastXmRAnalysis
      : null;

  const xbarsCandidate =
    typeof lastXbarSAnalysis !== "undefined"
      ? lastXbarSAnalysis
      : null;

  const attributeCandidate =
    typeof lastAttributeAnalysis !== "undefined"
      ? lastAttributeAnalysis
      : null;

  const rareCandidate =
    typeof lastRareAnalysis !== "undefined"
      ? lastRareAnalysis
      : null;

    const paretoCandidate =
    typeof lastParetoAnalysis !==
      "undefined"
      ? lastParetoAnalysis
      : null;

  const isAttributeChart =
    ["c", "p", "pprime", "u", "uprime"].includes(chartType);

  const isRareChart =
    ["t", "g"].includes(chartType);

  return {
    chartType,

    run:
      chartType === "run"
        ? runCandidate
        : null,

    xmr:
      chartType === "xmr"
        ? xmrCandidate
        : null,

    xbars:
      chartType === "xbars"
        ? xbarsCandidate
        : null,

    attribute:
      isAttributeChart &&
      attributeCandidate &&
      attributeCandidate.chartType === chartType
        ? attributeCandidate
        : null,

    rare:
      isRareChart &&
      rareCandidate &&
      rareCandidate.chartType === chartType
        ? rareCandidate
        : null,

    pareto:
      chartType === "pareto"
        ? paretoCandidate
        : null,

    xmrPeriods:
      chartType === "xmr"
        ? (window.lastXmRPeriods || null)
        : null
  };
}

function describeRulesForCurrentChart(chartType) {
  switch (chartType) {
    case "p":
    case "pprime":
    case "u":
    case "uprime":
    case "c":
      return "For this chart, SimpleSPC uses beyond-limit signals and the run rule. Zone rules are blocked because they may create misleading alerts on attribute charts.";

    case "t":
    case "g":
      return "For this chart, beyond-limit signals are used by default. Run and trend rules are advanced options only, because rare-event charts can naturally look irregular.";

    case "xmr":
      return "For this XmR chart, beyond-limit signals and the run rule are used by default. Trend and zone-style rules are advanced options because adding more rules can increase false alerts.";

    case "xbars":
      return "For this X̄–S chart, beyond-limit signals and the run rule are used by default. Trend and zone rules are available as advanced options.";

    case "run":
      return "For this run chart, SimpleSPC uses run-chart rules such as shifts and, where enabled, trends. Control-limit rules do not apply to run charts.";

    case "pareto":
  return (
    "SPC signal rules are not used on a Pareto chart. " +
    "A Pareto chart ranks categories rather than showing process behaviour over time."
  );

    default:
      return "SimpleSPC uses a conservative, chart-aware rule set. Not every rule is available for every chart type.";
  }
}

  function getSuggestedQuestions(hasChart) {
  const chartType = getChartTypeSafe();

  

  const chartQuestions = [
    "What is my chart telling me?",
    "Has anything changed?",
    "Is my process stable?",
    "Is this likely to be real improvement?",
    "What does this mean for our target?",
    "What should I do next?",
    "Should I add a split here?",

    chartType === "xmr"
      ? "Why are my XmR limits different?"
      : null,

    chartType === "pprime"
      ? "Why am I using a P′ chart?"
      : null,

    chartType === "uprime"
      ? "Why am I using a U′ chart?"
      : null,

        (
      chartType === "pprime" ||
      chartType === "uprime"
    )
      ? "What does Sigma Z mean?"
      : null,

    ["xmr", "c", "p", "pprime", "u", "uprime", "xbars"].includes(chartType)
      ? "When should I exclude a point from limit calculations?"
      : null
  ].filter(Boolean);

  return {
    general: [
      "What is SPC?",
      "What is the difference between common and special cause variation?",
      "How do I choose the right chart?",
      "What is a run chart?",
      "What is an XmR chart?",
      "What does stable mean?",
      "What is the difference between a baseline and a split?",
      "How do control limits work?",
      "What should I do if I see a signal?",
      "What if my process is stable but performance is poor?",
      "Why are some SPC rules unavailable?",
      "Why are T/G run and trend rules off by default?",
      "Why can’t I use zone rules on P, P′, U, U′, or C charts?"
    ],

        chart:
      hasChart
        ? (
            chartType === "pareto"
              ? [
                  "What is my chart telling me?",
                  "What does the 80% line mean?",
                  "Are the biggest categories the causes?",
                  "What should I do next?"
                ]
              : chartQuestions
          )
        : []
  };
}

  const FAQ = [
    {
      id: "rule-availability",
      priority: 9,
      aliases: [
        "why are some spc rules unavailable",
        "why are some rules unavailable",
        "why can't i use some rules",
        "why are rules blocked"
      ],
      keywords: [
        ["rules", "unavailable"],
        ["rules", "blocked"],
        ["some", "rules"]
      ],
      answer:
        "Some SPC rules are only suitable for certain chart types.\n\n" +
        "This tool uses a conservative, chart-aware rule policy. That means it shows safer default rules first and blocks or hides rules that could create misleading alerts.\n\n" +
        "This is especially important for healthcare and public-service use, where false signals can lead to wasted effort, unnecessary concern, or alert fatigue."
    },
    {
      id: "tg-run-trend-off",
      priority: 9,
      aliases: [
        "why are tg run and trend rules off by default",
        "why are t g run trend rules off",
        "why are rare event run trend rules off",
        "why are run and trend rules off for t charts",
        "why are run and trend rules off for g charts"
      ],
      keywords: [
        ["t", "g", "run", "trend"],
        ["rare", "event", "run", "trend"],
        ["run", "trend", "off"]
      ],
      answer:
        "T and G charts are rare-event charts. They can naturally look irregular even when the process has not really changed.\n\n" +
        "For that reason, this tool keeps run and trend rules off by default for T/G charts. You can enable them in advanced settings, but they should be interpreted cautiously because they may create false signals."
    },
    {
  id: "puc-zone-rules-blocked",
  priority: 9,

  aliases: [
    "why can't i use zone rules on p p prime u u prime or c charts",
    "why are zone rules blocked on p charts",
    "why are zone rules blocked on p prime charts",
    "why are zone rules blocked on u charts",
    "why are zone rules blocked on u prime charts",
    "why are zone rules blocked on c charts",
    "zone rules p u c"
  ],

  keywords: [
    ["zone", "rules", "p"],
    ["zone", "rules", "u"],
    ["zone", "rules", "c"],
    ["zone", "rules", "blocked"]
  ],

  answer:
    "Zone rules are not available for P, P′, U, U′, or C charts because they can create misleading alerts on these chart types.\n\n" +
    "These charts are based on count, proportion, or rate data. Their limits may vary, and their data may not behave like the continuous data that classic zone rules were designed for.\n\n" +
    "The safer default is to use beyond-limit signals and run rules, rather than adding extra zone-rule alerts."
},
{
  id: "xmr-moving-range-screening",
  priority: 10,
  aliases: [
    "what is moving range screening",
    "what is mr screening",
    "what is the nelson method",
    "nelson moving range",
    "why are my xmr limits different",
    "why are xmr limits different",
    "screen large jumps",
    "3.5 moving range"
  ],
  keywords: [
    ["moving", "range", "screening"],
    ["mr", "screening"],
    ["nelson"],
    ["xmr", "limits", "different"],
    ["large", "jump"],
    ["3.5"]
  ],
  answer:
    "XmR control limits estimate routine variation from the differences between consecutive points, called moving ranges.\n\n" +
    "A very large point-to-point jump can inflate the average moving range and make the control limits unusually wide. By default, SimpleSPC screens unusually large moving ranges when estimating routine variation. The original observations and moving ranges are still shown on the chart.\n\n" +
    "SimpleSPC uses the Nelson method: moving ranges greater than 3.5 times the initial average moving range are not used in the variation estimate. You can turn this screening off in Advanced detection options if you want the unscreened calculation."
},
    {
      id: "exclude-point-from-limits",
      priority: 10,
      aliases: [
        "when should i exclude a point from limit calculations",
        "exclude a point from limit calculations",
        "exclude point from limits",
        "ghost a point",
        "ghosting",
        "causing out"
      ],
      keywords: [
        ["exclude", "point", "limit"],
        ["exclude", "calculation"],
        ["ghosting"],
        ["causing", "out"]
      ],
      answer:
        "Exclude a point from the limit calculations only after you have investigated it and have good reason to believe it reflects a known special cause rather than the routine variation of the process.\n\n" +
        "The point is not deleted. It stays visible on the chart, but SimpleSPC does not use it to estimate the centre line or control limits. Do not exclude a point simply because it is high, low, or outside a control limit. First investigate what happened.\n\n" +
        "If the change represents a lasting new way of working or a new process level, a split is usually more appropriate than excluding the point.\n\n" +
        "For XmR charts, excluding an observation also leaves out the moving ranges immediately before and after that observation when routine variation is estimated. SimpleSPC does not create an artificial moving range across the excluded point. For X̄–S charts, exclusion applies to the whole plotted subgroup.\n\n" +
        "Technical note: some SPC texts call this ghosting or causing out a point. It is a way of estimating the process's routine common-cause variation — sometimes called the voice of the process — without allowing an investigated special cause to distort the limits."
    },

    {
      id: "what-is-spc",
      priority: 8,
      aliases: ["what is spc", "define spc", "explain spc"],
      keywords: [["what", "spc"], ["define", "spc"], ["explain", "spc"]],
      answer:
        "Statistical Process Control (SPC) helps you use data over time to understand whether a process is behaving as usual, or whether something may have changed.\n\n" +
        "In healthcare improvement, that helps teams distinguish routine ups and downs from signals that may reflect a real change in the service, pathway, staffing, demand, coding, or measurement system."
    },
    {
      id: "common-vs-special",
      priority: 8,
      aliases: ["common cause", "special cause", "difference between common and special cause variation"],
      keywords: [
        ["common", "cause"],
        ["special", "cause"],
        ["difference", "common", "special", "cause"]
      ],
      answer:
        "Common-cause variation is the routine fluctuation built into the current system. Special-cause variation means the pattern suggests something unusual may have happened.\n\n" +
        "In practice:\n" +
        "• common cause = the system behaving as it usually does\n" +
        "• special cause = a prompt to investigate what may have changed"
    },
    {
  id: "choose-chart",
  priority: 8,
  aliases: [
    "choose the right chart",
    "which chart should i use",
    "how do i choose the right chart"
  ],
  keywords: [
    ["choose", "chart"],
    ["which", "chart"],
    ["right", "chart"]
  ],
  answer:
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
    "For proportion data, SimpleSPC normally starts with a P chart and checks whether P′ better matches the observed variation. " +
    "For rate data with changing opportunity, it normally starts with a U chart and checks whether U′ better matches the observed variation."
},
    {
      id: "run-chart",
      priority: 7,
      aliases: ["what is a run chart", "run chart"],
      keywords: [["run", "chart"]],
      answer:
        "A run chart shows data over time with a centre line, usually the median. It is a simple way to see whether the pattern looks broadly stable or whether there may be a shift or trend.\n\n" +
        "It is often a very good starting point for healthcare improvement teams."
    },
    {
      id: "xmr-chart",
      priority: 7,
      aliases: ["what is an xmr chart", "xmr chart"],
      keywords: [["xmr", "chart"]],
      answer:
        "An XmR chart is used for continuous data when you have one observation per time point.\n\n" +
        "It usually has:\n" +
        "• an Individuals (X) chart showing the values over time\n" +
        "• a Moving Range (MR) chart showing how much consecutive points differ\n\n" +
        "It helps separate routine variation from possible special-cause signals."
    },
    {
      id: "stable-meaning",
      priority: 7,
      aliases: ["what does stable mean", "what is a stable process", "stability"],
      keywords: [["stable"], ["stability"]],
      answer:
        "A stable process shows only routine common-cause variation. In simple terms, it is behaving predictably within its current level of performance.\n\n" +
        "Stable does not necessarily mean good. A stable process can still be consistently underperforming."
    },
    {
      id: "control-limits",
      priority: 8,
      aliases: ["control limits", "ucl", "lcl", "how do control limits work"],
      keywords: [["control", "limit"], ["ucl"], ["lcl"], ["upper", "lower", "limit"]],
      answer:
        "Control limits show the range you would expect from routine common-cause variation. They are not the same as targets, standards, or pass/fail thresholds.\n\n" +
        "A point outside the limits suggests something may have changed and is worth investigating. Some chart types also support selected pattern rules, but the tool only offers those where they are reasonably defensible."
    },
    {
      id: "target",
      priority: 6,
      aliases: ["what is a target", "how should i use a target", "target"],
      keywords: [["target"]],
      answer:
        "A target is the level of performance you are aiming for. In SPC, it should be interpreted alongside stability.\n\n" +
        "A stable process below target often means the system needs redesign or improvement. An unstable process meeting target intermittently may still not be reliable."
    },
    {
      id: "baseline",
      priority: 7,
      aliases: ["what is a baseline", "when should i use a baseline"],
      keywords: [["baseline"]],
      answer:
        "A baseline is the initial set of points used to calculate the centre line and, where relevant, the limits.\n\n" +
        "It is useful when you want to compare later performance against an earlier period that represents the original system."
    },
    {
      id: "baseline-vs-split",
      priority: 8,
      aliases: ["difference between a baseline and a split", "baseline and split", "baseline vs split"],
      keywords: [
        ["baseline", "split"],
        ["difference", "baseline", "split"]
      ],
      answer:
        "A baseline fixes the calculation using an earlier reference period. A split recalculates the centre line and limits from a chosen point onward.\n\n" +
        "In simple terms:\n" +
        "• baseline = compare later performance against the old system\n" +
        "• split = treat the later period as a new system or new normal"
    },
    {
      id: "what-to-do-signal",
      priority: 8,
      aliases: ["what should i do if i see a signal", "what do i do if i see a signal", "what should i do next"],
      keywords: [
        ["what", "do", "signal"],
        ["what", "should", "i", "do", "next"],
        ["see", "signal"]
      ],
      answer:
        "A special-cause signal is usually a prompt to investigate what changed around that time.\n\n" +
        "In a healthcare service, that might include a pathway change, staffing change, demand pressure, case-mix change, coding change, data-definition change, equipment issue, or an intentional improvement intervention.\n\n" +
        "The aim is not to blame a person or overreact to one point. The aim is to understand whether something in the system changed, whether that change is real, and whether it is likely to continue."
    },
    {
      id: "stable-but-poor",
      priority: 8,
      aliases: ["stable but poor", "stable but below target", "stable but not good enough"],
      keywords: [
        ["stable", "poor"],
        ["stable", "below", "target"],
        ["stable", "not", "good", "enough"]
      ],
      answer:
        "If the process is stable but the level of performance is not good enough, that usually means the system is delivering a predictable but unsatisfactory result.\n\n" +
        "In improvement terms, that often points toward redesigning or improving the system rather than reacting to individual good or bad days."
    },
    {
      id: "data-definition-change",
      priority: 7,
      aliases: ["what if our data definition changed", "coding change", "case definition changed"],
      keywords: [
        ["data", "definition", "changed"],
        ["coding", "change"],
        ["case", "definition", "changed"]
      ],
      answer:
        "A change in coding, inclusion criteria, denominator rules, or case definition can create a chart signal even if the real service has not changed.\n\n" +
        "That is why SPC charts should always be interpreted alongside local knowledge about data collection and operational context."
    },
    {
      id: "annotations",
      priority: 6,
      aliases: ["how should i use annotations", "annotations"],
      keywords: [["annotation"], ["annotations"]],
      answer:
        "Use annotations to mark events that could plausibly explain a change in the pattern, such as a new pathway, staffing change, policy change, or data-definition change.\n\n" +
        "Keep them short and factual. They support interpretation, but they do not prove causation on their own."
    },
    {
      id: "improvement-vs-one-good-point",
      priority: 7,
      aliases: ["are we improving", "is this improvement", "one good point"],
      keywords: [
        ["improv"],
        ["getting", "better"],
        ["one", "good", "point"]
      ],
      answer:
        "In SPC, improvement is more convincing when the chart shows a sustained change in the desired direction, not just one good point.\n\n" +
        "A single good result can happen by chance. A sustained shift or other clear signal is more consistent with real change."
    },
{
  id: "runchart-median",
  priority: 6,
  aliases: ["why does a run chart use the median", "why median run chart"],
  keywords: [
    ["run", "chart", "median"],
    ["why", "median"]
  ],
  answer:
    "Run charts often use the median because it is simple and robust, especially when you are starting improvement work or have limited data.\n\n" +
    "It gives a practical centre line for judging runs and trends."
},
{
  id: "xmr-mean",
  priority: 6,
  aliases: ["why does an xmr chart use the mean", "why mean xmr"],
  keywords: [
    ["xmr", "mean"],
    ["why", "xmr", "mean"]
  ],
  answer:
    "An XmR chart uses the mean together with the moving range to estimate routine variation and calculate control limits.\n\n" +
    "That gives stronger statistical signals than a simple run chart when the data suit that chart type."
},
{
  id: "should-add-split",
  priority: 8,
  aliases: ["should i add a split here", "should i split the chart", "add split"],
  keywords: [
    ["add", "split"],
    ["split", "chart"],
    ["should", "split"]
  ],
  answer:
    "Usually only add a split when there is good evidence of a real and sustained process change, and you expect the new pattern to continue.\n\n" +
    "A split should help interpretation, not hide instability or make the chart look better."
},
{
  id: "above-target-solved",
  priority: 7,
  aliases: ["we are above target now", "is the problem solved", "above target solved"],
  keywords: [
    ["above", "target"],
    ["problem", "solved"]
  ],
  answer:
    "Not necessarily. Crossing the target once is not the same as showing a stable, sustained level of performance.\n\n" +
    "A better question is whether the process now looks predictably at or beyond the target, and whether that change is likely to last."
},
{
  id: "below-target-stable",
  priority: 8,
  aliases: ["below target but stable", "stable but below target"],
  keywords: [
    ["below", "target", "stable"],
    ["stable", "below", "target"]
  ],
  answer:
    "This usually means the system is delivering a consistent result, but not a good enough one.\n\n" +
    "That often points toward improvement or redesign work rather than reacting to individual points."
},
{
  id: "caused-improvement",
  priority: 7,
  aliases: ["did this cause the improvement", "can i say this caused improvement"],
  keywords: [
    ["cause", "improvement"],
    ["caused", "improvement"]
  ],
  answer:
    "Be cautious. A chart can show that performance changed, and annotations can help link timing to an intervention, but the chart alone does not prove causation.\n\n" +
    "It is usually better to say the change is consistent with the intervention, supported by local knowledge."
},
{
  id: "case-mix",
  priority: 6,
  aliases: ["what if patient case mix changed", "case mix changed"],
  keywords: [
    ["case", "mix"],
    ["patient", "mix"]
  ],
  answer:
    "A real change in the type or complexity of patients can affect the chart, even if the service process stayed the same.\n\n" +
    "That does not make the chart wrong, but it does affect interpretation."
},
{
  id: "denominator-changed",
  priority: 6,
  aliases: ["what if our denominator changed", "denominator changed a lot"],
  keywords: [
    ["denominator", "changed"],
    ["denominator"]
  ],
  answer:
    "If the amount of opportunity changes a lot over time, that affects how the chart should be set up and interpreted.\n\n" +
    "That is one reason P and U charts are useful: they allow for changing totals or opportunities."
},
{
  id: "rare-event-improvement",
  priority: 6,
  aliases: ["what does rare event improvement look like", "rare event improvement"],
  keywords: [
    ["rare", "event", "improvement"],
    ["rare", "event"]
  ],
  answer:
    "For rare-event charts, improvement often means longer times between unwanted events or more opportunities between events.\n\n" +
    "Because rare-event data can look dramatic, it helps to focus on sustained pattern changes rather than one unusually long gap."
}
  ];

  function scoreFaqItem(item, q) {
    let score = item.priority || 0;

    (item.aliases || []).forEach(alias => {
      const a = norm(alias);
      if (q === a) score += 100;
      else if (q.includes(a)) score += 30;
    });

    (item.keywords || []).forEach(k => {
      if (Array.isArray(k)) {
        if (includesAll(q, k)) score += 15;
      } else {
        const kk = norm(k);
        if (q.includes(kk)) score += 10;
      }
    });

    return score;
  }

  function findBestFaq(q) {
    let best = null;
    let bestScore = 0;

    FAQ.forEach(item => {
      const score = scoreFaqItem(item, q);
      if (score > bestScore) {
        best = item;
        bestScore = score;
      }
    });

    return bestScore >= 15 ? best : null;
  }

function isXmRPeriodAssessable(period) {
  if (!period) return false;

  // Newer SimpleSPC analysis explicitly records this.
  if (period.hasEnoughPointsForInterpretation === false) {
    return false;
  }

  // null means "not yet assessed", not "unstable".
  if (period.isStable === null) {
    return false;
  }

  return typeof period.isStable === "boolean";
}

  function describeXmRPeriods(periods) {
  if (!Array.isArray(periods) || periods.length === 0) return null;

  const pieces = periods.map((p, idx) => {
    const label = `Period ${idx + 1}`;
    const start = (p.startIndex ?? 0) + 1;
    const end = (p.endIndex ?? 0) + 1;
    const meanTxt = finiteNumber(p.mean) ? p.mean.toFixed(2) : null;

    if (!isXmRPeriodAssessable(p)) {
      const count = finiteNumber(p.limitEstimatePointCount)
        ? p.limitEstimatePointCount
        : null;

      const minimum = finiteNumber(p.minimumPointsForInterpretation)
        ? p.minimumPointsForInterpretation
        : 12;

      return (
        `${label} (points ${start}–${end}) has provisional XmR limits` +
        `${meanTxt ? ` around a mean of ${meanTxt}` : ""}. ` +
        `${count !== null ? `Only ${count} points were used to estimate the limits, so ` : ""}` +
        `SimpleSPC will wait until at least ${minimum} points are available before describing this period as stable or showing special-cause variation`
      );
    }

    const sigs = signalTextList(p.signals);

    if (p.isStable === true) {
      return `${label} (points ${start}–${end}) looks stable${meanTxt ? ` around a mean of ${meanTxt}` : ""}`;
    }

    return `${label} (points ${start}–${end}) shows ${joinNice(sigs) || "special-cause signals"}${meanTxt ? ` around a mean of ${meanTxt}` : ""}`;
  });

  return pieces.join(". ") + ".";
}

  function buildXmRChangeAnswer(ctx) {
  const latest = ctx.xmr;
  const periods = ctx.xmrPeriods;

  // Helper used only inside this answer:
  // look for early warning features without treating them as a
  // formal stability conclusion when the period is still provisional.
  function provisionalHasWarningSigns() {
    if (!latest) return false;

    return (
      (finiteNumber(latest.nBeyond) && latest.nBeyond > 0) ||
      !!latest.hasRunViolation ||
      !!latest.hasTrend ||
      !!latest.hasAstronomical
    );
  }

  if (Array.isArray(periods) && periods.length > 1) {
    const latestPeriod = periods[periods.length - 1];
    const latestAssessable = isXmRPeriodAssessable(latestPeriod);

    const earlierAssessablePeriods =
      periods
        .slice(0, -1)
        .filter(isXmRPeriodAssessable);

    const earlierUnstable =
      earlierAssessablePeriods.some(p => p.isStable === false);

    const latestStart = (latestPeriod.startIndex ?? 0) + 1;
    const latestEnd = (latestPeriod.endIndex ?? 0) + 1;

    let out = "";

    if (earlierUnstable) {
      out +=
        "Looking across the XmR chart, one or more periods with enough data show special-cause signals. That is evidence that the process may have changed.\n\n";
    } else {
      out +=
        "This XmR chart has been divided into separate process periods using one or more splits. A split marks a point where you believe the process may have changed and tells SimpleSPC to recalculate the mean and limits. The split itself does not prove that a change occurred.\n\n";
    }

    out += describeXmRPeriods(periods) + "\n\n";

    // Latest period is still provisional
    if (!latestAssessable) {
      const count = finiteNumber(latestPeriod.limitEstimatePointCount)
        ? latestPeriod.limitEstimatePointCount
        : null;

      const minimum = finiteNumber(latestPeriod.minimumPointsForInterpretation)
        ? latestPeriod.minimumPointsForInterpretation
        : 12;

      out +=
        `For the most recent period (points ${latestStart}–${latestEnd}), there are not yet enough data to make a firm judgement about stability. ` +
        `${count !== null ? `There are currently ${count} points in this period, and ` : ""}` +
        `SimpleSPC uses at least ${minimum} points before making a formal XmR stability interpretation.`;

      if (provisionalHasWarningSigns()) {
        out +=
          "\n\nThere are already some features in the latest points that are worth watching. That means the early picture is not yet clearly consistent with a settled process. More data will help show whether those features disappear as temporary fluctuation or develop into a sustained signal.";
      } else {
        out +=
          "\n\nSo far, the latest points look fairly consistent around their current level and there are no obvious warning signs in this short period. That is encouraging, but it is still provisional rather than proof of stability. If the next points continue in a similar pattern, that would strengthen the evidence that the process has settled into a new stable level.";
      }

      return out;
    }

    // Latest period has enough data for a formal interpretation
    if (latestPeriod.isStable === true) {
      out +=
        `In the most recent period (points ${latestStart}–${latestEnd}), I cannot see a clear special-cause signal. With enough data now available for interpretation, the latest section is consistent with routine variation around its current level.`;
    } else {
      const latestSignals =
        joinNice(signalTextList(latestPeriod.signals));

      out +=
        `In the most recent period (points ${latestStart}–${latestEnd}), there are still signs of possible change — specifically ${latestSignals || "special-cause signals"}.`;
    }

    return out;
  }

  // No current XmR analysis
  if (!latest) {
    return "Please generate an XmR chart first, then ask again.";
  }

  // Single provisional period
  if (!isXmRPeriodAssessable(latest)) {
    const count = finiteNumber(latest.limitEstimatePointCount)
      ? latest.limitEstimatePointCount
      : null;

    const minimum = finiteNumber(latest.minimumPointsForInterpretation)
      ? latest.minimumPointsForInterpretation
      : 12;

    let out =
      "There are not yet enough data to make a firm judgement about stability. " +
      `${count !== null ? `There are currently ${count} points being used to estimate the limits, and ` : ""}` +
      `SimpleSPC uses at least ${minimum} points before making a formal XmR stability interpretation.`;

    if (provisionalHasWarningSigns()) {
      out +=
        "\n\nThere are already some features in the pattern that are worth watching, so I would not yet describe the process as settled. More observations will help show whether these develop into a sustained signal.";
    } else {
      out +=
        "\n\nSo far, the pattern looks fairly consistent and there are no obvious warning signs. That is encouraging, but still provisional. If future points continue in a similar pattern, confidence that the process is stable will become stronger.";
    }

    return out;
  }

  // Single period with enough data
  if (latest.isStable === true) {
    return (
      "Looking at the current XmR chart, I cannot see a clear special-cause signal. " +
      "There are enough data for interpretation, so the pattern is consistent with routine variation around its current level. " +
      "That does not necessarily mean performance is good enough; stability and performance against a target are separate questions."
    );
  }

  return (
    "This XmR chart shows evidence that something may have changed. " +
    "The pattern is not fully explained by routine variation alone, so it would be worth investigating what changed in the system or the data around that time."
  );
}

function buildPPrimeExplanation(ctx) {
  const a = ctx.attribute;

  if (!a || ctx.chartType !== "pprime") {
    return (
      "A P′ chart is a Laney-adjusted version of a P chart. " +
      "It is used for proportion data when the amount of variation is noticeably different from what a standard P chart expects."
    );
  }

  const sigmaZ =
    a.stats && finiteNumber(a.stats.sigmaZ)
      ? a.stats.sigmaZ
      : null;

  let out =
    "This P′ chart uses the same numerator, denominator, observed proportions and centre line as a standard P chart. " +
    "The difference is that it adjusts the width of the control limits so they better match the amount of variation actually seen in the data.";

  if (sigmaZ != null) {
    if (sigmaZ > 1) {
      out +=
        "\n\nYour data vary more than a standard P chart expects, so P′ uses wider control limits. " +
        "This helps avoid treating ordinary variation as a possible process change.";
    } else if (sigmaZ < 1) {
      out +=
        "\n\nYour data vary less than a standard P chart expects, so P′ uses narrower control limits. " +
        "This can make genuinely unusual variation easier to detect.";
    } else {
      out +=
        "\n\nThe amount of variation is very close to what a standard P chart expects, so the P′ adjustment is very small.";
    }

    out +=
      `\n\nThe Sigma Z adjustment factor for this chart is ${sigmaZ.toFixed(3)}. ` +
      "A value above 1 widens the standard P-chart limits; a value below 1 narrows them.";
  }

  return out;
}

function buildUPrimeExplanation(ctx) {
  const a = ctx.attribute;

  if (!a || ctx.chartType !== "uprime") {
    return (
      "A U′ chart is a Laney-adjusted version of a U chart. " +
      "It is used for rate data when the amount of variation is noticeably different from what a standard U chart expects."
    );
  }

  const sigmaZ =
    a.stats && finiteNumber(a.stats.sigmaZ)
      ? a.stats.sigmaZ
      : null;

  let out =
    "This U′ chart uses the same event counts, opportunities, observed rates and centre line as a standard U chart. " +
    "The difference is that it adjusts the width of the control limits so they better match the amount of variation actually seen in the data.";

  if (sigmaZ != null) {
    if (sigmaZ > 1) {
      out +=
        "\n\nYour data vary more than a standard U chart expects, so U′ uses wider control limits. " +
        "This helps avoid treating ordinary variation as a possible process change.";
    } else if (sigmaZ < 1) {
      out +=
        "\n\nYour data vary less than a standard U chart expects, so U′ uses narrower control limits. " +
        "This can make genuinely unusual variation easier to detect.";
    } else {
      out +=
        "\n\nThe amount of variation is very close to what a standard U chart expects, so the U′ adjustment is very small.";
    }

    out +=
      `\n\nThe Sigma Z adjustment factor for this chart is ${sigmaZ.toFixed(3)}. ` +
      "A value above 1 widens the standard U-chart limits; a value below 1 narrows them.";
  }

  return out;
}

function buildSigmaZExplanation(ctx) {
  const chartType = ctx.chartType;
  const a = ctx.attribute;

  const isPPrime =
    chartType === "pprime";

  const isUPrime =
    chartType === "uprime";

  if (!isPPrime && !isUPrime) {
    return (
      "Sigma Z is used on Laney P′ and U′ charts. " +
      "It is an adjustment factor that tells SimpleSPC whether the data are varying more or less than a standard P or U chart would normally expect."
    );
  }

  const sigmaZ =
    a?.stats &&
    finiteNumber(a.stats.sigmaZ)
      ? a.stats.sigmaZ
      : null;

  const standardChart =
    isPPrime
      ? "P chart"
      : "U chart";

  const primeChart =
    isPPrime
      ? "P′ chart"
      : "U′ chart";

  let out =
    "Sigma Z is the adjustment factor used by a Laney chart. " +
    `It tells SimpleSPC whether your data are varying more or less than a standard ${standardChart} expects.\n\n` +
    "A Sigma Z of about 1 means the amount of variation is close to what the standard chart expects. " +
    "A value above 1 means there is more variation than expected, so the control limits are made wider. " +
    "A value below 1 means there is less variation than expected, so the control limits are made narrower.\n\n" +
    "Sigma Z does not change your original observations or the centre line. " +
    `It only adjusts the width of the control limits on the ${primeChart}.`;

  if (sigmaZ != null) {
    out +=
      `\n\nFor your current chart, Sigma Z = ${sigmaZ.toFixed(3)}.`;

    if (sigmaZ > 1) {
      out +=
        ` This means the data are showing more variation than the standard ${standardChart} expects, so the ${primeChart} uses wider limits.`;
    } else if (sigmaZ < 1) {
      out +=
        ` This means the data are showing less variation than the standard ${standardChart} expects, so the ${primeChart} uses narrower limits.`;
    } else {
      out +=
        ` This is exactly 1, so the Laney adjustment does not change the width of the standard ${standardChart} limits.`;
    }
  }

  out +=
    "\n\nTechnical detail: SimpleSPC first allows for the amount of variation the standard chart expects at each point. " +
    "It then looks at the point-to-point changes in those standardised values. " +
    "Sigma Z summarises that extra or reduced variation into a single adjustment factor.";

  return out;
}

  function buildLatestPeriodSummary(ctx) {
  const chartType = ctx.chartType;

  if (
    chartType === "pareto"
  ) {
    return (
      buildParetoAnswer(ctx) +
      "\n\nA Pareto chart does not assess process stability because the categories are not plotted in time order."
    );
  }

  if (chartType === "xmr" && ctx.xmr) {
    const a = ctx.xmr;
    const start = (a.startIndex ?? 0) + 1;
    const end = (a.endIndex ?? 0) + 1;

    const periodText =
      (a.periodCount && a.periodCount > 1)
        ? `Looking at the latest period (Period ${a.periodIndex} of ${a.periodCount}, points ${start}–${end})`
        : `Looking at the chart (points ${start}–${end})`;

    if (!isXmRPeriodAssessable(a)) {
      const count =
        finiteNumber(a.limitEstimatePointCount)
          ? a.limitEstimatePointCount
          : null;

      const minimum =
        finiteNumber(a.minimumPointsForInterpretation)
          ? a.minimumPointsForInterpretation
          : 12;

      const provisionalHasWarning =
        (finiteNumber(a.nBeyond) && a.nBeyond > 0) ||
        !!a.hasRunViolation ||
        !!a.hasTrend ||
        !!a.hasAstronomical;

      let out =
        `${periodText}, there are not yet enough data to make a firm judgement about stability. ` +
        `${count !== null ? `There are currently ${count} points in this period, and ` : ""}` +
        `SimpleSPC uses at least ${minimum} points before making a formal XmR stability interpretation.`;

      if (provisionalHasWarning) {
        out +=
          "\n\nThere are already some features in this short period that are worth watching, so I would not yet describe the process as stable. More data will help show whether these are temporary fluctuations or part of a real pattern.";
      } else {
        out +=
          "\n\nSo far, the pattern looks fairly consistent and there are no obvious warning signs in this short period. That is encouraging, but it is still provisional rather than proof of stability.";
      }

      return out;
    }

    if (a.isStable === true) {
      return (
        `${periodText}, I cannot see a clear special-cause signal using the current rule settings. ` +
        "This suggests the process is behaving broadly consistently around its current level."
      );
    }

    const sigs =
      joinNice(
        signalTextList(a.signals)
      );

    return (
      `${periodText}, this XmR chart shows possible special-cause variation — specifically ` +
      `${sigs || "one or more SPC signals"}.`
    );
  }

  if (chartType === "xbars" && ctx.xbars) {
    const xStable =
      !!ctx.xbars.xbar?.isStable;

    const sStable =
      !!ctx.xbars.s?.isStable;

    if (xStable && sStable) {
      return (
        "Looking at the latest X̄–S period, both the subgroup averages and the within-subgroup variation look stable. " +
        "I cannot see a clear special-cause signal in either part of the chart."
      );
    }

    if (!xStable && sStable) {
      return (
        "Looking at the latest X̄–S period, the subgroup averages show possible change, " +
        "but the within-subgroup variation looks stable."
      );
    }

    if (xStable && !sStable) {
      return (
        "Looking at the latest X̄–S period, the subgroup averages look stable, " +
        "but the within-subgroup variation shows possible change."
      );
    }

    return (
      "Looking at the latest X̄–S period, both the subgroup averages and the within-subgroup variation show possible change."
    );
  }

  if (ctx.attribute) {
    const sigs =
      joinNice(
        signalTextList(
          ctx.attribute.signals
        )
      );

    if (ctx.attribute.isStable === true) {
      return (
        `Looking at the latest ${titleCaseChartType(chartType).toLowerCase()} period, ` +
        "I cannot see a clear special-cause signal using the current rule settings."
      );
    }

    return (
      `Looking at the latest ${titleCaseChartType(chartType).toLowerCase()} period, ` +
      `there are signs of possible special-cause variation — specifically ${sigs || "one or more SPC signals"}.`
    );
  }

  if (ctx.rare) {
    const sigs =
      joinNice(
        signalTextList(
          ctx.rare.signals
        )
      );

    if (ctx.rare.isStable === true) {
      return (
        `Looking at the latest ${titleCaseChartType(chartType).toLowerCase()} period, ` +
        "I cannot see a clear special-cause signal using the current rare-event rule settings."
      );
    }

    return (
      `Looking at the latest ${titleCaseChartType(chartType).toLowerCase()} period, ` +
      `there are signs of possible special-cause variation — specifically ${sigs || "one or more SPC signals"}.`
    );
  }

  if (ctx.run) {
    if (ctx.run.isStable === true) {
      return (
        "Looking at the current run chart, I cannot see a clear signal that the pattern has changed."
      );
    }

    return (
      "Looking at the current run chart, there are signs of possible change in the pattern."
    );
  }

  return (
    "Generate the currently selected chart first and I can give you a chart-specific interpretation."
  );
}

  function buildTargetAnswer(ctx) {
  const chartType = ctx.chartType;

  if (chartType === "xmr" && ctx.xmr) {
    const a = ctx.xmr;

    const hasTarget =
      finiteNumber(a.target);

    const hasMean =
      finiteNumber(a.mean);

    if (!hasTarget) {
      return (
        "There is no target currently set on this XmR chart. " +
        "If you add one, SimpleSPC can help interpret performance against the target alongside process stability."
      );
    }

    const target =
      a.target;

    const direction =
      a.direction === "above"
        ? "above"
        : "below";

    let currentPositionText = "";

    if (hasMean) {
      const meanMeetsTarget =
        direction === "above"
          ? a.mean >= target
          : a.mean <= target;

      currentPositionText =
        meanMeetsTarget
          ? `The current period mean (${a.mean.toFixed(2)}) is on the favourable side of the target (${target}).`
          : `The current period mean (${a.mean.toFixed(2)}) is not yet on the favourable side of the target (${target}).`;
    }

    if (!isXmRPeriodAssessable(a)) {
      const count =
        finiteNumber(a.limitEstimatePointCount)
          ? a.limitEstimatePointCount
          : null;

      const minimum =
        finiteNumber(a.minimumPointsForInterpretation)
          ? a.minimumPointsForInterpretation
          : 12;

      let out =
        "The latest XmR period is still provisional, so it is too early to make a firm judgement about whether the process can reliably meet the target.";

      if (currentPositionText) {
        out +=
          `\n\n${currentPositionText}`;
      }

      out +=
        `\n\n${count !== null ? `There are currently ${count} points in this period, and ` : ""}` +
        `SimpleSPC uses at least ${minimum} points before making a formal stability or capability interpretation.`;

      if (hasMean) {
        const meanMeetsTarget =
          direction === "above"
            ? a.mean >= target
            : a.mean <= target;

        if (meanMeetsTarget) {
          out +=
            "\n\nThe early results are encouraging, but they do not yet show that this level of performance is predictable or sustainable.";
        } else {
          out +=
            "\n\nThe early pattern may still be settling. More data will show whether the process level is genuinely different and whether further improvement is needed to reach the target reliably.";
        }
      }

      return out;
    }

    if (a.isStable === true) {
      let out =
        "The latest XmR period has enough data for interpretation and currently looks stable.";

      if (currentPositionText) {
        out +=
          `\n\n${currentPositionText}`;
      }

      if (
        a.capability &&
        finiteNumber(a.capability.prob)
      ) {
        out +=
          `\n\nBased on the current stable process, SimpleSPC estimates that about ${(a.capability.prob * 100).toFixed(1)}% of future results would meet the target, assuming the process continues to behave in the same way.`;
      } else {
        out +=
          "\n\nBecause the process appears stable, the target comparison is more meaningful: the main question is whether the current system is predictably good enough.";
      }

      return out;
    }

    let out =
      "The latest XmR period has enough data for interpretation, but it shows special-cause variation. " +
      "That means target performance should be interpreted cautiously because the process is not yet behaving predictably.";

    if (currentPositionText) {
      out +=
        `\n\n${currentPositionText}`;
    }

    out +=
      "\n\nThe priority is to understand what is changing in the system, demand, staffing, pathway, coding, or measurement before relying on a capability estimate.";

    return out;
  }

  let target = null;

  try {
    if (typeof getTargetValue === "function") {
      target = getTargetValue();
    }
  } catch (e) {}

  if (!finiteNumber(target)) {
    return (
      `There is no target currently set on this ${titleCaseChartType(chartType).toLowerCase()}. ` +
      "If you add one, SimpleSPC can help you interpret the target alongside the pattern of variation."
    );
  }

  let direction = "above";

  try {
    if (
      typeof targetDirectionInput !== "undefined" &&
      targetDirectionInput &&
      targetDirectionInput.value === "below"
    ) {
      direction = "below";
    }
  } catch (e) {}

  let centre = null;
  let stable = null;

  if (
    ctx.attribute &&
    ctx.attribute.stats &&
    finiteNumber(ctx.attribute.stats.cl)
  ) {
    centre =
      ctx.attribute.stats.cl;

    stable =
      ctx.attribute.isStable;
  } else if (
    ctx.rare &&
    ctx.rare.stats &&
    finiteNumber(ctx.rare.stats.cl)
  ) {
    centre =
      ctx.rare.stats.cl;

    stable =
      ctx.rare.isStable;
  } else if (
    ctx.run &&
    finiteNumber(ctx.run.median)
  ) {
    centre =
      ctx.run.median;

    stable =
      ctx.run.isStable;
  } else if (
    ctx.xbars &&
    ctx.xbars.stats &&
    finiteNumber(ctx.xbars.stats.xbarbar)
  ) {
    centre =
      ctx.xbars.stats.xbarbar;

    stable =
      !!ctx.xbars.xbar?.isStable &&
      !!ctx.xbars.s?.isStable;
  }

  const betterText =
    direction === "above"
      ? "higher values are treated as better"
      : "lower values are treated as better";

  let out =
    `The target is ${target}, and ${betterText}.`;

  if (finiteNumber(centre)) {
    const centreMeetsTarget =
      direction === "above"
        ? centre >= target
        : centre <= target;

    out +=
      centreMeetsTarget
        ? `\n\nThe current centre line (${centre.toFixed(3)}) is on the favourable side of the target.`
        : `\n\nThe current centre line (${centre.toFixed(3)}) is not yet on the favourable side of the target.`;
  }

  if (stable === true) {
    out +=
      "\n\nThe latest period does not show a clear special-cause signal, so the comparison with the target is more meaningful. " +
      "Remember that being stable does not automatically mean performance is good enough.";
  } else if (stable === false) {
    out +=
      "\n\nThe latest period shows possible special-cause variation, so target performance should be interpreted cautiously. " +
      "The priority is to understand what may have changed before assuming the current level will continue.";
  } else {
    out +=
      "\n\nInterpret the target alongside the pattern of variation rather than judging individual points on their own.";
  }

  return out;
}

function buildXmRLimitsAnswer(ctx) {
  const a = ctx.xmr;

  if (!a) {
    return (
      "XmR limits are estimated from the moving ranges between consecutive points. " +
      "Generate an XmR chart and I can explain the calculation used for your current data."
    );
  }

  let out =
    "XmR control limits depend on the estimated routine variation. " +
    "SimpleSPC estimates this from the moving ranges between consecutive observations.";

  if (a.mrScreeningEnabled === true) {
    out +=
      "\n\nFor your current chart, moving-range screening is ON.";

    if (
      finiteNumber(a.excludedMRCount) &&
      a.excludedMRCount > 0
    ) {
      const count = a.excludedMRCount;

      out +=
        ` SimpleSPC found ${count} unusually large point-to-point ` +
        `${count === 1 ? "change" : "changes"} and did not use ` +
        `${count === 1 ? "it" : "them"} when estimating routine variation. ` +
        "The original data points and moving ranges remain on the chart.";

      out +=
        "\n\nThis prevents an unusually large jump from making the estimated variation — and therefore the XmR control limits — artificially wide.";
    } else {
      out +=
        " No moving ranges in the current period exceeded the screening threshold, so screening has not changed the variation estimate for this period.";
    }

    if (
      finiteNumber(a.initialAvgMR) &&
      finiteNumber(a.mrScreeningThreshold) &&
      finiteNumber(a.avgMR)
    ) {
      out +=
        "\n\nFor the current period:\n" +
        `• initial average MR = ${a.initialAvgMR.toFixed(3)}\n` +
        `• Nelson screening threshold = ${a.mrScreeningThreshold.toFixed(3)}\n` +
        `• average MR used for the limits = ${a.avgMR.toFixed(3)}`;
    }
  } else if (a.mrScreeningEnabled === false) {
    out +=
      "\n\nFor your current chart, moving-range screening is OFF. " +
      "All baseline moving ranges are therefore used to estimate routine variation. " +
      "If one point-to-point jump is unusually large, this can increase the average moving range and produce wider control limits.";

    if (finiteNumber(a.avgMR)) {
      out +=
        `\n\nThe average MR currently being used for the limits is ${a.avgMR.toFixed(3)}.`;
    }
  } else {
    out +=
      "\n\nI cannot determine the moving-range screening setting from the current chart snapshot.";
  }

  if (a.periodCount && a.periodCount > 1) {
    out +=
      "\n\nThis chart also contains a split. Each period is treated as a separate process period with its own mean and limits, and the moving range across the split is not used. This can also make the limits different from an unsplit version of the same chart.";
  }

  if (!isXmRPeriodAssessable(a)) {
    const count = finiteNumber(a.limitEstimatePointCount)
      ? a.limitEstimatePointCount
      : null;

    const minimum = finiteNumber(a.minimumPointsForInterpretation)
      ? a.minimumPointsForInterpretation
      : 12;

    out +=
      "\n\nThe latest period is still provisional. " +
      `${count !== null ? `Only ${count} points are currently being used to estimate its limits. ` : ""}` +
      `SimpleSPC waits until at least ${minimum} points are available before making a formal stability interpretation.`;
  }

  return out;
}

function buildParetoAnswer(ctx) {
  const p =
    ctx.pareto;

  if (!p) {
    return (
      "A Pareto chart ranks categories from most to least frequent. " +
      "It can help you decide where further investigation or improvement effort may be most useful."
    );
  }

  const topName =
    p.topCategory ||
    p.items?.[0]
      ?.category ||
    "the first category";

  const topCount =
    finiteNumber(
      p.topCount
    )
      ? p.topCount
      : p.items?.[0]
          ?.count;

  const topPct =
    finiteNumber(
      p.topShare
    )
      ? p.topShare * 100
      : null;

  let out =
    `This Pareto chart contains ${p.total} occurrences across ${p.categoryCount} categories. ` +
    `The categories are ordered from most frequent to least frequent.`;

  if (
    topName &&
    finiteNumber(topCount)
  ) {
    out +=
      `\n\nThe largest category is ${topName}, with ${topCount} occurrence${
        topCount === 1
          ? ""
          : "s"
      }${
        finiteNumber(topPct)
          ? ` (${topPct.toFixed(1)}% of the total)`
          : ""
      }.`;
  }

  if (
    finiteNumber(
      p.cutoff80Count
    ) &&
    finiteNumber(
      p.cutoff80Pct
    )
  ) {
    out +=
      `\n\nThe first ${p.cutoff80Count} categor${
        p.cutoff80Count === 1
          ? "y"
          : "ies"
      } account for about ${p.cutoff80Pct.toFixed(1)}% of the occurrences. ` +
      "That can help you decide where to investigate first, but the 80% line is a prioritisation guide rather than a statistical limit.";
  }

  if (
    finiteNumber(p.total) &&
    p.total < 30
  ) {
    out +=
      "\n\nThere are fewer than 30 occurrences in this chart, so the ranking may change quite easily as more data are collected.";
  }

  out +=
    "\n\nDo not assume that the largest categories are necessarily the causes of the problem. " +
    "They are ways of grouping the data and should normally prompt further investigation.";

  out +=
    "\n\nA Pareto chart is not time ordered, so it cannot tell you whether a process is stable or whether a change is statistically significant. " +
    "Use a run or Shewhart chart alongside it when that question matters.";

  return out;
}


  function buildWhatChartTellingMe(ctx) {
  const chartType = ctx.chartType;

  if (
  chartType === "pareto"
) {
  return buildParetoAnswer(
    ctx
  );
}

if (
  chartType === "xmr" &&
  ctx.xmr
) {
  return buildXmRChangeAnswer(
    ctx
  );
}

return buildLatestPeriodSummary(
  ctx
);
}

function buildChangeAnswer(ctx) {
  const chartType = ctx.chartType;

  if (
    chartType === "pareto"
  ) {
    return (
      "A Pareto chart cannot tell whether something has changed over time because it deliberately removes the time order and ranks categories by frequency.\n\n" +
      buildParetoAnswer(ctx)
    );
  }

  if (chartType === "xmr" && ctx.xmr) {
    return buildXmRChangeAnswer(ctx);
  }

  if (chartType === "xbars" && ctx.xbars) {
    const xStable =
      !!ctx.xbars.xbar?.isStable;

    const sStable =
      !!ctx.xbars.s?.isStable;

    if (xStable && sStable) {
      return (
        "I cannot see clear evidence that the process has changed in the latest X̄–S period. " +
        "Neither the subgroup averages nor the within-subgroup variation shows a special-cause signal."
      );
    }

    if (!xStable && sStable) {
      return (
        "There are signs that the process level may have changed: the X̄ chart shows a signal, while the S chart remains stable."
      );
    }

    if (xStable && !sStable) {
      return (
        "There are signs that the process variation may have changed: the S chart shows a signal, while the X̄ chart remains stable."
      );
    }

    return (
      "There are signs that the process may have changed: both the X̄ and S charts show special-cause signals."
    );
  }

  const a =
    ctx.attribute ||
    ctx.rare;

  if (a) {
    const sigs =
      joinNice(
        signalTextList(a.signals)
      );

    if (a.isStable === true) {
      return (
        `I cannot see clear evidence that the process has changed in the latest ${titleCaseChartType(chartType).toLowerCase()} period. ` +
        "No special-cause signal is detected using the current rule settings."
      );
    }

    return (
      `There are signs that the process may have changed in the latest ${titleCaseChartType(chartType).toLowerCase()} period — specifically ` +
      `${sigs || "one or more SPC signals"}. ` +
      "That is a prompt to investigate what changed around that time."
    );
  }

  if (ctx.run) {
    if (ctx.run.isStable === true) {
      return (
        "I cannot see clear evidence that the pattern has changed on the current run chart."
      );
    }

    return (
      "There are signs that the pattern may have changed on the current run chart. " +
      "It would be worth checking what changed in the process or data around that time."
    );
  }

  return (
    "Generate the currently selected chart first and I can look for evidence of change."
  );
}

function buildImprovementAnswer(ctx) {
  const chartType = ctx.chartType;

  if (chartType === "xmr" && ctx.xmr) {
    if (!isXmRPeriodAssessable(ctx.xmr)) {
      return (
        "There are not yet enough data in the latest XmR period to make a firm judgement about stability or improvement. " +
        "The early pattern may be encouraging, but more points are needed before calling it a sustained change."
      );
    }

    if (ctx.xmr.isStable === true) {
      return (
        "The current XmR period does not show a clear special-cause signal, so the chart alone does not provide strong evidence of a new improvement. " +
        "Real improvement is more convincing when there is a sustained change in the desired direction and the new pattern is maintained."
      );
    }

    return (
      "The XmR chart shows evidence that the process may have changed. " +
      "That is not automatically the same as improvement. Check whether the change is in the desired direction, whether it is sustained, and whether local knowledge supports a plausible improvement explanation. " +
      "The chart can show evidence of change, but it cannot prove that an intervention caused it."
    );
  }

  let stable = null;

  if (ctx.attribute) {
    stable =
      ctx.attribute.isStable;
  } else if (ctx.rare) {
    stable =
      ctx.rare.isStable;
  } else if (ctx.run) {
    stable =
      ctx.run.isStable;
  } else if (ctx.xbars) {
    stable =
      !!ctx.xbars.xbar?.isStable &&
      !!ctx.xbars.s?.isStable;
  }

  if (stable === true) {
    return (
      `The current ${titleCaseChartType(chartType).toLowerCase()} does not show a clear special-cause signal, so the chart alone does not provide strong evidence of a new improvement. ` +
      "Improvement is more convincing when there is a sustained change in the desired direction, not just one favourable point."
    );
  }

  if (stable === false) {
    return (
      `The current ${titleCaseChartType(chartType).toLowerCase()} shows evidence that the process may have changed. ` +
      "That could be improvement, deterioration, or another change in the system. " +
      "To call it improvement, check that the change is in the desired direction, is sustained, and makes sense in the local context. " +
      "The chart alone does not prove causation."
    );
  }

  return (
    "Generate the currently selected chart first and I can help assess whether the pattern is consistent with real improvement."
  );
}

function buildDecisionAnswer(ctx) {

  if (
    ctx.chartType === "pareto"
  ) {
    return (
      buildParetoAnswer(ctx) +
      "\n\nA sensible next step is to investigate the leading categories in more detail, using local knowledge and, where useful, further stratification. Do not assume the category itself is the underlying cause."
    );
  }

  const latestSummary =
    buildLatestPeriodSummary(ctx);

  return latestSummary + "\n\n" +
    "From a healthcare improvement point of view, the next step is usually:\n" +
    "• if the process looks stable but performance is not good enough, think about redesigning the system\n" +
    "• if the chart shows a signal, investigate what changed in the service, pathway, demand, staffing, coding, or measurement\n" +
    "• avoid reacting to individual points as if each one proves success or failure";
}

  function answerQuestion(question) {
  const qRaw =
    String(question || "").trim();

  const q =
    norm(qRaw);

  if (!q) {
    return "Please type a question about SPC or your chart.";
  }

  const ctx =
    getCurrentChartContext();

  const asksSigmaZ =
    includesAny(q, [
      "sigma z",
      "what does sigma z mean",
      "what is sigma z",
      "explain sigma z"
    ]);

  const asksPPrime =
    includesAny(q, [
      "p prime",
      "laney p",
      "why p prime",
      "why am i using p prime",
      "why are these limits different"
    ]);

  const asksUPrime =
    includesAny(q, [
      "u prime",
      "laney u",
      "why u prime",
      "why am i using u prime",
      "why are these limits different"
    ]);

  const asksWhatChartSays =
    q === "what is my chart telling me" ||
    q === "what is the chart telling me" ||
    q === "interpret my chart" ||
    q === "interpret this chart";

  const asksChanged =
    q === "has anything changed" ||
    q === "has something changed" ||
    q === "has my process changed" ||
    q === "has the process changed" ||
    q === "is there evidence of change" ||
    q === "is there evidence that something changed" ||
    q === "has there been a significant change" ||
    q === "is this a significant change";

  const asksStable =
    q === "is my process stable" ||
    q === "is this process stable" ||
    q === "does my process look stable" ||
    q === "does this process look stable" ||
    q === "is my chart stable";

  const asksImprovement =
    q === "is this likely to be real improvement" ||
    q === "is this real improvement" ||
    q === "are we improving" ||
    q === "is this improvement" ||
    q === "does this show improvement";

  const asksTargetMeaning =
    q === "what does this mean for our target" ||
    q === "what does this mean for my target" ||
    q === "what does my chart mean for our target" ||
    q === "are we meeting the target" ||
    q === "are we meeting our target" ||
    q === "can this process meet the target" ||
    q === "how are we doing against the target";

  const asksRulesUsed =
    q === "which rules are being used" ||
    q === "what rules are being used" ||
    q === "which spc rules are being used" ||
    q === "what rules apply to this chart" ||
    q === "what is the rule policy for this chart";

  const asksXmRLimits =
    ctx.chartType === "xmr" &&
    (
      includesAny(q, [
        "moving range screening",
        "mr screening",
        "nelson",
        "large jump",
        "large jumps",
        "limits different",
        "control limits different",
        "why are my limits",
        "limit calculation",
        "3.5"
      ]) ||
      (
        q.includes("xmr") &&
        includesAny(
          q,
          [
            "limit",
            "limits",
            "calculation"
          ]
        )
      )
    );

  const asksDecision =
    q === "what should i do next" ||
    q === "what should we do next" ||
    q === "what is the next step" ||
    q === "what should the next step be";

    const hasChart =
    !!ctx.run ||
    !!ctx.xmr ||
    !!ctx.xbars ||
    !!ctx.attribute ||
    !!ctx.rare ||
    !!ctx.pareto;

  // Pareto-specific questions
  if (hasChart && ctx.chartType === "pareto") {

    if (q === "what does the 80% line mean") {
      const p = ctx.pareto;

      let answer =
        "The dashed 80% line is a guide to help prioritise " +
        "further investigation. It is not a statistical control limit. " +
        "The red cumulative line shows how the percentage of " +
        "occurrences builds up as categories are added from left to right.";

      if (
        Number.isFinite(p.cutoff80Count) &&
        Number.isFinite(p.cutoff80Pct)
      ) {
        answer +=
          `\n\nIn your chart, the first ${p.cutoff80Count} ` +
          `categories account for ${p.cutoff80Pct.toFixed(1)}% ` +
          "of all occurrences.";
      }

      answer +=
        "\n\nThe 80% line is a useful reference, not a rule " +
        "that 80% of problems must come from 20% of categories.";

      return answer;
    }

    if (q === "are the biggest categories the causes") {
      return (
        "Not necessarily. A Pareto chart shows which categories " +
        "contain the most recorded occurrences, but it does not " +
        "establish what caused them.\n\n" +
        "The largest categories are often useful starting points " +
        "for investigation. Look more closely at what happens " +
        "within those categories, consider how the data were " +
        "collected, and use local knowledge to identify possible causes.\n\n" +
        "A frequent category is not automatically a root cause, " +
        "and a less frequent category may still deserve priority " +
        "if its consequences are particularly serious."
      );
    }
  }

  if (
    hasChart &&
    asksXmRLimits
  ) {
    return buildXmRLimitsAnswer(ctx);
  }

  if (
    hasChart &&
    asksSigmaZ &&
    (
      ctx.chartType === "pprime" ||
      ctx.chartType === "uprime"
    )
  ) {
    return buildSigmaZExplanation(ctx);
  }

  if (
    hasChart &&
    ctx.chartType === "pprime" &&
    asksPPrime
  ) {
    return buildPPrimeExplanation(ctx);
  }

  if (
    hasChart &&
    ctx.chartType === "uprime" &&
    asksUPrime
  ) {
    return buildUPrimeExplanation(ctx);
  }

  if (
    hasChart &&
    asksRulesUsed
  ) {
    return describeRulesForCurrentChart(
      ctx.chartType
    );
  }

  if (
    hasChart &&
    asksWhatChartSays
  ) {
    return buildWhatChartTellingMe(ctx);
  }

  if (
    hasChart &&
    asksChanged
  ) {
    return buildChangeAnswer(ctx);
  }

  if (
    hasChart &&
    asksStable
  ) {
    return buildLatestPeriodSummary(ctx);
  }

  if (
    hasChart &&
    asksImprovement
  ) {
    return buildImprovementAnswer(ctx);
  }

  if (
    hasChart &&
    asksTargetMeaning
  ) {
    return buildTargetAnswer(ctx);
  }

  if (
    hasChart &&
    asksDecision
  ) {
    return buildDecisionAnswer(ctx);
  }

  const faq =
    findBestFaq(q);

  if (faq) {
    return faq.answer;
  }

  if (hasChart) {
    return (
      "I’m not fully sure which answer fits best. Try one of these:\n" +
      "• What is my chart telling me?\n" +
      "• Has anything changed?\n" +
      "• Is my process stable?\n" +
      "• Is this likely to be real improvement?\n" +
      "• What should I do next?\n" +
      "• What does this mean for our target?"
    );
  }

  return (
    "I’m not fully sure which answer fits best. Try one of these:\n" +
    "• What is SPC?\n" +
    "• How do I choose the right chart?\n" +
    "• What is the difference between a baseline and a split?\n" +
    "• What should I do if I see a signal?\n" +
    "• What if my process is stable but performance is poor?"
  );
}

  window.SPC_HELPER_LIBRARY = {
    getSuggestedQuestions,
    answerQuestion,
    faq: FAQ
  };

  window.matchSpcFaq = function matchSpcFaq(items, text) {
    const q = norm(text);
    let best = null;
    let bestScore = 0;

    (items || []).forEach(item => {
      const score = scoreFaqItem(item, q);
      if (score > bestScore) {
        best = item;
        bestScore = score;
      }
    });

    return best && bestScore >= 15 ? best.answer : null;
  };
})();
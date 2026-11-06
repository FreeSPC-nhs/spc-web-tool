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
  const t = helperDynamicText;

  // Use the same chart-family policy as before.
  // Translate only the explanation.

  switch (chartType) {
    case "p":
    case "pprime":
    case "u":
    case "uprime":
    case "c":
      return t(
        "helper.rules.attribute",
        "For this chart, SimpleSPC uses beyond-limit signals and the run rule. Zone rules are blocked because they may create misleading alerts on attribute charts."
      );

    case "t":
    case "g":
      return t(
        "helper.rules.rare",
        "For this chart, beyond-limit signals are used by default. Run and trend rules are advanced options only, because rare-event charts can naturally look irregular."
      );

    case "xmr":
      return t(
        "helper.rules.xmr",
        "For this XmR chart, beyond-limit signals and the run rule are used by default. Trend and zone-style rules are advanced options because adding more rules can increase false alerts."
      );

    case "xbars":
      return t(
        "helper.rules.xbars",
        "For this X̄–S chart, beyond-limit signals and the run rule are used by default. Trend and zone rules are available as advanced options."
      );

    case "run":
      return t(
        "helper.rules.run",
        "For this run chart, SimpleSPC uses run-chart rules such as shifts and, where enabled, trends. Control-limit rules do not apply to run charts."
      );

    case "pareto":
      return t(
        "helper.rules.pareto",
        "SPC signal rules are not used on a Pareto chart. A Pareto chart ranks categories rather than showing process behaviour over time."
      );

    default:
      return t(
        "helper.rules.default",
        "SimpleSPC uses a conservative, chart-aware rule set. Not every rule is available for every chart type."
      );
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

    chartType === "xmr"
      ? "Why do XmR signals disappear with a short baseline?"
      : null,

    chartType === "xmr"
      ? "Why does XmR need 12 usable points?"
      : null,

    chartType === "xmr"
      ? "Why is excluding this point from the limits disabled?"
      : null,

    chartType === "xmr"
      ? "Why doesn't Nelson screening flag changes after the baseline?"
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


/*
 * Translate a matched FAQ answer.
 *
 * Matching still uses the original English aliases,
 * keywords and priority scores. Only the returned
 * answer text changes with the selected language.
 */
function localisedFaqAnswer(faq) {
  if (!faq) return "";

  const key = `helper.faq.${faq.id}`;
  const i18n = window.SimpleSPCi18n;

  // Preserve the original answer when translation
  // support is unavailable (including older portable files).
  if (
    !i18n ||
    typeof i18n.translate !== "function"
  ) {
    return faq.answer;
  }

  const translated = i18n.translate(key);

  // A missing dictionary key must not appear as
  // user-facing text.
  if (!translated || translated === key) {
    return faq.answer;
  }

  return translated;
}


/*
 * Translate dynamic helper text while preserving
 * numbers, statistical decisions and English fallbacks.
 */
function helperDynamicText(key, englishFallback, values = {}) {
  const i18n = window.SimpleSPCi18n;

  if (i18n && typeof i18n.translate === "function") {
    const translated = i18n.translate(key, values);

    if (translated && translated !== key) {
      return translated;
    }
  }

  return String(englishFallback).replace(
    /\{([a-zA-Z0-9_]+)\}/g,
    (match, name) =>
      Object.prototype.hasOwnProperty.call(values, name)
        ? String(values[name])
        : match
  );
}

/*
 * Translate only the signal descriptions already produced
 * by the existing signalTextList() function.
 *
 * Unrecognised descriptions are preserved unchanged.
 * The original signals array is never modified.
 */

function helperXmRSignalDescription(text) {
  const original = String(text ?? "").trim();
  const normalised = original.toLowerCase();

  // Existing fixed descriptions from signalTextList().
  const keys = {
    "one or more points above the upper limit":
      "helper.xmr.signal.above",
    "one or more points below the lower limit":
      "helper.xmr.signal.below",
    "a sustained run on one side of the centre line":
      "helper.xmr.signal.run",
    "a sustained trend":
      "helper.xmr.signal.trend",
    "an unusual outlier":
      "helper.xmr.signal.outlier",
    "an unusual clustering of points far from the centre line":
      "helper.xmr.signal.cluster",
    "a sustained bias away from the centre line":
      "helper.xmr.signal.bias"
  };

  if (keys[normalised]) {
    return helperDynamicText(
      keys[normalised],
      original
    );
  }

  // Original XmR shift descriptor.
  if (
    normalised ===
    "a sustained shift (many points on the same side of the mean)"
  ) {
    return helperDynamicText(
      "helper.signal.shiftMean",
      original
    );
  }

  // Attribute-chart shift descriptors.
  // Retain the rule threshold and original date range.
  const shift =
    /^Shift: (\d+)\+ points in a row (above|below) the centre line \(from (.+?) to (.+)\)$/i
      .exec(original);

  if (shift) {
    const direction = helperDynamicText(
      shift[2].toLowerCase() === "above"
        ? "helper.signal.sideAbove"
        : "helper.signal.sideBelow",
      shift[2].toLowerCase()
    );

    return helperDynamicText(
      "helper.signal.shiftDetailed",
      "Shift: {count}+ points in a row {direction} the centre line (from {start} to {end})",
      {
        count: shift[1],
        direction,
        start: shift[3],
        end: shift[4]
      }
    );
  }

  // Attribute-chart trend descriptors.
  const trend =
    /^Trend: (\d+)\+ points steadily (increasing|decreasing) \(from (.+?) to (.+)\)$/i
      .exec(original);

  if (trend) {
    const direction = helperDynamicText(
      trend[2].toLowerCase() === "increasing"
        ? "helper.signal.directionIncreasing"
        : "helper.signal.directionDecreasing",
      trend[2].toLowerCase()
    );

    return helperDynamicText(
      "helper.signal.trendDetailed",
      "Trend: {count}+ points steadily {direction} (from {start} to {end})",
      {
        count: trend[1],
        direction,
        start: trend[3],
        end: trend[4]
      }
    );
  }

  // Do not invent translations for unknown messages.
  return original;
}

function helperXmRSignalList(signals) {
  const items = signalTextList(signals)
    .map(helperXmRSignalDescription)
    .filter(Boolean);

  if (!items.length) {
    return helperDynamicText(
      "helper.xmr.signal.other",
      "special-cause signals"
    );
  }

  if (items.length === 1) {
    return items[0];
  }

  if (items.length === 2) {
    return helperDynamicText(
      "helper.xmr.signal.joinTwo",
      "{first} and {last}",
      { first: items[0], last: items[1] }
    );
  }

  return helperDynamicText(
    "helper.xmr.signal.joinMany",
    "{first}, and {last}",
    {
      first: items.slice(0, -1).join(", "),
      last: items[items.length - 1]
    }
  );
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
  if (!Array.isArray(periods) || periods.length === 0) {
    return null;
  }

  const t = helperDynamicText;

  const pieces = periods.map((p, idx) => {
    const label = t(
      "helper.xmr.period.label",
      "Period {number}",
      { number: idx + 1 }
    );

    const start = (p.startIndex ?? 0) + 1;
    const end = (p.endIndex ?? 0) + 1;

    const mean = finiteNumber(p.mean)
      ? t(
          "helper.xmr.period.mean",
          " around a mean of {mean}",
          { mean: p.mean.toFixed(2) }
        )
      : "";

    const values = { label, start, end, mean };

    // Keep the original provisional-limit decision.
    if (!isXmRPeriodAssessable(p)) {
      const count =
        finiteNumber(p.limitEstimatePointCount)
          ? p.limitEstimatePointCount
          : null;

      const minimum =
        finiteNumber(p.minimumPointsForInterpretation)
          ? p.minimumPointsForInterpretation
          : 12;

      const countText = count !== null
        ? t(
            "helper.xmr.period.count",
            "Only {count} points contributed to the limit estimate, so ",
            { count }
          )
        : "";

      return t(
        "helper.xmr.period.provisional",
        "{label} (points {start}–{end}) has provisional XmR limits{mean}. {count}At least {minimum} usable points are required before SimpleSPC interprets stability and special-cause signals for this period",
        { ...values, count: countText, minimum }
      );
    }

    // Keep the original stable/special-cause decision.
    if (p.isStable === true) {
      return t(
        "helper.xmr.period.stable",
        "{label} (points {start}–{end}) appears stable{mean}",
        values
      );
    }

    return t(
      "helper.xmr.period.signals",
      "{label} (points {start}–{end}) shows {signals}{mean}",
      {
        ...values,
        signals: helperXmRSignalList(p.signals)
      }
    );
  });

  return pieces.join(". ") + ".";
}


 
function buildXmRChangeAnswer(ctx) {
  const latest = ctx.xmr;
  const periods = ctx.xmrPeriods;
  const t = helperDynamicText;

  // Preserve the existing early-warning checks.
  // These do not establish a formal stability conclusion.
  function provisionalHasWarningSigns() {
    if (!latest) return false;

    return (
      (finiteNumber(latest.nBeyond) && latest.nBeyond > 0) ||
      !!latest.hasRunViolation ||
      !!latest.hasTrend ||
      !!latest.hasAstronomical
    );
  }

  // MULTIPLE PERIODS
  if (Array.isArray(periods) && periods.length > 1) {
    const latestPeriod = periods[periods.length - 1];

    const latestAssessable =
      isXmRPeriodAssessable(latestPeriod);

    const earlierAssessablePeriods = periods
      .slice(0, -1)
      .filter(isXmRPeriodAssessable);

    const earlierUnstable = earlierAssessablePeriods
      .some(p => p.isStable === false);

    const start = (latestPeriod.startIndex ?? 0) + 1;
    const end = (latestPeriod.endIndex ?? 0) + 1;

    let out = earlierUnstable
      ? t(
          "helper.xmr.multi.earlier",
          "Across the XmR chart, at least one assessable period shows special-cause signals. This suggests the process may have changed.\n\n"
        )
      : t(
          "helper.xmr.multi.splits",
          "This XmR chart has multiple periods separated by splits. Each split starts a new calculation of the mean and limits, but the presence of a split does not itself prove a process change.\n\n"
        );

    out += describeXmRPeriods(periods) + "\n\n";

    // Latest period: provisional limits.
    if (!latestAssessable) {
      const count =
        finiteNumber(latestPeriod.limitEstimatePointCount)
          ? latestPeriod.limitEstimatePointCount
          : null;

      const minimum =
        finiteNumber(latestPeriod.minimumPointsForInterpretation)
          ? latestPeriod.minimumPointsForInterpretation
          : 12;

      const countText = count !== null
        ? t(
            "helper.xmr.multi.count",
            "There are currently {count} usable points in this period, and ",
            { count }
          )
        : "";

      out += t(
        "helper.xmr.multi.provisional",
        "The latest period (points {start}–{end}) does not yet have enough data for a firm stability conclusion. {count}SimpleSPC requires at least {minimum} usable points before interpreting XmR stability.",
        { start, end, count: countText, minimum }
      );

      out += provisionalHasWarningSigns()
        ? t(
            "helper.xmr.multi.warning",
            "\n\nSome early warning features are present, but this period is still provisional. More observations are needed to see whether the pattern persists."
          )
        : t(
            "helper.xmr.multi.calm",
            "\n\nNo obvious early warning features appear in this short period. That is encouraging, but does not yet establish stability."
          );

      return out;
    }

    // Latest period: enough points for interpretation.
    if (latestPeriod.isStable === true) {
      out += t(
        "helper.xmr.multi.stable",
        "In the latest period (points {start}–{end}), no clear special-cause signal is detected. With sufficient data, this is consistent with routine variation around the current level.",
        { start, end }
      );
    } else {
      out += t(
        "helper.xmr.multi.signals",
        "In the latest period (points {start}–{end}), possible special-cause variation remains: {signals}.",
        {
          start,
          end,
          signals: helperXmRSignalList(latestPeriod.signals)
        }
      );
    }

    return out;
  }

  // NO CURRENT XmR ANALYSIS
  if (!latest) {
    return t(
      "helper.xmr.single.noChart",
      "Generate an XmR chart first, then ask again."
    );
  }

  // SINGLE PERIOD: PROVISIONAL LIMITS
  if (!isXmRPeriodAssessable(latest)) {
    const count =
      finiteNumber(latest.limitEstimatePointCount)
        ? latest.limitEstimatePointCount
        : null;

    const minimum =
      finiteNumber(latest.minimumPointsForInterpretation)
        ? latest.minimumPointsForInterpretation
        : 12;

    const countText = count !== null
      ? t(
          "helper.xmr.single.count",
          "There are currently {count} usable points for limit estimation, and ",
          { count }
        )
      : "";

    let out = t(
      "helper.xmr.single.provisional",
      "There are not yet enough data to make a firm judgement about stability. {count}SimpleSPC requires at least {minimum} usable points for a formal XmR interpretation.",
      { count: countText, minimum }
    );

    out += provisionalHasWarningSigns()
      ? t(
          "helper.xmr.single.warning",
          "\n\nThere are features worth watching, but more observations are needed to determine whether they form a sustained signal."
        )
      : t(
          "helper.xmr.single.calm",
          "\n\nThe pattern currently looks reasonably consistent, with no obvious early warnings. This remains provisional until enough observations are available."
        );

    return out;
  }

  // SINGLE PERIOD: STABLE
  if (latest.isStable === true) {
    return t(
      "helper.xmr.single.stable",
      "The current XmR chart has enough data for interpretation and shows no clear special-cause signals. This is consistent with routine variation, although stable performance is not necessarily satisfactory performance."
    );
  }

  // SINGLE PERIOD: SPECIAL-CAUSE SIGNALS
  return t(
    "helper.xmr.single.signals",
    "The current XmR chart suggests that the process may have changed. The pattern is not fully explained by routine variation, so investigate what happened in the process or its measurement."
  );
}



/*
 * Batch 33 — Laney P′, U′ and Sigma Z explanations.
 *
 * Preserve the original sigmaZ values and the existing
 * > 1, < 1 and === 1 interpretation branches.
 * Only the explanatory text is translated.
 */

function buildPPrimeExplanation(ctx) {
  const a = ctx.attribute;
  const t = helperDynamicText;

  if (!a || ctx.chartType !== "pprime") {
    return t(
      "helper.laney.pprime.noChart",
      "A P′ chart is a Laney-adjusted version of a P chart. " +
      "It is used for proportion data when the amount of variation is noticeably different from what a standard P chart expects."
    );
  }

  // Use exactly the same calculated Sigma Z value.
  const sigmaZ =
    a.stats && finiteNumber(a.stats.sigmaZ)
      ? a.stats.sigmaZ
      : null;

  let out = t(
    "helper.laney.pprime.intro",
    "This P′ chart uses the same numerator, denominator, observed proportions and centre line as a standard P chart. " +
    "The difference is that it adjusts the width of the control limits so they better match the amount of variation actually seen in the data."
  );

  if (sigmaZ != null) {
    if (sigmaZ > 1) {
      out += t(
        "helper.laney.pprime.more",
        "\n\nYour data vary more than a standard P chart expects, so P′ uses wider control limits. " +
        "This helps avoid treating ordinary variation as a possible process change."
      );

    } else if (sigmaZ < 1) {
      out += t(
        "helper.laney.pprime.less",
        "\n\nYour data vary less than a standard P chart expects, so P′ uses narrower control limits. " +
        "This can make genuinely unusual variation easier to detect."
      );

    } else {
      out += t(
        "helper.laney.pprime.equal",
        "\n\nThe amount of variation is very close to what a standard P chart expects, so the P′ adjustment is very small."
      );
    }

    out += t(
      "helper.laney.pprime.factor",
      "\n\nThe Sigma Z adjustment factor for this chart is {sigmaZ}. " +
      "A value above 1 widens the standard P-chart limits; a value below 1 narrows them.",
      { sigmaZ: sigmaZ.toFixed(3) }
    );
  }

  return out;
}


function buildUPrimeExplanation(ctx) {
  const a = ctx.attribute;
  const t = helperDynamicText;

  if (!a || ctx.chartType !== "uprime") {
    return t(
      "helper.laney.uprime.noChart",
      "A U′ chart is a Laney-adjusted version of a U chart. " +
      "It is used for rate data when the amount of variation is noticeably different from what a standard U chart expects."
    );
  }

  // Preserve the original calculated value.
  const sigmaZ =
    a.stats && finiteNumber(a.stats.sigmaZ)
      ? a.stats.sigmaZ
      : null;

  let out = t(
    "helper.laney.uprime.intro",
    "This U′ chart uses the same event counts, opportunities, observed rates and centre line as a standard U chart. " +
    "The difference is that it adjusts the width of the control limits so they better match the amount of variation actually seen in the data."
  );

  if (sigmaZ != null) {
    if (sigmaZ > 1) {
      out += t(
        "helper.laney.uprime.more",
        "\n\nYour data vary more than a standard U chart expects, so U′ uses wider control limits. " +
        "This helps avoid treating ordinary variation as a possible process change."
      );

    } else if (sigmaZ < 1) {
      out += t(
        "helper.laney.uprime.less",
        "\n\nYour data vary less than a standard U chart expects, so U′ uses narrower control limits. " +
        "This can make genuinely unusual variation easier to detect."
      );

    } else {
      out += t(
        "helper.laney.uprime.equal",
        "\n\nThe amount of variation is very close to what a standard U chart expects, so the U′ adjustment is very small."
      );
    }

    out += t(
      "helper.laney.uprime.factor",
      "\n\nThe Sigma Z adjustment factor for this chart is {sigmaZ}. " +
      "A value above 1 widens the standard U-chart limits; a value below 1 narrows them.",
      { sigmaZ: sigmaZ.toFixed(3) }
    );
  }

  return out;
}


function buildSigmaZExplanation(ctx) {
  const chartType = ctx.chartType;
  const a = ctx.attribute;
  const t = helperDynamicText;

  const isPPrime = chartType === "pprime";
  const isUPrime = chartType === "uprime";

  // Generic explanation when the selected chart is
  // not a Laney P′ or U′ chart.
  if (!isPPrime && !isUPrime) {
    return t(
      "helper.laney.sigma.general",
      "Sigma Z is used on Laney P′ and U′ charts. " +
      "It is an adjustment factor that tells SimpleSPC whether the data are varying more or less than a standard P or U chart would normally expect."
    );
  }

  const sigmaZ =
    a?.stats && finiteNumber(a.stats.sigmaZ)
      ? a.stats.sigmaZ
      : null;

  // Translate the displayed names only.
  const standardChart = isPPrime
    ? t("helper.laney.chart.p", "P chart")
    : t("helper.laney.chart.u", "U chart");

  const primeChart = isPPrime
    ? t("helper.laney.chart.pprime", "P′ chart")
    : t("helper.laney.chart.uprime", "U′ chart");

  let out = t(
    "helper.laney.sigma.intro",
    "Sigma Z is the adjustment factor used by a Laney chart. " +
    "It tells SimpleSPC whether your data are varying more or less than a standard {standard} expects.",
    { standard: standardChart }
  );

  out += t(
    "helper.laney.sigma.guidance",
    "\n\nA Sigma Z of about 1 means the amount of variation is close to what the standard chart expects. " +
    "A value above 1 means there is more variation than expected, so the control limits are made wider. " +
    "A value below 1 means there is less variation than expected, so the control limits are made narrower.\n\n"
  );

  out += t(
    "helper.laney.sigma.unchanged",
    "Sigma Z does not change your original observations or the centre line. " +
    "It only adjusts the width of the control limits on the {prime}.",
    { prime: primeChart }
  );

  if (sigmaZ != null) {
    out += t(
      "helper.laney.sigma.value",
      "\n\nFor your current chart, Sigma Z = {sigmaZ}.",
      { sigmaZ: sigmaZ.toFixed(3) }
    );

    // Retain all original greater-than,
    // less-than and equal-to-one decisions.
    if (sigmaZ > 1) {
      out += t(
        "helper.laney.sigma.more",
        " This means the data are showing more variation than the standard {standard} expects, so the {prime} uses wider limits.",
        {
          standard: standardChart,
          prime: primeChart
        }
      );

    } else if (sigmaZ < 1) {
      out += t(
        "helper.laney.sigma.less",
        " This means the data are showing less variation than the standard {standard} expects, so the {prime} uses narrower limits.",
        {
          standard: standardChart,
          prime: primeChart
        }
      );

    } else {
      out += t(
        "helper.laney.sigma.equal",
        " This is exactly 1, so the Laney adjustment does not change the width of the standard {standard} limits.",
        { standard: standardChart }
      );
    }
  }

  out += t(
    "helper.laney.sigma.technical",
    "\n\nTechnical detail: SimpleSPC first allows for the amount of variation the standard chart expects at each point. " +
    "It then looks at the point-to-point changes in those standardised values. " +
    "Sigma Z summarises that extra or reduced variation into a single adjustment factor."
  );

  return out;
}



function buildRunChartAnswer(ctx) {
  const a = ctx.run;
  const t = helperDynamicText;

  if (!a) {
    return t(
      "helper.stability.run.noChart",
      "Generate a Run chart first so I can interpret it."
    );
  }

  const splitCount =
    Array.isArray(a.splits) ? a.splits.length : 0;

  // Preserve existing split-period interpretation.
  if (splitCount > 0) {
    const periodCount = splitCount + 1;

    let answer = t(
      "helper.stability.run.splitIntro",
      "This Run chart has been divided into {count} periods using splits. Each period has its own median, and SimpleSPC checks for shifts and trends separately within each period.\n\n",
      { count: periodCount }
    );

    // No within-period signal does not prove that
    // the process stayed unchanged across splits.
    if (a.isStable === true) {
      answer += t(
        "helper.stability.run.splitStable",
        "No shift or trend has been detected within the individual periods using the current rule settings. However, this does not mean that the overall process has remained unchanged.\n\n"
      );
    } else {
      answer += t(
        "helper.stability.run.splitSignals",
        "One or more periods contain a shift or trend signal that may warrant investigation.\n\n"
      );
    }

    answer += t(
      "helper.stability.run.splitAdvice",
      "Compare the medians reported for the different periods. A difference in their levels may indicate an important change in performance, even if there are no further signals within each period.\n\nA split is a boundary chosen for analysis; it does not by itself prove a statistically significant change. Consider why the split was introduced, what happened in the process around that time, and whether the new level is sustained."
    );

    return answer;
  }

  // Preserve existing unsplit Run-chart interpretation.
  if (a.isStable === true) {
    return t(
      "helper.stability.run.stable",
      "Looking at the current Run chart, no shift or trend has been detected using the selected run-chart rules."
    );
  }

  return t(
    "helper.stability.run.signals",
    "The current Run chart shows one or more shift or trend signals, suggesting that the pattern may have changed. Investigate what happened in the process around that time."
  );
}


  
function buildLatestPeriodSummary(ctx) {
  const chartType = ctx.chartType;
  const t = helperDynamicText;

  // Translate the chart name for shared attribute
  // and rare-event explanations.
  const chartName = t(
    `helper.stability.chart.${chartType}`,
    titleCaseChartType(chartType).toLowerCase()
  );

  // Translate existing signal descriptions for display.
  // Do not modify the underlying analysis or signals.
  function displaySignals(signals) {
    const items = signalTextList(signals);

    if (!items.length) {
      return t(
        "helper.stability.signalFallback",
        "one or more SPC signals"
      );
    }

    return helperXmRSignalList(signals);
  }

  // PARETO: categorisation, not stability assessment.
  if (chartType === "pareto") {
    return (
      buildParetoAnswer(ctx) +
      "\n\n" +
      t(
        "helper.stability.pareto",
        "A Pareto chart does not assess process stability because the categories are not plotted in time order."
      )
    );
  }

  // XmR: retain provisional, stable and special-cause
  // outcomes exactly as before.
  if (chartType === "xmr" && ctx.xmr) {
    const a = ctx.xmr;

    const start = (a.startIndex ?? 0) + 1;
    const end = (a.endIndex ?? 0) + 1;

    const periodText =
      (a.periodCount && a.periodCount > 1)
        ? t(
            "helper.stability.xmr.latest",
            "Looking at the latest period (Period {period} of {total}, points {start}–{end})",
            {
              period: a.periodIndex,
              total: a.periodCount,
              start,
              end
            }
          )
        : t(
            "helper.stability.xmr.single",
            "Looking at the chart (points {start}–{end})",
            { start, end }
          );

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

      const countText = count !== null
        ? t(
            "helper.stability.xmr.count",
            "There are currently {count} points in this period, and ",
            { count }
          )
        : "";

      let out = t(
        "helper.stability.xmr.provisional",
        "{scope}, there are not yet enough data to make a firm judgement about stability. {count}SimpleSPC uses at least {minimum} points before making a formal XmR stability interpretation.",
        {
          scope: periodText,
          count: countText,
          minimum
        }
      );

      if (provisionalHasWarning) {
        out += t(
          "helper.stability.xmr.warning",
          "\n\nThere are already some features in this short period that are worth watching, so I would not yet describe the process as stable. More data will help show whether these are temporary fluctuations or part of a real pattern."
        );
      } else {
        out += t(
          "helper.stability.xmr.calm",
          "\n\nSo far, the pattern looks fairly consistent and there are no obvious warning signs in this short period. That is encouraging, but it is still provisional rather than proof of stability."
        );
      }

      return out;
    }

    if (a.isStable === true) {
      return t(
        "helper.stability.xmr.stable",
        "{scope}, I cannot see a clear special-cause signal using the current rule settings. This suggests the process is behaving broadly consistently around its current level.",
        { scope: periodText }
      );
    }

    return t(
      "helper.stability.xmr.signals",
      "{scope}, this XmR chart shows possible special-cause variation — specifically {signals}.",
      {
        scope: periodText,
        signals: displaySignals(a.signals)
      }
    );
  }

  // X̄–S: preserve all four existing combinations.
  if (chartType === "xbars" && ctx.xbars) {
    const xStable =
      !!ctx.xbars.xbar?.isStable;

    const sStable =
      !!ctx.xbars.s?.isStable;

    if (xStable && sStable) {
      return t(
        "helper.stability.xbars.bothStable",
        "Looking at the latest X̄–S period, both the subgroup averages and the within-subgroup variation look stable. I cannot see a clear special-cause signal in either part of the chart."
      );
    }

    if (!xStable && sStable) {
      return t(
        "helper.stability.xbars.xChanged",
        "Looking at the latest X̄–S period, the subgroup averages show possible change, but the within-subgroup variation looks stable."
      );
    }

    if (xStable && !sStable) {
      return t(
        "helper.stability.xbars.sChanged",
        "Looking at the latest X̄–S period, the subgroup averages look stable, but the within-subgroup variation shows possible change."
      );
    }

    return t(
      "helper.stability.xbars.bothChanged",
      "Looking at the latest X̄–S period, both the subgroup averages and the within-subgroup variation show possible change."
    );
  }

  // C, P, P′, U and U′: preserve the original verdict.
  if (ctx.attribute) {
    if (ctx.attribute.isStable === true) {
      return t(
        "helper.stability.attribute.stable",
        "Looking at the latest {chart} period, I cannot see a clear special-cause signal using the current rule settings.",
        { chart: chartName }
      );
    }

    return t(
      "helper.stability.attribute.signals",
      "Looking at the latest {chart} period, there are signs of possible special-cause variation — specifically {signals}.",
      {
        chart: chartName,
        signals: displaySignals(ctx.attribute.signals)
      }
    );
  }

  // T and G: use the existing rare-event rule verdict.
  if (ctx.rare) {
    if (ctx.rare.isStable === true) {
      return t(
        "helper.stability.rare.stable",
        "Looking at the latest {chart} period, I cannot see a clear special-cause signal using the current rare-event rule settings.",
        { chart: chartName }
      );
    }

    return t(
      "helper.stability.rare.signals",
      "Looking at the latest {chart} period, there are signs of possible special-cause variation — specifically {signals}.",
      {
        chart: chartName,
        signals: displaySignals(ctx.rare.signals)
      }
    );
  }

  // Run charts share their answer with the change question.
  if (ctx.run) {
    return buildRunChartAnswer(ctx);
  }

  return t(
    "helper.stability.noChart",
    "Generate the currently selected chart first and I can give you a chart-specific interpretation."
  );
}


  
function buildTargetAnswer(ctx) {
  const chartType = ctx.chartType;
  const t = helperDynamicText;

  const chartName = t(
    `helper.stability.chart.${chartType}`,
    titleCaseChartType(chartType).toLowerCase()
  );

  // XmR: preserve all existing target, provisional-limit
  // and capability decisions.
  if (chartType === "xmr" && ctx.xmr) {
    const a = ctx.xmr;

    const hasTarget = finiteNumber(a.target);
    const hasMean = finiteNumber(a.mean);

    if (!hasTarget) {
      return t(
        "helper.target.xmr.none",
        "There is no target currently set on this XmR chart. If you add one, SimpleSPC can help interpret performance against the target alongside process stability."
      );
    }

    const target = a.target;

    const direction =
      a.direction === "above" ? "above" : "below";

    let currentPositionText = "";

    if (hasMean) {
      const meanMeetsTarget =
        direction === "above"
          ? a.mean >= target
          : a.mean <= target;

      currentPositionText = meanMeetsTarget
        ? t(
            "helper.target.xmr.meanFavourable",
            "The current period mean ({mean}) is on the favourable side of the target ({target}).",
            { mean: a.mean.toFixed(2), target }
          )
        : t(
            "helper.target.xmr.meanUnfavourable",
            "The current period mean ({mean}) is not yet on the favourable side of the target ({target}).",
            { mean: a.mean.toFixed(2), target }
          );
    }

    // XmR: provisional limits. Never estimate capability
    // or report a firm stability conclusion here.
    if (!isXmRPeriodAssessable(a)) {
      const count =
        finiteNumber(a.limitEstimatePointCount)
          ? a.limitEstimatePointCount
          : null;

      const minimum =
        finiteNumber(a.minimumPointsForInterpretation)
          ? a.minimumPointsForInterpretation
          : 12;

      let out = t(
        "helper.target.xmr.provisional",
        "The latest XmR period is still provisional, so it is too early to make a firm judgement about whether the process can reliably meet the target."
      );

      if (currentPositionText) {
        out += "\n\n" + currentPositionText;
      }

      const countText = count !== null
        ? t(
            "helper.target.xmr.count",
            "There are currently {count} points in this period, and ",
            { count }
          )
        : "";

      out += "\n\n" + t(
        "helper.target.xmr.minimum",
        "{count}SimpleSPC uses at least {minimum} points before making a formal stability or capability interpretation.",
        { count: countText, minimum }
      );

      if (hasMean) {
        const meanMeetsTarget =
          direction === "above"
            ? a.mean >= target
            : a.mean <= target;

        out += "\n\n" + (
          meanMeetsTarget
            ? t(
                "helper.target.xmr.earlyFavourable",
                "The early results are encouraging, but they do not yet show that this level of performance is predictable or sustainable."
              )
            : t(
                "helper.target.xmr.earlyUnfavourable",
                "The early pattern may still be settling. More data will show whether the process level is genuinely different and whether further improvement is needed to reach the target reliably."
              )
        );
      }

      return out;
    }

    // XmR: assessable and stable.
    if (a.isStable === true) {
      let out = t(
        "helper.target.xmr.stable",
        "The latest XmR period has enough data for interpretation and currently looks stable."
      );

      if (currentPositionText) {
        out += "\n\n" + currentPositionText;
      }

      if (
        a.capability &&
        finiteNumber(a.capability.prob)
      ) {
        // Keep the existing percentage calculation.
        const percentage =
          (a.capability.prob * 100).toFixed(1);

        out += "\n\n" + t(
          "helper.target.xmr.capability",
          "Based on the current stable process, SimpleSPC estimates that about {percentage}% of future results would meet the target, assuming the process continues to behave in the same way.",
          { percentage }
        );
      } else {
        out += "\n\n" + t(
          "helper.target.xmr.stableNoCapability",
          "Because the process appears stable, the target comparison is more meaningful: the main question is whether the current system is predictably good enough."
        );
      }

      return out;
    }

    // XmR: assessable, but special-cause signals exist.
    let out = t(
      "helper.target.xmr.unstable",
      "The latest XmR period has enough data for interpretation, but it shows special-cause variation. That means target performance should be interpreted cautiously because the process is not yet behaving predictably."
    );

    if (currentPositionText) {
      out += "\n\n" + currentPositionText;
    }

    out += "\n\n" + t(
      "helper.target.xmr.unstableAction",
      "The priority is to understand what is changing in the system, demand, staffing, pathway, coding, or measurement before relying on a capability estimate."
    );

    return out;
  }

  // Other charts: preserve the existing source of
  // the target setting.
  let target = null;

  try {
    if (typeof getTargetValue === "function") {
      target = getTargetValue();
    }
  } catch (e) {}

  if (!finiteNumber(target)) {
    return t(
      "helper.target.general.none",
      "There is no target currently set on this {chart}. If you add one, SimpleSPC can help you interpret the target alongside the pattern of variation.",
      { chart: chartName }
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

  // Keep the original chart-specific centre line
  // and stability sources unchanged.
  let centre = null;
  let stable = null;

  if (
    ctx.attribute &&
    ctx.attribute.stats &&
    finiteNumber(ctx.attribute.stats.cl)
  ) {
    centre = ctx.attribute.stats.cl;
    stable = ctx.attribute.isStable;
  } else if (
    ctx.rare &&
    ctx.rare.stats &&
    finiteNumber(ctx.rare.stats.cl)
  ) {
    centre = ctx.rare.stats.cl;
    stable = ctx.rare.isStable;
  } else if (
    ctx.run &&
    finiteNumber(ctx.run.median)
  ) {
    centre = ctx.run.median;
    stable = ctx.run.isStable;
  } else if (
    ctx.xbars &&
    ctx.xbars.stats &&
    finiteNumber(ctx.xbars.stats.xbarbar)
  ) {
    centre = ctx.xbars.stats.xbarbar;

    stable =
      !!ctx.xbars.xbar?.isStable &&
      !!ctx.xbars.s?.isStable;
  }

  const betterText =
    direction === "above"
      ? t(
          "helper.target.general.higher",
          "higher values are treated as better"
        )
      : t(
          "helper.target.general.lower",
          "lower values are treated as better"
        );

  let out = t(
    "helper.target.general.intro",
    "The target is {target}, and {direction}.",
    { target, direction: betterText }
  );

  if (finiteNumber(centre)) {
    const centreMeetsTarget =
      direction === "above"
        ? centre >= target
        : centre <= target;

    out += "\n\n" + (
      centreMeetsTarget
        ? t(
            "helper.target.general.centreFavourable",
            "The current centre line ({centre}) is on the favourable side of the target.",
            { centre: centre.toFixed(3) }
          )
        : t(
            "helper.target.general.centreUnfavourable",
            "The current centre line ({centre}) is not yet on the favourable side of the target.",
            { centre: centre.toFixed(3) }
          )
    );
  }

  if (stable === true) {
    out += "\n\n" + t(
      "helper.target.general.stable",
      "The latest period does not show a clear special-cause signal, so the comparison with the target is more meaningful. Remember that being stable does not automatically mean performance is good enough."
    );
  } else if (stable === false) {
    out += "\n\n" + t(
      "helper.target.general.unstable",
      "The latest period shows possible special-cause variation, so target performance should be interpreted cautiously. The priority is to understand what may have changed before assuming the current level will continue."
    );
  } else {
    out += "\n\n" + t(
      "helper.target.general.unknown",
      "Interpret the target alongside the pattern of variation rather than judging individual points on their own."
    );
  }

  return out;
}


/*
 * Batch 37 — Educational XmR explanations.
 *
 * Reads the existing chart context without changing it.
 * Does not perform any new SPC calculations.
 */
function buildXmREducationAnswer(ctx, topic) {
  const t = helperDynamicText;

  const a =
    ctx.chartType === "xmr"
      ? ctx.xmr
      : null;

  const minimum =
    a && finiteNumber(a.minimumPointsForInterpretation)
      ? a.minimumPointsForInterpretation
      : 12;

  const count =
    a && finiteNumber(a.limitEstimatePointCount)
      ? a.limitEstimatePointCount
      : null;

  const prefix = a
    ? ""
    : t(
        "helper.xmrEducation.noChart",
        "Generate an XmR chart to see information about your current period. The general explanation is below.\n\n"
      );

  // 1. Why signals disappear with a short baseline.
  if (topic === "short") {
    let answer = prefix + t(
      "helper.xmrEducation.short.intro",
      "SimpleSPC deliberately does not give formal special-cause interpretations when the XmR control limits are provisional. The limits may still be displayed, but an apparent signal is not reported as a confirmed SPC signal until there are enough usable observations in the limit-estimation period."
    );

    if (a && !isXmRPeriodAssessable(a)) {
      answer += "\n\n" + (
        count !== null
          ? t(
              "helper.xmrEducation.short.provisional",
              "Your latest XmR period currently has {count} usable points for estimating the limits; at least {minimum} are required for interpretation. This explains why a signal may disappear when you shorten the baseline, even if later plotted values look unusual.",
              { count, minimum }
            )
          : t(
              "helper.xmrEducation.short.provisionalUnknown",
              "The latest XmR period does not yet meet the minimum of {minimum} usable points. Its limits and stability interpretation therefore remain provisional.",
              { minimum }
            )
      );
    } else if (a && count !== null) {
      answer += "\n\n" + t(
        "helper.xmrEducation.short.assessable",
        "Your latest period has {count} usable points and meets the {minimum}-point minimum, so the selected SPC rules can now be interpreted. Not finding a signal does not necessarily mean performance is satisfactory.",
        { count, minimum }
      );
    }

    return answer;
  }

  // 2. Why twelve usable observations are required.
  if (topic === "minimum") {
    let answer = prefix + t(
      "helper.xmrEducation.minimum",
      "SimpleSPC uses a conservative minimum of {minimum} usable points for interpreting XmR limits, special-cause signals and capability. This is a safeguard against treating limits estimated from very little information as sufficiently reliable. It is a rule adopted by this tool, not a guarantee that any process with 12 points is stable.\n\n" +
      "The distinction matters: having {minimum} points somewhere on the chart is not enough. They must contribute to the relevant period's limit estimate. If the initial baseline contains only 10 points, or a split starts a short new period, the limits may remain provisional. Excluded observations are not automatically replaced by later points.",
      { minimum }
    );

    if (a && count !== null) {
      answer += t(
        "helper.xmrEducation.minimum.current",
        "\n\nFor your current period, {count} usable points contribute to the limit estimate; the minimum is {minimum}.",
        { count, minimum }
      );
    }

    return answer;
  }

  // 3. Why the right-click exclusion may be disabled.
  // No specific clicked-point context is assumed.
  if (topic === "exclusion") {
    return prefix + t(
      "helper.xmrEducation.exclusion",
      "Excluding a point from control-limit calculations is only useful if that point contributes to the limits being estimated. If you fixed the initial baseline, a later point outside that baseline does not contribute to those limits, so excluding it would have no effect. A later period created by a split has its own limit-estimation window.\n\n" +
      "SimpleSPC may also disable a new exclusion if it would leave too few included observations, or — for an XmR chart — no consecutive pair of included observations for estimating moving-range variation.\n\n" +
      "An excluded observation remains visible. Its adjacent moving ranges are omitted from the variation estimate; the tool does not join the observations across the gap to create an artificial moving range.\n\n" +
      "Only exclude an observation after investigating a known special cause. If you believe the process has moved to a lasting new level, consider whether a justified phase split is more appropriate. The right-click menu provides the specific reason when an exclusion is unavailable; this helper answer does not identify which point you clicked."
    );
  }

  // 4. Why Nelson screening does not detect later changes.
  if (topic === "nelson") {
    let answer = prefix + t(
      "helper.xmrEducation.nelson",
      "Nelson moving-range screening is a method for estimating XmR control limits, not a rule for flagging changes anywhere on the chart.\n\n" +
      "SimpleSPC examines consecutive moving ranges within the period used to estimate the limits. When screening is enabled, moving ranges strictly greater than 3.5 times the initial average moving range are omitted from the variation estimate. The observations and original moving ranges remain visible.\n\n" +
      "A jump after a fixed baseline is not part of that baseline's variation estimate, so Nelson screening will not remove or flag it. Later observations are assessed separately using the selected SPC signal rules, provided the relevant period has enough usable data for interpretation. A split starts a separate period with its own limit estimate.\n\n" +
      "Nelson screening also differs from manually excluding a point: screening omits qualifying moving ranges from the variation estimate; a manual exclusion omits an investigated observation and its adjacent moving ranges."
    );

    if (a && a.mrScreeningEnabled === true) {
      answer += t(
        "helper.xmrEducation.nelson.on",
        "\n\nOn your current XmR chart, moving-range screening is enabled."
      );
    } else if (a && a.mrScreeningEnabled === false) {
      answer += t(
        "helper.xmrEducation.nelson.off",
        "\n\nOn your current XmR chart, moving-range screening is disabled."
      );
    }

    return answer;
  }

  return prefix;
}


function buildXmRLimitsAnswer(ctx) {
  const a = ctx.xmr;
  const t = helperDynamicText;

  // No current XmR analysis.
  if (!a) {
    return t(
      "helper.xmrLimits.noChart",
      "XmR limits are estimated from the moving ranges between consecutive points. Generate an XmR chart and I can explain the calculation used for your current data."
    );
  }

  let out = t(
    "helper.xmrLimits.intro",
    "XmR control limits depend on the estimated routine variation. SimpleSPC estimates this from the moving ranges between consecutive observations."
  );

  // Screening enabled: preserve the existing setting.
  if (a.mrScreeningEnabled === true) {
    out += t(
      "helper.xmrLimits.screeningOn",
      "\n\nFor your current chart, moving-range screening is ON."
    );

    if (
      finiteNumber(a.excludedMRCount) &&
      a.excludedMRCount > 0
    ) {
      const count = a.excludedMRCount;

      // Retain the original one-versus-many decision.
      out += count === 1
        ? t(
            "helper.xmrLimits.screenedOne",
            " SimpleSPC found {count} unusually large point-to-point change and did not use it when estimating routine variation. The original data points and moving ranges remain on the chart.",
            { count }
          )
        : t(
            "helper.xmrLimits.screenedMany",
            " SimpleSPC found {count} unusually large point-to-point changes and did not use them when estimating routine variation. The original data points and moving ranges remain on the chart.",
            { count }
          );

      out += t(
        "helper.xmrLimits.whyScreen",
        "\n\nThis prevents an unusually large jump from making the estimated variation — and therefore the XmR control limits — artificially wide."
      );

    } else {
      out += t(
        "helper.xmrLimits.noneScreened",
        " No moving ranges in the current period exceeded the screening threshold, so screening has not changed the variation estimate for this period."
      );
    }

    // Preserve the exact existing three-number condition
    // and formatting to three decimal places.
    if (
      finiteNumber(a.initialAvgMR) &&
      finiteNumber(a.mrScreeningThreshold) &&
      finiteNumber(a.avgMR)
    ) {
      out += t(
        "helper.xmrLimits.statistics",
        "\n\nFor the current period:\n" +
        "• initial average MR = {initial}\n" +
        "• Nelson screening threshold = {threshold}\n" +
        "• average MR used for the limits = {average}",
        {
          initial: a.initialAvgMR.toFixed(3),
          threshold: a.mrScreeningThreshold.toFixed(3),
          average: a.avgMR.toFixed(3)
        }
      );
    }

  } else if (a.mrScreeningEnabled === false) {

    // Screening disabled: use the existing unscreened result.
    out += t(
      "helper.xmrLimits.screeningOff",
      "\n\nFor your current chart, moving-range screening is OFF. All baseline moving ranges are therefore used to estimate routine variation. If one point-to-point jump is unusually large, this can increase the average moving range and produce wider control limits."
    );

    if (finiteNumber(a.avgMR)) {
      out += t(
        "helper.xmrLimits.offAverage",
        "\n\nThe average MR currently being used for the limits is {average}.",
        { average: a.avgMR.toFixed(3) }
      );
    }

  } else {

    // Older/incomplete chart snapshots may not specify
    // the screening setting.
    out += t(
      "helper.xmrLimits.unknown",
      "\n\nI cannot determine the moving-range screening setting from the current chart snapshot."
    );
  }

  // Phase splits: preserve the existing period-count check.
  if (a.periodCount && a.periodCount > 1) {
    out += t(
      "helper.xmrLimits.splits",
      "\n\nThis chart also contains a split. Each period is treated as a separate process period with its own mean and limits, and the moving range across the split is not used. This can also make the limits different from an unsplit version of the same chart."
    );
  }

  // Provisional limits: preserve the existing safeguard.
  if (!isXmRPeriodAssessable(a)) {
    const count =
      finiteNumber(a.limitEstimatePointCount)
        ? a.limitEstimatePointCount
        : null;

    const minimum =
      finiteNumber(a.minimumPointsForInterpretation)
        ? a.minimumPointsForInterpretation
        : 12;

    out += t(
      "helper.xmrLimits.provisionalIntro",
      "\n\nThe latest period is still provisional. "
    );

    if (count !== null) {
      out += t(
        "helper.xmrLimits.provisionalCount",
        "Only {count} points are currently being used to estimate its limits. ",
        { count }
      );
    }

    out += t(
      "helper.xmrLimits.provisionalMinimum",
      "SimpleSPC waits until at least {minimum} points are available before making a formal stability interpretation.",
      { minimum }
    );
  }

  return out;
}



function buildParetoAnswer(ctx) {
  const p = ctx.pareto;
  const t = helperDynamicText;

  // No Pareto analysis currently available.
  if (!p) {
    return t(
      "helper.pareto.noChart",
      "A Pareto chart ranks categories from most to least frequent. It can help you decide where further investigation or improvement effort may be most useful."
    );
  }

  // Keep the original category and count sources.
  // Do not translate user-supplied category names.
  const topName =
    p.topCategory ||
    p.items?.[0]?.category ||
    t(
      "helper.pareto.firstCategory",
      "the first category"
    );

  const topCount =
    finiteNumber(p.topCount)
      ? p.topCount
      : p.items?.[0]?.count;

  const topPct =
    finiteNumber(p.topShare)
      ? p.topShare * 100
      : null;

  let out = t(
    "helper.pareto.summary",
    "This Pareto chart contains {total} occurrences across {categories} categories. The categories are ordered from most frequent to least frequent.",
    {
      total: p.total,
      categories: p.categoryCount
    }
  );

  // Leading category: preserve count and percentage.
  if (
    topName &&
    finiteNumber(topCount)
  ) {
    const shareText = finiteNumber(topPct)
      ? t(
          "helper.pareto.topShare",
          " ({percentage}% of the total)",
          {
            percentage: topPct.toFixed(1)
          }
        )
      : "";

    out += "\n\n" + t(
      topCount === 1
        ? "helper.pareto.topOne"
        : "helper.pareto.topMany",
      topCount === 1
        ? "The largest category is {category}, with {count} occurrence{share}."
        : "The largest category is {category}, with {count} occurrences{share}.",
      {
        category: topName,
        count: topCount,
        share: shareText
      }
    );
  }

  // Preserve the existing 80% cutoff and percentage.
  if (
    finiteNumber(p.cutoff80Count) &&
    finiteNumber(p.cutoff80Pct)
  ) {
    out += "\n\n" + t(
      p.cutoff80Count === 1
        ? "helper.pareto.cutoffOne"
        : "helper.pareto.cutoffMany",
      p.cutoff80Count === 1
        ? "The first {count} category accounts for about {percentage}% of the occurrences. That can help you decide where to investigate first, but the 80% line is a prioritisation guide rather than a statistical limit."
        : "The first {count} categories account for about {percentage}% of the occurrences. That can help you decide where to investigate first, but the 80% line is a prioritisation guide rather than a statistical limit.",
      {
        count: p.cutoff80Count,
        percentage: p.cutoff80Pct.toFixed(1)
      }
    );
  }

  // Keep the original small-sample condition.
  if (
    finiteNumber(p.total) &&
    p.total < 30
  ) {
    out += "\n\n" + t(
      "helper.pareto.smallSample",
      "There are fewer than 30 occurrences in this chart, so the ranking may change quite easily as more data are collected."
    );
  }

  out += "\n\n" + t(
    "helper.pareto.causeCaution",
    "Do not assume that the largest categories are necessarily the causes of the problem. They are ways of grouping the data and should normally prompt further investigation."
  );

  out += "\n\n" + t(
    "helper.pareto.notTimeOrdered",
    "A Pareto chart is not time ordered, so it cannot tell you whether a process is stable or whether a change is statistically significant. Use a run or Shewhart chart alongside it when that question matters."
  );

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
  const t = helperDynamicText;

  // Pareto: preserve the translation from Batch 34.
  if (chartType === "pareto") {
    return (
      t(
        "helper.pareto.changeIntro",
        "A Pareto chart cannot tell whether something has changed over time because it deliberately removes the time order and ranks categories by frequency.\n\n"
      ) +
      buildParetoAnswer(ctx)
    );
  }

  // XmR: preserve the detailed Batch 29 interpretation,
  // including provisional limits and multiple periods.
  if (chartType === "xmr" && ctx.xmr) {
    return buildXmRChangeAnswer(ctx);
  }

  // X-bar and S: retain the four original combinations.
  if (chartType === "xbars" && ctx.xbars) {
    const xStable =
      !!ctx.xbars.xbar?.isStable;

    const sStable =
      !!ctx.xbars.s?.isStable;

    if (xStable && sStable) {
      return t(
        "helper.change.xbars.bothStable",
        "I cannot see clear evidence that the process has changed in the latest X̄–S period. Neither the subgroup averages nor the within-subgroup variation shows a special-cause signal."
      );
    }

    if (!xStable && sStable) {
      return t(
        "helper.change.xbars.xOnly",
        "There are signs that the process level may have changed: the X̄ chart shows a signal, while the S chart remains stable."
      );
    }

    if (xStable && !sStable) {
      return t(
        "helper.change.xbars.sOnly",
        "There are signs that the process variation may have changed: the S chart shows a signal, while the X̄ chart remains stable."
      );
    }

    return t(
      "helper.change.xbars.bothSignals",
      "There are signs that the process may have changed: both the X̄ and S charts show special-cause signals."
    );
  }

  // Attribute and rare-event charts:
  // retain the original analysis and stable/unstable test.
  const a =
    ctx.attribute ||
    ctx.rare;

  if (a) {
    const isRare = !!ctx.rare;

    const chartName = t(
      `helper.stability.chart.${chartType}`,
      titleCaseChartType(chartType).toLowerCase()
    );

    // Batch 29 and the Batch 30 corrections already
    // translate known signal descriptions.
    // Preserve original flags and signal arrays.
    const existingSignals =
      signalTextList(a.signals).filter(Boolean);

    const sigs = existingSignals.length
      ? helperXmRSignalList(a.signals)
      : t(
          "helper.stability.signalFallback",
          "one or more SPC signals"
        );

    if (a.isStable === true) {
      return t(
        isRare
          ? "helper.change.rare.stable"
          : "helper.change.attribute.stable",
        "I cannot see clear evidence that the process has changed in the latest {chart} period. No special-cause signal is detected using the current rule settings.",
        { chart: chartName }
      );
    }

    const assessment = t(
      isRare
        ? "helper.change.rare.signals"
        : "helper.change.attribute.signals",
      "There are signs that the process may have changed in the latest {chart} period — specifically {signals}.",
      {
        chart: chartName,
        signals: sigs
      }
    );

    return assessment + " " + t(
      "helper.change.investigate",
      "That is a prompt to investigate what changed around that time."
    );
  }

  // Run chart: reuse the already translated answer.
  if (ctx.run) {
    return buildRunChartAnswer(ctx);
  }

  return t(
    "helper.change.noChart",
    "Generate the currently selected chart first and I can look for evidence of change."
  );
}



function buildImprovementAnswer(ctx) {
  const chartType = ctx.chartType;
  const t = helperDynamicText;

  const chartName = t(
    `helper.stability.chart.${chartType}`,
    titleCaseChartType(chartType).toLowerCase()
  );

  // XmR: preserve the minimum-point safeguard.
  if (chartType === "xmr" && ctx.xmr) {
    if (!isXmRPeriodAssessable(ctx.xmr)) {
      return t(
        "helper.improvement.xmr.provisional",
        "There are not yet enough data in the latest XmR period to make a firm judgement about stability or improvement. The early pattern may be encouraging, but more points are needed before calling it a sustained change."
      );
    }

    if (ctx.xmr.isStable === true) {
      return t(
        "helper.improvement.xmr.stable",
        "The current XmR period does not show a clear special-cause signal, so the chart alone does not provide strong evidence of a new improvement. Real improvement is more convincing when there is a sustained change in the desired direction and the new pattern is maintained."
      );
    }

    return t(
      "helper.improvement.xmr.unstable",
      "The XmR chart shows evidence that the process may have changed. That is not automatically the same as improvement. Check whether the change is in the desired direction, whether it is sustained, and whether local knowledge supports a plausible improvement explanation. The chart can show evidence of change, but it cannot prove that an intervention caused it."
    );
  }

  // Other chart families: retain the original stability
  // decision and its source.
  let stable = null;

  if (ctx.attribute) {
    stable = ctx.attribute.isStable;
  } else if (ctx.rare) {
    stable = ctx.rare.isStable;
  } else if (ctx.run) {
    stable = ctx.run.isStable;
  } else if (ctx.xbars) {
    stable =
      !!ctx.xbars.xbar?.isStable &&
      !!ctx.xbars.s?.isStable;
  }

  if (stable === true) {
    return t(
      "helper.improvement.general.stable",
      "The current {chart} does not show a clear special-cause signal, so the chart alone does not provide strong evidence of a new improvement. Improvement is more convincing when there is a sustained change in the desired direction, not just one favourable point.",
      { chart: chartName }
    );
  }

  if (stable === false) {
    return t(
      "helper.improvement.general.unstable",
      "The current {chart} shows evidence that the process may have changed. That could be improvement, deterioration, or another change in the system. To call it improvement, check that the change is in the desired direction, is sustained, and makes sense in the local context. The chart alone does not prove causation.",
      { chart: chartName }
    );
  }

  return t(
    "helper.improvement.general.noChart",
    "Generate the currently selected chart first and I can help assess whether the pattern is consistent with real improvement."
  );
}


function buildDecisionAnswer(ctx) {
  const t = helperDynamicText;

  // Preserve the existing Pareto-specific advice.
  // The main Pareto analysis will be translated in Batch 34.
  if (ctx.chartType === "pareto") {
    return (
      buildParetoAnswer(ctx) +
      t(
        "helper.decision.paretoNext",
        "\n\nA sensible next step is to investigate the leading categories in more detail, using local knowledge and, where useful, further stratification. Do not assume the category itself is the underlying cause."
      )
    );
  }

  // Preserve the original chart-specific assessment.
  const latestSummary =
    buildLatestPeriodSummary(ctx);

  // Only translate the general improvement advice.
  return latestSummary + "\n\n" + t(
    "helper.decision.general",
    "From a healthcare improvement point of view, the next step is usually:\n" +
    "• if the process looks stable but performance is not good enough, think about redesigning the system\n" +
    "• if the chart shows a signal, investigate what changed in the service, pathway, demand, staffing, coding, or measurement\n" +
    "• avoid reacting to individual points as if each one proves success or failure"
  );
}

  function answerQuestion(question) {
  const qRaw =
    String(question || "").trim();

  const q =
    norm(qRaw);

  
  if (!q) {
    return helperDynamicText(
      "helper.response.empty",
      "Please type a question about SPC or your chart."
    );
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

    // Exclusion questions must not be intercepted
    // by the broad "limit calculation" matcher.
    !/\b(exclude|excluded|excluding|exclusion|ghost|ghosting)\b/.test(q) &&

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

  
  // Pareto-specific questions.
  // Preserve the original English question identifiers.
  if (hasChart && ctx.chartType === "pareto") {

    if (q === "what does the 80% line mean") {
      const p = ctx.pareto;
      const t = helperDynamicText;

      let answer = t(
        "helper.pareto.eighty.explain",
        "The dashed 80% line is a guide to help prioritise further investigation. It is not a statistical control limit. The red cumulative line shows how the percentage of occurrences builds up as categories are added from left to right."
      );

      // Use the existing calculated cutoff.
      if (
        Number.isFinite(p.cutoff80Count) &&
        Number.isFinite(p.cutoff80Pct)
      ) {
        answer += "\n\n" + t(
          p.cutoff80Count === 1
            ? "helper.pareto.eighty.countOne"
            : "helper.pareto.eighty.countMany",
          p.cutoff80Count === 1
            ? "In your chart, the first {count} category accounts for {percentage}% of all occurrences."
            : "In your chart, the first {count} categories account for {percentage}% of all occurrences.",
          {
            count: p.cutoff80Count,
            percentage: p.cutoff80Pct.toFixed(1)
          }
        );
      }

      answer += "\n\n" + t(
        "helper.pareto.eighty.notRule",
        "The 80% line is a useful reference, not a rule that 80% of problems must come from 20% of categories."
      );

      return answer;
    }

    if (q === "are the biggest categories the causes") {
      return helperDynamicText(
        "helper.pareto.causes",
        "Not necessarily. A Pareto chart shows which categories contain the most recorded occurrences, but it does not establish what caused them.\n\n" +
        "The largest categories are often useful starting points for investigation. Look more closely at what happens within those categories, consider how the data were collected, and use local knowledge to identify possible causes.\n\n" +
        "A frequent category is not automatically a root cause, and a less frequent category may still deserve priority if its consequences are particularly serious."
      );
    }
  }


  // Batch 37 — Match the four new questions exactly.
  //
  // Handle these before the broad XmR/Nelson matcher,
  // which intentionally recognises many related words.
  // Existing English question identifiers are preserved.

  
  // Batch 37 matching correction.
  // Recognise a small set of common alternative phrasings.
  // These all point to the existing bilingual answers.

  const xmrEducationQuestions = new Map([

    // Why signals disappear
    [
      norm("Why do XmR signals disappear with a short baseline?"),
      "short"
    ],
    [
      norm("Why have signals disappeared?"),
      "short"
    ],
    [
      norm("Why have my signals disappeared?"),
      "short"
    ],
    [
      norm("Why did my XmR signals disappear?"),
      "short"
    ],
    [
      norm("Why did the SPC signals disappear?"),
      "short"
    ],

    // Minimum number of points
    [
      norm("Why does XmR need 12 usable points?"),
      "minimum"
    ],
    [
      norm("Why XmR needs 12 usable points"),
      "minimum"
    ],
    [
      norm("Why do I need 12 points for XmR?"),
      "minimum"
    ],
    [
      norm("Why are 12 points needed for XmR?"),
      "minimum"
    ],

    // Disabled exclusion
    [
      norm("Why is excluding this point from the limits disabled?"),
      "exclusion"
    ],
    [
      norm("Why can't I exclude this point?"),
      "exclusion"
    ],
    [
      norm("Why is point exclusion disabled?"),
      "exclusion"
    ],

    // Nelson screening
    [
      norm("Why doesn't Nelson screening flag changes after the baseline?"),
      "nelson"
    ],
    [
      norm("Why doesn't Nelson screening detect changes after the baseline?"),
      "nelson"
    ],
    [
      norm("Why doesn't Nelson catch later changes?"),
      "nelson"
    ]

  ]);


  const xmrEducationTopic =
    xmrEducationQuestions.get(q);

  if (xmrEducationTopic) {
    return buildXmREducationAnswer(
      ctx,
      xmrEducationTopic
    );
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
    return localisedFaqAnswer(faq);
  }


  
  // Batch 36 — Unmatched-question guidance.
  //
  // Reuse the existing suggested-question keys from
  // Batch 26. This keeps the French question wording
  // consistent with the clickable helper suggestions.
  //
  // Do not change the FAQ scoring or matching logic.

  const fallbackQuestions = hasChart
    ? [
        [
          "helper.q.myChart",
          "What is my chart telling me?"
        ],
        [
          "helper.q.changed",
          "Has anything changed?"
        ],
        [
          "helper.q.processStable",
          "Is my process stable?"
        ],
        [
          "helper.q.improvement",
          "Is this likely to be real improvement?"
        ],
        [
          "helper.q.next",
          "What should I do next?"
        ],
        [
          "helper.q.target",
          "What does this mean for our target?"
        ]
      ]
    : [
        [
          "helper.q.whatSpc",
          "What is SPC?"
        ],
        [
          "helper.q.chooseChart",
          "How do I choose the right chart?"
        ],
        [
          "helper.q.baselineSplit",
          "What is the difference between a baseline and a split?"
        ],
        [
          "helper.q.signalAction",
          "What should I do if I see a signal?"
        ],
        [
          "helper.q.stablePoor",
          "What if my process is stable but performance is poor?"
        ]
      ];

  const fallbackHeading = helperDynamicText(
    "helper.response.unsure",
    "I’m not fully sure which answer fits best. Try one of these:"
  );

  const fallbackSuggestions =
    fallbackQuestions.map(([key, english]) =>
      "• " + helperDynamicText(key, english)
    );

  return [
    fallbackHeading,
    ...fallbackSuggestions
  ].join("\n");

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
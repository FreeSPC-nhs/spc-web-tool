
/* SimpleSPC interface translations and locale-aware display/input utilities. */
(function () {
  "use strict";

  const STORAGE_KEY = "simplespc.language";
  const dictionaries = Object.create(null);
  let currentLanguage = "en";
  // To add another language, include its dictionary script; registerLanguage
  // will add it to the selector automatically. Native names/locales are optional.
  const languageMetadata = {
    en: { name: "English", locale: "en-GB" },
    fr: { name: "Français", locale: "fr-FR" },
    es: { name: "Español", locale: "es-ES" },
    de: { name: "Deutsch", locale: "de-DE" },
    it: { name: "Italiano", locale: "it-IT" },
    pt: { name: "Português (Portugal)", locale: "pt-PT" }
  };

  // English is the default; honour a previous manual choice when one exists.
  try {
    const saved = window.localStorage.getItem(STORAGE_KEY);
    if (saved && /^[a-z]{2}(?:-[A-Za-z]{2})?$/.test(saved)) {
      currentLanguage = saved;
    }
  } catch (error) {
    // Portable HTML files may not have access to browser storage.
  }

  
  function registerLanguage(code, entries) {
    if (!code || !entries || typeof entries !== "object") return;

    dictionaries[code] = {
      ...(dictionaries[code] || {}),
      ...entries
    };

    if (code === currentLanguage) {
      document.documentElement.lang = code;
    }

    // Also supports translations loaded after the page is ready,
    // including when scripts are restored from portable HTML.
    if (document.readyState !== "loading") {
      updateLanguageSelector();
      updatePageText();
    }
  }


  function getLocale() {
    return languageMetadata[currentLanguage]?.locale || currentLanguage || "en-GB";
  }

  function formatNumber(value, options = {}) {
    const n = Number(value);
    if (!Number.isFinite(n)) return String(value ?? "");
    return new Intl.NumberFormat(getLocale(), options).format(n);
  }

  function formatFixed(value, digits = 3) {
    const places = Math.min(12, Math.max(0, Math.trunc(Number(digits) || 0)));
    return formatNumber(value, {
      minimumFractionDigits: places,
      maximumFractionDigits: places,
      useGrouping: false
    });
  }

  // Parse numbers using the locale in which the SOURCE DATA was entered,
  // not necessarily the language currently shown in the user interface.
  // All internal calculations continue to use JavaScript Numbers.
  // Ambiguity is unavoidable: "1.234" can mean 1234 in de/es/it/pt or
  // 1.234 in English. The selected source-data language resolves this case.
  function parseNumber(value, options = {}) {
    if (typeof value === "number") return Number.isFinite(value) ? value : NaN;
    if (value == null) return NaN;
    let s = String(value).trim().replace(/\u2212/g, "-");
    if (!s) return NaN;
    s = s.replace(/\s*%\s*$/, "").trim(); // retain percentage-point semantics
    if (!s) return NaN;

    // Accept grouping spaces only in properly grouped thousands (including NBSP).
    if (/[\s\u00a0\u202f]/u.test(s)) {
      if (!/^[+-]?\d{1,3}(?:[\s\u00a0\u202f]\d{3})+(?:[.,]\d+)?$/u.test(s)) return NaN;
      s = s.replace(/[\s\u00a0\u202f]/gu, "");
    }

    const language = String(options.language || currentLanguage).split("-")[0].toLowerCase();
    const dotGroups = ["es", "de", "it", "pt"].includes(language);
    const commaGroups = language === "en";
    const dot = s.lastIndexOf(".");
    const comma = s.lastIndexOf(",");

    if (dot >= 0 && comma >= 0) {
      // Rightmost punctuation is decimal, the other must be grouped in threes.
      const decimal = dot > comma ? "." : ",";
      const group = decimal === "." ? "," : ".";
      const parts = s.split(decimal);
      if (parts.length !== 2 || !/^\d+$/.test(parts[1])) return NaN;
      const groupPattern = new RegExp("^[+-]?\\d{1,3}(?:\\" + group + "\\d{3})+$");
      if (!groupPattern.test(parts[0])) return NaN;
      s = parts[0].split(group).join("") + "." + parts[1];
    } else if (comma >= 0) {
      if (/^[+-]?\d{1,3}(?:,\d{3}){2,}$/.test(s)) {
        s = s.replace(/,/g, "");
      } else if (/^[+-]?\d{1,3},\d{3}$/.test(s) && commaGroups) {
        s = s.replace(/,/g, "");
      } else if (/^[+-]?\d+,\d+$/.test(s)) {
        s = s.replace(",", ".");
      } else return NaN;
    } else if (dot >= 0) {
      if (/^[+-]?\d{1,3}(?:\.\d{3}){2,}$/.test(s)) {
        s = s.replace(/\./g, "");
      } else if (/^[+-]?\d{1,3}\.\d{3}$/.test(s) && dotGroups) {
        s = s.replace(".", "");
      } else if (!/^[+-]?(?:\d+(?:\.\d+)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(s)) {
        return NaN;
      }
    }
    if (!/^[+-]?(?:\d+(?:\.\d+)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(s)) return NaN;
    const parsed = Number(s);
    return Number.isFinite(parsed) ? parsed : NaN;
  }

  // Input is already in percentage points: 12.5 means 12.5%, NOT 1250%.
  function formatPercent(value, digits = 1) {
    const places = Math.min(12, Math.max(0, Math.trunc(Number(digits) || 0)));
    const n = Number(value);
    if (!Number.isFinite(n)) return String(value ?? "");
    return new Intl.NumberFormat(getLocale(), {
      style: "percent",
      minimumFractionDigits: places,
      maximumFractionDigits: places,
      useGrouping: false
    }).format(n / 100);
  }

  function updateLanguageSelector() {
    const selector = document.getElementById("languageSelect");
    if (!selector) return;
    Object.keys(dictionaries).forEach(code => {
      if (Array.from(selector.options).some(option => option.value === code)) return;
      const option = document.createElement("option");
      option.value = code;
      option.textContent = languageMetadata[code]?.name || code;
      selector.appendChild(option);
    });
    // Remove unavailable languages so the menu reflects loaded dictionaries.
    Array.from(selector.options).forEach(option => {
      if (!Object.prototype.hasOwnProperty.call(dictionaries, option.value)) option.remove();
    });
    selector.value = currentLanguage;
  }

  function translate(key, values = {}) {
    const text = dictionaries[currentLanguage]?.[key]
      ?? dictionaries.en?.[key]
      ?? key;

    return String(text).replace(/\{([a-zA-Z0-9_]+)\}/g, (match, name) =>
      Object.prototype.hasOwnProperty.call(values, name)
        ? String(values[name])
        : match
    );
  }

  function setLanguage(code, options = {}) {
    if (!Object.prototype.hasOwnProperty.call(dictionaries, code)) return false;

    currentLanguage = code;
    document.documentElement.lang = code;

    if (options.remember !== false) {
      try {
        window.localStorage.setItem(STORAGE_KEY, code);
      } catch (error) {
        // Continue without persistence if browser storage is unavailable.
      }
    }

    document.dispatchEvent(new CustomEvent("simplespc:languagechange", {
      detail: { language: code }
    }));
    return true;
  }


  
  // Update interface text and translated attributes.
  function updatePageText() {

    // Simple text elements only: do not mark elements
    // containing icons, inputs or other nested controls.
    if (document.querySelector("title[data-i18n]")) {
      document.title = translate(document.querySelector("title[data-i18n]").getAttribute("data-i18n"));
    }
    document.querySelectorAll("[data-i18n]").forEach(element => {
      const key = element.getAttribute("data-i18n");

      if (key) {
        element.textContent = translate(key);
      }
    });

    // Translate tooltips, placeholders and accessibility labels.
    ["title", "placeholder", "aria-label"].forEach(attribute => {
      const marker = "data-i18n-" + attribute;

      document.querySelectorAll("[" + marker + "]").forEach(element => {
        const key = element.getAttribute(marker);

        if (key) {
          element.setAttribute(attribute, translate(key));
        }
      });
    });

    // Only languages still under translation show a draft notice.
    const draftNotice = document.getElementById("translationDraftNotice");
    if (draftNotice) {
      const warning = languageMetadata[currentLanguage]?.draftNotice;
      draftNotice.textContent = warning || "";
      draftNotice.hidden = !warning;
    }

    // Keep the language selector synchronised.
    const selector = document.getElementById("languageSelect");

    if (selector) {
      selector.value = currentLanguage;
    }
  }

  function initialisePageLanguage() {

    // Use English if a previously remembered language
    // is not available in this version of the application.
    if (
      !Object.prototype.hasOwnProperty.call(dictionaries, currentLanguage) &&
      Object.prototype.hasOwnProperty.call(dictionaries, "en")
    ) {
      currentLanguage = "en";
    }

    document.documentElement.lang = currentLanguage;

    updateLanguageSelector();
    const selector = document.getElementById("languageSelect");

    if (selector) {
      selector.value = currentLanguage;

      selector.addEventListener("change", () => {
        if (!setLanguage(selector.value)) {
          selector.value = currentLanguage;
        }
      });
    }

    updatePageText();
  }

  // Initialise once the HTML has loaded.
  if (document.readyState === "loading") {
    document.addEventListener(
      "DOMContentLoaded",
      initialisePageLanguage
    );
  } else {
    initialisePageLanguage();
  }

  // Update the interface whenever the language changes.
  document.addEventListener(
    "simplespc:languagechange",
    updatePageText
  );


  window.SimpleSPCi18n = {
    registerLanguage,
    translate,
    setLanguage,
    getLanguage: () => currentLanguage,
    getLocale,
    formatNumber,
    formatFixed,
    formatPercent,
    parseNumber,
    getAvailableLanguages: () => Object.keys(dictionaries)
  };
})();


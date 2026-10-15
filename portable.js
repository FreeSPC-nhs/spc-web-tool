/*
 * FreeSPC portable HTML support
 * --------------------------------
 * This file adds a "Save portable chart" capability to the SPC tool.
 *
 * It is designed to:
 *   1. Reuse the existing collectProjectFile() function to collect the
 *      current SPC data and settings.
 *   2. Package the application's HTML, CSS and JavaScript into one .html file.
 *   3. Embed the current SPC project inside that file.
 *   4. Reopen the saved project automatically when the .html file is opened.
 *   5. Allow a portable file to be edited and saved again while offline.
 *
 * Expected HTML button:
 *   <button id="savePortableBtn" type="button" disabled>
 *     Save portable chart
 *   </button>
 *
 * Expected script order near the end of index.html:
 *   <script src="spc-helper-library.js"></script>
 *   <script src="./js/modals.js"></script>
 *   <script src="spc.js"></script>
 *   <script src="./portable.js"></script>
 *
 * The normal JSON save/open functions are not changed by this file.
 */

(function () {
  "use strict";

  var PORTABLE_FORMAT = "freespc-portable-html";
  var PORTABLE_VERSION = 1;

  var BUNDLE_ELEMENT_ID = "spcPortableBundle";
  var PROJECT_ELEMENT_ID = "embeddedSpcProject";
  var BOOTSTRAP_ELEMENT_ID = "spcPortableBootstrap";

  var PORTABLE_BUTTON_ID = "savePortableBtn";
  var EXISTING_SAVE_BUTTON_ID = "exportSettingsBtn";

  /*
   * The bootstrap code is written into every portable HTML file.
   *
   * The portable file contains:
   *   - a clean copy of the application HTML;
   *   - JSON containing all CSS/JavaScript source files;
   *   - JSON containing the saved SPC project;
   *   - this small bootstrap script.
   *
   * The bootstrap recreates the original CSS and JavaScript in the same
   * order when the portable file is opened.
   */
  var PORTABLE_BOOTSTRAP_SOURCE = String.raw`
(function () {
  "use strict";

  function parseJsonElement(id) {
    var element = document.getElementById(id);

    if (!element) {
      throw new Error("Portable file is missing " + id + ".");
    }

    return JSON.parse(element.textContent || "");
  }

  function addStyles(bundle) {
    (bundle.styles || []).forEach(function (item) {
      var style = document.createElement("style");

      if (item.media) {
        style.media = item.media;
      }

      style.setAttribute("data-portable-source", item.href || "embedded-style");
      style.textContent = item.code || "";
      document.head.appendChild(style);
    });
  }

  function addScripts(bundle) {
    (bundle.scripts || []).forEach(function (item) {
      var script = document.createElement("script");
      var attributes = item.attributes || {};

      Object.keys(attributes).forEach(function (name) {
        /*
         * External URLs are deliberately not restored. The code itself is
         * already embedded in the portable file.
         *
         * async/defer are also omitted so scripts execute in their original
         * order as they are appended.
         */
        if (
          name.toLowerCase() === "src" ||
          name.toLowerCase() === "async" ||
          name.toLowerCase() === "defer"
        ) {
          return;
        }

        script.setAttribute(name, attributes[name]);
      });

      script.setAttribute(
        "data-portable-source",
        item.src || item.label || "embedded-script"
      );

      script.textContent = item.code || "";
      document.body.appendChild(script);
    });
  }

  try {
    var bundle = parseJsonElement("spcPortableBundle");

    if (
      !bundle ||
      bundle.format !== "freespc-portable-bundle" ||
      !Array.isArray(bundle.scripts)
    ) {
      throw new Error("This is not a recognised FreeSPC portable bundle.");
    }

    /*
     * portable.js uses this when a portable file is saved again offline.
     * Reusing the stored clean bundle means no network access is required.
     */
    window.__SPC_PORTABLE_BUNDLE__ = bundle;

    addStyles(bundle);
    addScripts(bundle);
  } catch (error) {
    console.error("FreeSPC portable bootstrap error:", error);

    window.setTimeout(function () {
      alert(
        "The saved FreeSPC file could not start correctly.\n\n" +
        "Technical detail: " +
        error.message
      );
    }, 0);
  }
})();
`;

  function isExecutableScript(script) {
    var type = (script.getAttribute("type") || "").trim().toLowerCase();

    /*
     * Leave JSON/data/template script elements in the shell.
     * Bundle only JavaScript that would normally execute.
     */
    if (!type) return true;

    return (
      type === "text/javascript" ||
      type === "application/javascript" ||
      type === "application/ecmascript" ||
      type === "text/ecmascript" ||
      type === "module"
    );
  }

  function attributesToObject(element) {
    var result = {};

    Array.prototype.forEach.call(element.attributes || [], function (attribute) {
      result[attribute.name] = attribute.value;
    });

    return result;
  }

  function makeJsonSafeForHtml(value) {
    /*
     * JSON placed inside a <script type="application/json"> element must not
     * contain a literal </script> sequence. Escaping "<" also protects any
     * user-entered text which happens to contain HTML-like content.
     */
    return JSON.stringify(value)
      .replace(/</g, "\\u003c")
      .replace(/>/g, "\\u003e")
      .replace(/&/g, "\\u0026")
      .replace(/\u2028/g, "\\u2028")
      .replace(/\u2029/g, "\\u2029");
  }

  function escapeClosingScriptSequences(source) {
    return String(source || "").replace(/<\/script/gi, "<\\/script");
  }

  async function fetchText(url, description) {
    var response = await fetch(url, {
      cache: "no-store"
    });

    if (!response.ok) {
      throw new Error(
        "Could not read " +
          description +
          " (" +
          response.status +
          " " +
          response.statusText +
          ")."
      );
    }

    return response.text();
  }

  function removeOldPortableElements(doc) {
    [
      BUNDLE_ELEMENT_ID,
      PROJECT_ELEMENT_ID,
      BOOTSTRAP_ELEMENT_ID
    ].forEach(function (id) {
      var element = doc.getElementById(id);

      if (element) {
        element.remove();
      }
    });
  }

  function normaliseIndexUrl() {
    var url = new URL(window.location.href);

    url.hash = "";

    /*
     * GitHub Pages commonly serves index.html when the address ends in "/".
     * Fetching that URL is fine, so there is no need to force "index.html".
     */
    return url.href;
  }

  async function createBundleFromOnlineApplication() {
    var indexUrl = normaliseIndexUrl();

    var indexSource = await fetchText(
      indexUrl,
      "the FreeSPC application page"
    );

    var parser = new DOMParser();
    var doc = parser.parseFromString(indexSource, "text/html");

    if (!doc || !doc.documentElement || !doc.body) {
      throw new Error("Could not create a clean copy of the FreeSPC page.");
    }

    removeOldPortableElements(doc);

    var styles = [];
    var stylesheetLinks = Array.prototype.slice.call(
      doc.querySelectorAll('link[rel~="stylesheet"][href]')
    );

    for (var i = 0; i < stylesheetLinks.length; i += 1) {
      var link = stylesheetLinks[i];
      var href = link.getAttribute("href");
      var absoluteHref = new URL(href, indexUrl).href;

      var cssCode = await fetchText(
        absoluteHref,
        "stylesheet " + href
      );

      styles.push({
        href: href,
        media: link.getAttribute("media") || "",
        code: cssCode
      });

      link.remove();
    }

    var scripts = [];
    var scriptElements = Array.prototype.slice.call(
      doc.querySelectorAll("script")
    );

    for (var j = 0; j < scriptElements.length; j += 1) {
      var script = scriptElements[j];

      if (!isExecutableScript(script)) {
        continue;
      }

      var src = script.getAttribute("src");
      var code = "";

      if (src) {
        var absoluteSrc = new URL(src, indexUrl).href;

        code = await fetchText(
          absoluteSrc,
          "script " + src
        );
      } else {
        code = script.textContent || "";
      }

      var attrs = attributesToObject(script);

      /*
       * "src", async and defer do not belong in the portable runtime.
       * They are removed again by the bootstrap as a safety measure.
       */
      delete attrs.src;
      delete attrs.async;
      delete attrs.defer;

      scripts.push({
        src: src || "",
        label: src ? "" : "inline-script-" + (j + 1),
        attributes: attrs,
        code: code
      });

      script.remove();
    }

    /*
     * shellHtml is a clean, non-running copy of index.html. It contains the
     * interface markup but not the external stylesheet/script tags that have
     * just been bundled.
     *
     * Keeping this clean shell inside the bundle is what allows a portable
     * chart to save another portable chart while completely offline.
     */
    var shellHtml = "<!DOCTYPE html>\n" + doc.documentElement.outerHTML;

    return {
      format: "freespc-portable-bundle",
      version: PORTABLE_VERSION,
      createdAt: new Date().toISOString(),
      sourceUrl: indexUrl,
      shellHtml: shellHtml,
      styles: styles,
      scripts: scripts
    };
  }

  var applicationBundlePromise = null;

async function getApplicationBundle() {
  /*
   * Portable files already contain the complete application bundle,
   * so reuse it directly when working offline.
   */
  if (
    window.__SPC_PORTABLE_BUNDLE__ &&
    window.__SPC_PORTABLE_BUNDLE__.format === "freespc-portable-bundle"
  ) {
    return window.__SPC_PORTABLE_BUNDLE__;
  }

  /*
   * On the normal web version, create the bundle only once.
   * Keeping the Promise also means multiple callers can wait for the
   * same background operation rather than rebuilding everything.
   */
  if (!applicationBundlePromise) {
    applicationBundlePromise = createBundleFromOnlineApplication();
  }

  return applicationBundlePromise;
}

  function collectCurrentProject() {
    if (typeof collectProjectFile !== "function") {
      throw new Error(
        "FreeSPC could not find collectProjectFile(). " +
        "Check that portable.js is loaded after spc.js."
      );
    }

    var project = collectProjectFile();

    if (!project || typeof project !== "object") {
      throw new Error(
        "FreeSPC could not collect the current chart data and settings."
      );
    }

    return project;
  }

  function createPortableFilename() {
    var date = new Date();
    var yyyy = String(date.getFullYear());
    var mm = String(date.getMonth() + 1).padStart(2, "0");
    var dd = String(date.getDate()).padStart(2, "0");

    /*
     * Keep the first version deliberately predictable. We can later replace
     * this with the chart title once the portable workflow has been tested.
     */
    return "FreeSPC-chart-" + yyyy + "-" + mm + "-" + dd + ".html";
  }

  function appendJsonElement(doc, id, value) {
    var element = doc.createElement("script");

    element.id = id;
    element.type = "application/json";
    element.textContent = makeJsonSafeForHtml(value);

    doc.body.appendChild(element);
  }

  function buildPortableHtml(bundle, project) {
    if (!bundle || !bundle.shellHtml) {
      throw new Error("The portable application bundle is incomplete.");
    }

    var parser = new DOMParser();
    var doc = parser.parseFromString(bundle.shellHtml, "text/html");

    if (!doc || !doc.documentElement || !doc.body) {
      throw new Error("Could not build the portable FreeSPC document.");
    }

    removeOldPortableElements(doc);

    /*
     * The two JSON elements are deliberately placed before the bootstrap:
     * the bootstrap needs the application bundle immediately, while
     * portable.js reads the project after the app has initialised.
     */
    appendJsonElement(doc, BUNDLE_ELEMENT_ID, bundle);
    appendJsonElement(doc, PROJECT_ELEMENT_ID, {
      format: PORTABLE_FORMAT,
      version: PORTABLE_VERSION,
      savedAt: new Date().toISOString(),
      project: project
    });

    var bootstrap = doc.createElement("script");

    bootstrap.id = BOOTSTRAP_ELEMENT_ID;
    bootstrap.textContent = escapeClosingScriptSequences(
      PORTABLE_BOOTSTRAP_SOURCE
    );

    doc.body.appendChild(bootstrap);

    return "<!DOCTYPE html>\n" + doc.documentElement.outerHTML;
  }

  function downloadTextFile(contents, filename, mimeType) {
    var blob = new Blob([contents], {
      type: mimeType || "text/plain;charset=utf-8"
    });

    var url = URL.createObjectURL(blob);
    var link = document.createElement("a");

    link.href = url;
    link.download = filename;

    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    window.setTimeout(function () {
      URL.revokeObjectURL(url);
    }, 1000);
  }

  function setPortableButtonBusy(button, busy) {
    if (!button) return;

    if (busy) {
      button.dataset.previousText = button.textContent;
      button.textContent = "Creating portable file…";
      button.disabled = true;
      return;
    }

    if (button.dataset.previousText) {
      button.textContent = button.dataset.previousText;
      delete button.dataset.previousText;
    }

    syncPortableButtonState();
  }

  async function savePortableChart() {
    var button = document.getElementById(PORTABLE_BUTTON_ID);

    try {
      setPortableButtonBusy(button, true);

      var project = collectCurrentProject();

      var confirmed = window.confirm(
        "Save an editable offline copy of this SPC chart?\n\n" +
        "The HTML file will contain the underlying chart data as well as " +
        "the FreeSPC application. Store and share the saved file in line " +
        "with your organisation's information-governance requirements."
      );

      if (!confirmed) {
        return;
      }

      /*
       * On the normal web application this gathers all local application
       * files. On an already-portable file it simply reuses the stored bundle.
       */
      var bundle = await getApplicationBundle();

      var html = buildPortableHtml(bundle, project);

      downloadTextFile(
        html,
        createPortableFilename(),
        "text/html;charset=utf-8"
      );
    } catch (error) {
      console.error("FreeSPC portable save error:", error);

      alert(
        "The portable FreeSPC file could not be created.\n\n" +
        "Technical detail: " +
        error.message
      );
    } finally {
      setPortableButtonBusy(button, false);
    }
  }

  function parseEmbeddedProject() {
    var element = document.getElementById(PROJECT_ELEMENT_ID);

    if (!element) {
      return null;
    }

    var wrapper = JSON.parse(element.textContent || "");

    if (
      !wrapper ||
      wrapper.format !== PORTABLE_FORMAT ||
      !wrapper.project
    ) {
      throw new Error(
        "This file does not contain a recognised FreeSPC portable project."
      );
    }

    return wrapper.project;
  }

  function loadEmbeddedProjectIfPresent() {
    var project;

    try {
      project = parseEmbeddedProject();
    } catch (error) {
      console.error("FreeSPC portable project read error:", error);

      alert(
        "The saved SPC chart could not be read.\n\n" +
        "The file may be damaged or may have been created by an incompatible version.\n\n" +
        "Technical detail: " +
        error.message
      );

      return;
    }

    if (!project) {
      return;
    }

    if (typeof loadProjectObject !== "function") {
      console.error(
        "FreeSPC portable project load error: loadProjectObject() was not found."
      );

      alert(
        "The saved SPC chart could not be loaded because the project loader " +
        "was not available."
      );

      return;
    }

    try {
      /*
       * loadProjectObject() may be synchronous or asynchronous. Wrapping the
       * result with Promise.resolve supports either form.
       */
      Promise.resolve(loadProjectObject(project)).catch(function (error) {
        console.error("FreeSPC portable project load error:", error);

        alert(
          "The saved SPC chart could not be loaded.\n\n" +
          "Technical detail: " +
          error.message
        );
      });
    } catch (error) {
      console.error("FreeSPC portable project load error:", error);

      alert(
        "The saved SPC chart could not be loaded.\n\n" +
        "Technical detail: " +
        error.message
      );
    }
  }

  function syncPortableButtonState() {
    var portableButton = document.getElementById(PORTABLE_BUTTON_ID);

    if (!portableButton) {
      return;
    }

    var normalSaveButton =
      document.getElementById(EXISTING_SAVE_BUTTON_ID);

    /*
     * Mirror the existing JSON "Save chart" button where possible. This means
     * the new portable button becomes available at the same point as the
     * existing save function, without needing changes inside spc.js.
     */
    if (normalSaveButton) {
      portableButton.disabled = !!normalSaveButton.disabled;
    } else {
      /*
       * Fallback for testing if the normal save button ID changes.
       */
      portableButton.disabled = false;
    }
  }

  function wirePortableButton() {
    var portableButton = document.getElementById(PORTABLE_BUTTON_ID);

    if (!portableButton) {
      console.warn(
        "FreeSPC portable support is loaded, but #" +
          PORTABLE_BUTTON_ID +
          " was not found."
      );
      return;
    }

    portableButton.addEventListener("click", savePortableChart);

    var normalSaveButton =
      document.getElementById(EXISTING_SAVE_BUTTON_ID);

    if (normalSaveButton && typeof MutationObserver !== "undefined") {
      var observer = new MutationObserver(function () {
        syncPortableButtonState();
      });

      observer.observe(normalSaveButton, {
        attributes: true,
        attributeFilter: ["disabled"]
      });
    }

    /*
     * Also resynchronise after common data-entry events. This is harmless
     * redundancy and helps if the existing application updates its button
     * state without changing the attribute immediately.
     */
    document.addEventListener("change", function () {
      window.setTimeout(syncPortableButtonState, 0);
    });

    document.addEventListener("input", function () {
      window.setTimeout(syncPortableButtonState, 0);
    });

    syncPortableButtonState();
  }

  function initialisePortableSupport() {
  wirePortableButton();

  /*
   * Let the normal SPC startup code complete first.
   */
  window.setTimeout(loadEmbeddedProjectIfPresent, 100);

  /*
   * On the normal online version, prepare the portable application bundle
   * quietly in the background. This makes the later Save portable chart
   * action much faster because the CSS and JavaScript have already been read.
   *
   * A portable/offline copy already contains its bundle and does not need
   * to do this.
   */
  if (!window.__SPC_PORTABLE_BUNDLE__) {
    window.setTimeout(function () {
      getApplicationBundle().catch(function (error) {
        /*
         * Do not interrupt normal use of the SPC tool if background
         * preparation fails. Save portable chart will try again later
         * and display an error if necessary.
         */
        console.warn(
          "FreeSPC portable bundle could not be prepared in advance:",
          error
        );

        applicationBundlePromise = null;
      });
    }, 500);
  }
}

  if (document.readyState === "loading") {
    document.addEventListener(
      "DOMContentLoaded",
      initialisePortableSupport
    );
  } else {
    initialisePortableSupport();
  }
})();
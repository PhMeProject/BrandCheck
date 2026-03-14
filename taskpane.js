/* =========================================================
   BrandCheck – task pane logic
   Depends on: Office.js (loaded by taskpane.html)
   ========================================================= */

"use strict";

// ── In-memory settings ────────────────────────────────────
// Persists for the lifetime of the task pane session.
// Exported as a plain object so the checker can read it directly.
//
//   fontName      – required font face; "" disables this rule
//   fontSize      – required point size; 0 disables this rule
//   allowedColors – hex strings (e.g. "#000000"); [] disables color rule
//   boldAllowed   – when false, bold paragraphs are flagged
//   italicAllowed – when false, italic paragraphs are flagged
const settings = {
  fontName:      "",
  fontSize:      0,
  allowedColors: [],
  boldAllowed:   true,
  italicAllowed: true,
};

// ── DOM refs ───────────────────────────────────────────────
const viewChecker      = document.getElementById("view-checker");
const viewSettings     = document.getElementById("view-settings");
const btnOpenSettings  = document.getElementById("btn-open-settings");
const btnCloseSettings = document.getElementById("btn-close-settings");
const btnCheck         = document.getElementById("btn-check");
const statusEl         = document.getElementById("status");
const resultsEl        = document.getElementById("results");

const sFontName        = document.getElementById("s-font-name");
const sFontSize        = document.getElementById("s-font-size");
const colorListEl      = document.getElementById("color-list");
const btnAddColor      = document.getElementById("btn-add-color");
const sAllowBold       = document.getElementById("s-allow-bold");
const sAllowItalic     = document.getElementById("s-allow-italic");
const btnSaveSettings  = document.getElementById("btn-save-settings");
const settingsStatusEl = document.getElementById("settings-status");

// ── Office initialisation ─────────────────────────────────
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    btnCheck.addEventListener("click", runBrandCheck);
    btnOpenSettings.addEventListener("click", openSettings);
    btnCloseSettings.addEventListener("click", closeSettings);
    btnSaveSettings.addEventListener("click", saveSettings);
    btnAddColor.addEventListener("click", () => appendColorEntry("#000000"));

    populateSettingsForm();
    setStatus("Configure settings (⚙), then click Check Document.");
  } else {
    setStatus("This add-in only works inside Microsoft Word.", true);
    btnCheck.disabled = true;
  }
});

// =========================================================
// View switching
// =========================================================

function openSettings() {
  populateSettingsForm();
  viewChecker.classList.add("hidden");
  viewSettings.classList.remove("hidden");
}

function closeSettings() {
  viewSettings.classList.add("hidden");
  viewChecker.classList.remove("hidden");
}

// =========================================================
// Settings: populate form from settings object
// =========================================================

function populateSettingsForm() {
  sFontName.value      = settings.fontName;
  sFontSize.value      = settings.fontSize > 0 ? settings.fontSize : "";
  sAllowBold.checked   = settings.boldAllowed;
  sAllowItalic.checked = settings.italicAllowed;

  colorListEl.innerHTML = "";
  for (const hex of settings.allowedColors) {
    appendColorEntry(hex);
  }
}

// =========================================================
// Settings: save form values back into the settings object
// =========================================================

function saveSettings() {
  settings.fontName      = sFontName.value.trim();
  settings.fontSize      = parseFloat(sFontSize.value) || 0;
  settings.boldAllowed   = sAllowBold.checked;
  settings.italicAllowed = sAllowItalic.checked;

  settings.allowedColors = Array.from(
    colorListEl.querySelectorAll(".color-entry__input")
  ).map((input) => input.value.toLowerCase());

  settingsStatusEl.textContent = "Settings saved.";
  settingsStatusEl.className   = "status status--ok";
  setTimeout(() => {
    settingsStatusEl.textContent = "";
    settingsStatusEl.className   = "status";
  }, 2000);
}

// =========================================================
// Color-list management
// =========================================================

function appendColorEntry(hex) {
  const entry = document.createElement("div");
  entry.className = "color-entry";

  const colorInput = document.createElement("input");
  colorInput.type      = "color";
  colorInput.value     = hex;
  colorInput.className = "color-entry__input";

  const label = document.createElement("span");
  label.className   = "color-entry__label";
  label.textContent = hex;

  colorInput.addEventListener("input", () => {
    label.textContent = colorInput.value;
  });

  const removeBtn = document.createElement("button");
  removeBtn.type      = "button";
  removeBtn.className = "btn btn--sm btn--ghost color-entry__remove";
  removeBtn.textContent = "✕";
  removeBtn.setAttribute("aria-label", "Remove color");
  removeBtn.addEventListener("click", () => entry.remove());

  entry.appendChild(colorInput);
  entry.appendChild(label);
  entry.appendChild(removeBtn);
  colorListEl.appendChild(entry);
}

// =========================================================
// Brand check – main entry point
// =========================================================

async function runBrandCheck() {
  btnCheck.disabled = true;
  clearResults();
  setStatus("Scanning document…");

  try {
    const violations = await Word.run(async (context) => {
      // ── 1. Load all paragraphs ─────────────────────────
      const paras = context.document.body.paragraphs;
      context.load(paras, "items");
      await context.sync();

      // ── 2. Load font properties for every paragraph ────
      // A single extra sync fetches everything in one round-trip.
      paras.items.forEach((para) =>
        context.load(para, "text, font/name, font/size, font/color, font/bold, font/italic")
      );
      await context.sync();

      // ── 3. Run rules against each non-empty paragraph ──
      const found = [];
      paras.items.forEach((para, idx) => {
        if (!para.text || para.text.trim() === "") return;
        checkParagraph(para, idx, found);
      });

      return found;
    });

    if (violations.length === 0) {
      setStatus("No violations found. Looks great!");
    } else {
      setStatus(
        `Found ${violations.length} violation${violations.length !== 1 ? "s" : ""}.`
      );
      renderViolations(violations);
    }
  } catch (err) {
    setStatus(`Error: ${err.message}`, true);
    console.error(err);
  } finally {
    btnCheck.disabled = false;
  }
}

// =========================================================
// Rule engine – check one paragraph against current settings
//
// Each violation pushed to `out` has the shape:
//   {
//     ruleId:         string   — machine-readable identifier
//     message:        string   — plain-English description
//     paragraphIndex: number   — 0-based paragraph position
//     fix: {
//       type:  "setFont" | "setFontSize" | "setFontColor" | "setBold" | "setItalic"
//       value: string | number | boolean  — the correct value from settings
//     }
//   }
// =========================================================

function checkParagraph(para, idx, out) {
  const f     = para.font;
  const label = `Paragraph ${idx + 1}`;

  // ── Font name ────────────────────────────────────────
  if (settings.fontName && f.name && f.name !== settings.fontName) {
    out.push({
      ruleId:         "wrong-font",
      message:        `${label} uses "${f.name}" — brand requires "${settings.fontName}"`,
      paragraphIndex: idx,
      fix: { type: "setFont", value: settings.fontName },
    });
  }

  // ── Font size ────────────────────────────────────────
  if (settings.fontSize > 0 && f.size !== null && f.size !== settings.fontSize) {
    out.push({
      ruleId:         "wrong-font-size",
      message:        `${label} is ${f.size}pt — brand requires ${settings.fontSize}pt`,
      paragraphIndex: idx,
      fix: { type: "setFontSize", value: settings.fontSize },
    });
  }

  // ── Font color ───────────────────────────────────────
  // Word returns "" for automatic/theme colour; skip those.
  if (settings.allowedColors.length > 0 && f.color && f.color !== "") {
    const actual  = normalizeHex(f.color);
    const allowed = settings.allowedColors.map(normalizeHex);
    if (!allowed.includes(actual)) {
      out.push({
        ruleId:         "wrong-font-color",
        message:        `${label} color (${f.color}) is not in the allowed palette`,
        paragraphIndex: idx,
        fix: { type: "setFontColor", value: settings.allowedColors[0] },
      });
    }
  }

  // ── Bold ─────────────────────────────────────────────
  if (!settings.boldAllowed && f.bold === true) {
    out.push({
      ruleId:         "bold-not-allowed",
      message:        `${label} contains bold text, which is not permitted`,
      paragraphIndex: idx,
      fix: { type: "setBold", value: false },
    });
  }

  // ── Italic ───────────────────────────────────────────
  if (!settings.italicAllowed && f.italic === true) {
    out.push({
      ruleId:         "italic-not-allowed",
      message:        `${label} contains italic text, which is not permitted`,
      paragraphIndex: idx,
      fix: { type: "setItalic", value: false },
    });
  }
}

// Strip leading "#" and lowercase for reliable comparison.
function normalizeHex(hex) {
  return hex.replace(/^#/, "").toLowerCase();
}

// =========================================================
// Render violations
// =========================================================

function renderViolations(violations) {
  const summary = document.createElement("div");
  summary.className   = "summary";
  summary.textContent = `${violations.length} violation${violations.length !== 1 ? "s" : ""} found`;
  resultsEl.appendChild(summary);

  for (const v of violations) {
    resultsEl.appendChild(buildCard(v));
  }
}

// Build a single violation card.
// Accept button is rendered but disabled — fix logic comes in the next step.
// The full violation object is stored as JSON on the element so the next
// step can read it without re-running the check.
function buildCard(v) {
  const card = document.createElement("div");
  card.className          = "card";
  card.dataset.violation  = JSON.stringify(v);   // payload for future Accept handler

  // ── Header row: rule badge + paragraph reference ──────
  const header = document.createElement("div");
  header.className = "card__header";
  header.innerHTML =
    `<span class="badge badge--rule">${escapeHtml(v.ruleId)}</span>` +
    `<span class="card__para">¶${v.paragraphIndex + 1}</span>`;
  card.appendChild(header);

  // ── Human-readable message ────────────────────────────
  const msg = document.createElement("div");
  msg.className   = "card__message";
  msg.textContent = v.message;
  card.appendChild(msg);

  // ── Fix preview ───────────────────────────────────────
  const fixEl = document.createElement("div");
  fixEl.className = "card__fix";
  fixEl.innerHTML =
    `Fix: <code>${escapeHtml(v.fix.type)}</code>` +
    ` → <strong>${escapeHtml(String(v.fix.value))}</strong>`;
  card.appendChild(fixEl);

  // ── Actions ───────────────────────────────────────────
  const actions = document.createElement("div");
  actions.className = "card__actions";

  // Accept is disabled until wired in the next step.
  const acceptBtn = document.createElement("button");
  acceptBtn.className   = "btn btn--accept";
  acceptBtn.textContent = "Accept";
  acceptBtn.disabled    = true;
  acceptBtn.title       = "Fix will be available in the next step";
  actions.appendChild(acceptBtn);

  const dismissBtn = document.createElement("button");
  dismissBtn.className   = "btn btn--dismiss";
  dismissBtn.textContent = "Dismiss";
  dismissBtn.addEventListener("click", () => {
    card.remove();
    updateSummaryCount();
  });
  actions.appendChild(dismissBtn);

  card.appendChild(actions);
  return card;
}

// =========================================================
// Helpers
// =========================================================

function setStatus(msg, isError = false) {
  statusEl.textContent = msg;
  statusEl.className   = "status" + (isError ? " status--error" : "");
}

function clearResults() {
  resultsEl.innerHTML = "";
}

function updateSummaryCount() {
  const remaining = resultsEl.querySelectorAll(".card").length;
  const summary   = resultsEl.querySelector(".summary");
  if (!summary) return;
  if (remaining === 0) {
    summary.remove();
    setStatus("All violations dismissed.");
  } else {
    summary.textContent = `${remaining} violation${remaining !== 1 ? "s" : ""} remaining`;
  }
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

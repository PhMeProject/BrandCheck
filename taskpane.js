/* =========================================================
   BrandCheck – task pane logic
   Depends on: Office.js (loaded by taskpane.html)
   ========================================================= */

"use strict";

// ── Brand / style rules ────────────────────────────────────
// Each rule describes one pattern to flag.
//
// Shape:
//   id        – unique slug
//   severity  – "error" | "warning" | "info"
//   message   – human-readable rule name shown in the card
//   pattern   – RegExp to search for in the document text
//   replace   – string or function(match, ...groups) => string
//               Used when the user clicks Accept.
//               Pass null if there is no automatic fix.
//
// Add your own rules here as the linter grows.
const RULES = [
  {
    id: "no-lorem",
    severity: "error",
    message: "Placeholder text detected",
    pattern: /lorem ipsum/gi,
    replace: "[CONTENT NEEDED]",
  },
  {
    id: "company-name",
    severity: "error",
    message: 'Use full company name "Acme Corp"',
    pattern: /\bacme\b(?!\s+corp)/gi,
    replace: "Acme Corp",
  },
  {
    id: "passive-voice-was",
    severity: "warning",
    message: 'Avoid weak passive "was [verb]ed"',
    pattern: /\bwas\s+\w+ed\b/gi,
    replace: null, // no auto-fix – rewrite manually
  },
  {
    id: "double-space",
    severity: "info",
    message: "Remove double spaces",
    pattern: /  +/g,
    replace: " ",
  },
  {
    id: "smart-quotes",
    severity: "info",
    message: 'Use curly quotes instead of straight " or \'',
    pattern: /"|'/g,
    replace: null,
  },
];

// ── DOM refs ───────────────────────────────────────────────
const btnCheck = document.getElementById("btn-check");
const statusEl = document.getElementById("status");
const resultsEl = document.getElementById("results");

// ── Office initialisation ─────────────────────────────────
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    btnCheck.addEventListener("click", runBrandCheck);
    setStatus("Ready. Click Brand Check to scan.");
  } else {
    setStatus("This add-in only works inside Microsoft Word.", true);
    btnCheck.disabled = true;
  }
});

// ── Main check ────────────────────────────────────────────
async function runBrandCheck() {
  btnCheck.disabled = true;
  clearResults();
  setStatus("Scanning document…");

  try {
    await Word.run(async (context) => {
      // Load the full body text so we can search it.
      const body = context.document.body;
      context.load(body, "text");
      await context.sync();

      const docText = body.text;
      const violations = findViolations(docText);

      if (violations.length === 0) {
        setStatus("No violations found. Looks great!");
      } else {
        setStatus(`Found ${violations.length} violation${violations.length !== 1 ? "s" : ""}.`);
        renderViolations(violations, context);
      }
    });
  } catch (err) {
    setStatus(`Error: ${err.message}`, true);
    console.error(err);
  } finally {
    btnCheck.disabled = false;
  }
}

// ── Rule engine ───────────────────────────────────────────
function findViolations(text) {
  const violations = [];

  for (const rule of RULES) {
    // Reset lastIndex so global regexps work on successive calls.
    rule.pattern.lastIndex = 0;

    let match;
    while ((match = rule.pattern.exec(text)) !== null) {
      violations.push({
        rule,
        matchText: match[0],
        index: match.index,
        suggestion: typeof rule.replace === "string"
          ? rule.replace
          : typeof rule.replace === "function"
          ? rule.replace(match[0])
          : null,
      });
    }
  }

  return violations;
}

// ── Render violation cards ────────────────────────────────
function renderViolations(violations, context) {
  // Summary bar
  const summary = document.createElement("div");
  summary.className = "summary";
  summary.textContent = `${violations.length} violation${violations.length !== 1 ? "s" : ""} found`;
  resultsEl.appendChild(summary);

  for (const v of violations) {
    resultsEl.appendChild(buildCard(v, context));
  }
}

function buildCard(violation, context) {
  const { rule, matchText, suggestion } = violation;

  const card = document.createElement("div");
  card.className = "card";

  // Rule name + severity badge
  const ruleLine = document.createElement("div");
  ruleLine.className = "card__rule";
  ruleLine.innerHTML =
    escapeHtml(rule.message) +
    `<span class="badge badge--${rule.severity}">${rule.severity}</span>`;
  card.appendChild(ruleLine);

  // Matched text
  const matchLine = document.createElement("div");
  matchLine.className = "card__match";
  matchLine.innerHTML = `Found: <mark>${escapeHtml(matchText)}</mark>`;
  card.appendChild(matchLine);

  // Suggestion (if any)
  if (suggestion !== null) {
    const sugLine = document.createElement("div");
    sugLine.className = "card__suggestion";
    sugLine.innerHTML = `Replace with: <span>${escapeHtml(suggestion)}</span>`;
    card.appendChild(sugLine);
  }

  // Action buttons
  const actions = document.createElement("div");
  actions.className = "card__actions";

  if (suggestion !== null) {
    const acceptBtn = document.createElement("button");
    acceptBtn.className = "btn btn--accept";
    acceptBtn.textContent = "Accept";
    acceptBtn.addEventListener("click", () =>
      applyFix(violation, card, acceptBtn)
    );
    actions.appendChild(acceptBtn);
  }

  const dismissBtn = document.createElement("button");
  dismissBtn.className = "btn btn--dismiss";
  dismissBtn.textContent = "Dismiss";
  dismissBtn.addEventListener("click", () => card.remove());
  actions.appendChild(dismissBtn);

  card.appendChild(actions);
  return card;
}

// ── Apply fix via Office.js search/replace ────────────────
async function applyFix(violation, card, acceptBtn) {
  acceptBtn.disabled = true;
  acceptBtn.textContent = "Applying…";

  try {
    await Word.run(async (context) => {
      // Word.Body.search returns all ranges matching the text.
      // We replace the first unfixed occurrence we find.
      const results = context.document.body.search(violation.matchText, {
        matchCase: false,
        matchWholeWord: false,
      });
      context.load(results, "items");
      await context.sync();

      if (results.items.length > 0) {
        results.items[0].insertText(violation.suggestion, "Replace");
        await context.sync();
      }
    });

    // Remove the card after a successful fix.
    card.remove();

    // Update the summary count.
    updateSummaryCount();
  } catch (err) {
    acceptBtn.disabled = false;
    acceptBtn.textContent = "Accept";
    setStatus(`Fix failed: ${err.message}`, true);
    console.error(err);
  }
}

// ── Helpers ───────────────────────────────────────────────
function setStatus(msg, isError = false) {
  statusEl.textContent = msg;
  statusEl.className = "status" + (isError ? " status--error" : "");
}

function clearResults() {
  resultsEl.innerHTML = "";
}

function updateSummaryCount() {
  const remaining = resultsEl.querySelectorAll(".card").length;
  const summary = resultsEl.querySelector(".summary");
  if (!summary) return;
  if (remaining === 0) {
    summary.remove();
    setStatus("All violations resolved.");
  } else {
    summary.textContent = `${remaining} violation${remaining !== 1 ? "s" : ""} remaining`;
  }
}

function escapeHtml(str) {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

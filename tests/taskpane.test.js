/**
 * Task-pane logic tests. Run with:  node tests/taskpane.test.js
 *
 * Node is needed ONLY for this file. The add-in itself has no build step and
 * no Node dependency — see README. GitHub's runners ship Node already, so CI
 * runs these without installing anything.
 *
 * taskpane.js is browser code that grabs DOM nodes at load time, so it is
 * evaluated inside a vm context against a stub document; that exposes its
 * top-level functions for direct testing.
 */
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const TASKPANE = path.join(__dirname, "..", "addin", "src", "taskpane");
const read = (f) => fs.readFileSync(path.join(TASKPANE, f), "utf8");
const js = read("taskpane.js");
const html = read("taskpane.html");
const css = read("taskpane.css");

let failures = 0;
function report(ok, msg, detail) {
  if (!ok) failures++;
  console.log(`  ${ok ? "PASS" : "FAIL"}  ${msg}`);
  if (!ok && detail) console.log(`        ${detail}`);
}
const check = (name, actual, expected) =>
  report(actual === expected, name, `expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
const checkIncludes = (name, haystack, needle) =>
  report(String(haystack).includes(needle), name, `missing ${JSON.stringify(needle)} in ${JSON.stringify(haystack)}`);

// ---------- load taskpane.js under a stub DOM ----------

function makeEl() {
  return {
    style: {},
    classList: { add() {}, remove() {}, toggle() {}, contains: () => false },
    children: [],
    parentNode: null,
    innerHTML: "",
    textContent: "",
    title: "",
    value: "",
    disabled: false,
    scrollHeight: 20,
    scrollTop: 0,
    addEventListener() {},
    append() {},
    appendChild(c) { this.children.push(c); c.parentNode = this; },
    remove() {},
    focus() {},
    querySelector: () => makeEl(),
  };
}

const ctx = {
  document: { getElementById: () => makeEl(), createElement: () => makeEl(), querySelectorAll: () => [] },
  Office: { onReady() {} }, // swallow the callback so nothing auto-runs
  Excel: {},
  console,
  fetch: () => {},
  AbortController: class { constructor() { this.signal = {}; } abort() {} },
  FileReader: class {},
  TextDecoder: class {},
};
vm.createContext(ctx);
vm.runInContext(js, ctx, { filename: "taskpane.js" });

// ---------- markdown ----------

console.log("\nMarkdown rendering");
check("bold", ctx.renderMarkdown("say **hi**"), "<p>say <strong>hi</strong></p>");
check("inline code", ctx.renderMarkdown("use `A1:B2`"), "<p>use <code>A1:B2</code></p>");
check("heading", ctx.renderMarkdown("## Totals"), "<h2>Totals</h2>");
check("bullets", ctx.renderMarkdown("- one\n- two"), "<ul><li>one</li><li>two</li></ul>");
check("numbered", ctx.renderMarkdown("1. first\n2. second"), "<ol><li>first</li><li>second</li></ol>");
check("paragraphs", ctx.renderMarkdown("a\n\nb"), "<p>a</p><p>b</p>");
checkIncludes("fenced code", ctx.renderMarkdown("text\n```js\nconst x = 1;\n```"), "<pre><code>const x = 1;</code></pre>");
checkIncludes("link", ctx.renderMarkdown("[docs](https://example.com)"),
  '<a href="https://example.com" target="_blank" rel="noopener">docs</a>');

// Claude's output is untrusted input to the DOM. These must never produce
// live markup, or a crafted spreadsheet could script the task pane.
console.log("\nSafety — model output must never inject markup");
check("script tag escaped", ctx.renderMarkdown("<script>alert(1)</script>"),
  "<p>&lt;script&gt;alert(1)&lt;/script&gt;</p>");
check("img onerror escaped", ctx.renderMarkdown('<img src=x onerror="alert(1)">'),
  '<p>&lt;img src=x onerror="alert(1)"&gt;</p>');
checkIncludes("javascript: link not linkified", ctx.renderMarkdown("[x](javascript:alert(1))"), "[x](javascript:alert(1))");
check("ampersand escaped", ctx.escapeHtml("Tom & Jerry <b>"), "Tom &amp; Jerry &lt;b&gt;");
check("code block contents escaped", ctx.renderMarkdown("```\n<b>hi</b>\n```"),
  "<pre><code>&lt;b&gt;hi&lt;/b&gt;</code></pre>");

// ---------- spreadsheet address arithmetic ----------

console.log("\nColumn arithmetic");
check("A -> 1", ctx.columnToNumber("A"), 1);
check("Z -> 26", ctx.columnToNumber("Z"), 26);
check("AA -> 27", ctx.columnToNumber("AA"), 27);
check("1 -> A", ctx.numberToColumn(1), "A");
check("26 -> Z", ctx.numberToColumn(26), "Z");
check("27 -> AA", ctx.numberToColumn(27), "AA");
check("702 -> ZZ", ctx.numberToColumn(702), "ZZ");
let roundTrip = true;
for (let i = 1; i <= 1000; i++) if (ctx.columnToNumber(ctx.numberToColumn(i)) !== i) roundTrip = false;
check("round-trips 1..1000", roundTrip, true);

console.log("\nCell references in the write preview");
check("anchor of Sheet1!B3:D9", JSON.stringify(ctx.parseAnchor("Sheet1!B3:D9")), '{"col":2,"row":3}');
check("anchor with $ absolute", JSON.stringify(ctx.parseAnchor("$B$3")), '{"col":2,"row":3}');
check("anchor of bare A1", JSON.stringify(ctx.parseAnchor("A1")), '{"col":1,"row":1}');
check("offset (0,0) of B3", ctx.cellRef(ctx.parseAnchor("Sheet1!B3:D9"), 0, 0), "B3");
check("offset (2,1) of B3", ctx.cellRef(ctx.parseAnchor("Sheet1!B3:D9"), 2, 1), "C5");
check("offset past Z", ctx.cellRef(ctx.parseAnchor("Y1"), 0, 3), "AB1");
check("null anchor fallback", ctx.cellRef(null, 0, 0), "r1c1");

console.log("\nDisplayed cell values");
check("empty string", ctx.displayValue(""), "(empty)");
check("null", ctx.displayValue(null), "(empty)");
check("zero is shown, not treated as empty", ctx.displayValue(0), "0");
check("false is shown", ctx.displayValue(false), "false");

// ---------- HTML/CSS/JS wiring ----------

console.log("\nIDs referenced by JS exist in the HTML");
for (const id of new Set([...js.matchAll(/getElementById\("([^"]+)"\)/g)].map((m) => m[1]))) {
  report(html.includes(`id="${id}"`), `#${id}`);
}

console.log("\nSelectors queried by JS exist in the HTML");
for (const sel of new Set([...js.matchAll(/querySelectorAll\("\.([^"]+)"\)/g)].map((m) => m[1]))) {
  report(html.includes(`${sel}"`), `.${sel}`);
}

console.log("\nCSS classes assigned by JS are styled");
const assigned = new Set();
for (const m of js.matchAll(/className = "([^"${}]+)"/g)) m[1].split(/\s+/).forEach((c) => c && assigned.add(c));
for (const m of js.matchAll(/classList\.(?:add|toggle|remove)\("([^"]+)"/g)) assigned.add(m[1]);
for (const c of [...assigned].sort()) report(css.includes(`.${c}`), `.${c}`);

console.log("\nHTML links the stylesheet and script");
report(html.includes('href="taskpane.css"'), "taskpane.css linked");
report(html.includes('src="taskpane.js"'), "taskpane.js linked");
report(html.includes("office.js"), "office.js loaded");

console.log("\nEvery CSS variable used is defined");
const defined = new Set([...css.matchAll(/^\s*(--[\w-]+):/gm)].map((m) => m[1]));
for (const v of new Set([...css.matchAll(/var\((--[\w-]+)\)/g)].map((m) => m[1]))) {
  report(defined.has(v), `var(${v})`);
}

console.log("\nDark theme overrides every light token");
const split = css.indexOf("prefers-color-scheme: dark");
const lightVars = new Set([...css.slice(0, split).matchAll(/^\s*(--[\w-]+):/gm)].map((m) => m[1]));
const darkVars = new Set([...css.slice(split).matchAll(/^\s*(--[\w-]+):/gm)].map((m) => m[1]));
for (const v of [...lightVars].sort()) report(darkVars.has(v), `dark override for ${v}`);

console.log(failures === 0 ? "\nALL TESTS PASSED" : `\n${failures} FAILURE(S)`);
process.exit(failures === 0 ? 0 : 1);

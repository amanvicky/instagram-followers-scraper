"use strict";
// UI script tests — requires Node >= 18 and jsdom (optional dependency).
// Skipped automatically when jsdom is not installed:
//    node tests/test_ui_js.js
const fs = require("fs");
const path = require("path");

let results = [];
function test(name, fn) {
  try { fn(); results.push(["PASS", name]); }
  catch (e) { results.push(["FAIL", name + " :: " + e.message]); }
}
function assert(cond, msg) { if (!cond) throw new Error(msg || "assertion failed"); }

let jsdom;
try { jsdom = require("jsdom"); } catch (e) { jsdom = null; }

if (!jsdom) {
  console.log("SKIP tests/test_ui_js.js — jsdom not installed (pip-only repo).");
  process.exit(0);
}

const { JSDOM } = jsdom;
const html = fs.readFileSync(
  path.join(__dirname, "..", "templates", "index.html"), "utf8");
const dom = new JSDOM(html, { runScripts: "dangerously",
                              url: "http://127.0.0.1/" });
const { window } = dom;
const { document } = window;

// Minimal fetch stub: /start 400 on empty target, else success.
window.fetch = async (url, opts = {}) => {
  const form = opts.body;
  const target = form && form.get ? form.get("target") : "";
  const json = async (obj) => obj;
  if (url === "/start") {
    if (!target) return { ok: false, status: 400,
                          json: async () => ({ error: "Please type a target Instagram account." }) };
    return { ok: true, status: 200,
             json: async () => ({ job_id: "1", reference_size: 0, mode: "logged_in" }) };
  }
  if (url === "/login-status") return { ok: true, json: async () => ({ status: "done" }) };
  return { ok: true, json: async () => ({}) };
};

test("max validation: '5hundred' is rejected with a clear message", () => {
  document.getElementById("target").value = "nasa";
  document.getElementById("max").value = "5hundred";
  document.getElementById("go").click();
  const err = document.getElementById("inputError");
  assert(!err.classList.contains("hidden"), "error box stays hidden");
  assert(/plain number/i.test(err.textContent), "message wrong: " + err.textContent);
});

test("max validation: '5x' is rejected (not silently capped at 5)", () => {
  document.getElementById("max").value = "5x";
  document.getElementById("go").click();
  assert(!document.getElementById("inputError").classList.contains("hidden"),
         "5x was accepted");
});

test("max validation: '0' is rejected", () => {
  document.getElementById("max").value = "0";
  document.getElementById("go").click();
  assert(!document.getElementById("inputError").classList.contains("hidden"),
         "0 was accepted");
});

test("max validation: empty max starts the job", () => {
  document.getElementById("max").value = "";
  document.getElementById("go").click();
  assert(!document.getElementById("progressCard").classList.contains("hidden"),
         "job did not start with empty max");
});

const failed = results.filter((r) => r[0] === "FAIL");
for (const [status, name] of results) console.log(status + " " + name);
console.log(`${results.length - failed.length}/${results.length} passed`);
process.exit(failed.length ? 1 : 0);

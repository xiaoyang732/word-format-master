"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const { spawn } = require("node:child_process");
const { chromium } = require(process.env.WFM_NODE_MODULES ? path.join(process.env.WFM_NODE_MODULES, "playwright") : "playwright");

let server;
async function dashboardUrl() {
  if (process.argv[2]) return process.argv[2];
  server = spawn(process.env.WFM_PYTHON || "python", ["-u", path.join(__dirname, "..", "scripts", "serve_dashboard.py"), "--no-open"],
    { stdio: ["ignore", "pipe", "pipe"], env: { ...process.env, PYTHONIOENCODING: "utf-8" } });
  return new Promise((resolve, reject) => {
    let output = "";
    let stderr = "";
    const timer = setTimeout(() => reject(new Error(`Dashboard startup timed out: ${stderr}`)), 30000);
    const finish = callback => { clearTimeout(timer); callback(); };
    server.stderr.on("data", data => { stderr += data.toString(); });
    server.once("error", error => finish(() => reject(error)));
    server.once("exit", code => finish(() => reject(new Error(`Dashboard exited (${code}): ${stderr}`))));
    server.stdout.on("data", data => {
      output += data.toString();
      const match = output.match(/Word Format Master: (http:\/\/127\.0\.0\.1:\d+\/)/);
      if (match) finish(() => resolve(match[1]));
    });
  });
}

(async () => {
  const screenshotDir = path.join(__dirname, "..", ".tmp", "dashboard-ui");
  fs.mkdirSync(screenshotDir, { recursive: true });
  const url = await dashboardUrl();
  const browser = await chromium.launch({ headless: true, ...(process.env.WFM_BROWSER_CHANNEL ? { channel: process.env.WFM_BROWSER_CHANNEL } : {}) });
  try {
    const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
    const errors = [];
    page.on("pageerror", error => errors.push(error.message));
    await page.goto(url, { waitUntil: "networkidle" });
    await page.waitForFunction(() => state.presets.length && state.capabilities);
    assert.equal(await page.locator("#bodyFirstLineIndentValueInput").inputValue(), "2");
    await page.evaluate(() => {
      const input = document.querySelector("#bodyFirstLineIndentValueInput");
      input.value = "0"; input.dispatchEvent(new Event("input", { bubbles: true }));
      const size = document.querySelector("#bodyFontSizeSelect");
      size.value = "14"; size.dispatchEvent(new Event("change", { bubbles: true }));
    });
    const body = await page.evaluate(() => state.spec.body);
    assert.equal(body.first_line_indent_chars, 0);
    assert.equal(body.first_line_indent_mm, undefined);
    await page.evaluate(() => {
      state.spec.body.first_line_indent_mm = 7.4;
      delete state.spec.body.first_line_indent_chars;
      document.querySelector("#bodyFirstLineIndentUnitSelect").value = "mm";
      updateInputs();
    });
    assert.equal(await page.locator("#bodyFirstLineIndentValueInput").inputValue(), "7.4");
    await page.evaluate(() => {
      state.customTemplate = { name: "test", result: { inferred_spec: clone(state.spec) } };
      state.analysis = { inferred_spec: clone(state.spec), source: { name: "test.docx" } };
      renderPresetSelect();
    });
    await page.locator("#presetSelect").selectOption("custom-template");
    const spec = await page.evaluate(() => state.spec);
    assert.equal(spec.mode, "parameterized");
    assert.equal(spec.template_required, false);
    const cap = await page.evaluate(() => state.capabilities.format_core.operations);
    assert.ok(cap.some(operation => operation.action === "paragraph.spacing_after.set"));
    assert.ok(cap.some(operation => operation.action === "header.font.size.set"));
    await page.getByRole("tab", { name: "页眉页码", exact: true }).click();
    await page.locator("#headerSectionNumberInput").fill("2");
    await page.locator('[data-spec-path="headers_footers.header.enabled"]').check();
    await page.locator('[data-spec-path="headers_footers.header.mode"]').selectOption("chapter_title");
    await page.locator('[data-spec-path="headers_footers.header.style_name"]').fill("Heading 1");
    await page.locator('[data-spec-path="headers_footers.header.font_size_pt"]').fill("9");
    await page.locator('[data-spec-path="headers_footers.different_odd_even"]').check();
    await page.locator('[data-spec-path="headers_footers.even_header.enabled"]').check();
    await page.locator('[data-spec-path="headers_footers.even_header.text"]').fill("论文题目");
    const headerSpec = await page.evaluate(() => state.spec.headers_footers);
    assert.equal(headerSpec.section_number, 2);
    assert.equal(headerSpec.header.mode, "chapter_title");
    assert.equal(headerSpec.header.font_size_pt, 9);
    assert.equal(headerSpec.even_header.text, "论文题目");
    assert.equal(headerSpec.different_odd_even, true);
    assert.match(await page.locator("#previewHeaderText").textContent(), /动态示意/);
    for (const viewport of [{ width: 1440, height: 1000 }, { width: 390, height: 844 }]) {
      await page.setViewportSize(viewport);
      await page.screenshot({ path: path.join(screenshotDir, `dashboard-${viewport.width}.png`), fullPage: true });
      assert.ok(await page.locator("#headerSectionNumberInput").isVisible());
    }
    assert.deepEqual(errors, []);
    console.log("Dashboard UI passed: zero indent, precise units, template reselection, section header fields, shared capabilities, desktop/mobile screenshots.");
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; }).finally(() => { if (server) server.kill(); });

#!/usr/bin/env node

import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import process from "node:process";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const skillRoot = path.resolve(__dirname, "..");
const stateDir = path.join(skillRoot, "state");
const profileDirDefault = path.join(stateDir, "chromium-profile");
const sessionFile = path.join(stateDir, "session.json");

const HUMAN_TEXTS = [
  "扫码登录",
  "微信快速安全登录",
  "企业身份登录",
  "个人身份登录",
  "请使用企业微信扫描二维码登录",
  "请使用微信扫描二维码登录",
  "验证码",
  "滑动验证",
  "安全验证",
  "异常登录",
  "账号存在风险",
  "申请权限",
  "无权访问",
  "仅内部成员可见",
  "需要申请",
  "需要审批",
];

const EXPORT_NAME = {
  docx: "本地Word文档(.docx)",
  pdf: "PDF",
  png: "图片 (.png)",
  xlsx: "本地Excel表格 (.xlsx)",
  csv: "本地CSV文件 (.csv, 当前工作表)",
  zip: "本地Excel表格和图片 (.zip)",
};

const SYSTEM_CHROME_CANDIDATES = [
  "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
  "/Applications/Chromium.app/Contents/MacOS/Chromium",
];

function parseArgs(argv) {
  const args = {
    format: "docx",
    login: false,
    convertMd: false,
    titleOnly: false,
    keepDuplicates: false,
    batchFile: null,
    outputDir: process.cwd(),
    timeoutMs: 30000,
    manualWaitMs: 180000,
    profileDir: profileDirDefault,
  };
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    const next = argv[index + 1];
    if (token === "--url") args.url = next, index += 1;
    else if (token === "--link-text") args.linkText = next, index += 1;
    else if (token === "--batch-file") args.batchFile = expandHome(next), index += 1;
    else if (token === "--output-dir") args.outputDir = expandHome(next), index += 1;
    else if (token === "--format") args.format = next, index += 1;
    else if (token === "--profile-dir") args.profileDir = expandHome(next), index += 1;
    else if (token === "--timeout-ms") args.timeoutMs = Number(next), index += 1;
    else if (token === "--manual-wait-ms") args.manualWaitMs = Number(next), index += 1;
    else if (token === "--login") args.login = true;
    else if (token === "--convert-md") args.convertMd = true;
    else if (token === "--title-only") args.titleOnly = true;
    else if (token === "--keep-duplicates") args.keepDuplicates = true;
    else if (token === "--help" || token === "-h") args.help = true;
    else throw new Error(`未知参数: ${token}`);
  }
  return args;
}

function expandHome(value) {
  if (!value) return value;
  return value.startsWith("~/") ? path.join(os.homedir(), value.slice(2)) : value;
}

function usage() {
  console.log(`用法:
  node wecom_doc_download.mjs --url "<企微文档链接>" [选项]

选项:
  --link-text "<子文件标题>"    在知识库页面点击指定链接
  --batch-file "<json/jsonl>"   批量任务清单，串行执行
  --output-dir "<目录>"         输出目录，默认当前目录
  --format docx|pdf|png|xlsx|csv|zip
                               导出格式；普通文档默认 docx，表格默认 xlsx
  --login                       允许弹出有界面浏览器，等待人工扫码/验证
  --convert-md                  在导出 docx 后调用 pandoc 转为 md
  --title-only                  只读取当前页面标题，不执行下载
  --keep-duplicates             同名文件不跳过，自动追加 2/3... 后缀
  --profile-dir "<目录>"        持久化登录态目录
  --timeout-ms 30000            常规等待超时
  --manual-wait-ms 180000       人工步骤等待时间
`);
}

function fail(code, message, extra = {}) {
  const payload = { ok: false, code, message, ...extra };
  console.error(JSON.stringify(payload, null, 2));
  process.exit(code);
}

function saveSessionMeta(meta) {
  fs.mkdirSync(stateDir, { recursive: true });
  fs.writeFileSync(
    sessionFile,
    JSON.stringify({ updatedAt: new Date().toISOString(), ...meta }, null, 2),
    "utf8",
  );
}

function commandExists(command) {
  const result = spawnSync("bash", ["-lc", `command -v ${command}`], { encoding: "utf8" });
  return result.status === 0;
}

function parseJsonPayload(text) {
  const raw = (text || "").trim();
  if (!raw) return null;
  const start = raw.lastIndexOf("\n{");
  const jsonText = (start >= 0 ? raw.slice(start + 1) : raw).trim();
  try {
    return JSON.parse(jsonText);
  } catch {
    return null;
  }
}

function readBatchTasks(batchFile) {
  if (!fs.existsSync(batchFile)) {
    fail(35, "批量任务文件不存在。", { batchFile });
  }
  const content = fs.readFileSync(batchFile, "utf8").trim();
  if (!content) {
    fail(36, "批量任务文件为空。", { batchFile });
  }
  let tasks = [];
  if (content.startsWith("[")) {
    tasks = JSON.parse(content);
  } else {
    tasks = content
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter((line) => line && !line.startsWith("#"))
      .map((line) => JSON.parse(line));
  }
  if (!Array.isArray(tasks) || tasks.length === 0) {
    fail(37, "批量任务文件中没有有效任务。", { batchFile });
  }
  return tasks.map((task, index) => {
    if (!task || typeof task !== "object" || !task.url) {
      fail(38, "批量任务项缺少 url。", { batchFile, taskIndex: index + 1, task });
    }
    return task;
  });
}

function buildChildArgs(baseArgs, task) {
  const childArgs = [__filename, "--url", task.url, "--profile-dir", baseArgs.profileDir];
  childArgs.push("--output-dir", expandHome(task.outputDir || baseArgs.outputDir));
  childArgs.push("--format", task.format || baseArgs.format);
  childArgs.push("--timeout-ms", String(task.timeoutMs || baseArgs.timeoutMs));
  childArgs.push("--manual-wait-ms", String(task.manualWaitMs || baseArgs.manualWaitMs));
  if (task.linkText) childArgs.push("--link-text", task.linkText);
  if (task.convertMd ?? baseArgs.convertMd) childArgs.push("--convert-md");
  if (task.titleOnly ?? baseArgs.titleOnly) childArgs.push("--title-only");
  if (task.keepDuplicates ?? baseArgs.keepDuplicates) childArgs.push("--keep-duplicates");
  if (baseArgs.login) childArgs.push("--login");
  return childArgs;
}

function runBatch(args) {
  const tasks = readBatchTasks(args.batchFile);
  const results = [];
  for (let index = 0; index < tasks.length; index += 1) {
    const task = tasks[index];
    const childArgs = buildChildArgs(args, task);
    const result = spawnSync(process.execPath, childArgs, {
      encoding: "utf8",
      maxBuffer: 10 * 1024 * 1024,
    });
    const payload =
      parseJsonPayload(result.stdout) ||
      parseJsonPayload(result.stderr) ||
      parseJsonPayload(`${result.stdout}\n${result.stderr}`);
    if (result.status !== 0) {
      fail(39, `批量任务在第 ${index + 1} 项停止。`, {
        taskIndex: index + 1,
        task,
        completedResults: results,
        childResult: payload || {
          status: result.status,
          stdout: result.stdout,
          stderr: result.stderr,
        },
      });
    }
    results.push({
      taskIndex: index + 1,
      task,
      result: payload,
    });
  }
  saveSessionMeta({
    batchFile: args.batchFile,
    batchCount: results.length,
    results,
    profileDir: args.profileDir,
  });
  console.log(
    JSON.stringify(
      {
        ok: true,
        batchFile: args.batchFile,
        count: results.length,
        results,
        profileDir: args.profileDir,
      },
      null,
      2,
    ),
  );
}

function normalizeTitle(value) {
  return (value || "")
    .replace(/\.[^.]+$/g, "")
    .replace(/\s+/g, "")
    .replace(/[()（）【】\[\]「」『』《》·:：,，'"“”‘’\-—_]/g, "")
    .trim()
    .toLowerCase();
}

async function loadPlaywright() {
  try {
    return await import("playwright");
  } catch {
    fail(
      10,
      "缺少 playwright 依赖。请先在当前 scripts 目录执行 npm install && npx playwright install chromium。",
      { scriptDir: __dirname },
    );
  }
}

async function textContent(page) {
  return page.evaluate(() => document.body?.innerText || "");
}

async function detectHumanGate(page) {
  const text = await textContent(page);
  return HUMAN_TEXTS.find((item) => text.includes(item)) || null;
}

async function waitForManualCompletion(page, timeoutMs) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    const humanGate = await detectHumanGate(page);
    if (!humanGate) return true;
    await page.waitForTimeout(1500);
  }
  return false;
}

async function ensureAuthenticated(page, args) {
  const humanGate = await detectHumanGate(page);
  if (!humanGate) return;
  if (!args.login) {
    fail(20, `检测到需要人工介入：${humanGate}。请重新用 --login 运行并由你完成扫码/验证。`, {
      url: page.url(),
    });
  }
  console.log(`检测到人工步骤：${humanGate}。请在弹出的浏览器中完成后等待继续。`);
  const done = await waitForManualCompletion(page, args.manualWaitMs);
  if (!done) {
    fail(21, `人工步骤超时，仍未完成：${humanGate}。请处理后重试。`, { url: page.url() });
  }
}

async function detectDocumentTitle(page) {
  const titleBox = page.locator('toolbar[aria-label="titlebar"] textbox').first();
  if (await titleBox.count()) {
    const value = await titleBox.inputValue().catch(() => "");
    if (value?.trim()) return value.trim();
  }
  const buttonTitle = page.getByRole("button").filter({ has: page.locator("textbox") }).first();
  if (await buttonTitle.count()) {
    const text = await buttonTitle.innerText().catch(() => "");
    if (text?.trim()) return text.trim();
  }
  const fallback = await page.title().catch(() => "");
  return fallback.trim();
}

function detectPageKind(url) {
  if (!url) return "doc";
  if (url.includes("/sheet/") || url.includes("/smartsheet/")) return "sheet";
  return "doc";
}

function resolveFormat(args, page) {
  if (args.format && args.format !== "docx") return args.format;
  return page.kind === "sheet" ? "xlsx" : "docx";
}

function getExecutablePath() {
  for (const candidate of SYSTEM_CHROME_CANDIDATES) {
    if (fs.existsSync(candidate)) return candidate;
  }
  return undefined;
}

async function openTargetPage(context, page, args) {
  await page.goto(args.url, { waitUntil: "domcontentloaded", timeout: args.timeoutMs });
  await page.waitForTimeout(1200);
  await ensureAuthenticated(page, args);
  if (!args.linkText) return page;

  const popupPromise = context.waitForEvent("page", { timeout: args.timeoutMs }).catch(() => null);
  const candidate = page.getByText(args.linkText, { exact: true }).first();
  if (await candidate.count()) {
    await candidate.click();
  } else {
    const fuzzy = page.getByText(args.linkText).first();
    if (!(await fuzzy.count())) {
      fail(22, `未找到链接标题：${args.linkText}`, { url: page.url() });
    }
    await fuzzy.click();
  }
  const nextPage = await popupPromise;
  if (!nextPage) {
    fail(23, `点击链接后没有打开新页面：${args.linkText}`, { url: page.url() });
  }
  await nextPage.waitForLoadState("domcontentloaded", { timeout: args.timeoutMs });
  await nextPage.waitForTimeout(1200);
  await ensureAuthenticated(nextPage, args);
  return nextPage;
}

async function openFileMenu(page, timeoutMs) {
  const candidates = [
    page.locator("#headerbar-filemenu").first(),
    page.locator("#main-menu-file").first(),
    page.locator('[aria-label="按钮:文件操作"]').first(),
    page.locator('[aria-label="file"]').first(),
    page.getByRole("button", { name: "file" }).first(),
  ];
  for (const candidate of candidates) {
    if (await candidate.count().catch(() => 0)) {
      await candidate.click({ timeout: timeoutMs, force: true });
      return;
    }
  }
  fail(25, "未找到文件操作入口。", { url: page.url() });
}

async function clickMenuItem(page, locator, timeoutMs) {
  await locator.waitFor({ timeout: timeoutMs });
  await locator.click({ timeout: timeoutMs, force: true });
}

function nextAvailablePath(destination, keepDuplicates) {
  if (!fs.existsSync(destination)) return destination;
  if (!keepDuplicates) return destination;
  const dirname = path.dirname(destination);
  const ext = path.extname(destination);
  const base = path.basename(destination, ext);
  let index = 2;
  while (true) {
    const candidate = path.join(dirname, `${base}${index}${ext}`);
    if (!fs.existsSync(candidate)) return candidate;
    index += 1;
  }
}

function findExistingOutput(outputDir, pageTitle, format) {
  const ext = format === "docx" ? ".docx" : format === "pdf" ? ".pdf" : format === "png" ? ".png" : format === "xlsx" ? ".xlsx" : format === "csv" ? ".csv" : format === "zip" ? ".zip" : "";
  if (!ext) return null;
  const candidate = path.join(outputDir, `${pageTitle}${ext}`);
  return fs.existsSync(candidate) ? candidate : null;
}

async function exportFile(page, args, runtimeFormat) {
  if (!EXPORT_NAME[runtimeFormat]) {
    fail(24, `不支持的导出格式：${runtimeFormat}`);
  }
  let lastError = null;
  for (let attempt = 1; attempt <= 3; attempt += 1) {
    try {
      await openFileMenu(page, args.timeoutMs);
      await page.waitForTimeout(400);
      const exportSubmenu = page.locator(".mainmenu-submenu-exportAs").first();
      if (await exportSubmenu.count().catch(() => 0)) {
        await exportSubmenu.hover({ timeout: args.timeoutMs, force: true });
      } else {
        const exportMenuItem = page.getByRole("menuitem", { name: /导出|导出为/ }).first();
        await clickMenuItem(page, exportMenuItem, args.timeoutMs);
      }
      await page.waitForTimeout(300);
      const exportCandidates = [
        page.locator(`.mainmenu-item-export-local`).first(),
        page.locator(`.mainmenu-item-export-csv`).first(),
        page.locator(`.mainmenu-item-export-image`).first(),
        page.locator(`.mainmenu-item-export-archive`).first(),
        page.getByRole("menuitem", { name: EXPORT_NAME[runtimeFormat] }).first(),
      ];
      let exportTarget = null;
      if (runtimeFormat === "xlsx") exportTarget = exportCandidates[0];
      else if (runtimeFormat === "csv") exportTarget = exportCandidates[1];
      else if (runtimeFormat === "png") exportTarget = exportCandidates[2];
      else if (runtimeFormat === "zip") exportTarget = exportCandidates[3];
      else exportTarget = exportCandidates[4];
      const [download] = await Promise.all([
        page.waitForEvent("download", { timeout: args.timeoutMs }),
        clickMenuItem(page, exportTarget, args.timeoutMs),
      ]);
      fs.mkdirSync(args.outputDir, { recursive: true });
      const filename = download.suggestedFilename();
      const destination = nextAvailablePath(path.join(args.outputDir, filename), args.keepDuplicates);
      await download.saveAs(destination);
      return destination;
    } catch (error) {
      lastError = error;
      if (attempt >= 3) break;
      await page.keyboard.press("Escape").catch(() => {});
      await page.waitForTimeout(600);
    }
  }
  throw lastError;
}

function assertTitleMatches(pageTitle, outputFile) {
  const docTitle = normalizeTitle(pageTitle);
  const fileTitle = normalizeTitle(path.basename(outputFile));
  if (!docTitle || !fileTitle) {
    fail(33, "标题核对失败：页面标题或下载文件名为空。", { pageTitle, outputFile });
  }
  if (docTitle.includes(fileTitle) || fileTitle.includes(docTitle)) {
    return;
  }
  fail(34, "标题核对失败：页面标题与下载文件名不一致。", {
    pageTitle,
    outputFile,
  });
}

function findExistingDocx(outputDir, pageTitle) {
  const candidate = path.join(outputDir, `${pageTitle}.docx`);
  if (fs.existsSync(candidate)) return candidate;
  return null;
}

function convertDocxToMarkdown(filePath) {
  if (!commandExists("pandoc")) {
    fail(30, "请求转换 Markdown，但系统中未找到 pandoc。");
  }
  const mdPath = filePath.replace(/\.docx$/i, ".md");
  const result = spawnSync(
    "pandoc",
    ["--track-changes=all", filePath, "-t", "gfm", "-o", mdPath],
    { encoding: "utf8" },
  );
  if (result.status !== 0) {
    fail(31, "pandoc 转 Markdown 失败。", { stderr: result.stderr });
  }
  return mdPath;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    usage();
    process.exit(0);
  }
  if (args.batchFile) {
    runBatch(args);
    return;
  }
  if (!args.url) {
    usage();
    process.exit(1);
  }

  fs.mkdirSync(args.profileDir, { recursive: true });
  fs.mkdirSync(args.outputDir, { recursive: true });

  const { chromium } = await loadPlaywright();
  const context = await chromium.launchPersistentContext(args.profileDir, {
    headless: !args.login,
    acceptDownloads: true,
    executablePath: getExecutablePath(),
  });

  try {
    const page = context.pages()[0] || (await context.newPage());
    const targetPage = await openTargetPage(context, page, args);
    const pageTitle = await detectDocumentTitle(targetPage);
    const pageKind = detectPageKind(targetPage.url());
    const runtimeFormat = resolveFormat(args, { kind: pageKind });
    if (args.titleOnly) {
      saveSessionMeta({
        url: args.url,
        linkText: args.linkText || null,
        pageTitle,
        pageKind,
        profileDir: args.profileDir,
      });
      console.log(JSON.stringify({ ok: true, pageTitle, pageKind, profileDir: args.profileDir }, null, 2));
      return;
    }
    let outputFile = null;
    const existingOutput = !args.keepDuplicates ? findExistingOutput(args.outputDir, pageTitle, runtimeFormat) : null;
    if (args.convertMd && runtimeFormat === "docx") {
      const existingDocx = findExistingDocx(args.outputDir, pageTitle);
      if (existingDocx) outputFile = existingDocx;
    }
    if (!outputFile && existingOutput) {
      outputFile = existingOutput;
    } else {
      outputFile = await exportFile(targetPage, args, runtimeFormat);
      if (pageKind === "doc") {
        assertTitleMatches(pageTitle, outputFile);
      }
    }
    let markdownFile = null;
    if (args.convertMd) {
      if (!/\.docx$/i.test(outputFile)) {
        fail(32, "--convert-md 仅支持 docx 导出结果。", { outputFile });
      }
      markdownFile = convertDocxToMarkdown(outputFile);
    }
    saveSessionMeta({
      url: args.url,
      linkText: args.linkText || null,
      pageTitle,
      pageKind,
      runtimeFormat,
      outputFile,
      markdownFile,
      profileDir: args.profileDir,
    });
    console.log(
      JSON.stringify(
        {
          ok: true,
          pageTitle,
          pageKind,
          runtimeFormat,
          outputFile,
          markdownFile,
          profileDir: args.profileDir,
          reusedLoginState: true,
        },
        null,
        2,
      ),
    );
  } finally {
    await context.close();
  }
}

main().catch((error) => {
  fail(99, error.message || "未知错误", { stack: error.stack });
});

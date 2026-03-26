---
name: wecom-doc-downloader
description: 当用户提到企微文档、企业微信文档、腾讯文档、知识库、`doc.weixin.qq.com` 链接、下载文档、导出 docx/pdf、下载某个链接文件、批量下载企微文档、批量转 md、或希望优先复用已保存登录态时使用。本 skill 负责打开企微文档或知识库，按标题定位子链接，导出为本地文件，并通过持久化浏览器 profile 复用登录态；遇到扫码、验证码、风控、权限申请等人工步骤时必须暂停并请求用户介入，禁止无上限重试。
---

# Wecom Doc Downloader

## Overview

这个 skill 用于稳定执行“企微文档/知识库下载”任务，覆盖直接文档链接、知识库中的子链接文件、`sheet/smartsheet` 表格链接、串行批量任务、导出为 `docx/pdf/png/xlsx/csv/zip`、以及可选的 `docx -> md` 转换。

核心目标有两个：
- 尽量后台执行，减少对台前工作的影响；
- 首次人工扫码后，把登录态保存在本地浏览器 profile 中，后续优先复用。

## Workflow

1. 判断用户给的是：
   - 企微文档直链；
   - 知识库/汇总页链接；
   - 知识库内某个子文件标题。
2. 优先复用 `scripts/` 中持久化 Chromium profile 的登录态。
3. 若是知识库页，按用户给的 `link_text` 精确或模糊定位子链接并打开。
4. 统一优先在右上角“文件操作/文档操作”中打开 `导出`：
   - 普通文档优先导出 `docx/pdf/png`；
   - 表格/智能表格优先导出 `xlsx/csv/png/zip`。
5. 如用户要求 Markdown，且导出格式是 `docx`，则再调用 `pandoc` 转为 `.md`。
6. 把结果保存到用户指定目录，并向用户返回最终路径。

批量模式下：
1. 从 `--batch-file` 读取任务清单；
2. 串行逐条执行，默认复用同一个 profile；
3. 任一任务遇到人工介入或失败时立即停下；
4. 返回已完成结果 + 当前阻塞项，不继续后续任务。

## Human-In-The-Loop Rules

以下情况必须立即停下，并明确告诉用户需要介入，不得盲目反复尝试：
- 扫码登录；
- 验证码、滑块、安全验证、风控提醒；
- 权限申请、仅内部成员可见、需要审批；
- 登录态失效且当前又不适合弹出前台窗口；
- 导出时出现网站侧确认弹窗，且无法稳定自动完成。

允许的小范围自动重试仅限：
- 页面尚未加载完成；
- 菜单被临时遮挡；
- 目标链接刚刷新导致元素短暂失效。

重试上限要低，通常 1-2 次；若仍失败，转为“向用户汇报当前阻塞点”。

批量任务额外规则：
- 默认只做串行，不做并行，避免多个页面同时抢登录态；
- 一旦某一项需要你扫码/验证，整批暂停在该项；
- 已完成项保留结果，不回滚、不重复下载。

## Login State Reuse

登录态通过持久化 Chromium profile 保存，默认目录：

`state/chromium-profile/`

首次使用：
- 用 `--login` 启动有界面的浏览器；
- 用户扫码/验证后，profile 内会保留 cookie 与站点状态；
- 后续任务默认可在无界面模式下直接复用。

这不是永久免登录。若网站主动失效、风控或 profile 损坏，仍要回退到人工登录。

## Scripts

- `scripts/wecom_doc_download.mjs`
  - 持久化浏览器 profile；
  - 打开企微文档或知识库；
  - 按标题点击子链接；
  - 导出 `docx/pdf/png`；
  - 可选 `docx -> md`；
  - 可选标题核对；
  - 支持 `json/jsonl` 批量清单串行执行；
  - 在遇到人工步骤时退出并给出明确提示。

- `scripts/package.json`
  - 声明 `playwright` 依赖，便于一次性安装。

## Setup

首次安装依赖：

```bash
cd ~/.codex/skills/wecom-doc-downloader/scripts
npm install
npx playwright install chromium
```

若只做 `docx` 导出可先不装 `pandoc`；若要转 Markdown，再安装：

```bash
brew install pandoc
```

## Quick Start

首次登录建档：

```bash
cd ~/.codex/skills/wecom-doc-downloader/scripts
node wecom_doc_download.mjs \
  --url "https://doc.weixin.qq.com/doc/xxx" \
  --output-dir "~/Downloads/wecom-docs" \
  --login
```

后续复用登录态，直接后台下载：

```bash
cd ~/.codex/skills/wecom-doc-downloader/scripts
node wecom_doc_download.mjs \
  --url "https://doc.weixin.qq.com/doc/xxx" \
  --output-dir "~/Downloads/wecom-docs"
```

下载知识库中的子文件：

```bash
cd ~/.codex/skills/wecom-doc-downloader/scripts
node wecom_doc_download.mjs \
  --url "https://doc.weixin.qq.com/doc/knowledge-base-link" \
  --link-text "情境行为性" \
  --output-dir "~/Downloads/wecom-docs"
```

导出后顺手转 Markdown：

```bash
cd ~/.codex/skills/wecom-doc-downloader/scripts
node wecom_doc_download.mjs \
  --url "https://doc.weixin.qq.com/doc/xxx" \
  --output-dir "~/Downloads/wecom-docs" \
  --convert-md
```

批量串行执行：

```bash
cd ~/.codex/skills/wecom-doc-downloader/scripts
node wecom_doc_download.mjs \
  --batch-file "~/Downloads/wecom-docs/tasks.jsonl" \
  --output-dir "~/Downloads/wecom-docs"
```

批量 + 首次登录：

```bash
cd ~/.codex/skills/wecom-doc-downloader/scripts
node wecom_doc_download.mjs \
  --batch-file "~/Downloads/wecom-docs/tasks.jsonl" \
  --output-dir "~/Downloads/wecom-docs" \
  --login
```

## Batch File Format

支持两种格式：
- `.json`：顶层是数组
- `.jsonl`：每行一个 JSON 对象

每个任务至少包含：
- `url`

可选字段：
- `linkText`
- `format`
- `convertMd`
- `outputDir`
- `titleOnly`
- `timeoutMs`
- `manualWaitMs`

`tasks.jsonl` 示例：

```jsonl
{"url":"https://doc.weixin.qq.com/doc/aaa","convertMd":true}
{"url":"https://doc.weixin.qq.com/doc/bbb","linkText":"情境行为性","convertMd":true}
{"url":"https://doc.weixin.qq.com/doc/ccc","format":"pdf","outputDir":"~/Downloads/wecom-pdf"}
```

`tasks.json` 示例：

```json
[
  {
    "url": "https://doc.weixin.qq.com/doc/aaa",
    "convertMd": true
  },
  {
    "url": "https://doc.weixin.qq.com/doc/bbb",
    "linkText": "情境行为性",
    "convertMd": true
  }
]
```

## Output Rules

- 默认输出目录由 `--output-dir` 指定。
- 导出文件名以站点返回下载名为准。
- 若启用 `--convert-md`，会在同目录生成同名 `.md`。
- 运行元信息会写入 `state/session.json`，记录最近一次成功时间与输出文件。
- 批量模式会在 `state/session.json` 中记录已完成项和当前批次结果。

## When Handling User Requests In Codex

- 用户只说“帮我下载这个企微文档”时，优先触发本 skill。
- 若用户还给了“知识库中的某个文件标题”，则传入 `--link-text`。
- 若用户说“批量下载”“一组企微链接”“批量转 md”，则优先使用 `--batch-file`。
- 若站点提示人工操作，先停下并告诉用户当前所需步骤，再等待用户完成。
- 若用户强调“不影响台前”，首次登录之外默认不要弹前台窗口，优先无界面模式。
- 若用户给的是 `sheet`/`smartsheet` 链接，默认按表格处理；未显式指定格式时自动使用 `xlsx`。
- 若批量任务里存在重复标题且用户要求“同名也保留”，应启用 `--keep-duplicates`，让脚本自动落为 `文件名2.xlsx`、`文件名3.xlsx`。

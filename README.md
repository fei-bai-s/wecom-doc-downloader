# wecom-doc-downloader

用于下载企微文档、腾讯文档、知识库子链接文件、智能表格中的表格文件，并优先复用本地登录态。

## 功能

- 下载单个 `doc.weixin.qq.com` 文档
- 下载单个 `doc.weixin.qq.com/sheet/...` / `.../smartsheet/...` 表格
- 下载知识库中的指定子文件
- 导出为 `docx` / `pdf` / `png` / `xlsx` / `csv` / `zip`
- 标题核对后将 `docx` 转为 `md`
- 支持串行批量任务
- 首次扫码后复用持久化浏览器 profile
- 默认优先使用系统 `Google Chrome` 无头模式，尽量不打断台前操作
- 遇到扫码、验证码、风控、权限申请时立即停下，等待人工介入
- 支持同名文件保留并自动追加 `2` / `3` 后缀

## 安装

```bash
cd scripts
npm install
npx playwright install chromium
```

若本机已安装 `Google Chrome`，脚本会优先复用系统 `Chrome` 做无头导出，不强依赖 Playwright 自带浏览器。

如果需要 `docx -> md`：

```bash
brew install pandoc
```

## 快速开始

首次运行并建立登录态：

```bash
cd scripts
node wecom_doc_download.mjs \
  --url "https://doc.weixin.qq.com/doc/xxx" \
  --output-dir "~/Downloads/wecom-docs" \
  --login
```

后续直接复用登录态下载：

```bash
cd scripts
node wecom_doc_download.mjs \
  --url "https://doc.weixin.qq.com/doc/xxx" \
  --output-dir "~/Downloads/wecom-docs"
```

后台下载表格时，未指定格式会自动按 `xlsx` 导出：

```bash
cd scripts
node wecom_doc_download.mjs \
  --url "https://doc.weixin.qq.com/sheet/xxx" \
  --output-dir "~/Downloads/wecom-sheet"
```

下载知识库中的子文件：

```bash
cd scripts
node wecom_doc_download.mjs \
  --url "https://doc.weixin.qq.com/doc/knowledge-base-link" \
  --link-text "情境行为性" \
  --output-dir "~/Downloads/wecom-docs"
```

下载后转 Markdown：

```bash
cd scripts
node wecom_doc_download.mjs \
  --url "https://doc.weixin.qq.com/doc/xxx" \
  --output-dir "~/Downloads/wecom-docs" \
  --convert-md
```

## 批量任务

支持：

- `examples/tasks.json`
- `examples/tasks.jsonl`

串行批量执行：

```bash
cd scripts
node wecom_doc_download.mjs \
  --batch-file "../examples/tasks.jsonl" \
  --output-dir "~/Downloads/wecom-docs" \
  --keep-duplicates
```

批量任务中任一项遇到人工步骤或错误时，整批会停在当前项，并返回已完成结果。

`--keep-duplicates` 适用于“表格里多行引用同一个文件，但仍要全部落盘”的场景；脚本会把重复文件名自动命名为 `文件名2.xlsx`、`文件名3.xlsx`。

## 文件结构

- `SKILL.md`：skill 触发与工作流说明
- `agents/openai.yaml`：UI 元数据
- `scripts/wecom_doc_download.mjs`：主执行脚本
- `examples/`：批量任务模板
- `.github/workflows/validate.yml`：基础校验流程

## 注意

- 登录态保存在 `state/chromium-profile/`，默认不提交到 Git
- 登录态不是永久有效，站点失效后仍需重新扫码
- 本项目默认以“尽量不影响台前”为优先策略，除首次登录外优先无界面执行
- 普通文档与表格在用户视角都走“文档操作 -> 导出”，但页面底层实现不同，脚本已分别兼容

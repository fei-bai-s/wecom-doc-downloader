# wecom-doc-downloader

用于下载企微文档、腾讯文档、知识库子链接文件，并优先复用本地登录态。

## 功能

- 下载单个 `doc.weixin.qq.com` 文档
- 下载知识库中的指定子文件
- 导出为 `docx` / `pdf` / `png`
- 标题核对后将 `docx` 转为 `md`
- 支持串行批量任务
- 首次扫码后复用持久化浏览器 profile
- 遇到扫码、验证码、风控、权限申请时立即停下，等待人工介入

## 安装

```bash
cd scripts
npm install
npx playwright install chromium
```

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
  --output-dir "~/Downloads/wecom-docs"
```

批量任务中任一项遇到人工步骤或错误时，整批会停在当前项，并返回已完成结果。

## 文件结构

- `SKILL.md`：skill 触发与工作流说明
- `agents/openai.yaml`：UI 元数据
- `scripts/wecom_doc_download.mjs`：主执行脚本
- `examples/`：批量任务模板

## 注意

- 登录态保存在 `state/chromium-profile/`，默认不提交到 Git
- 登录态不是永久有效，站点失效后仍需重新扫码
- 本项目默认以“尽量不影响台前”为优先策略，除首次登录外优先无界面执行

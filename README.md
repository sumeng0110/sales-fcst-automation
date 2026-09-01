# Sales Forecast Automation

Utilities for rolling a monthly sales forecast package forward across forecast
periods, updating Excel external links, cutting downstream handoff workbooks, and
refreshing calculated outputs.

This public repository contains code and sanitized examples only. It does not
include source workbooks, finance-share paths, customer data, or local caches.

## What It Automates

The workflow is designed around three core forecast workbooks:

```text
monthly feedback + FX rate workbook
        ↓
1. Intercompany sales workbook
        ↓
2. Greater China sales workbook
        ↓
3. China sales workbook
        ↓
SCM handoff workbooks
```

The scripts support:

- copying a period folder and renaming the core workbooks;
- rewriting Excel external-link targets without opening Excel;
- rolling recurring monthly inputs such as FX rate workbooks and input templates;
- checking Smart View / HFM grids before recalculation;
- refreshing workbooks through Excel COM;
- cutting `- wo adj` snapshots for downstream SCM use;
- relinking numbered SCM handoff files in dependency order;
- updating P120-style FX formulas on `Cover` for RMB/TWD GAAP rates.

## Repository Layout

| Path | Purpose |
| --- | --- |
| `src/run_month.py` | Main monthly roll: copy folder, rewrite links, optionally refresh in Excel |
| `src/relink.py` | External-link planning and OOXML rewrite logic |
| `src/p120_fx.py` | Rolls P120 `Cover` FX formulas for RMB/TWD GAAP rates |
| `src/cut_wo_adj.py` | Creates `- wo adj` snapshots and blanks configured adjustment rows |
| `src/scm.py` | Creates/relinks SCM handoff folders and numbered workbooks |
| `src/excel_refresh.py` | Opens workbooks in Excel, updates links, recalculates, and saves |
| `src/hfm_guard.py` | Checks whether HFM/Smart View grids cover the target actual month |
| `tools/` | Workbook inspection and comparison utilities |
| `web/` | Small local FastAPI UI for plan/apply workflows |
| `config/roll.json` | Sanitized example configuration with placeholder paths |

## Configuration

Edit `config/roll.json` for your environment before running the tools. The public
version uses placeholder paths such as:

```json
{
  "share_prefix": "\\\\fileserver.example.com\\Finance",
  "period_root": "\\\\fileserver.example.com\\Finance\\Shared Finance Root\\Sales Forecast\\A-P1\\CY26",
  "scm_root": "\\\\fileserver.example.com\\Finance\\Shared Finance Root\\Sales Forecast\\SCM"
}
```

For a real deployment, keep private paths and local overrides out of git. One
simple pattern is to maintain a local copy such as `config/roll.local.json` and
pass it with `--config`; this file is ignored by `.gitignore`.

## Monthly Usage

Preview the folder copy and link rewrite:

```powershell
python src\run_month.py --from '6+6' --to '7+5'
```

Apply the folder copy and external-link rewrite:

```powershell
python src\run_month.py --from '6+6' --to '7+5' --apply
```

Roll the P120 `Cover` FX formulas for the target period:

```powershell
python src\p120_fx.py --period '7+5' --apply
```

After Smart View/HFM grids have been refreshed manually, recalculate the core
workbooks:

```powershell
python src\run_month.py --from '6+6' --to '7+5' --apply --refresh
```

After sign-off, prepare SCM handoff files:

```powershell
python src\scm.py --from '6+6' --to '7+5' --create --apply
python src\cut_wo_adj.py --period '7+5' --apply
python src\scm.py --from '6+6' --to '7+5' --apply
```

## Local Web UI

The `web/` folder contains a small local FastAPI app that exposes the same
workflow as browser-friendly APIs. It is meant to run on the same machine that
can reach the finance share and automate Excel.

Start it from the project root:

```powershell
.\.venv\Scripts\python.exe -m uvicorn web.app:app --port 8765
```

Then open:

```text
http://127.0.0.1:8765
```

The UI/API supports:

- listing existing forecast periods and suggesting the next period;
- previewing and applying the period-folder copy;
- previewing and applying workbook external-link rewrites;
- checking HFM / Smart View grid freshness;
- triggering Excel recalculation;
- comparing two workbooks;
- uploading newly received input files into the target period folder.

Long-running actions are started as background jobs. The browser polls
`/api/jobs/{job_id}` for progress logs and final results, so copy/refresh tasks
can keep running without blocking the page.

## Safety Notes

- The scripts default to preview mode where possible; use `--apply` to write.
- Source workbooks are not committed and should remain on a controlled file share.
- Excel workbooks are excluded by `.gitignore`.
- Local caches, virtual environments, and sandbox folders are excluded.
- HFM/Smart View refresh remains manual because stale actuals can silently produce
  plausible but wrong forecast outputs.

## Development

Python 3.13 was used during development. Excel automation requires Microsoft
Excel and `pywin32`; the local web UI additionally uses FastAPI/Uvicorn.

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

---

## 中文版

这是一个用于月度销售预测滚动的 Excel 自动化工具集，覆盖期间推进、外部链接更新、下游交接表生成，以及计算结果刷新等重复性步骤。

这个公开仓库只包含代码和脱敏示例，不包含原始工作簿、真实共享盘路径、客户数据或本地缓存。

## 自动化范围

流程围绕三张核心预测工作簿展开：

```text
月度业务反馈 + FX Rate 工作簿
        ↓
1. Intercompany sales 工作簿
        ↓
2. Greater China sales 工作簿
        ↓
3. China sales 工作簿
        ↓
SCM 下游交接工作簿
```

脚本支持：

- 复制期间文件夹，并按目标期间重命名核心工作簿；
- 不启动 Excel，直接改写工作簿里的外部链接目标；
- 推进 FX Rate、输入模板等带月份戳的周期性输入；
- 在重算前检查 Smart View / HFM 网格是否覆盖目标实际月份；
- 通过 Excel COM 更新链接、全量重算并保存；
- 生成给 SCM 使用的 `- wo adj` 快照；
- 按依赖顺序重指 SCM 编号交接表；
- 更新 P120 风格 `Cover` 表里 RMB/TWD GAAP Rate 的月度汇率公式。

## 目录结构

| 路径 | 说明 |
| --- | --- |
| `src/run_month.py` | 月度主入口：复制文件夹、重指链接，可选启动 Excel 重算 |
| `src/relink.py` | 外部链接计划生成与 OOXML 改写逻辑 |
| `src/p120_fx.py` | 滚动 P120 `Cover` 表的 RMB/TWD GAAP Rate 公式 |
| `src/cut_wo_adj.py` | 生成 `- wo adj` 快照，并清空配置里的调整行 |
| `src/scm.py` | 创建 / 重指 SCM 交接文件夹和编号工作簿 |
| `src/excel_refresh.py` | 通过 Excel 打开工作簿、更新链接、重算并保存 |
| `src/hfm_guard.py` | 检查 HFM / Smart View 网格是否覆盖目标实际月份 |
| `tools/` | 工作簿检查、缓存比对、逐格比较等工具 |
| `web/` | 本地 FastAPI 小界面，用于预览和执行流程 |
| `config/roll.json` | 使用占位路径的脱敏配置示例 |

## 配置

运行前需要按自己的环境修改 `config/roll.json`。公开版使用的是占位路径，例如：

```json
{
  "share_prefix": "\\\\fileserver.example.com\\Finance",
  "period_root": "\\\\fileserver.example.com\\Finance\\Shared Finance Root\\Sales Forecast\\A-P1\\CY26",
  "scm_root": "\\\\fileserver.example.com\\Finance\\Shared Finance Root\\Sales Forecast\\SCM"
}
```

真实路径和本地配置不要提交到 git。建议维护一份本地配置，例如 `config/roll.local.json`，运行时通过 `--config` 指定；该文件已被 `.gitignore` 忽略。

## 月度用法

预览文件夹复制和链接更新：

```powershell
python src\run_month.py --from '6+6' --to '7+5'
```

执行文件夹复制和外部链接更新：

```powershell
python src\run_month.py --from '6+6' --to '7+5' --apply
```

滚动 P120 `Cover` 表的 FX Rate 公式：

```powershell
python src\p120_fx.py --period '7+5' --apply
```

人工刷新 Smart View / HFM 网格后，重算核心工作簿：

```powershell
python src\run_month.py --from '6+6' --to '7+5' --apply --refresh
```

定稿后准备 SCM 交接表：

```powershell
python src\scm.py --from '6+6' --to '7+5' --create --apply
python src\cut_wo_adj.py --period '7+5' --apply
python src\scm.py --from '6+6' --to '7+5' --apply
```

## 本地网页界面

`web/` 文件夹里是一个本地 FastAPI 应用，把同一套自动化流程包装成浏览器可以调用的 API。它应该运行在能访问共享盘、也能调用本机 Excel 的电脑上。

在项目根目录启动：

```powershell
.\.venv\Scripts\python.exe -m uvicorn web.app:app --port 8765
```

然后浏览器打开：

```text
http://127.0.0.1:8765
```

网页 / API 支持：

- 列出已有预测期间，并建议下一个期间；
- 预览和执行期间文件夹复制；
- 预览和执行工作簿外部链接重指；
- 检查 HFM / Smart View 网格是否覆盖目标实际月份；
- 触发 Excel 重算；
- 比较两个工作簿；
- 把新收到的输入文件上传到目标期间目录。

复制、重算、比较这类耗时动作会作为后台任务运行。前端通过 `/api/jobs/{job_id}` 轮询进度日志和最终结果，所以页面不用一直卡住等 Excel 或共享盘。

<img width="1606" height="1000" alt="Local Web UI workflow" src="https://github.com/user-attachments/assets/b6db7ae7-8eb6-4049-9b03-806a4edff43b" />

<img width="1605" height="1003" alt="Local Web UI actions" src="https://github.com/user-attachments/assets/e34429b9-40af-4083-bbe7-9bbbd828d4c2" />

## 安全说明

- 脚本尽量默认只预览，真正写入需要加 `--apply`。
- 原始 Excel 工作簿不应提交到仓库，应保留在受控文件共享环境中。
- `.gitignore` 已排除 Excel 文件、本地缓存、虚拟环境和沙盒目录。
- HFM / Smart View 刷新仍保留人工操作，因为实际数如果取错月份，结果可能看起来合理但实际错误。

## 开发环境

开发时使用 Python 3.13。Excel 自动化需要本机安装 Microsoft Excel 和 `pywin32`；本地 Web UI 还需要 FastAPI / Uvicorn。

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

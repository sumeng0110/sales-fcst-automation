# Workflow Notes

This document describes the sanitized workflow encoded by the scripts. It avoids
private file-share names, internal department names, people names, customer names,
and workbook data values.

## Forecast Chain

The monthly roll is a dependency chain:

```text
monthly feedback + FX rate workbook
        ↓
intercompany sales workbook
        ↓
regional sales workbook
        ↓
final sales workbook
        ↓
SCM handoff workbooks
```

Each step should be refreshed in dependency order. Downstream workbooks can keep
valid-looking cached values even when their external links point at stale files,
so link targets and refresh order both matter.

## Period Rules

Forecast periods use `n+(12-n)` notation. Examples:

- `6+6` means six actual months and six forecast months.
- `7+5` means seven actual months and five forecast months.

When rolling from one period to the next:

- files in the managed forecast period tree advance by one period;
- baseline periods remain pinned;
- prior-year comparison links remain pinned;
- links outside the managed forecast tree are pinned unless a specific SCM rule
  says otherwise;
- ambiguous matches are reported instead of guessed.

## Recurring Monthly Inputs

Some inputs live outside the period folder but carry a month stamp in the file
name, such as:

```text
Rate YYYYMMDD.xlsx
YYYYMM input template.xlsx
```

Only the month stamp in the file name is rolled. Version folders or other parent
folder names are not modified.

## FX Formula Rule

The P120-style `Cover` sheet has RMB/TWD GAAP rate rows that require more than a
plain external-link rewrite.

For a target period:

- actual months use `GC Avg Rate`;
- the first forecast month uses `GC Corporate Rate` for the latest actual month;
- later forecast months equal the previous month.

Example for `7+5`:

```text
D:J  -> GC Avg Rate, Jan-Jul
K    -> GC Corporate Rate, latest actual month
L:O  -> previous month
```

Use:

```powershell
python src\p120_fx.py --period '7+5' --apply
```

The script writes through Excel COM so Excel preserves formulas, shared formulas,
and cached values correctly.

## HFM / Smart View

The scripts do not refresh Smart View grids. Actuals must be refreshed manually in
Excel before recalculation. `src/hfm_guard.py` checks whether grids appear to cover
the required actual month:

```powershell
python src\hfm_guard.py <workbook> --period '7+5'
```

## Wo Adj Snapshots

After forecast sign-off, downstream handoff uses `- wo adj` snapshots. Most
workbooks are copied unchanged; configured adjustment rows are blanked in Excel and
the workbook is recalculated.

```powershell
python src\cut_wo_adj.py --period '7+5' --apply
```

The adjustment rows are configured in `config/roll.json`.

## SCM Handoff

SCM handoff workbooks are numbered and must be processed in order. A typical
sequence is:

```powershell
python src\scm.py --from '6+6' --to '7+5' --create --apply
python src\cut_wo_adj.py --period '7+5' --apply
python src\scm.py --from '6+6' --to '7+5' --apply
```

The SCM rules differ from the core forecast tree:

- numbered handoff workbooks move with the target period;
- links between numbered SCM workbooks are kept in dependency order;
- target-period transfer-ratio and forecast-data files are carried from the prior
  SCM folder when they have already been prepared there;
- forecast-data files can keep a source-data `YYMM` stamp while rolling only the
  period token in the file name.

## Publishing Safety

Before publishing this project:

- do not commit Excel workbooks;
- do not commit `.venv`, `.cache`, or `.sandbox`;
- keep real share paths in local config only;
- remove names of people, customers, internal departments, and servers from docs;
- search for UNC paths, drive-letter paths, IP addresses, and email-style tokens.

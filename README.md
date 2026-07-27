# MilnePM Invoice Schedule — Release Artifacts

This repo hosts the VBA source files for the MilnePM Invoice Schedule macro.
It is the canonical, version-controlled source that the macro's **Check for Updates** button pulls from at runtime.

## Files

| File | Purpose |
|---|---|
| `modGenerateSchedule.bas` | The macro that generates the PAYMENTS sheet. Replaced on every release. The first few lines hold a `Public Const MODULE_VERSION` that the updater reads. |
| `modUpdater.bas` | Self-update bootstrap shipped once with Update 2. Immutable — never replaced by the updater. |

## How updates flow

```
Josiah pushes new modGenerateSchedule.bas to main
                  │
                  ▼
   raw.githubusercontent.com  (public, no auth)
                  │
                  ▼
User clicks "Check for Updates" on REPORT_SETTINGS
                  │
                  ▼
Macro is swapped in place; user saves the workbook
```

## Releasing a new version

1. In the development repo, edit `modGenerateSchedule.bas` and bump `MODULE_VERSION` (semver-ish, e.g. `3.0` → `3.1`). If the change alters the input-sheet layout, also bump `SCHEMA_VERSION` and add a matching `EnsureSchema` migration step.
2. Run the release gate from the dev repo: `python tmp/release.py --message "Update X.Y — <summary>"` — it verifies the version advanced, runs the full regression suite (including the upgrade-from-shipped-workbook scenario), copies the file here, and commits.
3. `git push origin main` (manual, always).
4. Notify the team with a one-line email: *"Update X.Y is out: <change>. Click Check for Updates."*

## ⚠ Rollbacks: FIX FORWARD ONLY (since 3.0)

Since 3.0 the module migrates workbook schema in place (`EnsureSchema`): 3.0 reordered the COMMITTED sheet (DISCIPLINE inserted at column B). A pre-3.0 module running on a migrated workbook mis-reads that layout — fee items parse as company names — so **never push a module older than 3.0 to `main`**. To undo a bad release, publish a *forward* commit with a **higher** version that reverts the behaviour. (The `Roll back to 2.0…` commits in this log predate schema migrations and were safe then; that door is now closed.)

## Source

Full development repo, plans, harness, and project docs are kept private at Saints Technologies. Contact josiah@saintstech.org for source access.

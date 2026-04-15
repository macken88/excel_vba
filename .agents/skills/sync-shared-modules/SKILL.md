---
name: sync-shared-modules
description: Sync shared VBA modules from `src/shared/` into each workbook's VBA source directory for this repository. Use when Codex needs to propagate changes to shared modules across all workbooks that declare them in `config/*.toml` under `[shared].modules`, keeping the text source in sync before any Excel import.
---

# Sync Shared Modules

Use this skill to copy updated shared modules from `src/shared/` into each workbook's tracked VBA directory.

## Workflow

1. Run [`scripts/sync-shared-modules.ps1`](scripts/sync-shared-modules.ps1) from the repository root.
2. The script must:
   - scan all `config/*.toml` files
   - for each config, read `[project].vba_directory` and `[shared].modules`
   - skip configs that have no `[shared].modules` entries
   - for each declared module, copy `src/shared/<module>` → `<vba_directory>/<module>`
   - verify the source file exists in `src/shared/` before copying; stop and report if missing
   - report each file copied (source → destination)
3. After the script completes, confirm which files were copied and remind the user to commit the changes before running `import-vba`.

## Constraints

- Do not modify `src/shared/` itself. This skill only copies **from** `src/shared/`.
- Do not run `import-vba` as part of this skill. Syncing text files and importing into Excel are separate steps.
- Do not change `config/*.toml` as part of the sync step.
- If a declared module is missing from `src/shared/`, stop immediately and report the missing file. Do not create placeholder files.

## Output Expectations

Summarize:

- how many configs were processed
- which files were copied (from → to) per workbook
- whether any modules were missing from `src/shared/`
- reminder that `git diff src/` should be reviewed and committed before running `import-vba`

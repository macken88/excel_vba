---
name: import-vba-all
description: Import VBA source from `src/` into the configured Excel workbooks for this repository. Use when Codex needs to push tracked `.bas` / `.cls` / `.frm` files into the `.xlsm` workbooks defined by `config/*.toml`, with a preflight check for Excel VBA Trust Access before import.
---

# Import VBA All

Use this skill for controlled Excel-side import from repository source.

## Workflow

1. Run [`scripts/import-vba-all.ps1`](scripts/import-vba-all.ps1) from the repository root.
2. The script must:
   - verify the repo-local venv and `excel-vba.exe`
   - run `excel-vba check`
   - stop if Trust Access is disabled
   - read `file` and `vba_directory` from each repository `[project]` config
   - import each configured workbook with explicit `excel-vba --file ... --vba-directory ...` arguments
3. Review the command output and report whether each workbook import succeeded.

## Constraints

- Do not edit `.xlsm` files directly outside `vba-edit`.
- Do not skip the Trust Access check.
- Do not change `config/*.toml` as part of the import step.
- Keep workbook import as a separate step from source editing.
- Do not assume `excel-vba --config` understands this repository's `config/*.toml` format.

## Output Expectations

Summarize:

- which config files were processed
- which workbook imports succeeded or failed
- whether Excel-side state now matches the tracked source

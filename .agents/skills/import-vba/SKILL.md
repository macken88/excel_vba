---
name: import-vba
description: Import VBA source from `src/` into selected Excel workbooks for this repository. Use when Codex needs to push tracked `.bas` / `.cls` / `.frm` files into one or more `.xlsm` workbooks defined by `config/*.toml`, with a preflight check for Excel VBA Trust Access before import.
---

# Import VBA

Use this skill for controlled Excel-side import into explicitly selected workbooks from repository source.

## Workflow

1. Run [`scripts/import-vba.ps1`](scripts/import-vba.ps1) from the repository root and pass one or more workbook file paths.
2. The script must:
   - verify the repo-local venv and `excel-vba.exe`
   - run `excel-vba check`
   - stop if Trust Access is disabled
   - read `file` and `vba_directory` from each repository `[project]` config
   - resolve only the requested workbook file paths
   - stop if any requested workbook path cannot be matched exactly
   - show candidate paths when the request is close but not exact, then ask the user to rerun with an exact value
   - import only the resolved workbooks with explicit `excel-vba --file ... --vba-directory ...` arguments
3. Review the command output and report whether each selected workbook import succeeded.

## Constraints

- Do not edit `.xlsm` files directly outside `vba-edit`.
- Do not skip the Trust Access check.
- Do not change `config/*.toml` as part of the import step.
- Keep workbook import as a separate step from source editing.
- Do not assume `excel-vba --config` understands this repository's `config/*.toml` format.
- Treat `[project].file` in `config/*.toml` as the canonical workbook identifier for matching.
- If a workbook argument is ambiguous or unmatched, do not guess. Stop and ask the user to confirm the exact workbook path.

## Output Expectations

Summarize:

- which requested workbook paths were resolved
- which config files were processed
- which workbook imports succeeded or failed
- whether Excel-side state now matches the tracked source

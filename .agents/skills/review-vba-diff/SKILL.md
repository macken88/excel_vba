---
name: review-vba-diff
description: Review the Git diff of VBA source files in `src/` before importing into Excel workbooks. Use when Codex needs to show pending text changes, confirm the diff is intentional, and produce a human-readable summary of what will change in each workbook after import.
---

# Review VBA Diff

Use this skill to inspect and summarize pending VBA source changes before committing or importing them into Excel.

## Workflow

1. Run [`scripts/review-vba-diff.ps1`](scripts/review-vba-diff.ps1) from the repository root.
2. The script must:
   - run `git diff HEAD -- src/` to show unstaged changes
   - run `git diff --cached HEAD -- src/` to show staged changes
   - combine both outputs if both exist
   - print the raw diff
   - list which workbook directories under `src/` have changes, based on `config/*.toml`
3. After the script output, provide a plain-language summary of each changed file: what was added, removed, or modified.
4. Ask the user to confirm the diff is intentional before proceeding to commit or `import-vba`.

## Constraints

- This skill is read-only. Do not stage, commit, or modify any files.
- Do not run `import-vba` automatically. The user must explicitly request it after reviewing the diff.
- If there are no changes in `src/`, report that clearly and stop.

## Output Expectations

Summarize:

- whether there are staged or unstaged changes in `src/`
- which source files changed and in which workbook directory
- a brief plain-language description of each changed file's diff
- whether shared modules in `src/shared/` were changed
- confirmation prompt before the user proceeds to commit or import

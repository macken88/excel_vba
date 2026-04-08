---
name: edit-vba
description: Edit VBA source files for this repository under `src/` only. Use when Codex needs to add or modify `.bas`, `.cls`, or `.frm` source files for a workbook, while keeping `.xlsm` binaries untouched and following the repository rule that VBA edits must go through a project-defined skill.
---

# Edit VBA

Use this skill for text-based VBA edits in `src/`.

## Workflow

1. Confirm the task is a VBA source edit, not direct workbook manipulation.
2. Edit only the relevant files under `src/`.
3. Keep changes minimal and local to the requested workbook unless the user explicitly asks for shared modules.
4. Preserve workbook-specific structure and module naming unless there is a clear need to introduce a new module.
5. After editing, report which source files changed and whether Excel-side import is required.

## Constraints

- Do not edit `.xlsm` files directly.
- Do not change `config/*.toml` unless the task explicitly requires it.
- If shared logic is needed, update `src/shared/` first, then sync into workbook-specific directories in a separate step.
- Prefer simple standard modules for new test macros unless the user asks for forms or class modules.

## Output Expectations

Summarize:

- changed files under `src/`
- what each new or changed macro does
- whether the workbook now needs re-import from source

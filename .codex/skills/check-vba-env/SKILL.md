---
name: check-vba-env
description: Verify the Windows-local VBA maintenance environment for this repository. Use when Codex needs to confirm `.venv-vba-tools`, `vba-edit`, `config/*.toml`, workbook paths, or Git initialization status before or after export/import/sync work.
---

# Check VBA Env

Use this skill for read-only validation of the repo's VBA maintenance prerequisites.

## Workflow

1. Run [`scripts/check-vba-env.ps1`](scripts/check-vba-env.ps1) from the repository root.
2. Review the reported status for:
   - Python and venv discovery
   - installed `vba-edit` version
   - config file presence
   - whether each config uses the repository `[project]` mapping expected by local skills
   - workbook path existence for each TOML file
   - Git repository initialization
3. Report any blockers concretely, especially workbook extension mismatches such as `.xlsx` files where `.xlsm` is expected.

## Constraints

- Keep this skill read-only.
- Do not "fix" workbook/config mismatches automatically unless explicitly asked.
- Use it before any sync/import/export action that depends on `vba-edit`.
- Remember that `config/*.toml` is repository metadata, not necessarily a `vba-edit --config` file that can be passed through directly.

## Output Expectations

Summarize:

- whether the local toolchain is usable
- whether Git is initialized
- which config entries point at missing workbook files
- whether the config shape matches the repository's expected `[project]` structure
- whether Excel-side re-import can proceed

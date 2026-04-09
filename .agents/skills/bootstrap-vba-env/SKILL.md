---
name: bootstrap-vba-env
description: Create and refresh the Windows-local Python environment used for VBA maintenance in this repository. Use when Codex needs to perform the documented setup flow for this repo: create `.venv-vba-tools`, upgrade `pip`, install `requirements-vba-tools.txt`, or re-bootstrap the fixed `vba-edit` toolchain before export/import work.
---

# Bootstrap VBA Env

Use this skill only for deterministic environment bootstrap on Windows. Treat the repository root as the working directory and keep the environment local to `.venv-vba-tools`.

## Workflow

1. Confirm that the task is environment bootstrap, not workbook editing.
2. Run [`scripts/bootstrap-vba-env.ps1`](scripts/bootstrap-vba-env.ps1) from the repository root.
3. Read the script output and verify:
   - `.venv-vba-tools` exists
   - `pip` upgrade completed
   - `requirements-vba-tools.txt` installed successfully
   - `vba-edit` can be queried from the venv
4. If bootstrap fails because Python is unavailable, stop and report the exact missing executable path.
5. After bootstrap, run `$check-vba-env` or [`../check-vba-env/scripts/check-vba-env.ps1`](../check-vba-env/scripts/check-vba-env.ps1) to verify the resulting state.

## Constraints

- Prefer the repo-local `.venv-vba-tools`; do not install tools globally.
- Do not change `requirements-vba-tools.txt` just to make bootstrap pass.
- Do not touch `.xlsm` files or invent workbook-sync steps here.
- Keep the bootstrap deterministic: Python path discovery, venv creation, pip upgrade, pinned dependency install, version check.

## Python Discovery

The bootstrap script already checks common Windows locations. Use the script as-is unless the repository documents a new standard Python location.

## Output Expectations

When reporting back, summarize:

- whether `.venv-vba-tools` was created or reused
- the Python executable selected
- whether `vba-edit==0.4.4` is installed
- whether follow-up Excel import/export work is blocked by workbook/config mismatches

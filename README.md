# VBA Maintenance Template

This repository template is intended for maintaining Excel workbooks with VBA code under Git.

## Structure

- `workbooks/`: Excel workbook binaries such as `.xlsm`
- `src/BookA_vba/`: exported VBA source for BookA
- `src/BookB_vba/`: exported VBA source for BookB
- `src/shared/`: shared VBA modules used across multiple workbooks
- `config/`: repository-side workbook mapping files consumed by Codex skills
- `AGENTS.md`: operating rules for Codex and human maintainers
- `requirements-vba-tools.txt`: pinned Python tool dependencies

## Setup example

Create a dedicated environment on Windows:

```powershell
python -m venv .venv-vba-tools
.\\.venv-vba-tools\\Scripts\\python.exe -m pip install --upgrade pip
.\\.venv-vba-tools\\Scripts\\python.exe -m pip install -r requirements-vba-tools.txt
```

## Important

- Edit VBA text sources under `src/`, not workbook binaries directly.
- Use Codex skills to standardize synchronization procedures.
- Update `config/*.toml` to match actual workbook filenames and source directories.
- `config/*.toml` uses this repository's `[project]` format. Treat it as project metadata for skills, not as a `vba-edit` CLI config file to pass directly to `excel-vba --config`.
- Before import/export, enable Excel's `Trust access to the VBA project object model`.
- Import/export skills should translate `config/*.toml` into explicit `excel-vba --file ... --vba-directory ...` arguments.

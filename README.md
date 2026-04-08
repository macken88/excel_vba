# VBA Maintenance Template

This repository template is intended for maintaining Excel workbooks with VBA code under Git.

## Structure

- `workbooks/`: Excel workbook binaries such as `.xlsm`
- `src/BookA_vba/`: exported VBA source for BookA
- `src/BookB_vba/`: exported VBA source for BookB
- `src/shared/`: shared VBA modules used across multiple workbooks
- `config/`: `vba-edit` TOML configuration files
- `AGENTS.md`: operating rules for Codex and human maintainers
- `requirements-vba-tools.txt`: pinned Python tool dependencies

## Setup example

Create a dedicated environment on Windows:

```powershell
py -3.11 -m venv .venv-vba-tools
.\\.venv-vba-tools\\Scripts\\python.exe -m pip install --upgrade pip
.\\.venv-vba-tools\\Scripts\\python.exe -m pip install -r requirements-vba-tools.txt
```

## Important

- Edit VBA text sources under `src/`, not workbook binaries directly.
- Use Codex skills to standardize synchronization procedures.
- Update `config/*.toml` to match actual workbook filenames.

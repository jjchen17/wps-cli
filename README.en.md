# WPS CLI

[![Python](https://img.shields.io/badge/python-%3E%3D3.10-3776AB?logo=python&logoColor=white)](https://python.org)
[![Commands](https://img.shields.io/badge/commands-36-important)]()
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows-0078D6?logo=windows&logoColor=white)]()
[![WPS](https://img.shields.io/badge/requires-WPS%20Office%202019%2B-C00000)]()

Drive WPS Office from the terminal — real formatting, real layout, 36 commands.

> Not "guess what you see" — **what you see is what you get**. Don't simulate the file format, just drive the actual WPS engine.

[中文](README.md) · [Changelog](CHANGELOG.md) · [Contributing](CONTRIBUTING.md)

---

## Why

If you've ever generated Chinese Word documents with `python-docx`, you've probably hit moments like these:

- The code says "**SimSun**" but the file opens as **Calibri**.
- English-only field names mangle Chinese paragraph layout — every indent feels like a roll of the dice.
- Table borders disappear after merging cells; headers and footers refuse to cooperate.
- `openpyxl` writes a formula, but Excel won't compute it; conditional formatting is a saga.

These libraries fundamentally **parse and assemble XML file formats**. They don't know how the font will actually render in WPS, so they can't guarantee the final output.

**wps-cli takes a different path: don't simulate the file format — drive the running WPS Office desktop process via COM automation.**

When you set "Heiti, size 3, bold," it actually executes that operation inside WPS. Fonts, spacing, page numbers, table of contents, headers and footers — every pixel is identical to manual operation, because the same program is doing the work. You can even watch the document grow on screen while debugging.

---

## Three highlights

|  | | |
|---|---|---|
| **100% formatting fidelity** | Not parsing files — directly commanding the WPS engine | What you see is what you get |
| **AI Agent native** | All commands support `--json` with a unified schema | Built for LLM agents and automation pipelines |
| **One tool for four apps** | Writer / Calc / Impress / PDF, 36 commands | Office automation in one place |

---

## Install

**Requirements**: Windows 10/11 · WPS Office 2019+ · Python 3.10+

```bash
pip install wps-cli
```

```bash
# Verify install
wps doctor
```

---

## Three-step demo

Three lines turn `budget.xlsx` into a chart + extracted data + PDF:

```bash
# Step 1: insert a bar chart
wps calc chart-create budget.xlsx --data "A1:C10" --type bar --title "2026 sales"

# Step 2: extract a value as JSON for downstream scripts or LLMs
wps calc cell-get budget.xlsx C12 --json

# Step 3: export to PDF with WPS-grade fidelity
wps export convert budget.xlsx pdf
```

---

## Commands

```
wps
├── writer    Word docs    (new, info, replace, count, table-insert, table-get,
│                            image-insert, page-setup, style-apply, export-pdf)
├── calc      Excel        (new, info, sheet-list, cell-get/-set/-range/-formula,
│                            chart-create, sort, export-csv)
├── impress   PPT          (new, info, slide-list/-add/-delete,
│                            text-set/-get, image-insert, export-pdf)
├── pdf       PDF          (info, merge, extract-pages, split, watermark)
├── export    Conversion   (convert, batch)
├── version   Print version
└── doctor    Environment diagnostics
```

All commands support `--json` for AI-friendly structured output:

```json
{ "success": true, "command": "calc.cell_get", "data": {"ref": "C12", "value": 285600} }
```

On error:

```json
{
  "success": false,
  "command": "calc.cell_set",
  "error": { "type": "ValidationError", "message": "...", "code": 50,
              "suggestion": "...", "context": {} }
}
```

---

## Exit codes

| Code | Meaning |
|----:|---------|
| 0  | Success |
| 1  | Generic error |
| 10 | WPS not installed / detected |
| 11 | Session management error |
| 20 | File operation failed |
| 21 | File not found |
| 30 | COM call failed |
| 40 | Unsupported format |
| 50 | Validation failed |
| 60 | Operation timed out |
| 61 | Partial batch failure |

---

## Security

By default, wps-cli applies several hardenings to avoid common Office automation abuse paths:

- **Macros disabled**: every document open call sets `AutomationSecurity = msoAutomationSecurityForceDisable`, blocking `Auto_Open` / `Document_Open` style macros.
- **Formula injection blocked**: `calc cell-formula` rejects formulas containing `SHELL`, `DDE`, `DDEAUTO`, `EXEC`, `CALL`, `REGISTER`, `HYPERLINK`. `calc cell-set` rejects values starting with `=` to prevent secondary injection.
- **Path constraints**: file arguments must not be UNC paths or symlinks. `export batch` rejects absolute glob patterns and confines matches to the current working directory.
- **Page-range bombs blocked**: `pdf extract-pages` enforces upper bounds on page numbers and range size.
- **Path redaction**: JSON error responses redact local paths so logs/agent contexts don't leak filesystem structure.

See [SECURITY.md](SECURITY.md) for the full policy.

---

## Architecture

```
CLI layer (Typer)
    │
    ▼
Service layer
    │
    ▼
COM backend (abstract base + WPS impl)
    │
    ▼
WPS Office desktop
```

The COM backend is pluggable — replacing it with a LibreOffice/WebOffice backend is on the roadmap.

---

## Develop

```bash
git clone https://github.com/jjchen17/wps-cli.git
cd wps-cli
pip install -e ".[dev]"
pytest
```

See [CONTRIBUTING.md](CONTRIBUTING.md).

---

## Feedback

Open an [Issue](https://github.com/jjchen17/wps-cli/issues) or email **[948881912@qq.com](mailto:948881912@qq.com)**.

## License

[MIT](LICENSE)

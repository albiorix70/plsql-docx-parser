# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Oracle PL/SQL packages for parsing DOCX files inside an Oracle APEX environment. The parsed content is intended for downstream PDF generation via pdfmake. All runtime dependencies (`apex_zip`, `dbms_xmldom`, `logger`) are provided by the Oracle/APEX platform.

## Build & Compile

Compile all packages in dependency order using SQL*Plus:

```sql
-- From SQL*Plus connected to the target schema:
@compile_docs_parser.sql
```

This runs: `docx_parser_util.pks` → `docx_parser_util.pkb` → `docx_parser.pks` → `docx_parser.pkb`

VS Code tasks (`.vscode/tasks.json`) reference CI scripts that are not yet in the repo:
- `./ci/run_compile.sh` (Bash) / `.\ci\run_compile.ps1` (PowerShell)
- `./ci/run_mcp_cli.sh` / `.\ci\run_mcp_cli.ps1`

## Running Tests

Tests use **utPLSQL**. Run individual suites from SQL*Plus:

```sql
-- Full parser suite
exec ut.run('test_docx_parser');

-- Rels and unpack suite
exec ut.run('tests.test_docx_parser_rels');

-- Styles JSON manual test (not utPLSQL, just runs inline)
@tests/test_styles_json.sql
```

Upload a real DOCX blob for integration tests:

```bash
python scripts/generate_test_docx.py        # generates tests/test_doc.docx
python scripts/upload_test_docx.py          # uploads to DB table test_docx_store (id=1)
# or PowerShell:
powershell scripts/upload_test_docx.ps1
```

Manual testlauf (ad-hoc execution in SQL*Plus):

```sql
@scripts/testlauf.sql
```

## Package Architecture

### `docx_parser` (top-level parser)
- **Public API**: `parse_content_xml(p_content_xml CLOB) RETURN CLOB`
- Iterates over `w:body/*` children using `XMLTABLE` with the `w:` namespace
- Currently handles `w:p` (paragraph) elements; calls internal `parse_paragraph_node` and `ppr_to_json`
- Package-level state: `lg_content_element` (JSON array), `lg_document_json` (JSON object)

### `docx_parser_util` (utility layer)
Stateful package — styles and the loaded DOCX are held in package-global variables for the session:

| Function/Procedure | Purpose |
|---|---|
| `xml_to_json(CLOB)` | Recursively converts any XML to `json_object_t` using `dbms_xmldom`. Attributes → `@localName`, text → `#text`, multiple same-name children → JSON array |
| `parse_styles_xml(CLOB)` | Parses `styles.xml`; populates `lg_style_json` keyed by style name |
| `get_styles_by_id(VARCHAR2)` | Looks up a parsed style by its `styleId` |
| `mark_style_used(VARCHAR2)` / `styles_to_json()` | Tracks which styles were referenced; exports used styles as pdfmake-compatible JSON |
| `parse_rels_xml(CLOB)` | Parses `document.xml.rels`; returns `t_rels_table` (associative array indexed by relationship `Id`) |
| `unpack_docx_clob(path, blob)` | Extracts a file from the DOCX ZIP as CLOB (uses `apex_zip`) |
| `unpack_docx_blob(path, blob)` | Extracts a file from the DOCX ZIP as BLOB |
| `load_docx_source(table, col, id_col, id_val)` | Dynamically loads a DOCX BLOB from any DB table into `g_loaded_docx` |
| `get_loaded_docx()` | Returns the session-cached DOCX BLOB |

### `pdfmake` / `pdf_make_util`
Separate packages for PDF generation using the pdfmake JavaScript library. The parsed DOCX content (JSON) feeds into these packages.

## Key Conventions

- XML parsing uses two approaches depending on context:
  - `XMLTABLE` with explicit namespace mappings for structured iteration (`docx_parser.pkb`)
  - `dbms_xmldom` DOM API for recursive tree traversal (`docx_parser_util.pkb`)
- Font sizes in DOCX are stored as half-points; the parser divides by 2 to get points.
- `dbms_assert.simple_sql_name` is used in `load_docx_source` to prevent SQL injection from dynamic table/column names.
- Errors are logged via the `logger` package (Oracle Logger); always wrap `logger` calls in their own exception handler since logger itself may not be installed in all environments.
- Style storage in `docx_parser_util` is keyed by **style name** (e.g. `"heading 1"`) in `lg_style_json`, but tracked by **style id** (e.g. `"Heading1"`) in `lg_used_styles`.

## DOCX Internal Structure (relevant paths)

| Path inside ZIP | Purpose |
|---|---|
| `word/document.xml` | Main document body |
| `word/styles.xml` | Paragraph and character style definitions |
| `word/_rels/document.xml.rels` | Relationships (images, hyperlinks, etc.) |
| `word/numbering.xml` | List/numbering definitions |

## Dependencies

- Oracle APEX (for `apex_zip` package — ZIP extraction)
- Oracle built-ins: `dbms_xmldom`, `dbms_assert`, `xmltype`, `json_object_t`, `json_array_t`
- `logger` package (Oracle Logger — optional, errors are silently swallowed if absent)
- utPLSQL (test framework)
- Python 3.8+ with `python-docx`, `Pillow` (for test DOCX generation only)

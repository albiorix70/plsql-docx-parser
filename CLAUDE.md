# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Oracle PL/SQL packages for parsing DOCX files inside an Oracle APEX environment. The parsed content is a pdfmake-compatible JSON array intended for downstream PDF generation. All runtime dependencies (`apex_zip`, `dbms_xmldom`) are provided by the Oracle/APEX platform.

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

**Public API**: `parse_docx(p_filename VARCHAR2) RETURN CLOB`

Loads the DOCX from `APEX_WORKSPACE_STATIC_FILES` (column `FILE_CONTENT`, keyed by `FILE_NAME`), unpacks `word/document.xml`, and returns the full document as a JSON array of pdfmake paragraph objects.

**Package-body state**

| Symbol | Kind | Purpose |
|---|---|---|
| `l_namespaces` | `t_ns` (assoc. array) | Namespace prefix → URI map, populated from root element attributes |
| `c_lf` | constant `varchar2(1)` | Line break — `chr(10)`; used for `w:br` (default/textWrapping) |
| `c_ff` | constant `varchar2(1)` | Page break — `chr(12)`; used for `w:br w:type="page"` |
| `c_tab` | constant `varchar2(1)` | Horizontal tab — `chr(9)`; used for `w:tab` |
| `c_exception_no_body` | exception | Raised (ORA-20001) when `w:body` is not found |
| `c_exception_no_childnodes` | exception | Raised (ORA-20002) when `w:body` has no children |

**Private functions**

| Function | Purpose |
|---|---|
| `set_namespaces(node)` | Reads `xmlns:*` attributes from the document root into `l_namespaces` |
| `parse_run_node(node)` → `json_object_t` | Parses a `w:r` element: `w:rPr` → character styles; `w:t` → text; `w:br` → `c_lf`/`c_ff`; `w:tab` → `c_tab` |
| `parse_paragraph(node)` → `json_object_t` | Parses a `w:p` element: collects `w:pPr` styles + array of `parse_run_node` results into `{"text":[...], ...pPrStyles}` |

**Body dispatch loop** iterates `w:body` children:
- `w:p` → `parse_paragraph` → appended to result array
- `w:tbl` → exits loop (table processing not yet implemented; avoids trailing empty paragraphs)

**Output shape** (one element per `w:p`):
```json
[
  { "style": "Normal", "alignment": "justify", "text": [
      { "text": "Hello ", "bold": true },
      { "text": "World",  "color": "#FF0000" }
  ]},
  { "text": [{ "text": "Second paragraph" }] }
]
```

---

### `docx_parser_util` (utility layer)

Stateful package — the loaded DOCX BLOB is held in the package-global `g_loaded_docx` for the session.

**Public API**

| Function/Procedure | Purpose |
|---|---|
| `unpack_docx_clob(path, blob)` | Extracts a ZIP entry as UTF-8 CLOB (`xmltype` conversion); returns NULL if not found |
| `unpack_docx_blob(path, blob)` | Extracts a ZIP entry as raw BLOB; returns NULL if not found |
| `load_docx_source(table, col, id_col, id_val)` | Dynamically loads a DOCX BLOB from any DB table into `g_loaded_docx`; raises ORA-20002 if no row found |
| `get_loaded_docx()` | Returns the session-cached DOCX BLOB |
| `get_style_attributes(node [, ns])` | Converts a `w:pPr` or `w:rPr` DOM node to a pdfmake-compatible `json_object_t` |
| `el_get_attribute(el, attr_name, ns)` | Reads a namespace-qualified attribute from a DOM element via `dbms_xmldom.getattribute(elem, ns, name)` |
| `el_is_toggle_on(el, attr_name, ns)` | Returns TRUE unless `w:val` is `"false"`, `"0"`, or `"off"` (handles toggle properties like `w:b`, `w:i`) |
| `dxa_to_pt(v)` | twips → points (page layout, spacing) |
| `halfpt_to_pt(v)` | half-points → points (`w:sz` font sizes) |
| `hundredthpt_to_pt(v)` | hundredths of a point → points (character spacing) |
| `eighthpt_to_pt(v)` | eighth-points → points (border widths `w:bdr`) |
| `emu_to_pt(v)` | EMU → points (DrawingML image dimensions) |
| `lineunit_to_pt(v)` | line units → points (`w:spacing/@w:line` as absolute pt) |
| `fiftieth_to_pt(v, page_w_dxa)` | fiftieths-of-percent → points (table widths `w:type="pct"`) |

**Private helper**

| | |
|---|---|
| `get_docx_file_blob(path, blob)` | Shared primitive for all ZIP extraction; centralises `apex_zip` interaction and exception handling |

**`get_style_attributes` — supported mappings**

| OOXML | pdfmake key | Notes |
|---|---|---|
| `w:rPr/w:b` | `bold` | boolean; toggle property |
| `w:rPr/w:i` | `italics` | boolean; toggle property |
| `w:rPr/w:u` | `decoration: "underline"` | skipped when `w:val="none"` |
| `w:rPr/w:strike` | `decoration: "lineThrough"` | toggle property |
| `w:rPr/w:color/@w:val` | `color` | prefixed `#RRGGBB`; skipped when `"auto"` |
| `w:rPr/w:sz/@w:val` | `fontSize` | half-points ÷ 2 |
| `w:rPr/w:rFonts` | *(disabled)* | font mapping commented out — requires pdfmake font registration |
| `w:rPr/w:spacing/@w:val` | `characterSpacing` | hundredths of a point ÷ 100 |
| `w:pPr/w:pStyle/@w:val` | `style` | style name pass-through |
| `w:pPr/w:jc/@w:val` | `alignment` | `"both"` → `"justify"`; others pass through |
| `w:pPr/w:spacing/@w:line` | `lineHeight` | **multiplier** (line units ÷ 240); 240 = 1.0× |
| `w:pPr/w:spacing/@w:before` | `marginTop` | dxa → pt |
| `w:pPr/w:spacing/@w:after` | `marginBottom` | dxa → pt |
| `w:pPr/w:ind/@w:left` | `marginLeft` | dxa → pt |
| `w:pPr/w:ind/@w:firstLine` | `indent` | dxa → pt |

## Key Conventions

- Both packages use `dbms_xmldom` DOM API exclusively for XML traversal — no `XMLTABLE`.
- Font sizes in DOCX are stored as half-points; use `halfpt_to_pt()` (divides by 2) to get points.
- `lineHeight` is a **multiplier** (`line_units / 240`), not an absolute point value — `lineunit_to_pt` is for absolute conversions only.
- `dbms_assert.simple_sql_name` is used in `load_docx_source` to prevent SQL injection from dynamic table/column names.
- `nls_charset_id('AL32UTF8')` is cached as the package constant `c_utf8_csid` (evaluated once at package init) and used in `unpack_docx_clob` for safe BLOB→CLOB conversion via `xmltype`.
- `el_get_attribute` uses the three-argument `dbms_xmldom.getattribute(elem, ns, name)` form — namespace is always passed explicitly. The default namespace is `c_ooxml_ns_w` (`http://schemas.openxmlformats.org/wordprocessingml/2006/main`).
- Non-element DOM nodes (text nodes, comment nodes) are skipped with a `getnodetype != element_node` guard before any `makeelement` call.
- Character format constants (`c_lf`, `c_ff`, `c_tab`) are declared at package-body level in `docx_parser` to avoid `chr()` literals scattered through run-parsing logic.

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
- utPLSQL (test framework)
- Python 3.8+ with `python-docx`, `Pillow` (for test DOCX generation only)

# CLAUDE.md

This file provides guidance to Claude Code when working with code in this repository.

## Project Overview

Oracle PL/SQL packages for parsing DOCX and ODT files inside an Oracle APEX environment. The parsed content is a pdfmake-compatible JSON object (`docDefinition`) intended for downstream PDF generation. All runtime dependencies are provided by the Oracle/APEX platform.

## Build & Compile

Compile all packages in dependency order using SQL*Plus:

```sql
-- From SQL*Plus connected to the target schema:
@compile_docs_parser.sql
```

Compile order:
1. `docx_parser_util.pks` → `docx_parser_util.pkb`
2. `docx_parser.pks` → `docx_parser.pkb`
3. `odt_parser.pks` → `odt_parser.pkb`

## PL/SQL Conventions

- Declare all variables **before** local subprograms in declaration sections — Oracle raises PLS-00stadard error otherwise
- `VARCHAR2` max size is **32767** in PL/SQL but **4000** in SQL contexts (XMLTABLE columns, SELECT INTO)
- Do not use `dbms_output` calls in production code
- Do not use deprecated DOM APIs — `getElementsByTagNameNS` is gone since 19c; use `getElementsByTagName(element, name, ns)` with a root element, not a document
- `dbms_xmldom.getElementsByTagName` signature: `(DOMElement, name, ns)` — element first, then name, then namespace URI

## Code Standards

- `XMLTABLE` column paths: keep all `varchar2` sizes at ≤ 4000 (SQL limit)
- DOM traversal: always guard with `getNodeType = ELEMENT_NODE` before calling `makeElement`
- Namespace URIs are declared as package-level constants — never inline string literals

---

## DOCX Parser

### Packages

| File | Role |
|---|---|
| `docx_parser.pks` / `.pkb` | Top-level parser — public API |
| `docx_parser_util.pks` / `.pkb` | Utility layer — ZIP extraction, DOM helpers, unit converters |

### Public API — `docx_parser`

```sql
function parse_docx(p_filename in varchar2) return clob;
```

Loads the DOCX BLOB from `APEX_WORKSPACE_STATIC_FILES` (keyed by `FILE_NAME`), unpacks `word/document.xml`, and returns the full document as a pdfmake JSON array.

### Public API — `docx_parser_util`

| Function / Procedure | Purpose |
|---|---|
| `unpack_docx_clob(path, blob)` | Extract ZIP entry as UTF-8 CLOB |
| `unpack_docx_blob(path, blob)` | Extract ZIP entry as raw BLOB |
| `load_docx_source(table, col, id_col, id_val)` | Load DOCX BLOB from any DB table into `g_loaded_docx`; uses `dbms_assert.simple_sql_name` to prevent SQL injection |
| `get_loaded_docx()` | Return session-cached DOCX BLOB |
| `get_style_attributes(node [, ns])` | Convert `w:pPr` / `w:rPr` DOM node → pdfmake `json_object_t` |
| `el_get_attribute(el, attr_name, ns)` | Read namespace-qualified attribute via `dbms_xmldom.getAttribute(elem, ns, name)` |
| `el_is_toggle_on(el, attr_name, ns)` | TRUE unless `w:val` is `"false"`, `"0"`, or `"off"` |
| `dxa_to_pt(v)` | twips → points |
| `halfpt_to_pt(v)` | half-points → points (`w:sz` font sizes) |
| `hundredthpt_to_pt(v)` | hundredths of a point → points (character spacing) |
| `eighthpt_to_pt(v)` | eighth-points → points (border widths) |
| `emu_to_pt(v)` | EMU → points (DrawingML dimensions) |
| `lineunit_to_pt(v)` | line units → points (absolute; 240 units = 12 pt) |
| `fiftieth_to_pt(v, page_w_dxa)` | fiftieths-of-percent → points (table widths `w:type="pct"`) |

### `get_style_attributes` — OOXML → pdfmake mappings

| OOXML attribute | pdfmake key | Notes |
|---|---|---|
| `w:rPr/w:b` | `bold` | toggle property |
| `w:rPr/w:i` | `italics` | toggle property |
| `w:rPr/w:u` | `decoration: "underline"` | skipped when `w:val="none"` |
| `w:rPr/w:strike` | `decoration: "lineThrough"` | toggle property |
| `w:rPr/w:color/@w:val` | `color` | `#RRGGBB`; skipped when `"auto"` |
| `w:rPr/w:sz/@w:val` | `fontSize` | half-points ÷ 2 |
| `w:rPr/w:spacing/@w:val` | `characterSpacing` | hundredths ÷ 100 |
| `w:pPr/w:pStyle/@w:val` | `style` | style name pass-through |
| `w:pPr/w:jc/@w:val` | `alignment` | `"both"` → `"justify"` |
| `w:pPr/w:spacing/@w:line` | `lineHeight` | multiplier (÷ 240); 240 = 1.0× |
| `w:pPr/w:spacing/@w:before` | `marginTop` | dxa → pt |
| `w:pPr/w:spacing/@w:after` | `marginBottom` | dxa → pt |
| `w:pPr/w:ind/@w:left` | `marginLeft` | dxa → pt |
| `w:pPr/w:ind/@w:firstLine` | `indent` | dxa → pt |

### DOCX ZIP paths

| Path | Purpose |
|---|---|
| `word/document.xml` | Main document body |
| `word/styles.xml` | Paragraph and character style definitions |
| `word/_rels/document.xml.rels` | Relationships (images, hyperlinks) |
| `word/numbering.xml` | List/numbering definitions |

### Output shape

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

## ODT Parser

### Package

| File | Role |
|---|---|
| `odt_parser.pks` / `.pkb` | Single self-contained package |

### Public API

```sql
-- Full parse from raw BLOB
function parse_odt(p_odt_blob in blob) return clob;

-- Low-level: pass already-extracted XML CLOBs (useful for testing)
function parse_xml(
  p_content_xml in clob,
  p_styles_xml  in clob default null
) return clob;
```

### Architecture

| Section | Procedure / Function | Purpose |
|---|---|---|
| Utility | `unzip_member(zip, name)` | Extract ODT ZIP entry → CLOB via `apex_zip` |
| Utility | `map_color(odf_color)` | ODF `fo:color` → pdfmake color string |
| Utility | `map_alignment(odf_align)` | ODF `fo:text-align` → pdfmake alignment |
| Utility | `odf_length_to_pt(len)` | ODF length string (pt/cm/mm/in/px) → points |
| Utility | `dom_get_attr(elem, ns, name)` | Null-safe `dbms_xmldom.getAttribute` wrapper |
| Utility | `dom_get_text(node)` | Recursive text collector; handles `text:s`, `text:tab`, `text:line-break` |
| Style | `write_style_props(...)` | Write pdfmake style key-value pairs into open `apex_json` object |
| Style | `write_styles(styles_xml, content_xml)` | Emit `"styles"` and `"defaultStyle"` JSON keys; merges named styles from `styles.xml` and automatic styles from `content.xml`; only emits styles referenced in the document |
| Content | `write_paragraph(node, style, level)` | Emit paragraph or heading JSON node |
| Content | `write_table(node)` | Emit pdfmake table node |
| Content | `write_list(node, style, is_ordered)` | Emit `ul` / `ol` node |
| Content | `write_content(content_xml)` | Iterate `office:text` children and dispatch to above writers |

### Style sourcing

Styles are collected from two sources and merged before emission:
- `styles.xml` — named styles (`office:styles` + `office:automatic-styles`)
- `content.xml` — document-local automatic styles (`office:automatic-styles`)

Only styles actually referenced via `@text:style-name` or `@table:style-name` in content.xml are written to the output. Ordered list detection checks style name for `NUMBER`, `NUMER`, `ENUM`, or `OL`.

### ODT ZIP paths

| Path | Purpose |
|---|---|
| `content.xml` | Document body + automatic styles |
| `styles.xml` | Named paragraph / character styles + default style |

### Output shape

```json
{
  "content": [
    { "text": "Heading", "style": "Heading1", "headlineLevel": 1 },
    { "text": "Plain paragraph", "style": "Text_Body" },
    { "text": [{ "text": "Bold " }, { "text": "word", "style": "T1" }], "style": "P1" },
    { "table": { "widths": ["*", "*"], "body": [[{ "text": "Cell" }]] } },
    { "ul": ["item 1", "item 2"] }
  ],
  "styles": {
    "Text_Body": { "fontSize": 12, "alignment": "justify" },
    "P1": { "basedOn": "Text_Body", "marginTop": 6 }
  },
  "defaultStyle": { "fontSize": 12 }
}
```

---

## Runtime Dependencies

| Dependency | Used by | Purpose |
|---|---|---|
| `apex_zip` | both parsers | ZIP extraction from BLOB |
| `apex_json` | `odt_parser` | JSON output generation |
| `dbms_xmldom` | both parsers | DOM-based XML traversal |
| `dbms_assert` | `docx_parser_util` | SQL injection prevention in dynamic queries |
| `xmltype` | both parsers | XML parsing |
| `json_object_t`, `json_array_t` | `docx_parser` | JSON construction |

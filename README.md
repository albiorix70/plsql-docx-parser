# doc-parser

Oracle PL/SQL package suite for parsing **DOCX** and **ODT** files stored as database BLOBs.  
Extracts document structure, styles, and text runs, and converts them to a **pdfmake-compatible JSON** structure for downstream PDF generation.

## Overview

Both DOCX and ODT are ZIP archives containing XML. The parsers unpack those archives in-database, traverse the XML using Oracle's native `DBMS_XMLDOM` API, resolve style inheritance, and emit structured JSON.

```
DOCX BLOB (ZIP)
  ├── word/document.xml      → paragraph / table / image nodes
  ├── word/styles.xml        → named paragraph and character styles
  ├── word/_rels/*.rels      → relationship map (images, hyperlinks)
  └── word/numbering.xml     → list definitions

ODT BLOB (ZIP)
  ├── content.xml            → document body + automatic styles
  └── styles.xml             → named styles + default style
```

## Requirements

| Dependency | Purpose |
|---|---|
| Oracle Database 19c+ | PL/SQL runtime |
| Oracle APEX (`apex_zip`, `apex_json`) | ZIP extraction, JSON output |
| `DBMS_XMLDOM` | DOM-based XML traversal |
| `JSON_OBJECT_T` / `JSON_ARRAY_T` | JSON type API (Oracle 12.2+) |

## Installation

Compile all packages in dependency order from SQL\*Plus:

```sql
@compile_docs_parser.sql
```

Compile order:

```
docx_parser_util.pks  →  docx_parser_util.pkb
docx_parser.pks       →  docx_parser.pkb
odt_parser.pks        →  odt_parser.pkb
```

The compiling user needs `EXECUTE` on `APEX_ZIP`, `APEX_JSON`, and `DBMS_XMLDOM`.

---

## DOCX Parser

### Packages

| File | Role |
|---|---|
| `docx_parser.pks` / `.pkb` | Top-level parser — public API |
| `docx_parser_util.pks` / `.pkb` | Utility layer — ZIP extraction, DOM helpers, unit converters |

### `docx_parser_util` – Public API

```plsql
-- Load a DOCX BLOB from any table row into session state.
-- p_table_name and column names are validated via DBMS_ASSERT.
-- Raises ORA-20002 when no row matches.
procedure load_docx_source(
   p_table_name in varchar2,
   p_blob_col   in varchar2,
   p_id_col     in varchar2,
   p_id_val     in varchar2
);

-- Return the session-cached DOCX BLOB (NULL if none loaded).
function get_loaded_docx return blob;

-- Extract a ZIP entry and return its content as UTF-8 CLOB.
-- Returns NULL when the path does not exist inside the ZIP.
function unpack_docx_clob(
   p_file_path in varchar2,   -- e.g. 'word/document.xml'
   p_docx_blob in blob
) return clob;

-- Extract a ZIP entry and return its raw BLOB content.
-- Use this for binary assets (images, fonts, …).
function unpack_docx_blob(
   p_file_path in varchar2,
   p_docx_blob in blob
) return blob;
```

### `docx_parser` – Public API

```plsql
-- Load DOCX from APEX_WORKSPACE_STATIC_FILES (keyed by FILE_NAME)
-- and return the full document as a pdfmake JSON array (CLOB).
function parse_docx(p_filename in varchar2) return clob;
```

### Usage

```plsql
declare
   l_blob blob;
   l_json clob;
begin
   -- 1. Load from any table
   docx_parser_util.load_docx_source(
      p_table_name => 'MY_DOCUMENTS',
      p_blob_col   => 'FILE_BLOB',
      p_id_col     => 'DOC_ID',
      p_id_val     => '42'
   );

   -- 2. Parse to pdfmake JSON
   l_json := docx_parser.parse_docx('my_document.docx');
   dbms_output.put_line(l_json);
end;
/
```

### Output shape

```json
[
  {
    "style": "Normal",
    "alignment": "justify",
    "fontSize": 11,
    "text": [
      { "text": "Hello ", "bold": true },
      { "text": "World",  "color": "#FF0000" }
    ]
  },
  { "style": "Normal", "text": "Second paragraph" }
]
```

### DOCX internal paths

| Path | Content |
|---|---|
| `word/document.xml` | Document body (paragraphs, tables, images) |
| `word/styles.xml` | Named paragraph and character styles |
| `word/_rels/document.xml.rels` | Relationships (images, hyperlinks) |
| `word/numbering.xml` | List and numbering definitions |
| `word/media/*` | Embedded images and binary assets |

---

## ODT Parser

### Package

| File | Role |
|---|---|
| `odt_parser.pks` / `.pkb` | Single self-contained package |

### Public API

```plsql
-- Parse an ODT BLOB and return a pdfmake-compatible JSON object (CLOB).
function parse_odt(p_odt_blob in blob) return clob;

-- Low-level entry point: pass pre-extracted XML strings.
-- Useful for unit testing without a real ZIP.
function parse_xml(
   p_content_xml in clob,
   p_styles_xml  in clob default null
) return clob;
```

### Usage

```plsql
declare
   l_blob blob;
   l_json clob;
begin
   -- Load BLOB from the database (any method)
   select file_blob into l_blob from my_documents where doc_id = 42;

   l_json := odt_parser.parse_odt(l_blob);
   dbms_output.put_line(l_json);
end;
/
```

### Output shape

```json
{
  "content": [
    {
      "style": "Heading1",
      "headlineLevel": 1,
      "fontSize": 16,
      "fontWeight": "bold",
      "text": "Introduction"
    },
    {
      "style": "Text_Body",
      "fontSize": 12,
      "alignment": "justify",
      "text": [
        { "text": "Plain text followed by " },
        { "style": "T1", "fontWeight": "bold", "text": "bold span" }
      ]
    },
    { "ul": ["item 1", "item 2"] },
    {
      "table": {
        "widths": ["*", "*"],
        "body": [[{ "text": "Cell A" }, { "text": "Cell B" }]]
      }
    }
  ],
  "styles": {
    "Text_Body": { "fontSize": 12, "alignment": "justify" },
    "Heading1":  { "fontSize": 16, "fontWeight": "bold" }
  },
  "defaultStyle": { "fontSize": 12 }
}
```

Style inheritance is fully resolved by the parser — each content node carries its effective style properties inline. The `"styles"` object contains only styles that are actually referenced in the document.

### ODT internal paths

| Path | Content |
|---|---|
| `content.xml` | Document body + automatic styles |
| `styles.xml` | Named paragraph / character styles + default style |

---

## Project structure

```
docx_parser_util.pks / .pkb   – DOCX utility package (ZIP, DOM helpers, unit converters)
docx_parser.pks / .pkb        – DOCX document parser
odt_parser.pks / .pkb         – ODT document parser (self-contained)
compile_docs_parser.sql        – SQL*Plus compilation script
pdfmake/                       – pdfmake runtime (client-side PDF generation)
tests/
  test_docx_parser.sql         – DOCX paragraph & style tests
  test_docx_parser_rels.sql    – DOCX unpack & relationship tests
  test_styles_json.sql         – DOCX styles smoke test
  odt_parser_examples.sql      – ODT parse examples
  test_doc.docx                – sample DOCX document
```

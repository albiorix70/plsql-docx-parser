# docx-parser

Oracle PL/SQL package suite for parsing Microsoft Word DOCX files stored as database BLOBs. Extracts document structure, styles, and relationships and converts them to JSON for further processing (e.g. PDF generation via pdfmake).

## Overview

DOCX files are ZIP archives containing XML files. This library unpacks those archives in-database, parses the XML using Oracle's native XML and JSON APIs, and produces structured JSON output that downstream packages (e.g. `pdfmake`) can consume directly.

```
DOCX BLOB (ZIP)
  └── word/document.xml     → docx_parser.parse_content_xml()  → JSON content
  └── word/styles.xml       → docx_parser_util (internal)       → style map
  └── word/_rels/...rels    → docx_parser_util (internal)       → relationship map
  └── word/media/*          → docx_parser_util.unpack_docx_blob → BLOB
```

## Requirements

| Dependency | Purpose | Optional |
|---|---|---|
| Oracle Database 19c+ | PL/SQL runtime | No |
| Oracle APEX (`apex_zip`) | ZIP extraction from DOCX | No |
| `DBMS_XMLDOM` | DOM-based XML parsing | No |
| `JSON_OBJECT_T` / `JSON_ARRAY_T` | JSON type API (Oracle 12.2+) | No |
| Oracle Logger | Error logging | Yes |
| utPLSQL | Unit testing | Yes (tests only) |
| Python 3.8+ + `python-docx` | Test DOCX generation | Yes (tests only) |

## Installation

Compile in dependency order using SQL\*Plus from the project root:

```sql
@compile_docs_parser.sql
```

This runs the following in sequence:

```
docx_parser_util.pks  →  docx_parser_util.pkb
docx_parser.pks       →  docx_parser.pkb
```

The compiling user needs `EXECUTE` on `APEX_ZIP` and `DBMS_XMLDOM`.

## Packages

### `docx_parser_util` – Utility layer

Handles DOCX loading, ZIP extraction, XML-to-JSON conversion, style and relationship parsing. Maintains session-level package state for the currently loaded document.

#### Types

```plsql
-- Single relationship entry from a .rels file
type t_rels_item is record (
   id          varchar2(200),   -- Relationship ID (e.g. 'rId1')
   rel_type    varchar2(1000),  -- Full type URI
   target      varchar2(2000),  -- Target path or URL
   target_mode varchar2(50)     -- 'External' or empty
);

-- Associative array indexed by Relationship/@Id
type t_rels_table is table of t_rels_item index by varchar2(200);
```

#### Public API

```plsql
-- Load a DOCX BLOB from a table row into package-internal session state.
-- Raises ORA-20002 when no row matches.
procedure load_docx_source(
   p_table_name in varchar2,   -- table name (validated via dbms_assert)
   p_blob_col   in varchar2,   -- BLOB column name
   p_id_col     in varchar2,   -- identifier column name
   p_id_val     in varchar2    -- identifier value
);

-- Return the session-cached DOCX BLOB (NULL if none loaded).
function get_loaded_docx return blob;

-- Extract a file from a DOCX ZIP and return its content as CLOB.
-- Returns NULL when the path does not exist inside the ZIP.
function unpack_docx_clob(
   p_file_path in varchar2,    -- e.g. 'word/document.xml'
   p_docx_blob in blob
) return clob;

-- Extract a file from a DOCX ZIP and return its raw BLOB content.
-- Use this for binary files (images, fonts, …).
function unpack_docx_blob(
   p_file_path in varchar2,
   p_docx_blob in blob
) return blob;
```

---

### `docx_parser` – Document parser

Parses `word/document.xml` and converts the document body to a JSON structure suitable for PDF rendering.

```plsql
-- Parse the main document XML and return a JSON CLOB.
function parse_content_xml(
   p_content_xml in clob    -- content of word/document.xml
) return clob;
```

## Usage examples

### Load a DOCX from the database and parse it

```plsql
declare
   l_blob blob;
   l_doc  clob;
   l_json clob;
begin
   -- 1. load from any table
   docx_parser_util.load_docx_source(
      p_table_name => 'MY_DOCUMENTS',
      p_blob_col   => 'FILE_BLOB',
      p_id_col     => 'DOC_ID',
      p_id_val     => '42'
   );

   -- 2. extract document body
   l_blob := docx_parser_util.get_loaded_docx();
   l_doc  := docx_parser_util.unpack_docx_clob('word/document.xml', l_blob);

   -- 3. parse to JSON
   l_json := docx_parser.parse_content_xml(l_doc);
   dbms_output.put_line(l_json);
end;
/
```

### Extract an image (binary file)

```plsql
declare
   l_img blob;
begin
   l_img := docx_parser_util.unpack_docx_blob(
      p_file_path => 'word/media/image1.png',
      p_docx_blob => docx_parser_util.get_loaded_docx()
   );
   -- store or process l_img …
end;
/
```

## DOCX internal paths (reference)

| Path | Content |
|---|---|
| `word/document.xml` | Document body (paragraphs, tables, images) |
| `word/styles.xml` | Named paragraph and character styles |
| `word/_rels/document.xml.rels` | Relationships (images, hyperlinks, …) |
| `word/numbering.xml` | List and numbering definitions |
| `word/fontTable.xml` | Embedded font declarations |
| `word/media/*` | Embedded images and other binary assets |

Font sizes in DOCX are stored in **half-points** (`w:sz w:val="24"` = 12 pt). Divide by 2 to obtain the point size.

## Testing

### Run unit tests (utPLSQL)

```sql
exec ut.run('test_docx_parser');
exec ut.run('test_docx_parser_rels');
```

Or run the manual styles test:

```sql
@tests/test_styles_json.sql
```

### Generate and upload test DOCX

```bash
# create tests/test_doc.docx
python scripts/generate_test_docx.py

# upload to the database
python scripts/upload_test_docx.py
```

The upload script inserts the DOCX into `test_docx_store` (id = 1). Tests fall back to embedded inline XML when that table is empty.

## Project structure

```
docx_parser_util.pks / .pkb   – utility package (ZIP, XML, JSON conversion)
docx_parser.pks / .pkb        – main document parser
compile_docs_parser.sql        – SQL*Plus compilation script
tests/
  test_docx_parser.sql         – utPLSQL: paragraph & style tests
  test_docx_parser_rels.sql    – utPLSQL: unpack & rels tests
  test_styles_json.sql         – manual styles smoke test
  test_doc.docx                – generated test document
scripts/
  generate_test_docx.py        – creates test_doc.docx via python-docx
  upload_test_docx.py          – uploads DOCX BLOB to Oracle
  requirements.txt             – Python dependencies
sample_data/                   – reference XML snippets from real DOCX files
```

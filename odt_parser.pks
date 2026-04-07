CREATE OR REPLACE PACKAGE odt_parser AS
  /*
   * odt_parser
   * ==========
   * Parses ODT files (stored as BLOBs) into a pdfmake-compatible docDefinition JSON.
   *
   * An ODT file is a ZIP archive containing:
   *   content.xml  – document body (paragraphs, headings, tables, lists)
   *   styles.xml   – named paragraph/character styles
   *
   * Uses APEX_ZIP for ZIP extraction and APEX_JSON for JSON generation.
   *
   * Output format (pdfmake docDefinition):
   *   {
   *     "content":      [ ...nodes... ],
   *     "styles":       { "StyleName": { ... }, ... },
   *     "defaultStyle": { ... }
   *   }
   *
   * Content node shapes
   * -------------------
   *   Paragraph  →  {"text":"Hello","style":"Text_Body"}
   *   Heading    →  {"text":"Title","style":"Heading1","headlineLevel":1}
   *   Mixed runs →  {"text":[{"text":"bold","style":"Bold"},{"text":" normal"}],
   *                  "style":"Text_Body"}
   *   Table      →  {"table":{"widths":["*","*"],"body":[[{cell},…],…]}}
   *   List ul    →  {"ul":["item 1","item 2"]}
   *   List ol    →  {"ol":["item 1","item 2"]}
   *
   * Main entry points
   * -----------------
   *   parse_odt  – unpacks the ODT ZIP and returns the full docDefinition
   *   parse_xml  – low-level: accepts already-extracted XML CLOBs
   */

  c_version CONSTANT VARCHAR2(10) := '2.0.0';

  -- -------------------------------------------------------------------------
  -- Public API
  -- -------------------------------------------------------------------------

  /**
   * Full parse of an ODT BLOB.
   * Returns a complete pdfmake docDefinition JSON as CLOB.
   *
   * @param p_odt_blob  The raw ODT file content as a BLOB
   * @return            pdfmake docDefinition: {"content":[…],"styles":{…},"defaultStyle":{…}}
   */
  FUNCTION parse_odt (p_odt_blob IN BLOB) RETURN CLOB;

  /**
   * Low-level: parse already-extracted content.xml and styles.xml CLOBs.
   * Useful when the caller has already unzipped the ODT, or for unit testing.
   *
   * @param p_content_xml  content.xml as CLOB
   * @param p_styles_xml   styles.xml as CLOB (optional; omit for no styles)
   * @return               pdfmake docDefinition CLOB
   */
  FUNCTION parse_xml (
    p_content_xml IN CLOB,
    p_styles_xml  IN CLOB DEFAULT NULL
  ) RETURN CLOB;

END odt_parser;
/

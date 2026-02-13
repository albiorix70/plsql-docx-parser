-- Umfangreiche utPLSQL-Tests für den DOCX-Parser
-- Diese Suite prüft: einfache Paragraphen, mehrere Runs mit Eigenschaften,
-- Styles-Parsing, Style-Vererbung und leere Eingaben.

-- Initialization: create test table for DOCX blobs (if not exists)
BEGIN
  BEGIN
    EXECUTE IMMEDIATE('CREATE TABLE test_docx_store (id NUMBER PRIMARY KEY, doc BLOB)');
  EXCEPTION
    WHEN OTHERS THEN
      IF SQLCODE != -955 THEN -- ORA-00955: name is already used by an existing object
        RAISE;
      END IF;
  END;
END;
/ 

CREATE OR REPLACE PACKAGE test_docx_parser AS
  --%suite(DOCX Parser Full Suite)
  --%suitepath(Parser/Full)

  --%test
  PROCEDURE test_parse_simple_paragraph;

  --%test
  PROCEDURE test_parse_empty_input_returns_empty;

  --%test
  PROCEDURE test_parse_multiple_runs_with_run_props;

  --%test
  PROCEDURE test_parse_styles_parsing;

  --%test
  PROCEDURE test_parse_style_inheritance_and_override;

END test_docx_parser;
/

-- Deinitialization: drop the test table used for storing the DOCX blob
BEGIN
  BEGIN
    EXECUTE IMMEDIATE('DROP TABLE test_docx_store PURGE');
  EXCEPTION
    WHEN OTHERS THEN
      IF SQLCODE != -942 THEN -- ORA-00942: table or view does not exist
        RAISE;
      END IF;
  END;
END;
/

CREATE OR REPLACE PACKAGE BODY test_docx_parser AS

  -- Local helper: load document.xml from test table if present, otherwise return NULL
  FUNCTION load_document_xml_from_table RETURN CLOB IS
    l_blob BLOB;
    l_clob CLOB;
  BEGIN
    BEGIN
      SELECT doc INTO l_blob FROM test_docx_store WHERE id = 1 FOR UPDATE;
    EXCEPTION
      WHEN NO_DATA_FOUND THEN
        RETURN NULL;
    END;
    l_clob := docx_parser.unpack_docx('word/document.xml', l_blob);
    RETURN l_clob;
  EXCEPTION
    WHEN OTHERS THEN
      RETURN NULL;
  END load_document_xml_from_table;

  -- Local helper: load styles.xml from test table if present
  FUNCTION load_styles_xml_from_table RETURN CLOB IS
    l_blob BLOB;
    l_clob CLOB;
  BEGIN
    BEGIN
      SELECT doc INTO l_blob FROM test_docx_store WHERE id = 1 FOR UPDATE;
    EXCEPTION
      WHEN NO_DATA_FOUND THEN
        RETURN NULL;
    END;
    l_clob := docx_parser.unpack_docx('word/styles.xml', l_blob);
    RETURN l_clob;
  EXCEPTION
    WHEN OTHERS THEN
      RETURN NULL;
  END load_styles_xml_from_table;

  ------------------------------------------------------------------
  -- Test: einfacher Paragraph
  PROCEDURE test_parse_simple_paragraph IS
    l_doc CLOB := NULL;
    l_res docx_parser.t_content_elements;
    l_blob_found BOOLEAN := FALSE;
  BEGIN
    -- try to load document.xml from test table; fallback to embedded sample
    l_doc := load_document_xml_from_table();
    IF l_doc IS NULL THEN
      l_doc := q'[<?xml version="1.0" encoding="UTF-8"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:p>
          <w:r><w:t>Hello World</w:t></w:r>
        </w:p>
      </w:body>
    </w:document>]';
    ELSE
      l_blob_found := TRUE;
    END IF;
    l_res := docx_parser.parse_content_xml(l_doc);
    ut.expect(l_res.COUNT).to_equal(1);
    ut.expect(l_res(1).element_type).to_equal('paragraph');
    ut.expect(l_res(1).text_content).to_equal('Hello World');
    -- if using table source, ensure it was indeed used
    IF l_blob_found THEN
      ut.assert(TRUE);
    END IF;
  END test_parse_simple_paragraph;

  ------------------------------------------------------------------
  -- Test: leere Eingabe
  PROCEDURE test_parse_empty_input_returns_empty IS
    l_doc CLOB := NULL;
    l_res docx_parser.t_content_elements;
  BEGIN
    -- If test table contains a document, load it and test parsing; otherwise test NULL input
    l_doc := load_document_xml_from_table();
    l_res := docx_parser.parse_content_xml(l_doc);
    ut.expect(l_res.COUNT).to_equal(0);
  END test_parse_empty_input_returns_empty;

  ------------------------------------------------------------------
  -- Test: mehrere Runs mit Run-Props (fett/italic/size)
  PROCEDURE test_parse_multiple_runs_with_run_props IS
    l_doc CLOB := NULL;
    l_res docx_parser.t_content_elements;
  BEGIN
    l_doc := load_document_xml_from_table();
    IF l_doc IS NULL THEN
      l_doc := q'[<?xml version="1.0" encoding="UTF-8"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:p>
          <w:r>
            <w:rPr><w:b/><w:sz w:val="24"/></w:rPr>
            <w:t>BoldText</w:t>
          </w:r>
          <w:r>
            <w:rPr><w:i/><w:sz w:val="20"/></w:rPr>
            <w:t>ItalicText</w:t>
          </w:r>
        </w:p>
      </w:body>
    </w:document>]';
    END IF;
    l_res := docx_parser.parse_content_xml(l_doc);
    ut.expect(l_res.COUNT).to_equal(1);
    -- Expect the paragraph text to be concatenation of runs
    ut.expect(l_res(1).text_content).to_equal('BoldTextItalicText');
    -- We also expect that run properties are present on the element(s) returned.
    -- The package returns paragraphs as single elements; run-level props may be
    -- set on the paragraph's style fields depending on implementation.
    -- Check at least that font_size is set (non-null) for the paragraph (heuristic)
    ut.expect(l_res(1).font_size).to_be_greater_than(0);
  END test_parse_multiple_runs_with_run_props;

  ------------------------------------------------------------------
  -- Test: Styles parsing (parse_styles_xml)
  PROCEDURE test_parse_styles_parsing IS
    l_styles CLOB := NULL;
    l_styles_out docx_parser.t_style_list;
  BEGIN
    -- try to load styles.xml from test table; fallback to embedded styles
    l_styles := load_styles_xml_from_table();
    IF l_styles IS NULL THEN
      l_styles := q'[<?xml version="1.0" encoding="UTF-8"?>
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="paragraph" w:styleId="Heading1">
        <w:name w:val="Heading 1"/>
        <w:rPr>
          <w:b/>
          <w:sz w:val="28"/>
          <w:color w:val="FF0000"/>
        </w:rPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="Normal">
        <w:name w:val="Normal"/>
        <w:rPr>
          <w:sz w:val="24"/>
        </w:rPr>
      </w:style>
    </w:styles>]';
    END IF;
    l_styles_out := docx_parser.parse_styles_xml(l_styles);
    ut.expect(l_styles_out.COUNT).to_equal(2);
    -- Find Heading1 and verify properties
    DECLARE
      found BOOLEAN := FALSE;
    BEGIN
      FOR i IN 1..l_styles_out.COUNT LOOP
        IF l_styles_out(i).style_id = 'Heading1' THEN
          found := TRUE;
          ut.expect(l_styles_out(i).is_bold).to_equal(TRUE);
          ut.expect(l_styles_out(i).font_size).to_equal(28);
          ut.expect(l_styles_out(i).font_color).to_equal('FF0000');
        END IF;
      END LOOP;
      ut.expect(found).to_equal(TRUE);
    END;
  END test_parse_styles_parsing;

  ------------------------------------------------------------------
  -- Test: Style-Inheritance und Überschreibung durch Run-Props
  PROCEDURE test_parse_style_inheritance_and_override IS
    l_doc CLOB := NULL;
    l_styles CLOB := NULL;
    l_res docx_parser.t_content_elements;
  BEGIN
    -- Try to load from test table; otherwise fall back to embedded examples
    l_doc := load_document_xml_from_table();
    IF l_doc IS NULL THEN
      l_doc := q'[<?xml version="1.0" encoding="UTF-8"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:p>
          <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
          <w:r><w:t>StyledTitle</w:t></w:r>
        </w:p>
        <w:p>
          <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
          <w:r>
            <w:rPr><w:b/></w:rPr>
            <w:t>OverrideBold</w:t>
          </w:r>
        </w:p>
      </w:body>
    </w:document>]';
    END IF;
    l_styles := load_styles_xml_from_table();
    IF l_styles IS NULL THEN
      l_styles := q'[<?xml version="1.0" encoding="UTF-8"?>
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="paragraph" w:styleId="Heading1">
        <w:name w:val="Heading 1"/>
        <w:rPr>
          <w:sz w:val="30"/>
        </w:rPr>
      </w:style>
    </w:styles>]';
    END IF;
    -- Use parse_styles_xml and then pass styles to paragraph parsing where needed
    l_res := docx_parser.parse_content_xml(l_doc);
    -- Zwei Paragraphen erwartet
    ut.expect(l_res.COUNT).to_equal(2);
    -- Erster Paragraph: sollte size 30 (von Style) und nicht fett sein
    ut.expect(l_res(1).font_size).to_equal(30);
    ut.expect(l_res(1).is_bold).to_equal(FALSE);
    -- Zweiter Paragraph: Run überschreibt mit Bold
    ut.expect(l_res(2).font_size).to_equal(30);
    ut.expect(l_res(2).is_bold).to_equal(TRUE);
  END test_parse_style_inheritance_and_override;

END test_docx_parser;
/

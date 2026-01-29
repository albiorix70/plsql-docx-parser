set serveroutput on
set feedback on

declare
  l_styles_xml clob := q'[<?xml version="1.0" encoding="UTF-8"?>
  <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:style w:type="paragraph" w:styleId="Heading1">
      <w:name w:val="heading 1"/>
      <w:rPr>
        <w:b/>
        <w:sz w:val="36"/>
      </w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="Normal">
      <w:name w:val="normal"/>
      <w:rPr>
        <w:sz w:val="24"/>
      </w:rPr>
    </w:style>
  </w:styles>]';
  l_json clob;
begin
  -- parse and persist styles in the util package
  docx_parser_util.parse_styles_xml(l_styles_xml);

  -- mark some styles as used (simulates parser activity)
  docx_parser_util.mark_style_used('Heading1');
  docx_parser_util.mark_style_used('Normal');

  l_json := docx_parser_util.styles_to_json;
  dbms_output.put_line('Styles JSON:');
  dbms_output.put_line(l_json);
end;
/


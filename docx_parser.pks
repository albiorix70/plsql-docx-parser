create or replace package docx_parser as

   -- Function to parse DOCX content.xml
   function parse_content_xml (
      p_content_xml in clob
   ) return clob;


end docx_parser;
/

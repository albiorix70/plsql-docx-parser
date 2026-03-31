create or replace package docx_parser as

   -- Function to parse DOCX content.xml
   function parse_docx (
      p_filename in varchar2
   ) return clob;
end docx_parser;
/
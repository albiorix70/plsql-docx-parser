create or replace package docx_parser as
   -- Record type to hold parsed content elements
   type t_content_element is record (
         element_type varchar2(50),
         text_content clob,
         style_name   varchar2(100),
         -- styling fields mapped to pdfmake properties
         font_name    varchar2(100),
         font_size    number,
         line_height  number,
         is_bold      boolean,
         is_italic    boolean,
         justify      varchar2(20),
         character_spacing number,
         font_color   varchar2(50),
         bgcolor      varchar2(50),
         decoration   varchar2(50),
         decoration_style varchar2(50),
         decoration_color varchar2(50)
   );
   
   -- Table type for multiple elements
   type t_content_elements is
      table of t_content_element;
   
   -- Function to parse DOCX content.xml
   function parse_content_xml (
      p_content_xml in clob
   ) return t_content_elements;
   
      -- Parse a single paragraph node (`w:p`). Extracts `w:pPr` as CLOB and concatenates runs.
   function parse_paragraph_node (
      p_node in xmltype
   ) return t_content_element;


end docx_parser;
/

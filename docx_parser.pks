create or replace package docx_parser as
   -- Record type to hold parsed content elements
   type t_content_element is record (
         element_type varchar2(50),
         text_content clob,
         style_name   varchar2(100),
         font_size    number,
         is_bold      boolean,
         is_italic    boolean,
         is_underline boolean,
         justify      varchar2(20), 
         font_color   varchar2(50),
         font_name    varchar2(100)
   );
   
   -- Table type for multiple elements
   type t_content_elements is
      table of t_content_element;
   
   -- Record type to hold style information
   type t_style_info is record (
         style_id     varchar2(100),
         style_name   varchar2(100),
         font_size    number,
         is_bold      boolean,
         is_italic    boolean,
         is_underline boolean,
         font_color   varchar2(50),
         font_name    varchar2(100)
   );
   
   -- Table type for multiple styles
   type t_style_list is
      table of t_style_info;
   
   -- Function to parse DOCX content.xml
   function parse_content_xml (
      p_content_xml in clob
   ) return t_content_elements;
   
   -- Function to parse content.xml with styles.xml applied
   function parse_content_xml_with_styles (
      p_content_xml in clob,
      p_styles_xml  in clob
   ) return t_content_elements;
   
   -- Function to parse DOCX styles.xml
   function parse_styles_xml (
      p_styles_xml in clob
   ) return t_style_list;

   -- Unpack a DOCX stored in view APEX_WORKFLOW_FILES using APEX_ZIP
   -- p_id_col: name of identifier column in the view (e.g. 'id' or 'file_name')
   -- p_id_val: value to match for the identifier column
   -- p_blob_col: name of the blob column containing the docx (default 'blob_content')
   -- Returns the document.xml and styles.xml as CLOBs when present
   procedure unpack_docx_from_apex (
      p_id_col       in varchar2,
      p_id_val       in varchar2,
      p_blob_col     in varchar2,
      p_document_xml out clob,
      p_styles_xml   out clob
   );
   
   -- Helper functions to extract specific elements
   function extract_text_elements (
      p_content_xml in clob
   ) return t_content_elements;
   function extract_paragraphs (
      p_content_xml in clob
   ) return t_content_elements;
   function extract_tables (
      p_content_xml in clob
   ) return t_content_elements;
   
   -- Function to get formatted text as CLOB
   function get_formatted_text (
      p_content_xml in clob
   ) return clob;

end docx_parser;
/

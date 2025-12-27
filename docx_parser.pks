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

   -- Record type for relationship entries (from .rels files)
   type t_rels_item is record (
         id          varchar2(200),
         rel_type    varchar2(1000),
         target      varchar2(2000),
         target_mode varchar2(50)
   );

   -- PL/SQL associative array (index by relationship Id)
   type t_rels_table is table of t_rels_item index by varchar2(200);

   -- Function to parse document.xml.rels (returns associative array indexed by Relationship/@Id)
   function parse_rels_xml (
      p_rels_xml in clob
   ) return t_rels_table;

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
      p_styles_xml   out clob,
      p_rels_xml     out clob
   );

end docx_parser;
/
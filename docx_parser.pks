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
   
   -- Record type to hold style information
   type t_style_info is record (
         style_id     varchar2(100),
         style_name   varchar2(100),
         -- mapped style properties for pdfmake
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
   
   -- Associative table (index by style id) for styles
   type t_style_list is table of t_style_info index by varchar2(100);
   
   -- Function to parse DOCX content.xml
   function parse_content_xml (
      p_content_xml in clob
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

   -- Parse a single paragraph node (`w:p`). Extracts `w:pPr` as CLOB and concatenates runs.
   function parse_paragraph_node (
      p_node in xmltype,
      p_styles in t_style_list
   ) return t_content_element;

   -- Unpack a file from a DOCX blob (returns CLOB for XML files)
   -- p_file_path: path inside the DOCX zip (e.g. 'word/document.xml')
   -- p_docx_blob: the DOCX content as BLOB
   -- returns: CLOB content when the target is XML
   function unpack_docx (
      p_file_path in varchar2,
      p_docx_blob in blob
   ) return clob;

   -- Overloaded: return raw BLOB content for non-XML targets
   function unpack_docx (
      p_file_path    in varchar2,
      p_docx_blob    in blob,
      p_return_blob  in boolean
   ) return blob;

   -- Load DOCX from a table row into a package-internal BLOB
   procedure load_docx_source (
      p_table_name in varchar2,
      p_blob_col   in varchar2,
      p_id_col     in varchar2,
      p_id_val     in varchar2
   );

   -- Return the previously loaded DOCX (as BLOB)
   function get_loaded_docx return blob;

end docx_parser;
/
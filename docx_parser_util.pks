create or replace package docx_parser_util as
   -- cached Namespaces object for the main OOXML namespace, to avoid reconstructing it on every attribute lookup.  
   c_ooxml_ns_w         constant varchar2(60) default 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

   /**
    * Record type for relationship entries parsed from `.rels` files.
    */
   type t_rels_item is record (
         id          varchar2(200),
         rel_type    varchar2(1000),
         target      varchar2(2000),
         target_mode varchar2(50)
   );

   /**
    * PL/SQL associative array (index by relationship id).
    */
   type t_rels_table is
      table of t_rels_item index by varchar2(200);

   
    /**
    * Unpack a file from a DOCX BLOB and return its content as CLOB for XML targets.
    * @param p_file_path path inside the DOCX zip (e.g. 'word/document.xml')
    * @param p_docx_blob DOCX file content as BLOB
    * @return CLOB content of the requested file when it is an XML file, otherwise NULL
    */
   function unpack_docx_clob (
      p_file_path in varchar2,
      p_docx_blob in blob
   ) return clob;

   /**
    * Overloaded: Unpack a file from a DOCX BLOB and return raw BLOB content.
    * @param p_file_path path inside the DOCX zip
    * @param p_docx_blob DOCX file content as BLOB
    * @param p_return_blob flag indicating caller expects a BLOB result
    * @return BLOB content of the requested file, or NULL if not found
    */
   function unpack_docx_blob (
      p_file_path in varchar2,
      p_docx_blob in blob
   ) return blob;

   /**
    * Load a DOCX BLOB from a database table row into package-internal storage.
    * @param p_table_name name of the table containing the DOCX BLOB
    * @param p_blob_col column name that holds the BLOB
    * @param p_id_col column name used as identifier in the WHERE clause
    * @param p_id_val value of the identifier to select the row
    */
   procedure load_docx_source (
      p_table_name in varchar2,
      p_blob_col   in varchar2,
      p_id_col     in varchar2,
      p_id_val     in varchar2
   );

   /**
    * Return the previously loaded DOCX stored in package-internal memory.
    * @return BLOB containing the loaded DOCX, or NULL if none loaded
    */
   function get_loaded_docx return blob;

   /**
    * Converts DXA (twips, twentieths of a point) to points.
    * Used for page layout dimensions and paragraph spacing (w:pgSz, w:pgMar, w:spacing).
    *
    * @param p_dxa  Value in DXA.
    * @return       Equivalent value in points.
    */
   function dxa_to_pt (
      p_dxa in number
   ) return number;

   /**
    * Converts half-points to points.
    * Used for font size attributes (w:sz, w:szCs).
    *
    * @param p_val  Value in half-points.
    * @return       Equivalent value in points.
    */
   function halfpt_to_pt (
      p_val in number
   ) return number;

   /**
    * Converts hundredths of a point to points.
    * Used for character spacing (w:spacing/@w:val in run properties).
    *
    * @param p_val  Value in hundredths of a point.
    * @return       Equivalent value in points.
    */
   function hundredthpt_to_pt (
      p_val in number
   ) return number;

   /**
    * Converts eighth-points to points.
    * Used for border widths (w:sz on w:bdr elements).
    *
    * @param p_val  Value in eighth-points.
    * @return       Equivalent value in points.
    */
   function eighthpt_to_pt (
      p_val in number
   ) return number;

   /**
    * Converts EMU (English Metric Units) to points.
    * Used for DrawingML image and shape dimensions (wp:extent, a:xfrm).
    * 914400 EMU = 1 inch = 72 pt, therefore 1 pt = 12700 EMU.
    *
    * @param p_emu  Value in EMU.
    * @return       Equivalent value in points.
    */
   function emu_to_pt (
      p_emu in number
   ) return number;

   /**
    * Converts line units to points.
    * Used for paragraph line spacing (w:spacing/@w:line).
    * 240 line units = 12 pt (single-spacing baseline).
    *
    * @param p_val  Value in line units.
    * @return       Equivalent value in points.
    */
   function lineunit_to_pt (
      p_val in number
   ) return number;

   /**
    * Converts a relative table width (fiftieths of a percent) to points.
    * A value of 5000 represents 100 % of the reference page width.
    * Used for table widths with w:type="pct" (w:tblW, w:tcW).
    *
    * @param p_val        Table width in fiftieths of a percent (0-5000).
    * @param p_page_w_dxa Reference page width in DXA (from sectPr/pgSz/@w:w).
    * @return             Absolute width in points.
    */
   function fiftieth_to_pt (
      p_val        in number,
      p_page_w_dxa in number
   ) return number;

   /**
    * Converts a w:style, w:pPr, or w:rPr DOM node to a pdfmake-compatible
    * JSON object.  Paragraph (pPr) and run (rPr) properties are merged into
    * a single flat object; run properties take precedence when both supply
    * the same key.
    *
    * Supported mappings
    *   w:rPr/w:b               → bold         (boolean)
    *   w:rPr/w:i               → italics       (boolean)
    *   w:rPr/w:u               → decoration    ('underline')
    *   w:rPr/w:strike          → decoration    ('lineThrough')
    *   w:rPr/w:color/@w:val    → color         ('#RRGGBB')
    *   w:rPr/w:sz/@w:val       → fontSize      (pt, half-points ÷ 2)
    *   w:rPr/w:rFonts          → font          (w:ascii or w:hAnsi)
    *   w:rPr/w:spacing/@w:val  → characterSpacing (pt, hundredths ÷ 100)
    *   w:pPr/w:jc/@w:val       → alignment     ('left'|'center'|'right'|'justify')
    *   w:pPr/w:spacing/@w:line → lineHeight    (multiplier; 240 = 1.0×)
    *   w:pPr/w:spacing/@w:before → marginTop   (pt)
    *   w:pPr/w:spacing/@w:after  → marginBottom (pt)
    *   w:pPr/w:ind/@w:left       → marginLeft  (pt)
    *   w:pPr/w:ind/@w:firstLine  → indent      (pt)
    *
    * @param p_style_node  DOM node to inspect (w:style, w:pPr, or w:rPr).
    * @return              pdfmake style object, or an empty object when the
    *                      node is NULL or carries no recognised properties.
    */
   function get_style_attributes (
      p_style_node in dbms_xmldom.domnode,
      p_ns         in varchar2 default 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
   ) return json_object_t;


 function el_get_attribute (
      p_el        in dbms_xmldom.domelement,
      p_attr_name in varchar2 default 'val',
      p_ns        in varchar2 default c_ooxml_ns_w
   ) return varchar2;

   function el_is_toggle_on (
      p_el        in dbms_xmldom.domelement,
      p_attr_name in varchar2 default 'val',
      p_ns        in varchar2 default c_ooxml_ns_w
   ) return boolean;
 
end docx_parser_util;
/
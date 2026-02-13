create or replace package docx_parser_util as
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
    * Parse DOCX `styles.xml` and populate the package-internal styles associative array.
    * @param p_styles_xml CLOB containing the contents of `styles.xml` from a DOCX archive
    */
   procedure parse_styles_xml (
      p_styles_xml in clob
   );
   /**
    * Retrieve a style record by its style id from the parsed styles collection.
    * @param p_style_id style id as found in the DOCX styles (e.g. 'Heading1')
    * @return t_style_info record with style attributes; empty record if not found
    */
   function get_styles_by_id (
      p_style_id in varchar2
   ) return json_object_t;

   /**
    * Parse `document.xml.rels` and return relationships indexed by `Relationship/@Id`.
    * @param p_rels_xml CLOB containing the contents of `document.xml.rels`
    * @return t_rels_table associative array keyed by relationship id
    */
   function parse_rels_xml (
      p_rels_xml in clob
   ) return t_rels_table;

   /**
    * Mark a style id as used by the parser (stored in package-internal map).
    * @param p_style_id style id to mark as used
    */
   procedure mark_style_used (
      p_style_id in varchar2
   );

   /**
    * Create a JSON object describing used styles in pdfmake format.
    * Example output structure:
    * {
    *   "heading1": {"fontSize":18, "bold":true, "margin":[0,12,0,6]},
    *   "normal":   {"fontSize":12, "margin":[0,0,0,12]}
    * }
    * @return CLOB JSON object of styles keyed by style id
    */
   function styles_to_json return clob;

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

end docx_parser_util;
/

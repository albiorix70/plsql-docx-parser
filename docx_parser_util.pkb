create or replace package body docx_parser_util as

   -- package-internal storage for a loaded DOCX BLOB
   g_loaded_docx blob;

   -- Private: extract a raw BLOB from a DOCX ZIP entry.
   -- Returns NULL when the path does not exist or the input is NULL.
   function get_docx_file_blob (
      p_file_path in varchar2,
      p_docx_blob in blob
   ) return blob is
      l_dir apex_zip.t_dir_entries;
   begin
      if p_docx_blob is null then
         return null;
      end if;
      l_dir := apex_zip.get_dir_entries(p_zipped_blob => p_docx_blob);
      if l_dir.exists(p_file_path) then
         return apex_zip.get_file_content(
            p_zipped_blob => p_docx_blob,
            p_dir_entry   => l_dir(p_file_path)
         );
      end if;
      return null;
   exception
      when others then
         begin
            logger.log_error('get_docx_file_blob', sqlerrm);
         exception
            when others then null;
         end;
         return null;
   end get_docx_file_blob;

   /**
    * Unpack a file from a DOCX BLOB and return its content as CLOB for XML targets.
    * @param p_file_path path inside the DOCX zip (e.g. 'word/document.xml')
    * @param p_docx_blob DOCX file content as BLOB
    * @return CLOB content of the requested file when it is an XML file, otherwise NULL
    */
   function unpack_docx_clob (
      p_file_path in varchar2,
      p_docx_blob in blob
   ) return clob is
      l_raw blob;
   begin
      l_raw := get_docx_file_blob(p_file_path, p_docx_blob);
      if l_raw is null then
         return null;
      end if;
      -- Use xmltype to handle the UTF-8 encoding declared in DOCX XML files correctly.
      return xmltype(l_raw, nls_charset_id('AL32UTF8')).getClobVal();
   exception
      when others then
         begin
            logger.log_error('unpack_docx_clob', sqlerrm);
         exception
            when others then null;
         end;
         return null;
   end unpack_docx_clob;

   /**
    * Overloaded: Unpack a file from a DOCX BLOB and return raw BLOB content.
    * @param p_file_path path inside the DOCX zip
    * @param p_docx_blob DOCX file content as BLOB
    * @return BLOB content of the requested file, or NULL if not found
    */
   function unpack_docx_blob (
      p_file_path in varchar2,
      p_docx_blob in blob
   ) return blob is
   begin
      return get_docx_file_blob(p_file_path, p_docx_blob);
   end unpack_docx_blob;

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
   ) is
      l_sql varchar2(2000);
      l_msg varchar2(2000);
   begin
      g_loaded_docx := null;
      l_sql := 'select '
               || dbms_assert.simple_sql_name(p_blob_col)
               || ' from '
               || dbms_assert.simple_sql_name(p_table_name)
               || ' where '
               || dbms_assert.simple_sql_name(p_id_col)
               || ' = :1';
      begin
         execute immediate l_sql
            into g_loaded_docx
            using p_id_val;
      exception
         when no_data_found then
            l_msg := 'No row found for '
                     || p_table_name || '.' || p_id_col || '=' || p_id_val;
            logger.log_error('load_docx_source', l_msg);
            raise_application_error(-20002, l_msg);
         when others then
            logger.log_error('load_docx_source', sqlerrm);
            raise;
      end;
   end load_docx_source;

   /**
    * Return the previously loaded DOCX stored in package-internal memory.
    * @return BLOB containing the loaded DOCX, or NULL if none loaded
    */
   function get_loaded_docx return blob is
   begin
      return g_loaded_docx;
   end get_loaded_docx;

end docx_parser_util;
/

create or replace package body docx_parser_util as

   -- package-internal storage for a loaded DOCX BLOB
   g_loaded_docx  blob;

   -- package-internal storage for parsed styles
   type t_used_styles is
      table of boolean index by varchar2(100);
   lg_used_styles t_used_styles;
   lg_style_json  json_object_t;

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
      l_dir      apex_zip.t_dir_entries;
      l_out_blob blob;
      l_clob     clob;
   begin
      if p_docx_blob is null then
         return null;
      end if;
      l_dir := apex_zip.get_dir_entries(p_zipped_blob => p_docx_blob);
      if l_dir.exists(p_file_path) then
         l_out_blob := apex_zip.get_file_content(
            p_zipped_blob => p_docx_blob,
            p_dir_entry   => l_dir(p_file_path)
         );
         if l_out_blob is not null then
            begin
               l_clob := to_clob(l_out_blob);
               return l_clob;
            exception
               when others then
                  logger.log_error(
                     'unpack_docx_clob',
                     sqlerrm
                  );
                  return null;
            end;
         end if;
      end if;
      return null;
   exception
      when others then
         begin
            logger.log_error(
               'unpack_docx_clob',
               sqlerrm
            );
         exception
            when others then
               null;
         end;
         return null;
   end unpack_docx_clob;

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
   ) return blob is
      l_dir      apex_zip.t_dir_entries;
      l_out_blob blob;
   begin
      if p_docx_blob is null then
         return null;
      end if;
      l_dir := apex_zip.get_dir_entries(p_zipped_blob => p_docx_blob);
      if l_dir.exists(p_file_path) then
         l_out_blob := apex_zip.get_file_content(
            p_zipped_blob => p_docx_blob,
            p_dir_entry   => l_dir(p_file_path)
         );
         return l_out_blob;
      end if;
      return null;
   exception
      when others then
         begin
            logger.log_error(
               'unpack_docx_blob',
               sqlerrm
            );
         exception
            when others then
               null;
         end;
         return null;
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
      l_sql  varchar2(2000);
      l_blob blob;
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
           into l_blob
            using p_id_val;
         g_loaded_docx := l_blob;
      exception
         when no_data_found then
            logger.log_error(
               'load_docx_source',
               'No row found for '
               || p_table_name
               || '.'
               || p_id_col
               || '='
               || p_id_val
            );
            raise_application_error(
               -20002,
               'No row found for '
               || p_table_name
               || '.'
               || p_id_col
               || '='
               || p_id_val
            );
         when others then
            logger.log_error(
               'load_docx_source',
               sqlerrm
            );
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
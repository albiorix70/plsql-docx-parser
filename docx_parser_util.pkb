create or replace package body docx_parser_util as

   -- package-internal storage for a loaded DOCX BLOB
   g_loaded_docx  blob;

   -- package-internal storage for parsed styles
   type t_used_styles is
      table of boolean index by varchar2(100);
   lg_used_styles t_used_styles;
   lg_style_json  json_object_t;

   /**
    * Parse DOCX `styles.xml` and populate the package-internal styles associative array.
    * @param p_styles_xml CLOB containing the contents of `styles.xml` from a DOCX archive
    */
   procedure parse_styles_xml (
      p_styles_xml in clob
   ) is
   begin
      lg_style_json := json_object_t('{}');
      if p_styles_xml is null then
         return;
      end if;

      -- Use XMLTable to extract styles and basic rPr values
      for s_rec in (
         select x.style_id,
                x.style_name,
                x.sz,
                x.bold,
                x.italic,
                x.u_val,
                x.u_color,
                x.color,
                x.p_charspacing,
                x.p_lineheight,
                x.justify,
                x.bg_color,
                x.font_names,
                x.strike
           from xmltable ( xmlnamespaces ( 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' as "w" ),
         '/w:styles/w:style'
               passing xmltype(p_styles_xml)
            columns
               style_id varchar2(100) path '@w:styleId',
               style_name varchar2(200) path 'w:name/@w:val',
               sz varchar2(20) path 'w:rPr/w:sz/@w:val',
               bold varchar2(1) path 'w:rPr/w:b',
               italic varchar2(1) path 'w:rPr/w:i',
               justify varchar2(20) path 'w:pPr/w:jc/@w:val',
               u_val varchar2(20) path 'w:rPr/w:u/@w:val',
               u_color varchar2(20) path 'w:rPr/w:u/@w:color',
               color varchar2(20) path 'w:rPr/w:color/@w:val',
               p_charspacing varchar2(20) path 'w:rPr/w:spacing/@w:val',
               p_lineheight varchar2(20) path 'w:pPr/w:spacing/@w:line',
               bg_color varchar2(20) path 'w:rPr/w:shd/@w:fill',
               font_names varchar2(100) path 'w:rPr/w:rFonts/@w:ascii',
               strike varchar2(1) path 'w:rPr/w:strike'
         ) x
      ) loop
         declare
            l_tmp json_object_t := json_object_t();
         begin
            -- map extracted values to t_style_info record
            l_tmp.put(
               'styleName',
               s_rec.style_name
            );
            l_tmp.put(
               'styleId',
               s_rec.style_id
            );
            -- font size (in half-points in DOCX)
            if s_rec.sz is not null then
               l_tmp.put(
                  'fontSize',
                  to_number(s_rec.sz default null on conversion error) / 2
               );
            end if;
            if s_rec.bold is not null then
               l_tmp.put(
                  'isBold',
                  true
               );
            end if;

            -- italics
            if s_rec.italic is not null then
               l_tmp.put(
                  'isItalic',
                  true
               );
            end if;

            -- paragraph justification
            if s_rec.justify is not null then
               l_tmp.put(
                  'alignment',
                  s_rec.justify
               );
            end if;

            -- character spacing
            if s_rec.p_charspacing is not null then
               l_tmp.put(
                  'characterSpacing',
                  to_number(s_rec.p_charspacing default null on conversion error)
               );
            end if;
            -- paragraph line height (as provided in the style)
            if s_rec.p_lineheight is not null then
               l_tmp.put(
                  'lineHeight',
                  to_number(s_rec.p_lineheight default null on conversion error)
               );
            end if;

            if s_rec.color is not null then
               l_tmp.put(
                  'color',
                  s_rec.color
               );
            end if;

            if s_rec.bg_color is not null then
               l_tmp.put(
                  'background',
                  s_rec.bg_color
               );
            end if;
            --l_style.font_name := s_rec.font_names;
            -- decoration: underline or strike
            if s_rec.strike is not null then
               s_rec.u_val := 'lineThrough';
            end if;
            if
               s_rec.u_val is not null
               and lower(s_rec.u_val) <> 'none'
            then
               l_tmp.put(
                  'decoration',
                  s_rec.u_val
               );
               if s_rec.u_color is not null then
                  l_tmp.put(
                     'decorationColor',
                     s_rec.u_color
                  );
               end if;
            end if;
            -- add style to json object
            lg_style_json.put(
               s_rec.style_name,
               l_tmp
            );
            lg_used_styles(s_rec.style_id) := false; -- initialize usage tracking
         end;
      end loop;

   end parse_styles_xml;

   procedure mark_style_used (
      p_style_id in varchar2
   ) is
   begin
      if p_style_id is not null then
         lg_used_styles(p_style_id) := true;
      end if;
   end mark_style_used;

   function styles_to_json return clob is
   begin
      return lg_style_json.to_string;
   exception
      when others then
         return '{}';
   end styles_to_json;

   /**
    * Retrieve a style record by its style id from the parsed styles collection.
    * @param p_style_id style id as found in the DOCX styles (e.g. 'Heading1')
    * @return t_style_info record with style attributes; empty record if not found
    */
   function get_styles_by_id (
      p_style_id in varchar2
   ) return json_object_t is
   begin
      return lg_style_json.get_object(p_style_id);
   exception
      when others then
         -- return empty record if style id not found
         return json_object_t('{}');
   end get_styles_by_id;

   -- Parse document.xml.rels into an associative PL/SQL table indexed by Relationship/@Id
   function parse_rels_xml (
      p_rels_xml in clob
   ) return t_rels_table is
      l_xml  xmltype;
      l_rels t_rels_table;
      l_item t_rels_item;
   begin
      if p_rels_xml is null then
         return l_rels; -- empty
      end if;
      l_xml := xmltype(p_rels_xml);
      for r in (
         select x.id,
                x.reltype,
                x.target,
                x.targetmode
           from xmltable ( xmlnamespaces ( 'http://schemas.openxmlformats.org/package/2006/relationships' as "r" ),
         '/r:Relationships/r:Relationship'
               passing l_xml
            columns
               id varchar2(200) path '@Id',
               reltype varchar2(1000) path '@Type',
               target varchar2(2000) path '@Target',
               targetmode varchar2(50) path '@TargetMode'
         ) x
      ) loop
         l_item.id := r.id;
         l_item.rel_type := r.reltype;
         l_item.target := r.target;
         l_item.target_mode := r.targetmode;
         -- store by Id
         l_rels(r.id) := l_item;
      end loop;

      return l_rels;
   exception
      when others then
         begin
            logger.log_error(
               'parse_rels_xml',
               sqlerrm
            );
         exception
            when others then
               null;
         end;
         return l_rels;
   end parse_rels_xml;


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

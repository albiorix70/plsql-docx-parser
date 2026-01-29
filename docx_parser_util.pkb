create or replace package body docx_parser_util as

   -- package-internal storage for a loaded DOCX BLOB
   g_loaded_docx  blob;

   -- package-internal storage for parsed styles
   type t_tyle_list is
      table of json_object_t index by varchar2(100);
   lg_style_list  t_style_list;
   type t_used_styles is
      table of boolean index by varchar2(100);
   lg_used_styles t_used_styles;

   /**
    * Parse DOCX `styles.xml` and populate the package-internal styles associative array.
    * @param p_styles_xml CLOB containing the contents of `styles.xml` from a DOCX archive
    */
   procedure parse_styles_xml (
      p_styles_xml in clob
   ) is
      l_styles t_style_list; -- associative table index by style id
      l_xml    xmltype;
   begin
      if p_styles_xml is null then
         return;
      end if;
      l_xml := xmltype(p_styles_xml);

      -- Use XMLTable to extract styles and basic rPr values
      for s_rec in (
         select x.style_id,
                x.style_name,
                x.sz,
                x.b,
                x.i,
                x.uval,
                x.u_color,
                x.color,
                x.char_spacing,
                x.p_line,
                x.pjc,
                x.shd_fill,
                x.rfonts,
                x.strike
           from xmltable ( xmlnamespaces ( 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' as "w" ),
         '/w:styles/w:style'
               passing l_xml
            columns
               style_id varchar2(100) path '@w:styleId',
               style_name varchar2(200) path 'w:name/@w:val',
               sz varchar2(20) path 'w:rPr/w:sz/@w:val',
               b varchar2(1) path 'w:rPr/w:b',
               i varchar2(1) path 'w:rPr/w:i',
               uval varchar2(20) path 'w:rPr/w:u/@w:val',
               u_color varchar2(20) path 'w:rPr/w:u/@w:color',
               color varchar2(20) path 'w:rPr/w:color/@w:val',
               char_spacing varchar2(20) path 'w:rPr/w:spacing/@w:val',
               p_line varchar2(20) path 'w:pPr/w:spacing/@w:line',
               pjc varchar2(20) path 'w:pPr/w:jc/@w:val',
               shd_fill varchar2(20) path 'w:rPr/w:shd/@w:fill',
               rfonts varchar2(100) path 'w:rPr/w:rFonts/@w:ascii',
               strike varchar2(1) path 'w:rPr/w:strike'
         ) x
      ) loop
         declare
            l_style t_style_info;
         begin
            l_style.style_id := s_rec.style_id;
            l_style.style_name := s_rec.style_name;
            if s_rec.sz is not null then
               l_style.font_size := to_number ( s_rec.sz default null on conversion error ) / 2;
            else
               l_style.font_size := null;
            end if;
            l_style.is_bold :=
               case
                  when s_rec.b is not null then
                     true
                  else
                     false
               end;
            l_style.is_italic :=
               case
                  when s_rec.i is not null then
                     true
                  else
                     false
               end;
            -- character spacing
            if s_rec.char_spacing is not null then
               l_style.character_spacing := to_number ( s_rec.char_spacing default null on conversion error );
            else
               l_style.character_spacing := null;
            end if;
            -- paragraph line height (as provided in the style)
            if s_rec.p_line is not null then
               l_style.line_height := to_number ( s_rec.p_line default null on conversion error );
            else
               l_style.line_height := null;
            end if;
            -- paragraph justification
            l_style.justify := s_rec.pjc;
            l_style.font_color := s_rec.color;
            l_style.bgcolor := s_rec.shd_fill;
            l_style.font_name := s_rec.rfonts;
            -- decoration: underline or strike
            if
               s_rec.uval is not null
               and lower(s_rec.uval) <> 'none'
            then
               l_style.decoration := 'underline';
               l_style.decoration_style := s_rec.uval;
               l_style.decoration_color := s_rec.u_color;
            elsif s_rec.strike is not null then
               l_style.decoration := 'lineThrough';
            else
               l_style.decoration := null;
               l_style.decoration_style := null;
               l_style.decoration_color := null;
            end if;
            -- store by style id
            l_styles(s_rec.style_id) := l_style;
         end;
      end loop;

   end parse_styles_xml;

   procedure mark_style_used (
      p_style_id in varchar2
   ) is
   begin
      if p_style_id is null then
         return;
      end if;
      lg_used_styles(p_style_id) := true;
   end mark_style_used;

   function styles_to_json return clob is
      l_idx    varchar2(100);
      l_json   clob := '{';
      l_first  boolean := true;
      l_props  clob;
      l_style  t_style_info;
      l_top    number;
      l_bottom number;
   begin
      l_idx := lg_used_styles.first;
      while l_idx is not null loop
         begin
            l_style := lg_style_list(l_idx);
         exception
            when others then
               l_idx := lg_used_styles.next(l_idx);
               continue;
         end;

         if not l_first then
            l_json := l_json || ',';
         else
            l_first := false;
         end if;

         l_props := '';
         if l_style.font_size is not null then
            l_props := l_props
                       || '"fontSize":'
                       || round(l_style.font_size)
                       || ',';
            l_top := round(l_style.font_size * 0.66);
            l_bottom := round(l_style.font_size * 0.5);
         else
            l_top := 0;
            l_bottom := 12;
         end if;
         if l_style.is_bold then
            l_props := l_props || '"bold":true,';
         end if;
         if l_style.is_italic then
            l_props := l_props || '"italics":true,';
         end if;
         if l_style.justify is not null then
            l_props := l_props
                       || '"alignment":"'
                       || lower(l_style.justify)
                       || '",';
         end if;
         if l_style.font_color is not null then
            l_props := l_props
                       || '"color":"'
                       || l_style.font_color
                       || '",';
         end if;
         if l_style.bgcolor is not null then
            l_props := l_props
                       || '"fillColor":"'
                       || l_style.bgcolor
                       || '",';
         end if;
         if l_style.character_spacing is not null then
            l_props := l_props
                       || '"characterSpacing":'
                       || l_style.character_spacing
                       || ',';
         end if;
         if l_style.decoration is not null then
            l_props := l_props
                       || '"decoration":"'
                       || l_style.decoration
                       || '",';
         end if;

         -- margin
         l_props := l_props
                    || '"margin":['
                    || 0
                    || ','
                    || l_top
                    || ','
                    || 0
                    || ','
                    || l_bottom
                    || '],';

         -- remove trailing comma if present
         if substr(
            l_props,
            -1,
            1
         ) = ',' then
            l_props := substr(
               l_props,
               1,
               length(l_props) - 1
            );
         end if;

         l_json := l_json
                   || '"'
                   || replace(
            l_idx,
            '"',
            '\"'
         )
                   || '":{'
                   || l_props
                   || '}';

         l_idx := lg_used_styles.next(l_idx);
      end loop;

      l_json := l_json || '}';
      return l_json;
   exception
      when others then
         return '{ }';
   end styles_to_json;

   /**
    * Retrieve a style record by its style id from the parsed styles collection.
    * @param p_style_id style id as found in the DOCX styles (e.g. 'Heading1')
    * @return t_style_info record with style attributes; empty record if not found
    */
   function get_styles_by_id (
      p_style_id in varchar2
   ) return t_style_info is
   begin
      return lg_style_list(p_style_id);
   exception
      when others then
         -- return empty record if style id not found
         return t_style_info();
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
   function unpack_docx (
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
                     'unpack_docx',
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
               'unpack_docx',
               sqlerrm
            );
         exception
            when others then
               null;
         end;
         return null;
   end unpack_docx;

   /**
    * Overloaded: Unpack a file from a DOCX BLOB and return raw BLOB content.
    * @param p_file_path path inside the DOCX zip
    * @param p_docx_blob DOCX file content as BLOB
    * @param p_return_blob flag indicating caller expects a BLOB result
    * @return BLOB content of the requested file, or NULL if not found
    */
   function unpack_docx (
      p_file_path   in varchar2,
      p_docx_blob   in blob,
      p_return_blob in boolean
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
               'unpack_docx(blob)',
               sqlerrm
            );
         exception
            when others then
               null;
         end;
         return null;
   end unpack_docx;

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
create or replace package body docx_parser as

   -- package-internal storage for a loaded DOCX BLOB
   g_loaded_docx blob;

   function parse_content_xml (
      p_content_xml in clob
   ) return t_content_elements is
      l_elements t_content_elements := t_content_elements();
      l_xml      xmltype;
      l_element  t_content_element;
   begin
      -- Use Oracle XML support (XMLTYPE + XMLTable) for robust parsing
      if p_content_xml is null then
         return l_elements;
      end if;
      l_xml := xmltype(p_content_xml);

        -- Iterate over child nodes of body
      for p_rec in (
         select x.p_node,
                x.x
           from xmltable ( xmlnamespaces ( 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' as "w" ),
         '/w:document/w:body/*'
               passing l_xml
            columns
               x varchar2(100) path 'local-name(.)',
               p_node xmltype path '.'
         ) x
      ) loop
      -- Trace output
         dbms_output.put_line('Found element: ' || p_rec.x);
      end loop;


      return l_elements;
   end parse_content_xml;

   function parse_styles_xml (
      p_styles_xml in clob
   ) return t_style_list is
      l_styles t_style_list; -- associative table index by style id
      l_xml    xmltype;
   begin
      if p_styles_xml is null then
         return l_styles; -- empty associative table
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
               begin
                  l_style.font_size := to_number ( s_rec.sz ) / 2;
               exception
                  when others then
                     begin
                        begin
                           logger.log_error(
                              'parse_styles_xml',
                              sqlerrm
                           );
                        exception
                           when others then
                              null;
                        end;
                        l_style.font_size := null;
                     end;
               end;
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
               begin
                  l_style.character_spacing := to_number(s_rec.char_spacing default null on conversion error);
               exception when others then l_style.character_spacing := null; end;
            else
               l_style.character_spacing := null;
            end if;
            -- paragraph line height (as provided in the style)
            if s_rec.p_line is not null then
               begin
                  l_style.line_height := to_number(s_rec.p_line default null on conversion error);
               exception when others then l_style.line_height := null; end;
            else
               l_style.line_height := null;
            end if;
            -- paragraph justification
            l_style.justify := s_rec.pjc;
            l_style.font_color := s_rec.color;
            l_style.bgcolor := s_rec.shd_fill;
            l_style.font_name := s_rec.rfonts;
            -- decoration: underline or strike
            if s_rec.uval is not null and lower(s_rec.uval) <> 'none' then
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

      return l_styles;
   end parse_styles_xml;

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
      -- parse a paragraph node: extract pPr and aggregate runs
   function parse_paragraph_node (
      p_node in xmltype,
      p_styles in t_style_list
   ) return t_content_element is
      l_element t_content_element;
      l_text    clob := '';
      l_first   boolean := true;
   begin
      if p_node is null then
         l_element.element_type := 'PARAGRAPH';
         return l_element;
      end if;

         -- pPr extraction removed (not stored on type)

         -- iterate runs and concatenate text; capture some run-level props from first run
      for r in (
         select x.t_text,
                x.sz,
                x.bold,
                x.italic,
                x.uval,
                x.u_color,
                x.color,
                x.char_spacing,
                x.shd_fill,
                x.strike,
                x.rfonts
           from xmltable ( xmlnamespaces ( 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' as "w" ),
         '/w:p/w:r'
               passing p_node
            columns
               t_text clob path 'w:t',
               sz varchar2(20) path 'w:rPr/w:sz/@w:val',
               bold varchar2(1) path 'w:rPr/w:b',
               italic varchar2(1) path 'w:rPr/w:i',
               uval varchar2(20) path 'w:rPr/w:u/@w:val',
               u_color varchar2(20) path 'w:rPr/w:u/@w:color',
               color varchar2(20) path 'w:rPr/w:color/@w:val',
               char_spacing varchar2(20) path 'w:rPr/w:spacing/@w:val',
               shd_fill varchar2(20) path 'w:rPr/w:shd/@w:fill',
               strike varchar2(1) path 'w:rPr/w:strike',
               rfonts varchar2(200) path 'w:rPr/w:rFonts/@w:ascii'
         ) x
      ) loop
         if r.t_text is not null then
            l_text := l_text || r.t_text;
         end if;
         if l_first then
            l_element.font_size :=
               case
                  when r.sz is not null then
                     to_number(r.sz default null on conversion error) / 2
                  else
                     null
               end;
            l_element.is_bold := case when r.bold is not null then true else null end;
            l_element.is_italic := case when r.italic is not null then true else null end;
            if r.char_spacing is not null then
               begin
                  l_element.character_spacing := to_number(r.char_spacing default null on conversion error);
               exception when others then l_element.character_spacing := null; end;
            end if;
            l_element.font_color := r.color;
            l_element.bgcolor := r.shd_fill;
            l_element.font_name := r.rfonts;
            -- decoration from run-level
            if r.uval is not null and lower(r.uval) <> 'none' then
               l_element.decoration := 'underline';
               l_element.decoration_style := r.uval;
               l_element.decoration_color := r.u_color;
            elsif r.strike is not null then
               l_element.decoration := 'lineThrough';
            end if;
            l_first := false;
         end if;
      end loop;

         -- paragraph properties: style, justification
      begin
         for p in (
            select x.pstyle,
                   x.pjc,
                   x.p_line,
                   x.p_shd
              from xmltable ( xmlnamespaces ( 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' as "w" ),
            '/w:p'
                  passing p_node
               columns
                  pstyle varchar2(100) path 'w:pPr/w:pStyle/@w:val',
                  pjc varchar2(20) path 'w:pPr/w:jc/@w:val',
                  p_line varchar2(20) path 'w:pPr/w:spacing/@w:line',
                  p_shd varchar2(20) path 'w:pPr/w:shd/@w:fill'
            ) x
         ) loop
            l_element.style_name := p.pstyle;
            l_element.justify := p.pjc;
            if p.p_line is not null then
               begin
                  l_element.line_height := to_number(p.p_line default null on conversion error);
               exception when others then l_element.line_height := null; end;
            end if;
            if p.p_shd is not null then
               l_element.bgcolor := p.p_shd;
            end if;
         end loop;
      exception
         when others then
            null;
      end;

      l_element.element_type := 'PARAGRAPH';
      l_element.text_content := l_text;
         -- If a style table was provided, compare element fields with the style
         if p_styles is not null and l_element.style_name is not null then
            begin
               if p_styles.exists(l_element.style_name) then
                  declare
                     l_style t_style_info := p_styles(l_element.style_name);
                  begin
                     -- font name
                     if l_style.font_name is not null and l_element.font_name is not null
                        and nvl(l_style.font_name,'') = nvl(l_element.font_name,'') then
                        l_element.font_name := null;
                     end if;
                     -- font size
                     if l_style.font_size is not null and l_element.font_size is not null
                        and l_style.font_size = l_element.font_size then
                        l_element.font_size := null;
                     end if;
                     -- line height
                     if l_style.line_height is not null and l_element.line_height is not null
                        and l_style.line_height = l_element.line_height then
                        l_element.line_height := null;
                     end if;
                     -- bold
                     if l_style.is_bold is not null and l_element.is_bold is not null
                        and l_style.is_bold = l_element.is_bold then
                        l_element.is_bold := null;
                     end if;
                     -- italic
                     if l_style.is_italic is not null and l_element.is_italic is not null
                        and l_style.is_italic = l_element.is_italic then
                        l_element.is_italic := null;
                     end if;
                     -- alignment/justify
                     if l_style.justify is not null and l_element.justify is not null
                        and nvl(l_style.justify,'') = nvl(l_element.justify,'') then
                        l_element.justify := null;
                     end if;
                     -- character spacing
                     if l_style.character_spacing is not null and l_element.character_spacing is not null
                        and l_style.character_spacing = l_element.character_spacing then
                        l_element.character_spacing := null;
                     end if;
                     -- color
                     if l_style.font_color is not null and l_element.font_color is not null
                        and nvl(l_style.font_color,'') = nvl(l_element.font_color,'') then
                        l_element.font_color := null;
                     end if;
                     -- background color
                     if l_style.bgcolor is not null and l_element.bgcolor is not null
                        and nvl(l_style.bgcolor,'') = nvl(l_element.bgcolor,'') then
                        l_element.bgcolor := null;
                     end if;
                     -- decoration
                     if l_style.decoration is not null and l_element.decoration is not null
                        and nvl(l_style.decoration,'') = nvl(l_element.decoration,'') then
                        l_element.decoration := null;
                     end if;
                     -- decoration style
                     if l_style.decoration_style is not null and l_element.decoration_style is not null
                        and nvl(l_style.decoration_style,'') = nvl(l_element.decoration_style,'') then
                        l_element.decoration_style := null;
                     end if;
                     -- decoration color
                     if l_style.decoration_color is not null and l_element.decoration_color is not null
                        and nvl(l_style.decoration_color,'') = nvl(l_element.decoration_color,'') then
                        l_element.decoration_color := null;
                     end if;
                  exception when others then null; end;
               end if;
            exception when others then null; end;
         end if;

      return l_element;
   end parse_paragraph_node;

   -- Unpack a file from a DOCX blob, return CLOB for XML targets
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

   -- Overloaded: return raw BLOB content
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

   -- Load a DOCX BLOB from a table row into package-internal storage
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

   function get_loaded_docx return blob is
   begin
      return g_loaded_docx;
   end get_loaded_docx;

end docx_parser;
/
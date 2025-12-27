create or replace package body docx_parser as

   function extract_paragraph_node (
      p_node in xmltype
   ) return t_content_element;

   function extract_table_node (
      p_node in xmltype
   ) return t_content_element;



   function extract_paragraph_node (
      p_node in xmltype
   ) return t_content_element is
      l_element t_content_element default t_content_element(
         element_type => 'PARAGRAPH',
         text_content => null,
         style_name   => null,
         font_size    => null,
         is_bold      => null,
         is_italic    => null,
         is_underline => null,
         font_color   => null,
         font_name    => null
      );
   begin
      for i in (
         select *
           from xmltable ( xmlnamespaces ( 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' as "w" ),
         '/w:p'
               passing p_node
            columns
               style varchar2(100) path 'w:pPr/w:pStyle/@w:val',
               fontsize varchar2(10) path 'w:rPr/w:sz/@w:val',
               justify varchar2(20) path 'w:pPr/w:jc/@w:val',
               ascii_fontname varchar2(100) path 'w:rPr/w:rFonts/@w:ascii',
               ansi_fontname varchar2(100) path 'w:rPr/w:rFonts/@w:hAnsi',
               utf_fontname varchar2(100) path 'w:rPr/w:rFonts/@w:cs',
               text_color varchar2(20) path 'w:rPr/w:color/@w:val',
               underline varchar2(20) path 'w:rPr/w:u/@w:val',
               bold varchar2(5) path 'w:rPr/w:b',
               italic varchar2(5) path 'w:rPr/w:i',
               text clob path './w:r/string()'
         )
      ) loop
         l_element.text_content := i.text;
         l_element.style_name := i.style;
         if i.fontsize is not null then
            l_element.font_size := to_number ( i.fontsize default null on conversion error ) / 2;
         end if;
         l_element.font_name := coalesce(
            i.utf_fontname,
            i.ansi_fontname,
            i.ascii_fontname
         );
         l_element.justify := i.justify;
         l_element.font_color := i.text_color;
         if i.bold is not null then
            l_element.is_bold := true;
         end if;
         if i.italic is not null then
            l_element.is_italic := true;
         end if;
         if
            i.underline is not null
            and lower(i.underline) <> 'none'
         then
            l_element.is_underline := true;
         end if;

      end loop;
      return l_element;
   end extract_paragraph_node;

   function extract_table_node (
      p_node in xmltype
   ) return t_content_element is
      l_element t_content_element default t_content_element(
         element_type => 'TABLE',
         text_content => null,
         style_name   => null,
         font_size    => null,
         is_bold      => null,
         is_italic    => null,
         is_underline => null,
         font_color   => null,
         font_name    => null
      );
   begin
      -- Placeholder implementation
      return l_element;
   end extract_table_node;

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
      /*
         case p_rec.x
            when 'p' then
               l_element := extract_paragraph_node(p_rec.p_node);
               l_element.element_type := 'PARAGRAPH';
            when 'tbl' then
               l_element := extract_table_node(p_rec.p_node);
               l_element.element_type := 'TABLE';
            else
            -- skip other elements for now
               continue;
         end case;
         l_elements.extend;
         l_elements(l_elements.count) := l_element;
         */
      end loop;


      return l_elements;
   end parse_content_xml;

   function parse_styles_xml (
      p_styles_xml in clob
   ) return t_style_list is
      l_styles t_style_list := t_style_list();
      l_xml    xmltype;
   begin
      if p_styles_xml is null then
         return l_styles;
      end if;
      l_xml := xmltype(p_styles_xml);

      -- Use XMLTable to extract styles and basic rPr values
      for s_rec in (
         select x.style_id,
                x.style_name,
                x.sz,
                x.b,
                x.i,
                x.u,
                x.color,
                x.rfonts
           from xmltable ( xmlnamespaces ( 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' as "w" ),
         '/w:styles/w:style'
               passing l_xml
            columns
               style_id varchar2(100) path '@w:styleId',
               style_name varchar2(200) path 'w:name/@w:val',
               sz varchar2(20) path 'w:rPr/w:sz/@w:val',
               b varchar2(1) path 'w:rPr/w:b',
               i varchar2(1) path 'w:rPr/w:i',
               u varchar2(20) path 'w:rPr/w:u/@w:val',
               color varchar2(20) path 'w:rPr/w:color/@w:val',
               rfonts varchar2(100) path 'w:rPr/w:rFonts/@w:ascii'
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
            l_style.is_underline :=
               case
                  when s_rec.u is not null
                     and lower(s_rec.u) <> 'none' then
                     true
                  else
                     false
               end;
            l_style.font_color := s_rec.color;
            l_style.font_name := s_rec.rfonts;
            l_styles.extend;
            l_styles(l_styles.count) := l_style;
         end;
      end loop;

      return l_styles;
   end parse_styles_xml;

   function extract_text_elements (
      p_content_xml in clob
   ) return t_content_elements is
      l_all_elements  t_content_elements;
      l_text_elements t_content_elements := t_content_elements();
   begin
      l_all_elements := parse_content_xml(p_content_xml);
      
      -- Filter for text elements only
      for i in 1..l_all_elements.count loop
         if l_all_elements(i).element_type = 'TEXT' then
            l_text_elements.extend;
            l_text_elements(l_text_elements.count) := l_all_elements(i);
         end if;
      end loop;

      return l_text_elements;
   end extract_text_elements;

   -- Parse document.xml.rels into an associative PL/SQL table indexed by Relationship/@Id
   function parse_rels_xml (
      p_rels_xml in clob
   ) return t_rels_table is
      l_xml   xmltype;
      l_rels  t_rels_table;
      l_item  t_rels_item;
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
           from xmltable (
                xmlnamespaces(
                   'http://schemas.openxmlformats.org/package/2006/relationships' as "r"
                ),
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
            logger.log_error('parse_rels_xml', sqlerrm);
         exception
            when others then
               null;
         end;
         return l_rels;
   end parse_rels_xml;

   function extract_paragraphs (
      p_content_xml in clob
   ) return t_content_elements is
      l_all_elements       t_content_elements;
      l_paragraph_elements t_content_elements := t_content_elements();
   begin
      l_all_elements := parse_content_xml(p_content_xml);
      
      -- Filter for paragraph elements only
      for i in 1..l_all_elements.count loop
         if l_all_elements(i).element_type = 'PARAGRAPH' then
            l_paragraph_elements.extend;
            l_paragraph_elements(l_paragraph_elements.count) := l_all_elements(i);
         end if;
      end loop;

      return l_paragraph_elements;
   end extract_paragraphs;

   function extract_tables (
      p_content_xml in clob
   ) return t_content_elements is
      l_elements t_content_elements := t_content_elements();
      -- In a full implementation, this would parse table elements
   begin
      -- Placeholder implementation
      return l_elements;
   end extract_tables;

   function get_formatted_text (
      p_content_xml in clob
   ) return clob is
      l_elements t_content_elements;
      l_result   clob := '';
   begin
      l_elements := extract_text_elements(p_content_xml);
      
      -- Concatenate all text elements
      for i in 1..l_elements.count loop
         l_result := l_result
                     || l_elements(i).text_content
                     ||
            case
               when i < l_elements.count then
                  chr(10)
            end;
      end loop;

      return l_result;
   end get_formatted_text;

   function parse_content_xml_with_styles (
      p_content_xml in clob,
      p_styles_xml  in clob
   ) return t_content_elements is
      l_elements t_content_elements := t_content_elements();
      l_xml      xmltype;
      l_element  t_content_element;
      l_styles   t_style_list := t_style_list();
   begin
      if p_content_xml is null then
         return l_elements;
      end if;
      l_xml := xmltype(p_content_xml);

      -- Parse styles if provided
      if p_styles_xml is not null then
         l_styles := parse_styles_xml(p_styles_xml);
      end if;

      -- Extract paragraph nodes first, then runs within each paragraph (avoids ancestor:: axis)
      for p_rec in (
         select x.pnode,
                x.pstyle
           from xmltable ( xmlnamespaces ( 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' as "w" ),
         '/w:document/w:body/w:p'
               passing l_xml
            columns
               pnode xmltype path '.',
               pstyle varchar2(100) path 'w:pPr/w:pStyle/@w:val'
         ) x
      ) loop
         -- iterate over runs under this paragraph
         for r_rec in (
            select x.t_text,
                   x.bold,
                   x.italic,
                   x.sz,
                   x.u as uval,
                   x.color,
                   x.rstyle,
                   x.rfonts
              from xmltable ( xmlnamespaces ( 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' as "w" ),
            'w:r'
                  passing p_rec.pnode
               columns
                  t_text clob path 'w:t',
                  bold varchar2(1) path 'w:rPr/w:b',
                  italic varchar2(1) path 'w:rPr/w:i',
                  sz varchar2(20) path 'w:rPr/w:sz/@w:val',
                  u varchar2(20) path 'w:rPr/w:u/@w:val',
                  color varchar2(20) path 'w:rPr/w:color/@w:val',
                  rstyle varchar2(100) path 'w:rPr/w:rStyle/@w:val',
                  rfonts varchar2(100) path 'w:rPr/w:rFonts/@w:ascii'
            ) x
         ) loop
            l_element.element_type := 'TEXT';
            l_element.text_content := r_rec.t_text;
            -- initialize
            l_element.style_name := null;
            l_element.font_size := null;
            l_element.is_bold := null;
            l_element.is_italic := null;
            l_element.is_underline := null;
            l_element.font_color := null;
            l_element.font_name := null;

            -- apply run-level properties if present, else paragraph style from outer loop
            if r_rec.rstyle is not null then
               l_element.style_name := r_rec.rstyle;
            elsif p_rec.pstyle is not null then
               l_element.style_name := p_rec.pstyle;
            end if;

            if r_rec.sz is not null then
               begin
                  l_element.font_size := to_number ( r_rec.sz ) / 2;
               exception
                  when others then
                     begin
                        begin
                           logger.log_error(
                              'parse_content_xml_with_styles',
                              sqlerrm
                           );
                        exception
                           when others then
                              null;
                        end;
                        l_element.font_size := null;
                     end;
               end;
            end if;
            l_element.is_bold :=
               case
                  when r_rec.bold is not null then
                     true
                  else
                     null
               end;
            l_element.is_italic :=
               case
                  when r_rec.italic is not null then
                     true
                  else
                     null
               end;
            l_element.is_underline :=
               case
                  when r_rec.uval is not null
                     and lower(r_rec.uval) <> 'none' then
                     true
                  else
                     null
               end;
            l_element.font_color := r_rec.color;
            l_element.font_name := r_rec.rfonts;

            -- If some properties are still null, try to fill from parsed styles
            if l_styles.count > 0 then
               declare
                  l_found boolean := false;
               begin
                  if r_rec.rstyle is not null then
                     for s in 1..l_styles.count loop
                        if l_styles(s).style_id = r_rec.rstyle then
                           if l_element.font_size is null then
                              l_element.font_size := l_styles(s).font_size;
                           end if;
                           if l_element.is_bold is null then
                              l_element.is_bold := l_styles(s).is_bold;
                           end if;
                           if l_element.is_italic is null then
                              l_element.is_italic := l_styles(s).is_italic;
                           end if;
                           if l_element.is_underline is null then
                              l_element.is_underline := l_styles(s).is_underline;
                           end if;
                           if l_element.font_color is null then
                              l_element.font_color := l_styles(s).font_color;
                           end if;
                           if l_element.font_name is null then
                              l_element.font_name := l_styles(s).font_name;
                           end if;
                           l_found := true;
                           exit;
                        end if;
                     end loop;
                  end if;
                  if
                     ( not l_found )
                     and p_rec.pstyle is not null
                  then
                     for s in 1..l_styles.count loop
                        if l_styles(s).style_id = p_rec.pstyle then
                           if l_element.font_size is null then
                              l_element.font_size := l_styles(s).font_size;
                           end if;
                           if l_element.is_bold is null then
                              l_element.is_bold := l_styles(s).is_bold;
                           end if;
                           if l_element.is_italic is null then
                              l_element.is_italic := l_styles(s).is_italic;
                           end if;
                           if l_element.is_underline is null then
                              l_element.is_underline := l_styles(s).is_underline;
                           end if;
                           if l_element.font_color is null then
                              l_element.font_color := l_styles(s).font_color;
                           end if;
                           if l_element.font_name is null then
                              l_element.font_name := l_styles(s).font_name;
                           end if;
                           exit;
                        end if;
                     end loop;
                  end if;
               end;
            end if;

            l_elements.extend;
            l_elements(l_elements.count) := l_element;
         end loop;
      end loop;

      return l_elements;
   end parse_content_xml_with_styles;

   procedure unpack_docx_from_apex (
      p_id_col       in varchar2,
      p_id_val       in varchar2,
      p_blob_col     in varchar2,
      p_document_xml out clob,
      p_styles_xml   out clob,
      p_rels_xml     out clob
   ) is

      l_blob      blob;
      l_out_blob  blob;
      l_sql       varchar2(1000);
      l_found     number := 0;
      l_dir       apex_zip.t_dir_entries;
      l_file_path varchar2(32767);
   begin
      p_document_xml := null;
      p_styles_xml := null;
      p_rels_xml := null;

         -- Build dynamic SQL to select the blob column from the view
      l_sql := 'select '
               || dbms_assert.simple_sql_name(p_blob_col)
               || ' from apex_application_files where '
               || dbms_assert.simple_sql_name(p_id_col)
               || ' = :1';

      begin
         execute immediate l_sql
           into l_blob
            using p_id_val;
         l_found := 1;
      exception
         when no_data_found then
            logger.log_error(
               'unpack_docx_from_apex',
               'No row found in apex_application_files for '
               || p_id_col
               || '='
               || p_id_val
            );
            raise_application_error(
               -20001,
               'No row found in apex_application_files for '
               || p_id_col
               || '='
               || p_id_val
            );
         when others then
            logger.log_error(
               'unpack_docx_from_apex',
               sqlerrm
            );
            raise;
      end;
 
 
         -- Try to extract 'word/document.xml' and 'word/styles.xml' using APEX_ZIP
      l_dir := apex_zip.get_dir_entries(p_zipped_blob => l_blob);
      if l_dir.exists('word/document.xml') then
         logger.log('DOCX parser: extracting word/document.xml');
         l_out_blob := apex_zip.get_file_content(
            p_zipped_blob => l_blob,
            p_dir_entry   => l_dir('word/document.xml')
         );

         if l_out_blob is not null then
            begin
               p_document_xml := to_clob(l_out_blob);
            exception
               when others then
                  logger.log_error(
                     'unpack_docx_from_apex',
                     sqlerrm
                  );
                  p_document_xml := null;
            end;
         end if;
      end if;
      -- Try to extract styles.xml if present
      if l_dir.exists('word/styles.xml') then
         l_out_blob := apex_zip.get_file_content(
            p_zipped_blob => l_blob,
            p_dir_entry   => l_dir('word/styles.xml')
         );

         if l_out_blob is not null then
            begin
               p_styles_xml := to_clob(l_out_blob);
            exception
               when others then
                  logger.log_error(
                     'unpack_docx_from_apex',
                     sqlerrm
                  );
                  p_styles_xml := null;
            end;
         end if;
      end if;

      -- Try to extract relationships for the main document
      if l_dir.exists('word/_rels/document.xml.rels') then
         l_out_blob := apex_zip.get_file_content(
            p_zipped_blob => l_blob,
            p_dir_entry   => l_dir('word/_rels/document.xml.rels')
         );

         if l_out_blob is not null then
            begin
               p_rels_xml := to_clob(l_out_blob);
            exception
               when others then
                  logger.log_error(
                     'unpack_docx_from_apex',
                     sqlerrm
                  );
                  p_rels_xml := null;
            end;
         end if;
      end if;

   end unpack_docx_from_apex;

end docx_parser;
/
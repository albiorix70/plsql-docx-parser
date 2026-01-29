create or replace package body docx_parser as

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

   
         -- parse a paragraph node: extract pPr and aggregate runs
   function parse_paragraph_node (
      p_node in xmltype
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
            l_element.is_bold :=
               case
                  when r.bold is not null then
                     true
                  else
                     null
               end;
            l_element.is_italic :=
               case
                  when r.italic is not null then
                     true
                  else
                     null
               end;
            if r.char_spacing is not null then
               begin
                  l_element.character_spacing := to_number ( r.char_spacing default null on conversion error );
               exception
                  when others then
                     l_element.character_spacing := null;
               end;
            end if;
            l_element.font_color := r.color;
            l_element.bgcolor := r.shd_fill;
            l_element.font_name := r.rfonts;
            -- decoration from run-level
            if
               r.uval is not null
               and lower(r.uval) <> 'none'
            then
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
               -- mark style as used in the util package
            if l_element.style_name is not null then
               begin
                  docx_parser_util.mark_style_used(l_element.style_name);
               exception
                  when others then
                     null;
               end;
            end if;
            if p.p_line is not null then
               l_element.line_height := to_number ( p.p_line default null on conversion error );
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

      return l_element;
   end parse_paragraph_node;


end docx_parser;
/
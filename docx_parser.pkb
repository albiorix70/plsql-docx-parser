create or replace package body docx_parser as

   lg_content_element json_array_t default json_array_t();
   lg_document_json   json_object_t default json_object_t();  
-- forwards for recursive parsing   
   procedure parse_paragraph_node (
      p_node in xmltype
   );
   function ppr_to_json (
      p_ppr_node in xmltype
   ) return json_object_t;
   -- parse the main document.xml content
   function parse_content_xml (
      p_content_xml in clob
   ) return clob is
      l_xml xmltype;
   begin
      -- Use Oracle XML support (XMLTYPE + XMLTable) for robust parsing
      if p_content_xml is null then
         return null;
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
         case p_rec.x
            when 'p' then
               parse_paragraph_node(p_rec.p_node);
            else
               null; -- For now, ignore other elements
         end case;
      end loop;

      return null;
   end parse_content_xml;

   -- parse a paragraph node: extract pPr and aggregate runs
   procedure parse_paragraph_node (
      p_node in xmltype
   ) is
      l_local_json json_object_t := json_object_t();
   begin
      if p_node is null then
         return;
      end if;
      for r in (
         select x.x,
                x.p_node
           from xmltable ( xmlnamespaces ( 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' as "w" ),
         '/w:p/*'
               passing p_node
            columns
               x varchar2(100) path 'local-name(.)',
               p_node xmltype path '.'
         ) x
      ) loop
         case r.x
            when 'pPr' then
               dbms_output.put_line('  Found pPr element');
               l_local_json := ppr_to_json(r.p_node);
            when 'r' then
               dbms_output.put_line('  Found run element');
         end case;
      end loop;
   exception
      when others then
         raise;
   end parse_paragraph_node;

   function ppr_to_json (
      p_ppr_node in xmltype
   ) return json_object_t is
      l_json json_object_t := json_object_t();
   begin
      if p_ppr_node is null then
         return l_json;
      end if;
      for r in (
         select x.x
           from xmltable ( xmlnamespaces ( 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' as "w" ),
         '/w:pPr/*'
               passing p_ppr_node
            columns
               x varchar2(100) path 'local-name(.)',
               p_node xmltype path '.'
         ) x
      ) loop
         dbms_output.put_line('    Found pPr child element: ' || r.x);
      end loop;
      return l_json;
   end ppr_to_json;

end docx_parser;
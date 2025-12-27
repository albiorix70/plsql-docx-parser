declare
   -- Sample DOCX content.xml content (simplified for demonstration)
   l_content_xml    clob;
      
   -- Sample DOCX styles.xml content (simplified for demonstration)
   l_styles_xml     clob;
   l_elements       docx_parser.t_content_elements;
   l_styles         docx_parser.t_style_list;
   l_formatted_text clob;
begin
    -- set security group to APEX workspace 'LINKE'
   apex_util.set_security_group_id(p_security_group_id => apex_util.find_security_group_id(p_workspace => 'LINKE'));
   -- Unpack a sample DOCX file from APEX application files 
   docx_parser.unpack_docx_from_apex(
      'filename',
      'Word2007RTFSpec9.docx',
--      'Telefon.docx',
      'blob_content',
      l_content_xml,
      l_styles_xml
   );
   -- Parse the content.xml
   dbms_output.put_line('Parsing DOCX content.xml...');
   l_elements := docx_parser.parse_content_xml(l_content_xml);
   
   -- Display parsed elements
   dbms_output.put_line('Found '
                        || l_elements.count || ' elements:');
   for i in 1..l_elements.count loop
      dbms_output.put_line('Element '
                           || i
                           || ': '
                           || l_elements(i).element_type
                           || ' - '
                           || dbms_lob.substr(
         l_elements(i).text_content,
         100,
         1
      )
                           || ' | '
                           || 'style='
                           || nvl(
         l_elements(i).style_name,
         'NULL'
      )
                           || ', '
                           || 'size='
                           || nvl(
         to_char(l_elements(i).font_size),
         'NULL'
      )
                           || ', '
                           || 'bold='
                           ||
         case
            when l_elements(i).is_bold is null then
               'NULL'
            when l_elements(i).is_bold then
               'Y'
            else
               'N'
         end
                           || ', '
                           || 'italic='
                           ||
         case
            when l_elements(i).is_italic is null then
               'NULL'
            when l_elements(i).is_italic then
               'Y'
            else
               'N'
         end
                           || ', '
                           || 'underline='
                           ||
         case
            when l_elements(i).is_underline is null then
               'NULL'
            when l_elements(i).is_underline then
               'Y'
            else
               'N'
         end
                           || ', '
                           || 'font='
                           || nvl(
         l_elements(i).font_name,
         'NULL'
      )
                           || ', '
                           || 'color=' || nvl(
         l_elements(i).font_color,
         'NULL'
      ));
   end loop;
   
   -- Parse the styles.xml
   dbms_output.put_line(chr(10) || 'Parsing DOCX styles.xml...');
   l_styles := docx_parser.parse_styles_xml(l_styles_xml);
   
   -- Display parsed styles
   dbms_output.put_line('Found '
                        || l_styles.count || ' styles:');
   for i in 1..l_styles.count loop
      dbms_output.put_line('Style '
                           || i
                           || ': '
                           || l_styles(i).style_id
                           || ' - ' || l_styles(i).style_name);
   end loop;
   
   -- Now parse content.xml with styles applied (run-level props fallback to styles)
   dbms_output.put_line(chr(10) || 'Parsing content.xml with styles applied...');
   l_elements := docx_parser.parse_content_xml_with_styles(
      l_content_xml,
      l_styles_xml
   );
   dbms_output.put_line('Found '
                        || l_elements.count || ' elements after applying styles:');
   for i in 1..l_elements.count loop
      dbms_output.put_line('Element '
                           || i
                           || ': '
                           || l_elements(i).element_type
                           || ' - '
                           || dbms_lob.substr(
         l_elements(i).text_content,
         200,
         1
      )
                           || ' | '
                           || 'style='
                           || nvl(
         l_elements(i).style_name,
         'NULL'
      )
                           || ', '
                           || 'size='
                           || nvl(
         to_char(l_elements(i).font_size),
         'NULL'
      )
                           || ', '
                           || 'bold='
                           ||
         case
            when l_elements(i).is_bold is null then
               'NULL'
            when l_elements(i).is_bold then
               'Y'
            else
               'N'
         end
                           || ', '
                           || 'italic='
                           ||
         case
            when l_elements(i).is_italic is null then
               'NULL'
            when l_elements(i).is_italic then
               'Y'
            else
               'N'
         end
                           || ', '
                           || 'underline='
                           ||
         case
            when l_elements(i).is_underline is null then
               'NULL'
            when l_elements(i).is_underline then
               'Y'
            else
               'N'
         end
                           || ', '
                           || 'font='
                           || nvl(
         l_elements(i).font_name,
         'NULL'
      )
                           || ', '
                           || 'color=' || nvl(
         l_elements(i).font_color,
         'NULL'
      ));
   end loop;

end;
/


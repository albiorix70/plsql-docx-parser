create or replace package body docx_parser as

   -- local variable to store namespaces (prefix -> URI)
   type t_ns is
      table of varchar2(4000) index by varchar2(100);
   l_namespaces t_ns;
   -- internal Exceptions
   c_exception_no_body exception;
   pragma exception_init ( c_exception_no_body,-20001 );
   c_exception_no_childnodes exception;
   pragma exception_init ( c_exception_no_childnodes,-20002 );

   c_lf  constant varchar2(1) default chr(10); -- line break (w:br default)
   c_ff  constant varchar2(1) default chr(12); -- page break (w:br w:type="page")
   c_tab constant varchar2(1) default chr(9);  -- horizontal tab (w:tab)

   procedure set_namespaces (
      p_node in dbms_xmldom.domnode
   ) is
      l_nodeattr dbms_xmldom.domnamednodemap;
      l_nsnode   dbms_xmldom.domnode;
   begin
      l_nodeattr := dbms_xmldom.getattributes(p_node);
      for i in 0..dbms_xmldom.getlength(l_nodeattr) - 1 loop
         l_nsnode := dbms_xmldom.item(
            l_nodeattr,
            i
         );
         if dbms_xmldom.getnodename(l_nsnode) like 'xmlns%' then
            l_namespaces(replace(
               dbms_xmldom.getnodename(l_nsnode),
               'xmlns:',
               ''
            )) := dbms_xmldom.getnodevalue(l_nsnode);
         end if;
      end loop;
   end set_namespaces;

   /**
    * Parses a w:r (run) node and returns a pdfmake-compatible JSON object.
    *
    * Handled children:
    *   w:rPr  → character formatting (bold, italics, color, fontSize, …)
    *   w:t    → text content (xml:space="preserve" is honoured by the XML parser)
    *   w:br   → c_lf (line break) or c_ff (page break) depending on w:type
    *   w:tab  → c_tab (horizontal tab)
    *
    * @param p_run_node  DOM node for the w:r element.
    * @return            pdfmake text-run object, e.g. {"text":"Hello","bold":true}
    */
   function parse_run_node (
      p_run_node in dbms_xmldom.domnode
   ) return json_object_t is
      l_result     json_object_t;
      l_rpr_styles json_object_t;
      l_childs     dbms_xmldom.domnodelist;
      l_child      dbms_xmldom.domnode;
      l_textchild  dbms_xmldom.domnode;
      l_el         dbms_xmldom.domelement;
      l_childname  varchar2(100);
      l_br_type    varchar2(20);
      l_text       varchar2(32767) := '';
   begin
      l_childs := dbms_xmldom.getchildnodes(p_run_node);
      for i in 0..dbms_xmldom.getlength(l_childs) - 1 loop
         l_child := dbms_xmldom.item(l_childs, i);
         if dbms_xmldom.getnodetype(l_child) != dbms_xmldom.element_node then
            continue;
         end if;
         l_el       := dbms_xmldom.makeelement(l_child);
         l_childname := dbms_xmldom.getlocalname(l_el);
         case l_childname
            when 'rPr' then
               l_rpr_styles := docx_parser_util.get_style_attributes(
                  l_child,
                  l_namespaces('w')
               );
            when 't' then
               -- text node is the first child of w:t
               l_textchild := dbms_xmldom.getfirstchild(l_child);
               if not dbms_xmldom.isnull(l_textchild) then
                  l_text := l_text || dbms_xmldom.getnodevalue(l_textchild);
               end if;
            when 'br' then
               -- w:type="page" → page break; default / "textWrapping" → line break
               l_br_type := docx_parser_util.el_get_attribute(l_el, 'type');
               if l_br_type = 'page' then
                  l_text := l_text || c_ff;
               else
                  l_text := l_text || c_lf;
               end if;
            when 'tab' then
               l_text := l_text || c_tab;
            else
               null;
         end case;
      end loop;

      -- Seed the result from rPr styles so character formatting is preserved,
      -- then overwrite / add the text content.
      if l_rpr_styles is not null then
         l_result := l_rpr_styles;
      else
         l_result := json_object_t();
      end if;
      l_result.put('text', l_text);
      return l_result;
   end parse_run_node;

   /**
    * Parses a w:p (paragraph) node and returns a pdfmake-compatible JSON object.
    *
    * Paragraph-level formatting from w:pPr (alignment, lineHeight, margins, style)
    * is placed at the top level; w:r children are collected into a "text" array.
    *
    * @param p_node  DOM node for the w:p element.
    * @return        pdfmake paragraph object,
    *                e.g. {"style":"Heading1","text":[{"text":"Hello","bold":true}]}
    */
   function parse_paragraph (
      p_node in dbms_xmldom.domnode
   ) return json_object_t is
      l_childs     dbms_xmldom.domnodelist;
      l_child      dbms_xmldom.domnode;
      l_childname  varchar2(100);
      l_ppr_styles json_object_t := json_object_t();
      l_runs       json_array_t   := json_array_t();
      l_result     json_object_t;
   begin
      l_childs := dbms_xmldom.getchildnodes(p_node);
      for i in 0..dbms_xmldom.getlength(l_childs) - 1 loop
         l_child := dbms_xmldom.item(l_childs, i);
         if dbms_xmldom.getnodetype(l_child) != dbms_xmldom.element_node then
            continue;
         end if;
         l_childname := dbms_xmldom.getlocalname(dbms_xmldom.makeelement(l_child));
         case l_childname
            when 'pPr' then
               l_ppr_styles := docx_parser_util.get_style_attributes(
                  l_child,
                  l_namespaces('w')
               );
            when 'r' then
               l_runs.append(parse_run_node(l_child));
            else
               null;
         end case;
      end loop;

      -- Paragraph object = pPr styles + "text" array of run objects
      l_result := l_ppr_styles;
      l_result.put('text', l_runs);
      return l_result;
   end parse_paragraph;

   function parse_docx (
      p_filename in varchar2
   ) return clob is
      l_doc       xmltype;
      l_dom       dbms_xmldom.domdocument;
      l_nodelist  dbms_xmldom.domnodelist;
      l_node      dbms_xmldom.domnode;
      l_node_name varchar2(100);
      l_content   json_array_t := json_array_t();
   begin
      docx_parser_util.load_docx_source(
         p_table_name => 'APEX_WORKSPACE_STATIC_FILES',
         p_blob_col   => 'FILE_CONTENT',
         p_id_col     => 'FILE_NAME',
         p_id_val     => p_filename
      );

      l_doc := xmltype(docx_parser_util.unpack_docx_clob(
         p_file_path => 'word/document.xml',
         p_docx_blob => docx_parser_util.get_loaded_docx
      ));

      l_dom  := dbms_xmldom.newdomdocument(l_doc);
      l_node := dbms_xmldom.makenode(dbms_xmldom.getdocumentelement(l_dom));
      set_namespaces(l_node);

      -- Locate w:body
      if dbms_xmldom.haschildnodes(l_node) then
         l_nodelist := dbms_xmldom.getchildnodes(l_node);
         for i in 0..dbms_xmldom.getlength(l_nodelist) - 1 loop
            if dbms_xmldom.getnodename(dbms_xmldom.item(l_nodelist, i)) = 'w:body' then
               l_node := dbms_xmldom.item(l_nodelist, i);
               dbms_xmldom.freenodelist(l_nodelist);
               exit;
            end if;
         end loop;
      else
         raise c_exception_no_body;
      end if;

      -- Iterate w:body children and dispatch by element type
      if dbms_xmldom.haschildnodes(l_node) then
         l_nodelist := dbms_xmldom.getchildnodes(l_node);
         for i in 0..dbms_xmldom.getlength(l_nodelist) - 1 loop
            l_node_name := dbms_xmldom.getnodename(dbms_xmldom.item(l_nodelist, i));
            case l_node_name
               when 'w:p' then
                  l_content.append(
                     parse_paragraph(dbms_xmldom.item(l_nodelist, i))
                  );
               when 'w:tbl' then
                  null; -- table processing not yet implemented
                  EXIT; -- IGNORE: stop after first table for now to avoid empty paragraphs at the end of the document 
               else
                  null;
            end case;
         end loop;
      else
         raise c_exception_no_childnodes;
      end if;

      dbms_xmldom.freedocument(l_dom);
      return l_content.to_clob;
   end parse_docx;

end docx_parser;

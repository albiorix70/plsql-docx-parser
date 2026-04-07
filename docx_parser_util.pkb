create or replace package body docx_parser_util as

   -- Unit conversion factors (OOXML value → points)
   c_dxa_per_pt         constant number default 20;       -- twips (page layout, spacing)
   c_halfpt_per_pt      constant number default 2;        -- half-points (font size w:sz)
   c_hundredthpt_per_pt constant number default 100;      -- hundredths of a point (character spacing)
   c_eighthpt_per_pt    constant number default 8;        -- eighth-points (border widths w:bdr)
   c_emu_per_pt         constant number default 12700;    -- English Metric Units (DrawingML)
   c_lineunit_per_pt    constant number default 240 / 12; -- line units (240 = 12 pt baseline)
   c_fiftieth_per_pct   constant number default 50;       -- fiftieths of percent (table widths)

   -- Cached charset ID for UTF-8; evaluated once at package initialisation.
   c_utf8_csid          constant number default nls_charset_id('AL32UTF8');

   -- Session-level storage for the currently loaded DOCX BLOB.
   g_loaded_docx        blob;

    /**
    * Extracts a single file entry from a DOCX ZIP archive and returns it as a raw BLOB.
    * Acts as the shared primitive for unpack_docx_clob and unpack_docx_blob,
    * centralising all apex_zip interaction and exception handling.
    *
    * @param p_file_path  Path of the entry inside the ZIP (e.g. 'word/document.xml').
    *                     Case-sensitive; use forward slashes.
    * @param p_docx_blob  BLOB containing the DOCX file (the ZIP archive).
    * @return             Raw BLOB of the requested entry, or NULL when p_docx_blob
    *                     is NULL or the entry does not exist in the archive.
    */
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
         return null;
   end get_docx_file_blob;

   /**
    * Unpack a file from a DOCX BLOB and return its content as CLOB for XML targets.
    * The BLOB-to-CLOB conversion explicitly uses UTF-8, matching the encoding
    * declared in all DOCX XML files (<?xml ... encoding="UTF-8"?>).
    *
    * @param p_file_path  Path inside the DOCX ZIP (e.g. 'word/document.xml').
    * @param p_docx_blob  DOCX file content as BLOB.
    * @return             CLOB content of the requested entry, or NULL if not found.
    */
   function unpack_docx_clob (
      p_file_path in varchar2,
      p_docx_blob in blob
   ) return clob is
      l_raw blob;
   begin
      l_raw := get_docx_file_blob(
         p_file_path,
         p_docx_blob
      );
      if l_raw is null then
         return null;
      end if;
      return xmltype(
         l_raw,
         c_utf8_csid
      ).getclobval();
   exception
      when others then
         return null;
   end unpack_docx_clob;

   /**
    * Unpack a file from a DOCX BLOB and return its raw BLOB content.
    * Use this for binary assets such as images or embedded fonts.
    *
    * @param p_file_path  Path inside the DOCX ZIP (e.g. 'word/media/image1.png').
    * @param p_docx_blob  DOCX file content as BLOB.
    * @return             Raw BLOB content of the requested entry, or NULL if not found.
    */
   function unpack_docx_blob (
      p_file_path in varchar2,
      p_docx_blob in blob
   ) return blob is
   begin
      return get_docx_file_blob(
         p_file_path,
         p_docx_blob
      );
   end unpack_docx_blob;

   /**
    * Load a DOCX BLOB from a database table row into package-internal session storage.
    * Raises ORA-20002 when no row matches the given identifier.
    *
    * @param p_table_name  Table name (validated with dbms_assert to prevent SQL injection).
    * @param p_blob_col    Column name that holds the BLOB.
    * @param p_id_col      Column name used as identifier in the WHERE clause.
    * @param p_id_val      Value of the identifier to select the row.
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
                     || p_table_name
                     || '.'
                     || p_id_col
                     || '='
                     || p_id_val;
            raise_application_error(
               -20002,
               l_msg
            );
         when others then
            raise;
      end;
   end load_docx_source;

   /**
    * Return the previously loaded DOCX stored in package-internal memory.
    * @return BLOB containing the loaded DOCX, or NULL if none loaded.
    */
   function get_loaded_docx return blob is
   begin
      return g_loaded_docx;
   end get_loaded_docx;

   -- -------------------------------------------------------------------------
   -- Unit conversion functions (OOXML → points)
   -- -------------------------------------------------------------------------

   function dxa_to_pt (
      p_dxa in number
   ) return number is
   begin
      return p_dxa / c_dxa_per_pt;
   end dxa_to_pt;

   function halfpt_to_pt (
      p_val in number
   ) return number is
   begin
      return p_val / c_halfpt_per_pt;
   end halfpt_to_pt;

   function hundredthpt_to_pt (
      p_val in number
   ) return number is
   begin
      return p_val / c_hundredthpt_per_pt;
   end hundredthpt_to_pt;

   function eighthpt_to_pt (
      p_val in number
   ) return number is
   begin
      return p_val / c_eighthpt_per_pt;
   end eighthpt_to_pt;

   function emu_to_pt (
      p_emu in number
   ) return number is
   begin
      return p_emu / c_emu_per_pt;
   end emu_to_pt;

   function lineunit_to_pt (
      p_val in number
   ) return number is
   begin
      return p_val / c_lineunit_per_pt;
   end lineunit_to_pt;

   function fiftieth_to_pt (
      p_val        in number,
      p_page_w_dxa in number
   ) return number is
   begin
      return ( p_val / ( c_fiftieth_per_pct * 100 ) ) * dxa_to_pt(p_page_w_dxa);
   end fiftieth_to_pt;

   -- -------------------------------------------------------------------------
   -- Extract style attributes from a w:style / w:pPr / w:rPr node into
   -- a pdfmake-compatible JSON object.
   -- -------------------------------------------------------------------------

   function get_style_attributes (
      p_style_node in dbms_xmldom.domnode,
      p_ns         in varchar2 default 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
   ) return json_object_t is
      -- return element
      r          json_object_t default json_object_t();
      l_children dbms_xmldom.domnodelist;
      l_nm       varchar2(200);

         -- Iterates child elements of a w:rPr node and populates r with character properties.
      procedure apply_rpr (
         p_node in dbms_xmldom.domnode
      ) is
         l_kids dbms_xmldom.domnodelist;
         l_kid  dbms_xmldom.domnode;
         l_el   dbms_xmldom.domelement;
         l_nm   varchar2(200);
         l_v    varchar2(1000);
      begin
         l_kids := dbms_xmldom.getchildnodes(p_node);
         for i in 0..dbms_xmldom.getlength(l_kids) - 1 loop
            l_kid := dbms_xmldom.item(
               l_kids,
               i
            );
            if dbms_xmldom.getnodetype(l_kid) != dbms_xmldom.element_node then
               continue;
            end if;
            l_el := dbms_xmldom.makeelement(l_kid);
            l_nm := dbms_xmldom.getlocalname(l_el);
            case l_nm
               when 'b' then
                  if el_is_toggle_on(l_el) then
                     r.put(
                        'bold',
                        true
                     );
                  end if;
               when 'i' then
                  if el_is_toggle_on(l_el) then
                     r.put(
                        'italics',
                        true
                     );
                  end if;
               when 'strike' then
                  if el_is_toggle_on(l_el) then
                     r.put(
                        'decoration',
                        'lineThrough'
                     );
                  end if;
               when 'u' then
                  l_v := el_get_attribute(l_el);
                  if
                     l_v is not null
                     and l_v != 'none'
                  then
                     r.put(
                        'decoration',
                        'underline'
                     );
                  end if;
               when 'color' then
                  l_v := el_get_attribute(l_el);
                  if
                     l_v is not null
                     and l_v != 'auto'
                  then
                     r.put(
                        'color',
                        '#' || l_v
                     );
                  end if;
               when 'sz' then
                  l_v := el_get_attribute(l_el);
                  if l_v is not null then
                     r.put(
                        'fontSize',
                        halfpt_to_pt(to_number(l_v))
                     );
                  end if;
               when 'rFonts' then
                  -- prefer w:ascii; fall back to w:hAnsi
                  l_v := el_get_attribute(
                     l_el,
                     'ascii'
                  );
                  if l_v is null then
                     l_v := el_get_attribute(
                        l_el,
                        'hAnsi'
                     );
                  end if;
                  if l_v is not null then
                  -- for simplicity, we assume the font name maps directly to a pdfmake font;
                  -- in practice this may require some mapping / normalization
                    null;
/*                     r.put(
                        'font',
                        l_v
                     );*/
                  end if;
               when 'spacing' then
                  -- character spacing (w:rPr/w:spacing); not to be confused with paragraph spacing
                  l_v := el_get_attribute(l_el);
                  if l_v is not null then
                     r.put(
                        'characterSpacing',
                        hundredthpt_to_pt(to_number(l_v))
                     );
                  end if;
               else
                  null;
            end case;
         end loop;
      end apply_rpr;

      -- Iterates child elements of a w:pPr node and populates r with paragraph properties.
      procedure apply_ppr (
         p_node in dbms_xmldom.domnode
      ) is
         l_kids dbms_xmldom.domnodelist;
         l_kid  dbms_xmldom.domnode;
         l_el   dbms_xmldom.domelement;
         l_nm   varchar2(200);
         l_v    varchar2(1000);
      begin
         l_kids := dbms_xmldom.getchildnodes(p_node);
         for i in 0..dbms_xmldom.getlength(l_kids) - 1 loop
            l_kid := dbms_xmldom.item(
               l_kids,
               i
            );
            if dbms_xmldom.getnodetype(l_kid) != dbms_xmldom.element_node then
               continue;
            end if;
            l_el := dbms_xmldom.makeelement(l_kid);
            l_nm := dbms_xmldom.getlocalname(l_el);
            case l_nm
               when 'pStyle' then
                  r.put(
                     'style',
                     el_get_attribute(l_el)
                  );
               when 'jc' then
                  l_v := el_get_attribute(l_el);
                  -- OOXML 'both' = justified; all other values pass through as-is
                  if l_v = 'both' then
                     r.put(
                        'alignment',
                        'justify'
                     );
                  elsif l_v is not null then
                     r.put(
                        'alignment',
                        l_v
                     );
                  end if;
               when 'spacing' then
                  -- line height: OOXML 240 line units = 1.0× (single spacing)
                  l_v := el_get_attribute(
                     l_el,
                     'line'
                  );
                  if l_v is not null then
                     r.put(
                        'lineHeight',
                        to_number(l_v) / 240
                     );
                  end if;
                  l_v := el_get_attribute(
                     l_el,
                     'before'
                  );
                  if l_v is not null then
                     r.put(
                        'marginTop',
                        dxa_to_pt(to_number(l_v))
                     );
                  end if;
                  l_v := el_get_attribute(
                     l_el,
                     'after'
                  );
                  if l_v is not null then
                     r.put(
                        'marginBottom',
                        dxa_to_pt(to_number(l_v))
                     );
                  end if;
               when 'ind' then
                  l_v := el_get_attribute(
                     l_el,
                     'left'
                  );
                  if l_v is not null then
                     r.put(
                        'marginLeft',
                        dxa_to_pt(to_number(l_v))
                     );
                  end if;
                  l_v := el_get_attribute(
                     l_el,
                     'firstLine'
                  );
                  if l_v is not null then
                     r.put(
                        'indent',
                        dxa_to_pt(to_number(l_v))
                     );
                  end if;
               when 'rPr' then
                  apply_rpr(l_kid);
               else
                  null;
            end case;
         end loop;
      end apply_ppr;

      -- Dispatches a node to the appropriate processor based on its local name.
      procedure dispatch (
         p_node in dbms_xmldom.domnode
      ) is
         l_nm varchar2(200);
      begin
         if dbms_xmldom.getnodetype(p_node) != dbms_xmldom.element_node then
            return;
         end if;
         l_nm := dbms_xmldom.getlocalname(dbms_xmldom.makeelement(p_node));
         case l_nm
            when 'pPr' then
               apply_ppr(p_node);
            when 'rPr' then
               apply_rpr(p_node);
            else
               null;
         end case;
      end dispatch;

   begin
      if dbms_xmldom.isnull(p_style_node) then
         return r;
      end if;

      -- If the node itself is pPr or rPr, process it directly.
      -- Otherwise iterate children and dispatch each pPr / rPr child.
      l_nm := dbms_xmldom.getlocalname(dbms_xmldom.makeelement(p_style_node));
      if l_nm in ( 'pPr',
                   'rPr' ) then
         dispatch(p_style_node);
      else
         l_children := dbms_xmldom.getchildnodes(p_style_node);
         for i in 0..dbms_xmldom.getlength(l_children) - 1 loop
            dispatch(dbms_xmldom.item(
               l_children,
               i
            ));
         end loop;
      end if;

      return r;
   end get_style_attributes;

   function el_get_attribute (
      p_el        in dbms_xmldom.domelement,
      p_attr_name in varchar2 default 'val',
      p_ns        in varchar2 default c_ooxml_ns_w
   ) return varchar2 is
   begin
      return dbms_xmldom.getattribute(
         elem => p_el,
         ns   => p_ns,
         name => p_attr_name
      );
   end el_get_attribute;

      -- Toggle properties (w:b, w:i, w:strike): present without w:val="false/0/off" means ON.
   function el_is_toggle_on (
      p_el        in dbms_xmldom.domelement,
      p_attr_name in varchar2,
      p_ns        in varchar2 default c_ooxml_ns_w
   ) return boolean is
      l_v varchar2(20) := lower(el_get_attribute(
         p_el,
         p_attr_name,
         p_ns
      ));
   begin
      return l_v not in ( 'false',
                          '0',
                          'off' );
   end el_is_toggle_on;
end docx_parser_util;
/
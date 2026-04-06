create or replace package body odt_parser as

  -- Namespace URI constants
  c_ns_office constant varchar2(100) := 'urn:oasis:names:tc:opendocument:xmlns:office:1.0';
  c_ns_style  constant varchar2(100) := 'urn:oasis:names:tc:opendocument:xmlns:style:1.0';
  c_ns_fo     constant varchar2(100) := 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0';
  c_ns_text   constant varchar2(100) := 'urn:oasis:names:tc:opendocument:xmlns:text:1.0';
  c_ns_table  constant varchar2(100) := 'urn:oasis:names:tc:opendocument:xmlns:table:1.0';

  -- ---------------------------------------------------------------------------
  -- Package-body-level style types and map.
  -- Declared here so write_paragraph can access resolved styles inline,
  -- even though it runs before write_styles outputs the JSON styles object.
  -- ---------------------------------------------------------------------------
  type t_sty is record (
        style_name  varchar2(200),
        family      varchar2(50),
        parent      varchar2(200),   -- null after inheritance is resolved
        text_align  varchar2(30),
        margin_top  varchar2(20),
        margin_bot  varchar2(20),
        margin_lft  varchar2(20),
        line_height varchar2(20),
        font_name   varchar2(100),
        font_size   varchar2(20),
        font_weight varchar2(20),
        font_style  varchar2(20),
        text_deco   varchar2(50),
        color       varchar2(20),
        bg_color    varchar2(20)
  );
  type t_sty_tab is table of t_sty;   -- bulk-collect target
  type t_sty_map is table of t_sty index by varchar2(200);
  -- Populated by load_styles; accessible to write_paragraph and write_styles
  g_style_map t_sty_map;

  -- ===========================================================================
  -- SECTION 1 – PRIVATE UTILITY HELPERS
  -- ===========================================================================

  /**
   * Unzip a named member from an ODT (ZIP) BLOB using APEX_ZIP.
   * Returns NULL when the member does not exist.
   */
   function unzip_member (
      p_zip  in blob,
      p_name in varchar2
   ) return clob is
      l_blob blob;
      l_clob clob;
      l_dest integer := 1;
      l_src  integer := 1;
      l_lang integer := 0;
      l_warn integer;
   begin
      l_blob := apex_zip.get_file_content(
         p_zip,
         p_name
      );
      if l_blob is null
      or dbms_lob.getlength(l_blob) = 0 then
         return null;
      end if;
      dbms_lob.createtemporary(
         l_clob,
         false,
         dbms_lob.session
      );
      dbms_lob.converttoclob(
         dest_lob     => l_clob,
         src_blob     => l_blob,
         amount       => dbms_lob.lobmaxsize,
         dest_offset  => l_dest,
         src_offset   => l_src,
         blob_csid    => nls_charset_id('AL32UTF8'),
         lang_context => l_lang,
         warning      => l_warn
      );
      return l_clob;
   exception
      when others then
         if dbms_lob.istemporary(l_clob) = 1 then
            dbms_lob.freetemporary(l_clob);
         end if;
         return null;
   end unzip_member;

  /**
   * Map an ODF fo:color value to a pdfmake color string.
   * Returns NULL for "transparent" / "none" / unrecognised values.
   */
   function map_color (
      p_odf_color in varchar2
   ) return varchar2 is
   begin
      if p_odf_color is null
      or lower(p_odf_color) in ( 'transparent',
                                 'none',
                                 'auto' ) then
         return null;
      end if;
      if regexp_like(
         p_odf_color,
         '^#[0-9A-Fa-f]{6}$'
      ) then
         return p_odf_color;
      end if;
      return null;
   end map_color;

  /** Map ODF fo:text-align to a pdfmake alignment string. */
   function map_alignment (
      p_odf_align in varchar2
   ) return varchar2 is
   begin
      return
         case lower(p_odf_align)
            when 'start'   then
               'left'
            when 'end'     then
               'right'
            when 'center'  then
               'center'
            when 'justify' then
               'justify'
            else
               null
         end;
   end map_alignment;

  /**
   * Convert an ODF length string (e.g. "12pt", "1.5cm", "10mm") to points.
   * Returns NULL when the input cannot be parsed.
   */
   function odf_length_to_pt (
      p_len in varchar2
   ) return number is
      l_val  number;
      l_unit varchar2(10);
   begin
      if p_len is null then
         return null;
      end if;
      l_val := to_number ( regexp_substr(
         p_len,
         '^([0-9]+\.?[0-9]*)',
         1,
         1,
         'i',
         1
      ),
      '9999990D9999',
      'NLS_NUMERIC_CHARACTERS=''.,''' );
      l_unit := lower(regexp_substr(
         p_len,
         '([a-z]+)$',
         1,
         1,
         'i',
         1
      ));
      return
         case l_unit
            when 'pt' then
               l_val
            when 'cm' then
               round(
                  l_val * 28.3465,
                  2
               )
            when 'mm' then
               round(
                  l_val * 2.83465,
                  2
               )
            when 'in' then
               round(
                  l_val * 72,
                  2
               )
            when 'px' then
               round(
                  l_val * 0.75,
                  2
               )
            else
               null
         end;
   exception
      when others then
         return null;
   end odf_length_to_pt;

  -- ===========================================================================
  -- SECTION 1b – DOM HELPERS
  -- ===========================================================================

  /**
   * Read a namespace-qualified attribute from a DOM element.
   * Returns NULL on any error.
   */
   function dom_get_attr (
      p_elem in dbms_xmldom.DOMElement,
      p_ns   in varchar2,
      p_name in varchar2
   ) return varchar2 is
   begin
      return dbms_xmldom.getAttribute(
         p_elem,
         p_ns,
         p_name
      );
   exception
      when others then
         return null;
   end dom_get_attr;

  /**
   * Recursively collect all text content of a DOM node.
   * Handles text:s (space), text:tab, text:line-break specially;
   * recurses into element children for everything else.
   */
   function dom_get_text (
      p_node in dbms_xmldom.DOMNode
   ) return varchar2 is
      l_result   varchar2(32767) := '';
      l_children dbms_xmldom.DOMNodeList;
      l_child    dbms_xmldom.DOMNode;
      l_ntype    pls_integer;
      l_lname    varchar2(200);
   begin
      l_ntype := dbms_xmldom.getNodeType(p_node);
      if l_ntype = dbms_xmldom.TEXT_NODE then
         return dbms_xmldom.getNodeValue(p_node);
      end if;
      if l_ntype != dbms_xmldom.ELEMENT_NODE then
         return '';
      end if;
      l_lname := dbms_xmldom.getLocalName(
         dbms_xmldom.makeElement(p_node)
      );
      if l_lname = 's'          then return ' ';     end if;
      if l_lname = 'tab'        then return chr(9);  end if;
      if l_lname = 'line-break' then return chr(10); end if;
      l_children := dbms_xmldom.getChildNodes(p_node);
      for i in 0..dbms_xmldom.getLength(l_children) - 1 loop
         l_child  := dbms_xmldom.item(
            l_children,
            i
         );
         l_result := l_result || dom_get_text(l_child);
      end loop;
      return l_result;
   end dom_get_text;

  -- ===========================================================================
  -- SECTION 2 – STYLE PROPERTY WRITER
  -- ===========================================================================

  /**
   * Write pdfmake style key-value pairs into the currently-open apex_json object.
   * Caller must have already called apex_json.open_object before invoking this.
   *
   * Mappings:
   *   font-name        → font
   *   font-size        → fontSize  (ODF length → pt)
   *   font-weight bold → bold: true
   *   font-style italic→ italics: true
   *   text-decoration  → decoration: "underline" | "lineThrough"
   *   color            → color
   *   background-color → fillColor
   *   text-align       → alignment
   *   margin-top/bot/left → margin: [left, 0, top, bottom]
   *   line-height %    → lineHeight  (multiplier, e.g. 120% → 1.2)
   */
   procedure write_style_props (
      p_font_name       in varchar2 default null,
      p_font_size       in varchar2 default null,
      p_font_weight     in varchar2 default null,
      p_font_style      in varchar2 default null,
      p_text_decoration in varchar2 default null,
      p_color           in varchar2 default null,
      p_bg_color        in varchar2 default null,
      p_text_align      in varchar2 default null,
      p_margin_top      in varchar2 default null,
      p_margin_bottom   in varchar2 default null,
      p_margin_left     in varchar2 default null,
      p_line_height     in varchar2 default null
   ) is
      l_fs  number;
      l_mt  number;
      l_mb  number;
      l_ml  number;
      l_lh  number;
      l_col varchar2(10);
      l_bgc varchar2(10);
      l_aln varchar2(10);
   begin
      if p_font_name is not null then
         apex_json.write(
            'font',
            p_font_name
         );
      end if;
      l_fs := odf_length_to_pt(p_font_size);
      if l_fs is not null then
         apex_json.write(
            'fontSize',
            l_fs
         );
      end if;
      if lower(p_font_weight) in ( 'bold',
                                   '700',
                                   '800',
                                   '900' ) then
         apex_json.write(
            'bold',
            true
         );
      end if;

      if lower(p_font_style) = 'italic' then
         apex_json.write(
            'italics',
            true
         );
      end if;

      if instr(
         lower(p_text_decoration),
         'underline'
      ) > 0 then
         apex_json.write(
            'decoration',
            'underline'
         );
      elsif instr(
         lower(p_text_decoration),
         'line-through'
      ) > 0 then
         apex_json.write(
            'decoration',
            'lineThrough'
         );
      end if;

      l_col := map_color(p_color);
      if l_col is not null then
         apex_json.write(
            'color',
            l_col
         );
      end if;
      l_bgc := map_color(p_bg_color);
      if l_bgc is not null then
         apex_json.write(
            'fillColor',
            l_bgc
         );
      end if;
      l_aln := map_alignment(p_text_align);
      if l_aln is not null then
         apex_json.write(
            'alignment',
            l_aln
         );
      end if;
      l_mt := odf_length_to_pt(p_margin_top);
      l_mb := odf_length_to_pt(p_margin_bottom);
      l_ml := odf_length_to_pt(p_margin_left);
      if l_mt is not null
      or l_mb is not null
      or l_ml is not null then
         apex_json.open_array('margin');
         apex_json.write(nvl(
            l_ml,
            0
         ));
         apex_json.write(0);
         apex_json.write(nvl(
            l_mt,
            0
         ));
         apex_json.write(nvl(
            l_mb,
            0
         ));
         apex_json.close_array;
      end if;

      if
         p_line_height is not null
         and instr(
            p_line_height,
            '%'
         ) > 0
      then
         l_lh := round(
            to_number(replace(
               p_line_height,
               '%',
               ''
            )) / 100,
            2
         );
         apex_json.write(
            'lineHeight',
            l_lh
         );
      end if;
   end write_style_props;

  -- ===========================================================================
  -- SECTION 3 – STYLE LOADER, RESOLVER, WRITER
  -- ===========================================================================

  /**
   * Resolve the full inheritance chain for one style in g_style_map.
   * Depth-first: resolves parent before merging into child.
   * Cycle safety: clears parent after resolution — revisiting stops immediately.
   */
   procedure resolve_inheritance (
      p_name in varchar2
   ) is
      l_sty   t_sty;
      l_par   t_sty;
      l_pname varchar2(200);
   begin
      if not g_style_map.exists(p_name) then
         return;
      end if;
      l_sty   := g_style_map(p_name);
      l_pname := l_sty.parent;
      if l_pname is null
      or not g_style_map.exists(l_pname) then
         return;
      end if;
      resolve_inheritance(l_pname);        -- depth-first
      l_par := g_style_map(l_pname);
      if l_sty.font_name   is null then l_sty.font_name   := l_par.font_name;   end if;
      if l_sty.font_size   is null then l_sty.font_size   := l_par.font_size;   end if;
      if l_sty.font_weight is null then l_sty.font_weight := l_par.font_weight; end if;
      if l_sty.font_style  is null then l_sty.font_style  := l_par.font_style;  end if;
      if l_sty.text_deco   is null then l_sty.text_deco   := l_par.text_deco;   end if;
      if l_sty.color       is null then l_sty.color       := l_par.color;       end if;
      if l_sty.bg_color    is null then l_sty.bg_color    := l_par.bg_color;    end if;
      if l_sty.text_align  is null then l_sty.text_align  := l_par.text_align;  end if;
      if l_sty.margin_top  is null then l_sty.margin_top  := l_par.margin_top;  end if;
      if l_sty.margin_bot  is null then l_sty.margin_bot  := l_par.margin_bot;  end if;
      if l_sty.margin_lft  is null then l_sty.margin_lft  := l_par.margin_lft;  end if;
      if l_sty.line_height is null then l_sty.line_height := l_par.line_height; end if;
      l_sty.parent := null;             -- mark resolved; breaks cycles
      g_style_map(p_name) := l_sty;
   end resolve_inheritance;

  /**
   * Load all styles from styles.xml and content.xml into g_style_map,
   * then resolve the full inheritance chain for every entry.
   *
   * Must be called before write_content so that write_paragraph can look up
   * resolved style properties and write them inline into content nodes.
   *
   * sources:
   *   styles.xml  – office:styles + office:automatic-styles  (named styles)
   *   content.xml – office:automatic-styles  (document-local; win on name collision)
   */
   procedure load_styles (
      p_styles_xml  in clob,
      p_content_xml in clob default null
   ) is
      l_buf t_sty_tab;
      l_nm  varchar2(200);
   begin
      -- Clear previous session state
      g_style_map.delete;

    -- ---- 1. Named styles from styles.xml ----
      if p_styles_xml is not null then
         select x.style_name,
                x.family,
                x.parent,
                x.text_align,
                x.margin_top,
                x.margin_bot,
                x.margin_lft,
                x.line_height,
                x.font_name,
                x.font_size,
                x.font_weight,
                x.font_style,
                x.text_deco,
                x.color,
                x.bg_color
         bulk collect
           into l_buf
           from xmltable (
                   xmlnamespaces (
                      'urn:oasis:names:tc:opendocument:xmlns:office:1.0' as "office",
                      'urn:oasis:names:tc:opendocument:xmlns:style:1.0'  as "style",
                      'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0' as "fo"
                   ),
                   '//(office:styles|office:automatic-styles)/style:style
                      [@style:family="paragraph" or @style:family="text"]'
                   passing xmltype(p_styles_xml)
                   columns
                      style_name  varchar2(200) path '@style:name',
                      family      varchar2(50)  path '@style:family',
                      parent      varchar2(200) path '@style:parent-style-name',
                      text_align  varchar2(30)  path 'style:paragraph-properties/@fo:text-align',
                      margin_top  varchar2(20)  path 'style:paragraph-properties/@fo:margin-top',
                      margin_bot  varchar2(20)  path 'style:paragraph-properties/@fo:margin-bottom',
                      margin_lft  varchar2(20)  path 'style:paragraph-properties/@fo:margin-left',
                      line_height varchar2(20)  path 'style:paragraph-properties/@fo:line-height',
                      font_name   varchar2(100) path 'style:text-properties/@style:font-name',
                      font_size   varchar2(20)  path 'style:text-properties/@fo:font-size',
                      font_weight varchar2(20)  path 'style:text-properties/@fo:font-weight',
                      font_style  varchar2(20)  path 'style:text-properties/@fo:font-style',
                      text_deco   varchar2(50)  path 'style:text-properties/@fo:text-decoration-line',
                      color       varchar2(20)  path 'style:text-properties/@fo:color',
                      bg_color    varchar2(20)  path 'style:text-properties/@fo:background-color'
                ) x;

         for i in 1..l_buf.count loop
            g_style_map(l_buf(i).style_name) := l_buf(i);
         end loop;
      end if;

    -- ---- 2. Automatic styles from content.xml (overwrite named styles on collision) ----
      if p_content_xml is not null then
         select x.style_name,
                x.family,
                x.parent,
                x.text_align,
                x.margin_top,
                x.margin_bot,
                x.margin_lft,
                x.line_height,
                x.font_name,
                x.font_size,
                x.font_weight,
                x.font_style,
                x.text_deco,
                x.color,
                x.bg_color
         bulk collect
           into l_buf
           from xmltable (
                   xmlnamespaces (
                      'urn:oasis:names:tc:opendocument:xmlns:office:1.0' as "office",
                      'urn:oasis:names:tc:opendocument:xmlns:style:1.0'  as "style",
                      'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0' as "fo"
                   ),
                   '//office:automatic-styles/style:style
                      [@style:family="paragraph" or @style:family="text"]'
                   passing xmltype(p_content_xml)
                   columns
                      style_name  varchar2(200) path '@style:name',
                      family      varchar2(50)  path '@style:family',
                      parent      varchar2(200) path '@style:parent-style-name',
                      text_align  varchar2(30)  path 'style:paragraph-properties/@fo:text-align',
                      margin_top  varchar2(20)  path 'style:paragraph-properties/@fo:margin-top',
                      margin_bot  varchar2(20)  path 'style:paragraph-properties/@fo:margin-bottom',
                      margin_lft  varchar2(20)  path 'style:paragraph-properties/@fo:margin-left',
                      line_height varchar2(20)  path 'style:paragraph-properties/@fo:line-height',
                      font_name   varchar2(100) path 'style:text-properties/@style:font-name',
                      font_size   varchar2(20)  path 'style:text-properties/@fo:font-size',
                      font_weight varchar2(20)  path 'style:text-properties/@fo:font-weight',
                      font_style  varchar2(20)  path 'style:text-properties/@fo:font-style',
                      text_deco   varchar2(50)  path 'style:text-properties/@fo:text-decoration-line',
                      color       varchar2(20)  path 'style:text-properties/@fo:color',
                      bg_color    varchar2(20)  path 'style:text-properties/@fo:background-color'
                ) x;

         for i in 1..l_buf.count loop
            g_style_map(l_buf(i).style_name) := l_buf(i);
         end loop;
      end if;

    -- ---- 3. Resolve full inheritance chains ----
      l_nm := g_style_map.first;
      while l_nm is not null loop
         resolve_inheritance(l_nm);
         l_nm := g_style_map.next(l_nm);
      end loop;
   end load_styles;

  /**
   * Write "styles" and "defaultStyle" JSON keys using the already-populated g_style_map.
   * load_styles must have been called first.
   * Only styles referenced in content.xml are emitted.
   */
   procedure write_styles (
      p_styles_xml  in clob,
      p_content_xml in clob default null
   ) is
      type t_name_set is
         table of varchar2(1) index by varchar2(200);
      l_used   t_name_set;
      l_nm     varchar2(200);
      l_s      t_sty;
      l_def_fn varchar2(100);
      l_def_fs varchar2(20);
      l_def_fw varchar2(20);
      l_def_fi varchar2(20);
      l_def_co varchar2(20);
      l_def_ta varchar2(30);
   begin
    -- ---- 1. Collect referenced style names from content.xml ----
      if p_content_xml is not null then
         for r in (
            select x.sname
              from xmltable (
                      xmlnamespaces (
                         'urn:oasis:names:tc:opendocument:xmlns:text:1.0' as "text"
                      ),
                      '//*[@text:style-name]'
                      passing xmltype(p_content_xml)
                      columns sname varchar2(200) path '@text:style-name'
                   ) x
             where x.sname is not null
            union
            select x.sname
              from xmltable (
                      xmlnamespaces (
                         'urn:oasis:names:tc:opendocument:xmlns:table:1.0' as "table"
                      ),
                      '//*[@table:style-name]'
                      passing xmltype(p_content_xml)
                      columns sname varchar2(200) path '@table:style-name'
                   ) x
             where x.sname is not null
         ) loop
            l_used(r.sname) := '1';
         end loop;
      end if;

    -- ---- 2. Emit "styles" – only referenced styles, flat (no basedOn) ----
      apex_json.open_object('styles');
      l_nm := g_style_map.first;
      while l_nm is not null loop
         l_s := g_style_map(l_nm);
         if p_content_xml is null
         or l_used.exists(l_nm)
         then
            apex_json.open_object(l_nm);
            write_style_props(
               p_font_name       => l_s.font_name,
               p_font_size       => l_s.font_size,
               p_font_weight     => l_s.font_weight,
               p_font_style      => l_s.font_style,
               p_text_decoration => l_s.text_deco,
               p_color           => l_s.color,
               p_bg_color        => l_s.bg_color,
               p_text_align      => l_s.text_align,
               p_margin_top      => l_s.margin_top,
               p_margin_bottom   => l_s.margin_bot,
               p_margin_left     => l_s.margin_lft,
               p_line_height     => l_s.line_height
            );
            apex_json.close_object;
         end if;
         l_nm := g_style_map.next(l_nm);
      end loop;
      apex_json.close_object;  -- close "styles"

    -- ---- 3. "defaultStyle" from styles.xml ----
      apex_json.open_object('defaultStyle');
      if p_styles_xml is not null then
         begin
            select x.fn,
                   x.fs,
                   x.fw,
                   x.fi,
                   x.co,
                   x.ta
              into l_def_fn,
                   l_def_fs,
                   l_def_fw,
                   l_def_fi,
                   l_def_co,
                   l_def_ta
              from xmltable (
                      xmlnamespaces (
                         'urn:oasis:names:tc:opendocument:xmlns:office:1.0' as "office",
                         'urn:oasis:names:tc:opendocument:xmlns:style:1.0'  as "style",
                         'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0' as "fo"
                      ),
                      '//office:styles/style:default-style[@style:family="paragraph"]'
                      passing xmltype(p_styles_xml)
                      columns
                         fn varchar2(100) path 'style:text-properties/@style:font-name',
                         fs varchar2(20)  path 'style:text-properties/@fo:font-size',
                         fw varchar2(20)  path 'style:text-properties/@fo:font-weight',
                         fi varchar2(20)  path 'style:text-properties/@fo:font-style',
                         co varchar2(20)  path 'style:text-properties/@fo:color',
                         ta varchar2(30)  path 'style:paragraph-properties/@fo:text-align'
                   ) x
             where rownum = 1;

            write_style_props(
               p_font_name   => l_def_fn,
               p_font_size   => l_def_fs,
               p_font_weight => l_def_fw,
               p_font_style  => l_def_fi,
               p_color       => l_def_co,
               p_text_align  => l_def_ta
            );
         exception
            when no_data_found then
               null;
         end;
      end if;
      apex_json.close_object;  -- close "defaultStyle"
   end write_styles;

  -- ===========================================================================
  -- SECTION 4 – PARAGRAPH WRITER
  -- ===========================================================================

  /**
   * Write a single paragraph or heading node into the currently-open apex_json array.
   * p_para_node is the text:p or text:h DOM node (owned by the caller's document).
   *
   * Simple paragraph  →  {"text":"Hello World","style":"Text_Body"}
   * Mixed-run para    →  {"text":[{"text":"Hello "},{"text":"World","style":"Bold"}],
   *                        "style":"Text_Body"}
   * Heading           →  {"text":"Title","style":"Heading1","headlineLevel":1}
   */
   procedure write_paragraph (
      p_para_node   in dbms_xmldom.DOMNode,
      p_para_style  in varchar2,
      p_heading_lvl in pls_integer default null
   ) is
      type t_run is record (
            rtype varchar2(20),
            txt   varchar2(32767),
            sty   varchar2(200)
      );
      type t_run_tab is
         table of t_run index by pls_integer;
      l_runs      t_run_tab;
      l_run_count pls_integer := 0;
      l_has_spans boolean := false;
      l_full_txt  varchar2(32767) := '';

      l_children  dbms_xmldom.DOMNodeList;
      l_child     dbms_xmldom.DOMNode;
      l_ntype     pls_integer;
      l_lname     varchar2(200);
      l_celem     dbms_xmldom.DOMElement;
      l_run       t_run;
   begin
      l_children := dbms_xmldom.getChildNodes(p_para_node);

      for i in 0..dbms_xmldom.getLength(l_children) - 1 loop
         l_child := dbms_xmldom.item(
            l_children,
            i
         );
         l_ntype := dbms_xmldom.getNodeType(l_child);

         if l_ntype = dbms_xmldom.TEXT_NODE then
            l_celem := dbms_xmldom.makeElement(l_child);
            l_run.rtype     := 'text';
            l_run.txt       := dbms_xmldom.getNodeValue(l_child);
            l_run.sty       := null;
            l_run_count     := l_run_count + 1;
            l_runs(l_run_count) := l_run;
         elsif l_ntype = dbms_xmldom.ELEMENT_NODE then
            l_celem := dbms_xmldom.makeElement(l_child);
            l_lname := dbms_xmldom.getLocalName(l_celem);
            case l_lname
               when 's' then
                  l_run.rtype     := 'text';
                  l_run.txt       := ' ';
                  l_run.sty       := null;
                  l_run_count     := l_run_count + 1;
                  l_runs(l_run_count) := l_run;
               when 'tab' then
                  l_run.rtype     := 'text';
                  l_run.txt       := chr(9);
                  l_run.sty       := null;
                  l_run_count     := l_run_count + 1;
                  l_runs(l_run_count) := l_run;
               when 'line-break' then
                  l_run.rtype     := 'text';
                  l_run.txt       := chr(10);
                  l_run.sty       := null;
                  l_run_count     := l_run_count + 1;
                  l_runs(l_run_count) := l_run;
               when 'span' then
                  l_run.rtype     := 'span';
                  l_run.txt       := dom_get_text(l_child);
                  l_run.sty       := dom_get_attr(
                     l_celem,
                     c_ns_text,
                     'style-name'
                  );
                  l_run_count     := l_run_count + 1;
                  l_runs(l_run_count) := l_run;
                  l_has_spans     := true;
               else
                  null;
            end case;
         end if;
      end loop;

      apex_json.open_object;
      if l_run_count = 0 then
      -- Empty paragraph
         apex_json.write(
            'text',
            ''
         );
      elsif not l_has_spans then
      -- Plain text: concatenate all fragments
         for i in 1..l_run_count loop
            l_full_txt := l_full_txt || l_runs(i).txt;
         end loop;
         apex_json.write(
            'text',
            l_full_txt
         );
      else
      -- Mixed inline runs: emit as an array of run objects
         apex_json.open_array('text');
         for i in 1..l_run_count loop
            apex_json.open_object;
            apex_json.write(
               'text',
               l_runs(i).txt
            );
            if
               l_runs(i).rtype = 'span'
               and l_runs(i).sty is not null
            then
               apex_json.write(
                  'style',
                  l_runs(i).sty
               );
            end if;
            apex_json.close_object;
         end loop;
         apex_json.close_array;
      end if;

    -- Paragraph / heading style metadata
      if p_heading_lvl is not null then
         apex_json.write(
            'style',
            'Heading' || to_char(p_heading_lvl)
         );
         apex_json.write(
            'headlineLevel',
            p_heading_lvl
         );
      elsif p_para_style is not null then
         apex_json.write(
            'style',
            p_para_style
         );
      end if;

    -- Write resolved style properties inline so each content node is self-contained
      if p_para_style is not null
      and g_style_map.exists(p_para_style)
      then
         declare
            l_s t_sty := g_style_map(p_para_style);
         begin
            write_style_props(
               p_font_name       => l_s.font_name,
               p_font_size       => l_s.font_size,
               p_font_weight     => l_s.font_weight,
               p_font_style      => l_s.font_style,
               p_text_decoration => l_s.text_deco,
               p_color           => l_s.color,
               p_bg_color        => l_s.bg_color,
               p_text_align      => l_s.text_align,
               p_margin_top      => l_s.margin_top,
               p_margin_bottom   => l_s.margin_bot,
               p_margin_left     => l_s.margin_lft,
               p_line_height     => l_s.line_height
            );
         end;
      end if;

      apex_json.close_object;
   end write_paragraph;

  -- ===========================================================================
  -- SECTION 5 – TABLE WRITER
  -- ===========================================================================

  /**
   * Write a pdfmake table node into the currently-open apex_json array.
   * p_table_node is the table:table DOM node (owned by the caller's document).
   *
   * Output structure:
   *   {"table":{"widths":["*","*"],"body":[[{cell},{cell}],…]}}
   *
   * Supports: column count, cell text, style-name, col/row span.
   */
   procedure write_table (
      p_table_node in dbms_xmldom.DOMNode
   ) is
      l_col_count  number := 0;
      l_children   dbms_xmldom.DOMNodeList;
      l_child      dbms_xmldom.DOMNode;
      l_ntype      pls_integer;
      l_lname      varchar2(200);
      l_elem       dbms_xmldom.DOMElement;
      l_rpt        number;
      l_row_children dbms_xmldom.DOMNodeList;
      l_row_child    dbms_xmldom.DOMNode;
      l_row_elem     dbms_xmldom.DOMElement;
      l_row_lname    varchar2(200);
      l_cspan      number;
      l_rspan      number;
      l_sty        varchar2(200);
      l_txt        varchar2(32767);
   begin
    -- Pass 1: count columns from table:table-column elements
      l_children := dbms_xmldom.getChildNodes(p_table_node);
      for i in 0..dbms_xmldom.getLength(l_children) - 1 loop
         l_child := dbms_xmldom.item(
            l_children,
            i
         );
         if dbms_xmldom.getNodeType(l_child) = dbms_xmldom.ELEMENT_NODE then
            l_elem  := dbms_xmldom.makeElement(l_child);
            l_lname := dbms_xmldom.getLocalName(l_elem);
            if l_lname = 'table-column' then
               l_rpt       := to_number(nvl(
                  dom_get_attr(
                     l_elem,
                     c_ns_table,
                     'number-columns-repeated'
                  ),
                  '1'
               ));
               l_col_count := l_col_count + l_rpt;
            end if;
         end if;
      end loop;
      if l_col_count = 0 then
         l_col_count := 1;
      end if;

      apex_json.open_object;
      apex_json.open_object('table');

    -- widths: ["*", "*", ...]
      apex_json.open_array('widths');
      for i in 1..l_col_count loop
         apex_json.write('*');
      end loop;
      apex_json.close_array;

    -- body: [[{cell},…],…]
      apex_json.open_array('body');

    -- Pass 2: rows and cells
      for i in 0..dbms_xmldom.getLength(l_children) - 1 loop
         l_child := dbms_xmldom.item(
            l_children,
            i
         );
         if dbms_xmldom.getNodeType(l_child) = dbms_xmldom.ELEMENT_NODE then
            l_elem  := dbms_xmldom.makeElement(l_child);
            l_lname := dbms_xmldom.getLocalName(l_elem);
            if l_lname = 'table-row' then
               apex_json.open_array;  -- row

               l_row_children := dbms_xmldom.getChildNodes(l_child);
               for j in 0..dbms_xmldom.getLength(l_row_children) - 1 loop
                  l_row_child := dbms_xmldom.item(
                     l_row_children,
                     j
                  );
                  if dbms_xmldom.getNodeType(l_row_child) = dbms_xmldom.ELEMENT_NODE then
                     l_row_elem  := dbms_xmldom.makeElement(l_row_child);
                     l_row_lname := dbms_xmldom.getLocalName(l_row_elem);
                     if l_row_lname in ( 'table-cell',
                                         'covered-table-cell' ) then
                        l_sty   := dom_get_attr(
                           l_row_elem,
                           c_ns_table,
                           'style-name'
                        );
                        l_cspan := to_number(nvl(
                           dom_get_attr(
                              l_row_elem,
                              c_ns_table,
                              'number-columns-spanned'
                           ),
                           '1'
                        ));
                        l_rspan := to_number(nvl(
                           dom_get_attr(
                              l_row_elem,
                              c_ns_table,
                              'number-rows-spanned'
                           ),
                           '1'
                        ));
                        l_txt   := dom_get_text(l_row_child);

                        apex_json.open_object;
                        apex_json.write(
                           'text',
                           l_txt
                        );
                        if l_sty is not null then
                           apex_json.write(
                              'style',
                              l_sty
                           );
                        end if;
                        if l_cspan > 1 then
                           apex_json.write(
                              'colSpan',
                              l_cspan
                           );
                        end if;
                        if l_rspan > 1 then
                           apex_json.write(
                              'rowSpan',
                              l_rspan
                           );
                        end if;
                        apex_json.close_object;
                     end if;
                  end if;
               end loop;

               apex_json.close_array;  -- row
            end if;
         end if;
      end loop;

      apex_json.close_array;   -- body
      apex_json.close_object;  -- table
      apex_json.close_object;  -- outer node
   end write_table;

  -- ===========================================================================
  -- SECTION 6 – LIST WRITER
  -- ===========================================================================

  /**
   * Write a pdfmake ul/ol node into the currently-open apex_json array.
   * p_list_node is the text:list DOM node (owned by the caller's document).
   *
   * Ordered detection: style name containing NUMBER / NUMER / ENUM / OL.
   *
   * Output:
   *   {"ul":["item 1","item 2"]}        (unordered)
   *   {"ol":["item 1","item 2"]}        (ordered)
   */
   procedure write_list (
      p_list_node  in dbms_xmldom.DOMNode,
      p_style_name in varchar2 default null,
      p_is_ordered in boolean default false
   ) is
      l_tag   varchar2(4) :=
         case
            when p_is_ordered then
               'ol'
            else
               'ul'
         end;
      l_children     dbms_xmldom.DOMNodeList;
      l_child        dbms_xmldom.DOMNode;
      l_elem         dbms_xmldom.DOMElement;
      l_lname        varchar2(200);
      l_item_ch      dbms_xmldom.DOMNodeList;
      l_item_child   dbms_xmldom.DOMNode;
      l_item_elem    dbms_xmldom.DOMElement;
      l_item_lname   varchar2(200);
      l_item_txt     varchar2(32767);
      l_first_p_done boolean;
   begin
      apex_json.open_object;
      apex_json.open_array(l_tag);

      l_children := dbms_xmldom.getChildNodes(p_list_node);
      for i in 0..dbms_xmldom.getLength(l_children) - 1 loop
         l_child := dbms_xmldom.item(
            l_children,
            i
         );
         if dbms_xmldom.getNodeType(l_child) = dbms_xmldom.ELEMENT_NODE then
            l_elem  := dbms_xmldom.makeElement(l_child);
            l_lname := dbms_xmldom.getLocalName(l_elem);
            if l_lname = 'list-item' then
            -- Get text of first text:p child
               l_item_txt     := null;
               l_first_p_done := false;
               l_item_ch      := dbms_xmldom.getChildNodes(l_child);
               for j in 0..dbms_xmldom.getLength(l_item_ch) - 1 loop
                  if not l_first_p_done then
                     l_item_child := dbms_xmldom.item(
                        l_item_ch,
                        j
                     );
                     if dbms_xmldom.getNodeType(l_item_child) = dbms_xmldom.ELEMENT_NODE then
                        l_item_elem  := dbms_xmldom.makeElement(l_item_child);
                        l_item_lname := dbms_xmldom.getLocalName(l_item_elem);
                        if l_item_lname = 'p' then
                           l_item_txt     := dom_get_text(l_item_child);
                           l_first_p_done := true;
                        end if;
                     end if;
                  end if;
               end loop;
               apex_json.write(nvl(
                  l_item_txt,
                  ''
               ));
            end if;
         end if;
      end loop;

      apex_json.close_array;
      if p_style_name is not null then
         apex_json.write(
            'style',
            p_style_name
         );
      end if;
      apex_json.close_object;
   end write_list;

  -- ===========================================================================
  -- SECTION 7 – CONTENT WRITER
  -- ===========================================================================

  /**
   * Write the "content" array key into the currently-open apex_json object.
   * Iterates all top-level body elements: p, h, table, list.
   */
   procedure write_content (
      p_content_xml in clob
   ) is

      l_xml       xmltype;
      l_doc       dbms_xmldom.DOMDocument;
      l_root      dbms_xmldom.DOMElement;
      l_text_list dbms_xmldom.DOMNodeList;
      l_text_node dbms_xmldom.DOMNode;
      l_children  dbms_xmldom.DOMNodeList;
      l_child     dbms_xmldom.DOMNode;
      l_elem      dbms_xmldom.DOMElement;
      l_lname     varchar2(200);
      l_sty       varchar2(200);
      l_olvl      pls_integer;

      function is_ordered (
         p_sty in varchar2
      ) return boolean is
      begin
         if p_sty is null then
            return false;
         end if;
         if instr(
            upper(p_sty),
            'NUMBER'
         ) > 0 then
            return true;
         end if;
         if instr(
            upper(p_sty),
            'NUMER'
         ) > 0 then
            return true;
         end if;
         if instr(
            upper(p_sty),
            'ENUM'
         ) > 0 then
            return true;
         end if;
         if instr(
            upper(p_sty),
            'OL'
         ) > 0 then
            return true;
         end if;
         return false;
      end;

   begin
      apex_json.open_array('content');
      if p_content_xml is null then
         apex_json.close_array;
         return;
      end if;

      l_xml  := xmltype(p_content_xml);
      l_doc  := dbms_xmldom.newDOMDocument(l_xml);
      l_root := dbms_xmldom.getDocumentElement(l_doc);

    -- Find the single office:text element
      l_text_list := dbms_xmldom.getElementsByTagName(
         l_root,
         'text',
         c_ns_office
      );
      if dbms_xmldom.getLength(l_text_list) = 0 then
         dbms_xmldom.freeDocument(l_doc);
         apex_json.close_array;
         return;
      end if;
      l_text_node := dbms_xmldom.item(
         l_text_list,
         0
      );

    -- Iterate direct children of office:text
      l_children := dbms_xmldom.getChildNodes(l_text_node);
      for i in 0..dbms_xmldom.getLength(l_children) - 1 loop
         l_child := dbms_xmldom.item(
            l_children,
            i
         );
         if dbms_xmldom.getNodeType(l_child) = dbms_xmldom.ELEMENT_NODE then
            l_elem  := dbms_xmldom.makeElement(l_child);
            l_lname := dbms_xmldom.getLocalName(l_elem);
            case l_lname
               when 'p' then
                  l_sty := dom_get_attr(
                     l_elem,
                     c_ns_text,
                     'style-name'
                  );
                  write_paragraph(
                     l_child,
                     l_sty
                  );
               when 'h' then
                  l_sty  := dom_get_attr(
                     l_elem,
                     c_ns_text,
                     'style-name'
                  );
                  l_olvl := to_number(nvl(
                     dom_get_attr(
                        l_elem,
                        c_ns_text,
                        'outline-level'
                     ),
                     '1'
                  ));
                  write_paragraph(
                     l_child,
                     l_sty,
                     l_olvl
                  );
               when 'table' then
                  write_table(l_child);
               when 'list' then
                  l_sty := dom_get_attr(
                     l_elem,
                     c_ns_text,
                     'style-name'
                  );
                  write_list(
                     l_child,
                     l_sty,
                     is_ordered(l_sty)
                  );
               else
                  null;
            end case;
         end if;
      end loop;

      dbms_xmldom.freeDocument(l_doc);
      apex_json.close_array;  -- content
   exception
      when others then
         if not dbms_xmldom.isNull(l_doc) then
            dbms_xmldom.freeDocument(l_doc);
         end if;
         raise;
   end write_content;

  -- ===========================================================================
  -- SECTION 8 – PUBLIC API
  -- ===========================================================================

   function parse_xml (
      p_content_xml in clob,
      p_styles_xml  in clob default null
   ) return clob is
      l_result clob;
   begin
      apex_json.initialize_clob_output;
      load_styles(p_styles_xml, p_content_xml);      -- populate g_style_map first
      apex_json.open_object;            -- {
      write_content(p_content_xml);                  --   "content": [...] (inline props via g_style_map)
      write_styles(p_styles_xml, p_content_xml);     --   "styles": {...}, "defaultStyle": {...}
      apex_json.close_object;           -- }
      l_result := apex_json.get_clob_output;
      apex_json.free_output;
      return l_result;
   exception
      when others then
         apex_json.free_output;
         raise;
   end parse_xml;

   function parse_odt (
      p_odt_blob in blob
   ) return clob is
      l_cxml   clob;
      l_sxml   clob;
      l_result clob;
   begin
      l_cxml := unzip_member(
         p_odt_blob,
         'content.xml'
      );
      l_sxml := unzip_member(
         p_odt_blob,
         'styles.xml'
      );
      l_result := parse_xml(
         l_cxml,
         l_sxml
      );
      if dbms_lob.istemporary(l_cxml) = 1 then
         dbms_lob.freetemporary(l_cxml);
      end if;
      if dbms_lob.istemporary(l_sxml) = 1 then
         dbms_lob.freetemporary(l_sxml);
      end if;
      return l_result;
   exception
      when others then
         if
            l_cxml is not null
            and dbms_lob.istemporary(l_cxml) = 1
         then
            dbms_lob.freetemporary(l_cxml);
         end if;
         if
            l_sxml is not null
            and dbms_lob.istemporary(l_sxml) = 1
         then
            dbms_lob.freetemporary(l_sxml);
         end if;
         raise;
   end parse_odt;

end odt_parser;
/

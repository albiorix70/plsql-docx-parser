create or replace package body odt_parser as

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
  -- SECTION 3 – STYLES WRITER
  -- ===========================================================================

  /**
   * Write "styles" and "defaultStyle" keys into the currently-open apex_json object.
   * Reads paragraph/text styles from styles.xml and emits pdfmake style definitions.
   */
   procedure write_styles (
      p_styles_xml in clob
   ) is

      l_xml    xmltype;
      type t_sty is record (
            style_name  varchar2(200),
            family      varchar2(50),
            parent      varchar2(200),
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
      type t_sty_tab is
         table of t_sty;
      l_styles t_sty_tab;

    -- defaultStyle fields
      l_def_fn varchar2(100);
      l_def_fs varchar2(20);
      l_def_fw varchar2(20);
      l_def_fi varchar2(20);
      l_def_co varchar2(20);
      l_def_ta varchar2(30);
   begin
    -- ---- "styles" object ----
      apex_json.open_object('styles');
      if p_styles_xml is not null then
         l_xml := xmltype(p_styles_xml);
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
           into l_styles
           from xmltable ( xmlnamespaces ( 'urn:oasis:names:tc:opendocument:xmlns:office:1.0' as "office",'urn:oasis:names:tc:opendocument:xmlns:style:1.0'
           as "style",'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0' as "fo" ),
         '//(office:styles|office:automatic-styles)/style:style
           [@style:family="paragraph" or @style:family="text"]'
               passing l_xml
            columns
               style_name varchar2(200) path '@style:name',
               family varchar2(50) path '@style:family',
               parent varchar2(200) path '@style:parent-style-name',
               text_align varchar2(30) path 'style:paragraph-properties/@fo:text-align',
               margin_top varchar2(20) path 'style:paragraph-properties/@fo:margin-top',
               margin_bot varchar2(20) path 'style:paragraph-properties/@fo:margin-bottom',
               margin_lft varchar2(20) path 'style:paragraph-properties/@fo:margin-left',
               line_height varchar2(20) path 'style:paragraph-properties/@fo:line-height',
               font_name varchar2(100) path 'style:text-properties/@style:font-name',
               font_size varchar2(20) path 'style:text-properties/@fo:font-size',
               font_weight varchar2(20) path 'style:text-properties/@fo:font-weight',
               font_style varchar2(20) path 'style:text-properties/@fo:font-style',
               text_deco varchar2(50) path 'style:text-properties/@fo:text-decoration-line',
               color varchar2(20) path 'style:text-properties/@fo:color',
               bg_color varchar2(20) path 'style:text-properties/@fo:background-color'
         ) x;

         if
            l_styles is not null
            and l_styles.count > 0
         then
            for i in l_styles.first..l_styles.last loop
               apex_json.open_object(l_styles(i).style_name);
               write_style_props(
                  p_font_name       => l_styles(i).font_name,
                  p_font_size       => l_styles(i).font_size,
                  p_font_weight     => l_styles(i).font_weight,
                  p_font_style      => l_styles(i).font_style,
                  p_text_decoration => l_styles(i).text_deco,
                  p_color           => l_styles(i).color,
                  p_bg_color        => l_styles(i).bg_color,
                  p_text_align      => l_styles(i).text_align,
                  p_margin_top      => l_styles(i).margin_top,
                  p_margin_bottom   => l_styles(i).margin_bot,
                  p_margin_left     => l_styles(i).margin_lft,
                  p_line_height     => l_styles(i).line_height
               );
               if l_styles(i).parent is not null then
                  apex_json.write(
                     'basedOn',
                     l_styles(i).parent
                  );
               end if;
               apex_json.close_object;
            end loop;
         end if;
      end if;

      apex_json.close_object;  -- close "styles"

    -- ---- "defaultStyle" object ----
      apex_json.open_object('defaultStyle');
      if p_styles_xml is not null then
         begin
            select x.fn,
                   x.fs,
                   x.fw,
                   x.fi,
                   x.co,
                   x.ta
              into
               l_def_fn,
               l_def_fs,
               l_def_fw,
               l_def_fi,
               l_def_co,
               l_def_ta
              from xmltable ( xmlnamespaces ( 'urn:oasis:names:tc:opendocument:xmlns:office:1.0' as "office",'urn:oasis:names:tc:opendocument:xmlns:style:1.0'
              as "style",'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0' as "fo" ),
            '//office:styles/style:default-style[@style:family="paragraph"]'
                  passing l_xml
               columns
                  fn varchar2(100) path 'style:text-properties/@style:font-name',
                  fs varchar2(20) path 'style:text-properties/@fo:font-size',
                  fw varchar2(20) path 'style:text-properties/@fo:font-weight',
                  fi varchar2(20) path 'style:text-properties/@fo:font-style',
                  co varchar2(20) path 'style:text-properties/@fo:color',
                  ta varchar2(30) path 'style:paragraph-properties/@fo:text-align'
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
   *
   * Simple paragraph  →  {"text":"Hello World","style":"Text_Body"}
   * Mixed-run para    →  {"text":[{"text":"Hello "},{"text":"World","style":"Bold"}],
   *                        "style":"Text_Body"}
   * Heading           →  {"text":"Title","style":"Heading1","headlineLevel":1}
   */
   procedure write_paragraph (
      p_para_xml    in xmltype,
      p_para_style  in varchar2,
      p_heading_lvl in pls_integer default null
   ) is
      type t_run is record (
            rtype varchar2(20),
            txt   varchar2(32767),
            sty   varchar2(200),
            pos   number
      );
      type t_run_tab is
         table of t_run;
      l_runs      t_run_tab;
      l_has_spans boolean := false;
      l_full_txt  varchar2(32767) := '';
   begin
    -- Collect inline content: direct text nodes + text:span children
      select rtype,
             txt,
             sty,
             pos
      bulk collect
        into l_runs
        from (
         select 'text' as rtype,
                x.tc as txt,
                null as sty,
                x.pos as pos
           from xmltable ( xmlnamespaces ( 'urn:oasis:names:tc:opendocument:xmlns:text:1.0' as "text" ),
         'text()|text:s|text:tab|text:line-break'
               passing p_para_xml
            columns
               tc varchar2(4000) path '.',
               pos for ordinality
         ) x
         union all
         select 'span' as rtype,
                x.stxt as txt,
                x.ssty as sty,
                x.pos + 1000000 as pos
           from xmltable ( xmlnamespaces ( 'urn:oasis:names:tc:opendocument:xmlns:text:1.0' as "text" ),
         'text:span'
               passing p_para_xml
            columns
               ssty varchar2(200) path '@text:style-name',
               stxt varchar2(4000) path 'string(.)',
               pos for ordinality
         ) x
      )
       order by pos;

    -- Detect mixed content
      if l_runs is not null then
         for i in l_runs.first..l_runs.last loop
            if l_runs(i).rtype = 'span' then
               l_has_spans := true;
               exit;
            end if;
         end loop;
      end if;

      apex_json.open_object;
      if l_runs is null
      or l_runs.count = 0 then
      -- Empty paragraph
         apex_json.write(
            'text',
            ''
         );
      elsif not l_has_spans then
      -- Plain text: concatenate all text-node fragments
         for i in l_runs.first..l_runs.last loop
            l_full_txt := l_full_txt || l_runs(i).txt;
         end loop;
         apex_json.write(
            'text',
            l_full_txt
         );
      else
      -- Mixed inline runs: emit as an array of run objects
         apex_json.open_array('text');
         for i in l_runs.first..l_runs.last loop
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

      apex_json.close_object;
   end write_paragraph;

  -- ===========================================================================
  -- SECTION 5 – TABLE WRITER
  -- ===========================================================================

  /**
   * Write a pdfmake table node into the currently-open apex_json array.
   *
   * Output structure:
   *   {"table":{"widths":["*","*"],"body":[[{cell},{cell}],…]}}
   *
   * Supports: column count, cell text, style-name, col/row span.
   */
   procedure write_table (
      p_table_xml in xmltype
   ) is

      type t_cell is record (
            row_idx number,
            col_idx number,
            cspan   number,
            rspan   number,
            sty     varchar2(200),
            txt     varchar2(32767)
      );
      type t_cell_tab is
         table of t_cell;
      l_cells     t_cell_tab;
      type t_col is record (
         rpt number
      );
      type t_col_tab is
         table of t_col;
      l_cols      t_col_tab;
      l_col_count number := 0;
      l_cur_row   number := 0;
      l_row_open  boolean := false;
   begin
    -- Column count from table:table-column
      select nvl(
         x.rpt,
         1
      )
      bulk collect
        into l_cols
        from xmltable ( xmlnamespaces ( 'urn:oasis:names:tc:opendocument:xmlns:table:1.0' as "table" ),
      'table:table-column'
            passing p_table_xml
         columns
            rpt number path '@table:number-columns-repeated'
      ) x;

      if l_cols is not null then
         for i in l_cols.first..l_cols.last loop
            l_col_count := l_col_count + l_cols(i).rpt;
         end loop;
      end if;
      if l_col_count = 0 then
         l_col_count := 1;
      end if;

    -- Collect cells
      select x.ri,
             x.ci,
             nvl(
                x.cs,
                1
             ),
             nvl(
                x.rs,
                1
             ),
             x.sty,
             x.txt
      bulk collect
        into l_cells
        from xmltable ( xmlnamespaces ( 'urn:oasis:names:tc:opendocument:xmlns:table:1.0' as "table",'urn:oasis:names:tc:opendocument:xmlns:text:1.0'
        as "text" ),
      'table:table-row/table:table-cell | table:table-row/table:covered-table-cell'
            passing p_table_xml
         columns
            ri number path 'count(../preceding-sibling::table:table-row)+1',
            ci number path 'count(preceding-sibling::table:table-cell)
                                  + count(preceding-sibling::table:covered-table-cell) + 1',
            cs number path '@table:number-columns-spanned',
            rs number path '@table:number-rows-spanned',
            sty varchar2(200) path '@table:style-name',
            txt varchar2(4000) path 'string(.)'
      ) x;

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
      if
         l_cells is not null
         and l_cells.count > 0
      then
         for i in l_cells.first..l_cells.last loop
              -- New row?
            if l_cells(i).row_idx != l_cur_row then
               if l_row_open then
                  apex_json.close_array;  -- close previous row
               end if;
               l_cur_row := l_cells(i).row_idx;
               l_row_open := true;
               apex_json.open_array;  -- open new row
            end if;

              -- Cell object
            apex_json.open_object;
            apex_json.write(
               'text',
               l_cells(i).txt
            );
            if l_cells(i).sty is not null then
               apex_json.write(
                  'style',
                  l_cells(i).sty
               );
            end if;
            if l_cells(i).cspan > 1 then
               apex_json.write(
                  'colSpan',
                  l_cells(i).cspan
               );
            end if;
            if l_cells(i).rspan > 1 then
               apex_json.write(
                  'rowSpan',
                  l_cells(i).rspan
               );
            end if;
            apex_json.close_object;
         end loop;

         if l_row_open then
            apex_json.close_array;  -- close last row
         end if;
      end if;

      apex_json.close_array;   -- body
      apex_json.close_object;    -- table
      apex_json.close_object;      -- outer node
   end write_table;

  -- ===========================================================================
  -- SECTION 6 – LIST WRITER
  -- ===========================================================================

  /**
   * Write a pdfmake ul/ol node into the currently-open apex_json array.
   *
   * Ordered detection: style name containing NUMBER / NUMER / ENUM / OL.
   * List items are taken from the first text:p of each text:list-item.
   *
   * Output:
   *   {"ul":["item 1","item 2"]}        (unordered)
   *   {"ol":["item 1","item 2"]}        (ordered)
   */
   procedure write_list (
      p_list_xml   in xmltype,
      p_style_name in varchar2 default null,
      p_is_ordered in boolean default false
   ) is
      type t_item is record (
            txt varchar2(32767),
            pos number
      );
      type t_item_tab is
         table of t_item;
      l_items t_item_tab;
      l_tag   varchar2(4) :=
         case
            when p_is_ordered then
               'ol'
            else
               'ul'
         end;
   begin
      select x.txt,
             x.pos
      bulk collect
        into l_items
        from xmltable ( xmlnamespaces ( 'urn:oasis:names:tc:opendocument:xmlns:text:1.0' as "text" ),
      'text:list-item'
            passing p_list_xml
         columns
            txt varchar2(4000) path 'string(text:p[1])',
            pos for ordinality
      ) x
       order by x.pos;

      apex_json.open_object;
      apex_json.open_array(l_tag);
      if
         l_items is not null
         and l_items.count > 0
      then
         for i in l_items.first..l_items.last loop
            apex_json.write(l_items(i).txt);
         end loop;
      end if;
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

      l_xml   xmltype;
      type t_node is record (
            kind varchar2(20),
            pos  number,
            sty  varchar2(200),
            olvl number,
            nxml xmltype
      );
      type t_node_tab is
         table of t_node;
      l_nodes t_node_tab;

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
      l_xml := xmltype(p_content_xml);

    -- Collect top-level body elements
      select x.kind,
             x.pos,
             x.sty,
             x.olvl,
             x.nxml
      bulk collect
        into l_nodes
        from xmltable ( xmlnamespaces ( 'urn:oasis:names:tc:opendocument:xmlns:office:1.0' as "office",'urn:oasis:names:tc:opendocument:xmlns:text:1.0'
        as "text",'urn:oasis:names:tc:opendocument:xmlns:table:1.0' as "table" ),
      '//office:body/office:text/*'
            passing l_xml
         columns
            kind varchar2(20) path 'local-name(.)',
            pos for ordinality,
            sty varchar2(200) path '@text:style-name',
            olvl number path '@text:outline-level',
            nxml xmltype path '.'
      ) x
       where x.kind in ( 'p',
                         'h',
                         'table',
                         'list' )
       order by x.pos;

      if
         l_nodes is not null
         and l_nodes.count > 0
      then
         for i in l_nodes.first..l_nodes.last loop
            case l_nodes(i).kind
               when 'p' then
                  write_paragraph(
                     l_nodes(i).nxml,
                     l_nodes(i).sty
                  );
               when 'h' then
                  write_paragraph(
                     l_nodes(i).nxml,
                     l_nodes(i).sty,
                     nvl(
                        l_nodes(i).olvl,
                        1
                     )
                  );
               when 'table' then
                  write_table(l_nodes(i).nxml);
               when 'list' then
                  write_list(
                     l_nodes(i).nxml,
                     l_nodes(i).sty,
                     is_ordered(l_nodes(i).sty)
                  );
               else
                  null;
            end case;
         end loop;
      end if;

      apex_json.close_array;  -- content
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
      apex_json.open_object;            -- {
      write_content(p_content_xml);   --   "content": [...]
      write_styles(p_styles_xml);     --   "styles": {...}, "defaultStyle": {...}
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
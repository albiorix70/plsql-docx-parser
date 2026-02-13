create or replace package body pdf_make_util as
   /*
   * This package provides utility functions for generating PDF content
   * using the pdfmake library. It includes functions for styling text,
   * adding paragraphs, and validating color formats.
   */
   g_content json_array_t := json_array_t();
   g_styles  json_object_t := json_object_t();

   -- This function checks if the provided color is in a valid RGB format or a hex code.
   function check_color (
      p_color in varchar2
   ) return varchar2;

   -- This function initializes the content for PDF generation.
   procedure init_content (
      p_reuse_styles in boolean default true
   ) is
   begin
      g_content := json_array_t();
      if not p_reuse_styles then
         g_styles := json_object_t();
      end if;
   end init_content;

   procedure free_content is
   begin
      init_content(p_reuse_styles => false);
   end free_content;

   -- This function generates a JSON object with style properties for PDF content.
   function style_properties (
      p_font             in varchar2 default null,
      p_font_size        in number default null,
      p_line_height      in number default null,
      p_bold             in boolean default null,
      p_italic           in boolean default null,
      p_alignment        in varchar2 default null,
      p_char_spacing     in number default null,
      p_color            in varchar2 default null,
      p_background       in varchar2 default null,
      p_marker_color     in varchar2 default null,
      p_decoration       in varchar2 default null,
      p_decoration_style in varchar2 default null,
      p_decoration_color in varchar2 default null
   ) return json_object_t is
      l_font             varchar2(32767) default p_font;
      l_font_size        number default p_font_size;
      l_line_height      number default p_line_height;
      l_bold             boolean default p_bold;
      l_italic           boolean default p_italic;
      l_alignment        varchar2(32767) default p_alignment;
      l_char_spacing     number default p_char_spacing;
      l_color            varchar2(32767) default check_color(p_color);
      l_background       varchar2(32767) default check_color(p_background);
      l_marker_color     varchar2(32767) default check_color(p_marker_color);
      l_decoration       varchar2(32767) default p_decoration;
      l_decoration_style varchar2(32767) default check_color(p_decoration_style);
      l_decoration_color varchar2(32767) default check_color(p_decoration_color);
      r                  json_object_t;
      sql_out            varchar2(32767);
   begin
      select
         json_object(
            key 'font' value p_font,
                     key 'fontSize' value l_font_size,
                     key 'lineHeight' value l_line_height,
                     key 'bold' value l_bold,
                     key 'italics' value l_italic,
                     key 'alignment' value l_alignment,
                     key 'characterSpacing' value l_char_spacing,
                     key 'color' value l_color,
                     key 'background' value l_background,
                     key 'markerColor' value l_marker_color,
                     key 'decoration' value p_decoration,
                     key 'decorationStyle' value l_decoration_style,
                     key 'decorationColor' value l_decoration_color
         absent on null)
        into sql_out
        from dual;
      r := json_object_t.parse(sql_out);
      return r;
   end style_properties;

   procedure add_paragraph (
      p_text       in varchar2,
      p_properties in json_object_t default null
   ) is
   begin
      null;
   end add_paragraph;

   -- noch in entwicklung
   function get_string (
      p_in in varchar2
   ) return varchar2 is
      r json_object_t;
   begin
      r := style_properties(
         p_color     => p_in,
         p_font      => 'Helvetica',
         p_font_size => 12
      );
      return r.stringify;
   end get_string;

   function get_content return clob is
      l_clob clob;
   begin
      l_clob := '{
      "content": ["Lorem ipsum dolor sit amet, consectetur adipisicing elit. Malit profecta versatur nomine ocurreret multavit, officiis viveremus aeternum superstitio suspicor alia nostram, quando nostros congressus susceperant concederetur leguntur iam, vigiliae democritea tantopere causae, atilii plerumque ipsas potitur pertineant multis rem quaeri pro, legendum didicisse credere ex maluisset per videtis. Cur discordans praetereat aliae ruinae dirigentur orestem eodem, praetermittenda divinum. Collegisti, deteriora malint loquuntur officii cotidie finitas referri doleamus ambigua acute. Adhaesiones ratione beate arbitraretur detractis perdiscere, constituant hostis polyaeno. Diu concederetur."]  
   }';
      return l_clob;
   end get_content;

   -- Private Functions 
   function check_color (
      p_color in varchar2
   ) return varchar2 is
      l_rgb_regexp constant varchar2(100) := '^(rgb\(\d{1,3},\s*\d{1,3},\s*\d{1,3}\)|#([0-9a-fA-F]{6}|[0-9a-fA-F]{3}))$';
      l_default    constant varchar2(20) := 'rgb(0, 0, 0)'; -- Default color
   begin
      if p_color is null
      or p_color = '' then
         return null; -- Default to black if no color is provided
      else
             -- Here you would typically validate the color format
         return
            case
               when regexp_like(
                  p_color,
                  l_rgb_regexp
               ) then
                  p_color
               else l_default -- If the color format is invalid, return default color
            end;
      end if;
      return l_default; -- Fallback in case of unexpected conditions 
   end check_color;

end pdf_make_util;
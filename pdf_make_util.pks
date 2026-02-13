create or replace package pdf_make_util as
   /*
   * This package provides utility functions for generating PDF content
   * using the pdfmake library. It includes functions for styling text,
   * adding paragraphs, and validating color formats.
   */

   procedure init_content (
      p_reuse_styles in boolean default true
   );

   procedure free_content;

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
   ) return json_object_t;

   function get_string (
      p_in in varchar2
   ) return varchar2;

   procedure add_paragraph (
      p_text       in varchar2,
      p_properties in json_object_t default null
   );

   function get_content return clob;
end pdf_make_util;
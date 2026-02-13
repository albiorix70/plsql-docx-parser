create or replace package pdfmake as
   procedure clear_content;

   function generate_pdf return clob;

   procedure add_text (
      p_text in varchar2
   );
   procedure add_image (
      p_image_data in blob
   );
   procedure add_table (
      p_table_data in clob
   );
   procedure set_metadata (
      p_title  in varchar2,
      p_author in varchar2
   );
   procedure add_style (
      p_style_name in varchar2,
      p_font_size  in number,
      p_bold       in boolean,
      p_italic     in boolean,
      p_color      in varchar2,
      p_underline  in boolean
   );
end pdfmake;
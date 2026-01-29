set echo off
set serveroutput off
set feedback on

@docx_parser_util.pks
@docx_parser_util.pkb

@docx_parser.pks
@docx_parser.pkb

show errors package docx_parser_util

-- exit

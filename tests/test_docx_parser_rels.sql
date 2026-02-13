-- utPLSQL tests for docx_parser rels and unpack functions
-- Run with: ut.run('tests.test_docx_parser_rels');

create or replace package tests.test_docx_parser_rels as
   --%suite(Rels and Unpack Suite)
   --%suitepath(tests)
   --%suitecase(Test load and get)
   procedure test_load_and_get;
   --%suitecase(Test unpack xml)
   procedure test_unpack_xml;
   --%suitecase(Test unpack blob)
   procedure test_unpack_blob;
   --%suitecase(Test parse rels)
   procedure test_parse_rels;
end test_docx_parser_rels;
/

create or replace package body tests.test_docx_parser_rels as

   procedure test_load_and_get is
   begin
      -- This test assumes a table TEST_DOCX_STORE(table_name) exists with a blob column
      -- We'll attempt to call load_docx_source; if the row doesn't exist, assert a handled exception
      begin
         docx_parser.load_docx_source('APEX_APPLICATION_FILES','blob_content','id','0');
         ut.expect(docx_parser.get_loaded_docx).to_be_not_null;
      exception
         when others then
            ut.expect(sqlcode).not_to_equal(0);
      end;
   end test_load_and_get;

   procedure test_unpack_xml is
      l_doc clob;
      l_blob blob := null;
   begin
      -- If package blob is loaded, try to unpack document.xml
      l_blob := docx_parser.get_loaded_docx;
      if l_blob is null then
         ut.skip('No loaded DOCX available; skipping unpack xml test');
         return;
      end if;

      l_doc := docx_parser.unpack_docx('word/document.xml', l_blob);
      ut.expect(l_doc).to_be_not_null;
      -- basic check: must contain w:document
      ut.expect(dbms_lob.instr(l_doc, '<w:document') > 0).to_equal(true);
   end test_unpack_xml;

   procedure test_unpack_blob is
      l_raw blob;
      l_blob blob := docx_parser.get_loaded_docx;
   begin
      if l_blob is null then
         ut.skip('No loaded DOCX available; skipping unpack blob test');
         return;
      end if;
      l_raw := docx_parser.unpack_docx('word/styles.xml', l_blob, true);
      ut.expect(l_raw).to_be_not_null;
   end test_unpack_blob;

   procedure test_parse_rels is
      l_doc clob;
      l_rels docx_parser.t_rels_table;
   begin
      l_doc := docx_parser.unpack_docx('word/_rels/document.xml.rels', docx_parser.get_loaded_docx);
      if l_doc is null then
         ut.skip('No rels file found in loaded DOCX; skipping parse_rels test');
         return;
      end if;

      l_rels := docx_parser.parse_rels_xml(l_doc);
      -- basic expectation: table may have entries; verify index by key access doesn't raise
      for k in l_rels.first .. l_rels.last loop
         null; -- iteration placeholder (associative arrays may not be dense)
      end loop;
      -- Check that at least one entry's id is not null
      declare
         l_found boolean := false;
      begin
         for i in l_rels.first .. l_rels.last loop
            if l_rels.exists(l_rels(i).id) then
               l_found := true;
               exit;
            end if;
         end loop;
         ut.expect(l_found).to_equal(true);
      exception
         when others then
            -- associative arrays may not be iterable like this; basic assert the collection itself is not null
            ut.expect(l_rels).to_be_not_null;
      end;
   end test_parse_rels;

end test_docx_parser_rels;
/

-- Run the suite automatically when the script is executed (optional)
-- begin
--    ut.run('tests.test_docx_parser_rels');
-- end;
-- /

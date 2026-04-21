create or replace package body xml_node_parser as

    -- -------------------------------------------------------------------------
    -- parse_namespaces
    -- Liest alle xmlns-Attribute des übergebenen Nodes und gibt eine
    -- Prefix→URI-Map zurück.
    -- xmlns="..."       → Eintrag unter Prefix ''  (Default-Namespace)
    -- xmlns:prefix="..."→ Eintrag unter 'prefix'
    -- -------------------------------------------------------------------------
    function parse_namespaces(
        p_root in dbms_xmldom.domnode
    ) return t_ns_map is
        l_result   t_ns_map;
        l_nm_map   dbms_xmldom.domnamednodemap;
        l_attr     dbms_xmldom.domnode;
        l_nm       varchar2(200);
        l_uri      varchar2(4000);
    begin
        l_nm_map := dbms_xmldom.getattributes(p_root);
        for i in 0 .. dbms_xmldom.getlength(l_nm_map) - 1 loop
            l_attr := dbms_xmldom.item(l_nm_map, i);
            l_nm   := dbms_xmldom.getnodename(l_attr);
            l_uri  := dbms_xmldom.getnodevalue(l_attr);
            if l_nm = 'xmlns' then
                -- Default-Namespace
                l_result('') := l_uri;
            elsif l_nm like 'xmlns:%' then
                -- Benannter Namespace-Prefix
                l_result(replace(l_nm, 'xmlns:', '')) := l_uri;
            end if;
        end loop;
        return l_result;
    end parse_namespaces;

    -- -------------------------------------------------------------------------
    -- parse_node  (rekursiv)
    -- Wandelt einen einzelnen ELEMENT_NODE in ein json_object_t um:
    --   name, ns, text (direkte TEXT-Kinder), attributes[], children[]
    -- -------------------------------------------------------------------------
    function parse_node(
        p_node in dbms_xmldom.domnode
    ) return json_object_t is
        l_obj      json_object_t;
        l_attrs    json_array_t;
        l_children json_array_t;
        l_text     varchar2(32767) := '';

        l_elem     dbms_xmldom.domelement;
        l_nm_map   dbms_xmldom.domnamednodemap;
        l_attr     dbms_xmldom.domnode;
        l_aobj     json_object_t;

        l_kids     dbms_xmldom.domnodelist;
        l_kid      dbms_xmldom.domnode;
        l_ntype    pls_integer;
        l_child_obj json_object_t;
    begin
        -- Nur ELEMENT_NODEs verarbeiten
        if dbms_xmldom.getnodetype(p_node) != dbms_xmldom.element_node then
            return null;
        end if;

        l_elem     := dbms_xmldom.makeelement(p_node);
        l_obj      := json_object_t();
        l_attrs    := json_array_t();
        l_children := json_array_t();

        -- Name und Namespace-URI des Elements
        l_obj.put('name', dbms_xmldom.getlocalname(l_elem));
        l_obj.put('ns',   dbms_xmldom.getnamespace(l_elem));

        -- Attribute einlesen
        l_nm_map := dbms_xmldom.getattributes(p_node);
        for i in 0 .. dbms_xmldom.getlength(l_nm_map) - 1 loop
            l_attr := dbms_xmldom.item(l_nm_map, i);
            l_aobj := json_object_t();
            l_aobj.put('name',  dbms_xmldom.getlocalname(dbms_xmldom.makeattr(l_attr)));
            l_aobj.put('ns',    dbms_xmldom.getnamespace(dbms_xmldom.makeattr(l_attr)));
            l_aobj.put('value', dbms_xmldom.getnodevalue(l_attr));
            l_attrs.append(l_aobj);
        end loop;

        -- Kinder verarbeiten
        l_kids := dbms_xmldom.getchildnodes(p_node);
        for i in 0 .. dbms_xmldom.getlength(l_kids) - 1 loop
            l_kid   := dbms_xmldom.item(l_kids, i);
            l_ntype := dbms_xmldom.getnodetype(l_kid);

            if l_ntype = dbms_xmldom.text_node then
                -- Direkte TEXT-Kinder zusammenfassen
                l_text := l_text || dbms_xmldom.getnodevalue(l_kid);

            elsif l_ntype = dbms_xmldom.element_node then
                -- Element-Kinder rekursiv parsen
                l_child_obj := parse_node(l_kid);
                if l_child_obj is not null then
                    l_children.append(l_child_obj);
                end if;

            end if;
            -- Andere Node-Typen (comment, cdata, processing-instruction) werden ignoriert
        end loop;

        l_obj.put('text',       l_text);
        l_obj.put('attributes', l_attrs);
        l_obj.put('children',   l_children);

        return l_obj;
    end parse_node;

    -- -------------------------------------------------------------------------
    -- parse
    -- Einstiegspunkt: XMLType → DOMDocument → Root parsen + Namespaces anfügen
    -- -------------------------------------------------------------------------
    function parse(
        p_xml in xmltype
    ) return json_object_t is
        l_doc    dbms_xmldom.domdocument;
        l_root   dbms_xmldom.domnode;
        l_ns_map t_ns_map;
        l_ns_obj json_object_t;
        l_result json_object_t;
        l_prefix varchar2(100);
    begin
        l_doc  := dbms_xmldom.newdomdocument(p_xml);
        l_root := dbms_xmldom.makenode(dbms_xmldom.getdocumentelement(l_doc));

        -- Namespace-Deklarationen aus dem Root-Element lesen
        l_ns_map := parse_namespaces(l_root);

        -- Namespace-Map in JSON-Objekt überführen
        l_ns_obj := json_object_t();
        l_prefix := l_ns_map.first;
        while l_prefix is not null loop
            l_ns_obj.put(l_prefix, l_ns_map(l_prefix));
            l_prefix := l_ns_map.next(l_prefix);
        end loop;

        -- Root-Node rekursiv parsen
        l_result := parse_node(l_root);
        l_result.put('namespaces', l_ns_obj);

        dbms_xmldom.freedocument(l_doc);
        return l_result;
    end parse;

end xml_node_parser;
/

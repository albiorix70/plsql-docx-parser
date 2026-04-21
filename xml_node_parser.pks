create or replace package xml_node_parser as

    -- Namespace-Prefix → URI Map  (z.B. 'w' → 'http://schemas...')
    type t_ns_map is table of varchar2(4000) index by varchar2(100);

    -- Namespace-Deklarationen aus dem Root-Element einlesen.
    -- Iteriert alle xmlns-Attribute und gibt eine Prefix→URI-Map zurück.
    -- Der Default-Namespace (xmlns="...") wird unter dem Prefix '' gespeichert.
    function parse_namespaces(
        p_root in dbms_xmldom.domnode
    ) return t_ns_map;

    -- Einen einzelnen DOM-Node in ein JSON-Objekt umwandeln.
    -- Gibt null zurück, wenn p_node kein ELEMENT_NODE ist.
    --
    -- Struktur des Rückgabeobjekts:
    -- {
    --   "name":       varchar2         (local name des Elements)
    --   "ns":         varchar2         (Namespace-URI, null wenn keiner)
    --   "text":       varchar2         (concat aller direkten TEXT_NODE-Kinder)
    --   "attributes": json_array_t     ([{"name": str, "ns": str, "value": str}, ...])
    --   "children":   json_array_t     ([json_object_t, ...] rekursiv)
    -- }
    function parse_node(
        p_node in dbms_xmldom.domnode
    ) return json_object_t;

    -- Top-Level-Einstieg: XMLType parsen, ab Root-Element.
    -- Gibt den vollständigen Node-Baum als json_object_t zurück.
    -- Zusätzlicher Key "namespaces": {"prefix": "uri", ...}
    function parse(
        p_xml in xmltype
    ) return json_object_t;

end xml_node_parser;
/

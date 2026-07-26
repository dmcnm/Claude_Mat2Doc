function cls_name = registry(clark_name)
% REGISTRY Custom element-class lookup: Clark tag name -> MATLAB class name.
%
%   cls_name = MAT2DOC.OXML.REGISTRY(clark_name) returns the fully qualified
%   MATLAB class name (string) registered for the element tag clark_name
%   ("{uri}local"), or "" when no custom class is registered -- the caller
%   (createElement / the parser) then falls back to plain XmlElement, mirroring
%   lxml's fallback to _Element for unregistered tags.
%
%   This is the MATLAB analogue of python-docx's parse-time class lookup:
%   `element_class_lookup = etree.ElementNamespaceClassLookup()` configured on
%   `oxml_parser` (docx/oxml/parser.py lines 18-20), populated by the
%   `register_element_cls(tag, cls)` calls. lxml keys the lookup by
%   (namespace URI, local part) -- register_element_cls lines 39-41 -- which is
%   exactly the Clark name, so the table here is keyed by Clark name.
%
%   REGISTRATION TABLE POLICY (design.md section 2 "Factories / registries"):
%   an explicit static table, one line per Python `register_element_cls` call,
%   kept in Python source order, audited line-by-line. Each WP that ports CT_*
%   classes appends its registration lines. Count guard (H10): 120
%   register_element_cls calls TOTAL for docx, all in docx/oxml/__init__.py
%   (catalogs\docx_catalog.json). The SEPARATE 5-class OPC lookup in
%   docx/opc/oxml.py (<Default>/<Override>/<Types>/<Relationship>/
%   <Relationships>, bound via ct_namespace[...] / pr_namespace[...] not
%   register_element_cls) is its OWN element_class_lookup in docx and lands
%   with the OPC-layer WP.
%
%   TABLE CONTENT: the 120 main-map rows (docx/oxml/__init__.py
%   register_element_cls calls) are appended by their CT_* WPs in
%   docx/oxml/__init__.py source order and remain EMPTY here until those WPs
%   land. SEPARATELY, P1-4 merges the 5 OPC element classes (CT_Default /
%   CT_Override / CT_Types / CT_Relationship / CT_Relationships) into this same
%   lookup, keyed by their RAW CLARK NAMES (plan-audit planaudit_2026-07-25
%   condition B2, option A). In docx these 5 live in a SEPARATE
%   element_class_lookup on the OPC parser (docx/opc/oxml.py:240-247), bound via
%   ct_namespace[...]/pr_namespace[...] rather than register_element_cls. Merging
%   them here is behavior-preserving because the ct/pr namespaces are DISJOINT
%   from the main w:* namespaces, so no tag resolves differently under one merged
%   table than under docx's two separate lookups (see +opc/+oxml/parse_xml.m).
%   The 5 OPC rows are added by raw Clark name via registerClark_ -- NOT via
%   registerElementCls_, which would resolve the ct/pr prefixes through the main
%   nsmap (mat2doc.oxml.nsmap has NO ct/pr) and hard-error.
%
%   COUNT GUARD (H10): 120 main-map rows (target, tracked as CT_* WPs land) + 5
%   OPC rows (present now) = 125 total. The two groups are tracked separately.
%
%   Inputs:  clark_name - (1,1) string, e.g. "{http://.../main}p"
%   Outputs: cls_name   - (1,1) string, e.g. "mat2doc.oxml.CT_P", or ""
%
%   Example:
%       cls = mat2doc.oxml.registry(mat2doc.oxml.qn("w:p"));   % "" until P1-x
%
%   Ported from python-docx v1.2.0: src/docx/oxml/parser.py::register_element_cls
%   + element_class_lookup (lines 18-41; registration blocks in
%   docx/oxml/__init__.py pending their CT_* WPs)

arguments
    clark_name (1,1) string
end
persistent map
if isempty(map)
    map = buildRegistry_();
end
if isKey(map, clark_name)
    cls_name = map(clark_name);
else
    cls_name = "";
end
end

function map = buildRegistry_()
% BUILDREGISTRY_ The explicit registration table (built once).
map = dictionary(string.empty(0, 1), string.empty(0, 1));
% -------------------------------------------------------------------------
% MAIN-MAP rows (docx/oxml/__init__.py): one registerElementCls_ line per Python
% register_element_cls call, in docx/oxml/__init__.py source order. Lines are
% appended by the WP that ports the corresponding CT_* class, e.g.:
%
%   map = registerElementCls_(map, "w:document", "mat2doc.oxml.CT_Document");
%
% The H10 dispatch-matrix probe (row count vs Python, target 120) applies as
% CT_* rows land. FIRST row added by P1-7 (coreprops): cp:coreProperties. The
% remaining 119 are appended by their CT_* WPs in docx/oxml/__init__.py order.
map = registerElementCls_(map, "cp:coreProperties", ...
    "mat2doc.oxml.coreprops.CT_CoreProperties");   % __init__.py:96 (P1-7)
% -------------------------------------------------------------------------
% OPC rows (P1-4; docx/opc/oxml.py:240-247): the 5 OPC element classes, keyed by
% RAW CLARK NAME (condition B2). The Clark URIs come from mat2doc.opc.NAMESPACE
% so they are IDENTICAL to the xmlns the CT_*.new factories emit and the parser
% resolves -- guaranteeing the parsed element's Clark key matches this row. NOT
% routed through registerElementCls_ (ct/pr are absent from the main nsmap).
NS = mat2doc.opc.NAMESPACE;
map = registerClark_(map, "{" + NS.OPC_CONTENT_TYPES + "}Default", ...
    "mat2doc.opc.oxml.CT_Default");        % ct_namespace["Default"]  opc/oxml.py:241
map = registerClark_(map, "{" + NS.OPC_CONTENT_TYPES + "}Override", ...
    "mat2doc.opc.oxml.CT_Override");       % ct_namespace["Override"] opc/oxml.py:242
map = registerClark_(map, "{" + NS.OPC_CONTENT_TYPES + "}Types", ...
    "mat2doc.opc.oxml.CT_Types");          % ct_namespace["Types"]    opc/oxml.py:243
map = registerClark_(map, "{" + NS.OPC_RELATIONSHIPS + "}Relationship", ...
    "mat2doc.opc.oxml.CT_Relationship");   % pr_namespace["Relationship"]  opc/oxml.py:246
map = registerClark_(map, "{" + NS.OPC_RELATIONSHIPS + "}Relationships", ...
    "mat2doc.opc.oxml.CT_Relationships");  % pr_namespace["Relationships"] opc/oxml.py:247
end

function map = registerClark_(map, clark, cls_name)
% REGISTERCLARK_ Register cls_name for a RAW Clark tag key (no prefix resolution).
%   Used for OPC classes whose ct/pr prefixes are not in the main nsmap, so the
%   NamespacePrefixedTag path in registerElementCls_ cannot be used.
arguments
    map dictionary
    clark (1,1) string
    cls_name (1,1) string
end
if isKey(map, clark)
    error("mat2doc:InternalError", ...
        "duplicate element-class registration for '%s'", clark);
end
map(clark) = cls_name;
end

function map = registerElementCls_(map, tag, cls_name) %#ok<DEFNU> -- invoked by table lines as CT_* WPs add them
% REGISTERELEMENTCLS_ Register cls_name to be constructed for elements having tag.
%
%   tag is a string of the form "nspfx:tagroot", e.g. "w:document".
%
%   Ported from python-docx v1.2.0: src/docx/oxml/parser.py::register_element_cls
%   (lines 32-41)
arguments
    map dictionary
    tag (1,1) string
    cls_name (1,1) string
end
% Python (lines 39-41): keys the class into the lookup by
% (nsmap[nspfx], tagroot) -- equivalent to the Clark name.
nsptag = mat2doc.oxml.NamespacePrefixedTag(tag);
key = nsptag.clark_name;
% Python dict assignment would silently overwrite; upstream never registers the
% same tag twice, so a duplicate here can only be a table-generation defect --
% fail loudly (build-time guard, not output-visible behavior).
if isKey(map, key)
    error("mat2doc:InternalError", ...
        "duplicate element-class registration for '%s'", tag);
end
map(key) = cls_name;
end

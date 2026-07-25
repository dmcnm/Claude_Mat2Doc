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
%   AT THIS WP (P1-2) THE TABLE IS EMPTY: no CT_* classes are ported yet
%   (xmlchemy/BaseOxmlElement is P1-3; the CT_* families are later WPs). Every
%   parsed/created element is therefore a plain XmlElement now, but the lookup
%   hook is fully wired. Rows are appended by their CT_* WPs in
%   docx/oxml/__init__.py source order.
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
% Registration table: one registerElementCls_ line per Python
% register_element_cls call, in docx/oxml/__init__.py source order. Lines are
% appended by the WP that ports the corresponding CT_* class, e.g.:
%
%   map = registerElementCls_(map, "w:document", "mat2doc.oxml.CT_Document");
%
% EMPTY at P1-2 (see file header). The H10 dispatch-matrix probe (row count vs
% Python, target 120) applies once CT_* rows start landing.
% -------------------------------------------------------------------------
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

function element = OxmlElement(nsptag_str, attrs, nsdecls)
% OXMLELEMENT Return a "loose" element having the tag specified by nsptag_str.
%
%   element = MAT2DOC.OXML.OXMLELEMENT(nsptag_str) -- nsptag_str must contain
%   the standard namespace prefix, e.g. "w:tbl". The resulting element is an
%   instance of the custom element class for this tag name if one is registered
%   (registry.m), and carries the namespace declaration for its prefix
%   (nsptag.nsmap), so a loose subtree serializes with its own xmlns:...
%   declaration when its ancestors don't already declare it.
%
%   element = MAT2DOC.OXML.OXMLELEMENT(nsptag_str, attrs) additionally SETS the
%   attributes in attrs on the new element. Because attribute names are Clark
%   names (e.g. qn("w:val")) which are NOT valid struct field names, attrs is
%   an Nx2 string array of [name, value] pairs, applied in row order (dict
%   insertion order, H11). Pass [] (None) for no attributes.
%
%   element = MAT2DOC.OXML.OXMLELEMENT(nsptag_str, attrs, nsdecls) uses the
%   declarations in scalar struct nsdecls (<prefix> -> URI) instead of the
%   single-prefix default nsptag.nsmap.
%
%   Inputs:  nsptag_str - (1,1) string, prefixed tag, e.g. "w:tbl"
%            attrs      - Nx2 string [name, value], optional; default [] (None,
%                         H13) -> no attributes set
%            nsdecls    - (1,1) struct, optional; default [] (None, H13)
%                         selects nsptag.nsmap, the single-prefix map of the tag
%   Outputs: element    - scalar XmlElement (or registered subclass) handle
%
%   NOTE (attrs whose namespace is not in nsdecls): lxml's makeelement would
%   eagerly declare a generated prefix for such an attribute at creation time;
%   this port instead invents the identical nsN prefix at SERIALIZE time (see
%   serialize_part_xml). Every docx OxmlElement call site passes an attribute
%   whose namespace equals the element's own prefix (e.g. qn("w:id") on "w:*"),
%   which is always in nsdecls, so the two are byte-equivalent on the used
%   surface; the differing-namespace attribute is an accepted-unreachable path.
%
%   Example:
%       end_ = mat2doc.oxml.OxmlElement("w:commentRangeEnd", ...
%           [mat2doc.oxml.qn("w:id"), string(commentId)]);
%
%   Ported from python-docx v1.2.0: src/docx/oxml/parser.py::OxmlElement
%   (lines 44-62)

arguments
    nsptag_str (1,1) string
    attrs = []      % Python default None (H13); an Nx2 string [name,value] when provided
    nsdecls = []    % Python default None (H13); a (1,1) struct when provided
end
% Python: nsptag = NamespacePrefixedTag(nsptag_str)  (line 59)
nsptag = mat2doc.oxml.NamespacePrefixedTag(nsptag_str);
% Python: if nsdecls is None: nsdecls = nsptag.nsmap  (lines 60-61)
if isequal(nsdecls, [])
    nsdecls = nsptag.nsmap;
end
% Python: return oxml_parser.makeelement(nsptag.clark_name, attrib=attrs,
%         nsmap=nsdecls)  (line 62)
% createElement resolves the same Clark name (nsptag.clark_name) from
% nsptag_str + nsdecls: for this factory path nsdecls always contains the tag's
% own prefix (nsptag.nsmap or a caller map of standard prefixes), so the
% own-nsdecls-first URI precedence in createElement is unobservable here and
% agrees with nsptag.clark_name.
element = mat2doc.oxml.createElement(nsptag_str, nsdecls);
% attrib=attrs: set each attribute in row (dict-insertion, H11) order.
if ~isequal(attrs, [])
    if ~(isstring(attrs) && ismatrix(attrs) && size(attrs, 2) == 2)
        error("mat2doc:TypeError", ...
            "attrs must be an Nx2 string array of [name, value] pairs or []");
    end
    for k = 1:size(attrs, 1)
        element.set(attrs(k, 1), attrs(k, 2));
    end
end
end

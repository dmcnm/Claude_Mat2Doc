function element = createElement(nsptag_str, nsmap)
% CREATEELEMENT Instantiate the element class registered for a tag.
%
%   element = MAT2DOC.OXML.CREATEELEMENT(nsptag_str) constructs a loose
%   element for the prefixed tag nsptag_str (e.g. "w:p"): an instance of the
%   registered custom class when one exists (see registry.m), otherwise a plain
%   XmlElement -- mirroring lxml, where `oxml_parser.makeelement` applies the
%   class lookup and falls back to plain `_Element` for unregistered tags.
%
%   element = MAT2DOC.OXML.CREATEELEMENT(nsptag_str, nsmap) additionally stores
%   the namespace declarations in scalar struct nsmap on the new element, as
%   lxml `makeelement(clark_name, nsmap=...)` does.
%
%   This is the makeelement-equivalent constructor of the XML layer (design.md
%   section 3). The parse walk (XmlParser.parseElement_) applies the same
%   registry-then-fallback rule per parsed element.
%
%   Inputs:  nsptag_str - (1,1) string, prefixed tag, e.g. "w:p"
%            nsmap      - (1,1) struct, optional: <prefix> -> URI declarations
%                         (field order preserved, H11). Default: none.
%   Outputs: element    - scalar XmlElement (or registered subclass) handle
%
%   Example:
%       p = mat2doc.oxml.createElement("w:p", struct("w", ...
%           "http://schemas.openxmlformats.org/wordprocessingml/2006/main"));
%
%   Ported from python-docx v1.2.0: src/docx/oxml/parser.py::oxml_parser
%   .makeelement with element_class_lookup applied (lines 18-20;
%   design-realization -- lxml has no single equivalent Python def)

arguments
    nsptag_str (1,1) string
    nsmap (1,1) struct = struct()
end
% -- resolve the Clark name for registry lookup, with the same prefix->URI
% -- precedence as the XmlElement constructor (own nsmap first, fixed map
% -- second) so factory and constructed element always agree on the tag URI
parts = split(nsptag_str, ":");
if numel(parts) == 2
    pfx = parts(1);
    local = parts(2);
    if isfield(nsmap, pfx)
        uri = string(nsmap.(pfx));
    else
        fixed = mat2doc.oxml.nsmap();
        if ~isfield(fixed, pfx)
            error("mat2doc:KeyError", "'%s'", pfx);
        end
        uri = fixed.(pfx);
    end
    clark_name = "{" + uri + "}" + local;
elseif numel(parts) == 1
    clark_name = nsptag_str;   % no-namespace tag (plain Clark form)
else
    error("mat2doc:ValueError", "Invalid tag name '%s'", nsptag_str);
end
cls_name = mat2doc.oxml.registry(clark_name);
if cls_name == ""
    % lxml fallback: unregistered tags construct the plain element class
    % (type is _Element, NOT ElementBase/BaseOxmlElement)
    element = mat2doc.oxml.XmlElement(nsptag_str, nsmap);
else
    element = feval(cls_name, nsptag_str, nsmap);
end
end

function clark_name = qn(tag)
% QN Clark-notation qualified tag name for tag, using the OPC-local nsmap.
%
%   clark_name = MAT2DOC.OPC.OXML.QN(tag)
%
%   'qn' stands for 'qualified name'. Turns a namespace-prefixed OPC tag name
%   like "pr:Relationship" or "ct:Default" into a Clark-notation qualified tag
%   name, resolving the prefix against the OPC-LOCAL nsmap (mat2doc.opc.oxml.
%   nsmap = {ct, pr, r}) -- NOT the WordprocessingML mat2doc.oxml.nsmap, which
%   has no ct/pr bindings. For example qn("ct:Default") returns
%       "{http://schemas.openxmlformats.org/package/2006/content-types}Default".
%
%   This is the OPC-layer twin of mat2doc.oxml.qn (docx keeps a separate qn in
%   docx/opc/oxml.py bound to the OPC nsmap). CT_Types.defaults/overrides and
%   CT_Relationships.Relationship_lst call THIS qn.
%
%   Inputs:  tag - (1,1) string, prefixed OPC tag, e.g. "pr:Relationship"
%   Outputs: clark_name - (1,1) string, "{<nsuri>}<tagroot>"
%
%   Example:
%       mat2doc.opc.oxml.qn("pr:Relationship")
%       % "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
%
%   Ported from python-docx v1.2.0: src/docx/opc/oxml.py::qn (lines 41-50)

arguments
    tag (1,1) string
end
% Python: prefix, tagroot = tag.split(":")  (opc/oxml.py line 48)
parts = split(tag, ":");
if numel(parts) < 2
    error("mat2doc:ValueError", ...
        "not enough values to unpack (expected 2, got %d)", numel(parts));
elseif numel(parts) > 2
    error("mat2doc:ValueError", "too many values to unpack (expected 2)");
end
prefix = parts(1);
tagroot = parts(2);
% Python: uri = nsmap[prefix]  (opc/oxml.py line 49)
map = mat2doc.opc.oxml.nsmap();
if ~isfield(map, prefix)
    error("mat2doc:KeyError", "'%s'", prefix);
end
uri = map.(prefix);
% Python: return "{%s}%s" % (uri, tagroot)  (opc/oxml.py line 50)
clark_name = "{" + uri + "}" + tagroot;
end

function clark_name = qn(tag)
% QN Return a Clark-notation qualified tag name for tag.
%
%   clark_name = MAT2DOC.OXML.QN(tag)
%
%   'qn' stands for 'qualified name'. This utility converts a familiar
%   namespace-prefixed tag name like "w:p" into a Clark-notation qualified tag
%   name for the tree layer. For example, qn("w:p") returns
%       "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p".
%
%   Implementation note (hot path): results are memoized in a persistent
%   dictionary keyed by the full prefixed tag. On a miss the value is computed
%   exactly as Python does -- split on ":", look up nsmap, format Clark form --
%   so behavior (including ValueError on malformed tags and KeyError on unknown
%   prefixes) is identical; only repeat-call cost differs. The memo is a pure
%   unordered cache over a fixed map, never iterated (H11 not implicated).
%
%   Inputs:  tag - string scalar, e.g. "w:p"
%   Outputs: clark_name - string scalar, "{<nsuri>}<tagroot>"
%
%   Example:
%       mat2doc.oxml.qn("w:p")
%       % "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"
%
%   Ported from python-docx v1.2.0: src/docx/oxml/ns.py::qn

arguments
    tag (1,1) string
end
persistent memo
if isempty(memo)
    memo = dictionary(string.empty(0, 1), string.empty(0, 1));
end
if isKey(memo, tag)
    clark_name = memo(tag);
    return
end
% Python: prefix, tagroot = tag.split(":")  (ns.py line 107)
parts = split(tag, ":");
if numel(parts) < 2
    error("mat2doc:ValueError", ...
        "not enough values to unpack (expected 2, got %d)", numel(parts));
elseif numel(parts) > 2
    error("mat2doc:ValueError", "too many values to unpack (expected 2)");
end
prefix = parts(1);
tagroot = parts(2);
% Python: uri = nsmap[prefix]  (ns.py line 108)
map = mat2doc.oxml.nsmap();
if ~isfield(map, prefix)
    error("mat2doc:KeyError", "'%s'", prefix);
end
uri = map.(prefix);
% Python: return "{%s}%s" % (uri, tagroot)  (ns.py line 109)
clark_name = "{" + uri + "}" + tagroot;
memo(tag) = clark_name;
end

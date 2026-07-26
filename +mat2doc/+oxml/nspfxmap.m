function nsmap_subset = nspfxmap(nspfxs)
% NSPFXMAP Subset namespace-prefix mappings specified by the given prefixes.
%
%   s = MAT2DOC.OXML.NSPFXMAP(pfx1, pfx2, ...) returns a 1x1 struct containing
%   the subset namespace prefix -> URI mappings specified by the prefixes. Any
%   number of namespace prefixes can be supplied, e.g. nspfxmap("a", "r", "p").
%
%   Representation: Python returns a dict; here a struct whose field order is
%   the argument order (H11: struct fields are insertion-ordered). A repeated
%   prefix keeps its first position, exactly like a Python dict comprehension.
%   Zero arguments -> struct with no fields (Python: {}).
%
%   NOTE: this is docx's subset function -- named nspfxmap in docx ns.py
%   (pptx's equivalent is `namespaces`, absent here). The FULL fixed map is
%   nsmap (nsmap.m).
%
%   Inputs:  nspfxs - repeating string scalars, namespace prefixes
%   Outputs: nsmap_subset - 1x1 struct, <prefix> -> namespace URI
%
%   Example:
%       s = mat2doc.oxml.nspfxmap("a", "r", "w");
%       string(fieldnames(s))'   % ["a"    "r"    "w"]  (argument order)
%       s.w                      % "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
%
%   Ported from python-docx v1.2.0: src/docx/oxml/ns.py::nspfxmap

arguments (Repeating)
    nspfxs (1,1) string
end
map = mat2doc.oxml.nsmap();
nsmap_subset = struct();
% Python: {pfx: nsmap[pfx] for pfx in nspfxs}  (ns.py line 97)
for k = 1:numel(nspfxs)
    pfx = nspfxs{k};
    if ~isfield(map, pfx)
        error("mat2doc:KeyError", "'%s'", pfx);
    end
    nsmap_subset.(pfx) = map.(pfx);
end
end

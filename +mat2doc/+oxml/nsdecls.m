function decls = nsdecls(prefixes)
% NSDECLS Namespace declaration attribute text for the given prefixes.
%
%   decls = MAT2DOC.OXML.NSDECLS(pfx1, pfx2, ...) returns a string like
%   'xmlns:w="http://..." xmlns:r="http://..."' -- one xmlns declaration per
%   prefix, space-separated, in argument order. Zero arguments -> "".
%   Handy for adding required namespace declarations to a tree root element.
%
%   Inputs:  prefixes - repeating string scalars, namespace prefixes
%   Outputs: decls - string scalar
%
%   Example:
%       mat2doc.oxml.nsdecls("w", "r")
%       % 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ...
%       %  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
%
%   Ported from python-docx v1.2.0: src/docx/oxml/ns.py::nsdecls

arguments (Repeating)
    prefixes (1,1) string
end
map = mat2doc.oxml.nsmap();
% Python: " ".join(['xmlns:%s="%s"' % (pfx, nsmap[pfx]) for pfx in prefixes])
parts = strings(1, numel(prefixes));
for k = 1:numel(prefixes)
    pfx = prefixes{k};
    if ~isfield(map, pfx)
        error("mat2doc:KeyError", "'%s'", pfx);
    end
    parts(k) = "xmlns:" + pfx + "=""" + map.(pfx) + """";
end
if isempty(parts)
    decls = "";
else
    decls = strjoin(parts, " ");
end
end

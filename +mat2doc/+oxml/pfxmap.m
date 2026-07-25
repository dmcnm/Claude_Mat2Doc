function map = pfxmap()
% PFXMAP Reverse map: namespace URI -> namespace prefix.
%
%   map = MAT2DOC.OXML.PFXMAP() returns an Nx2 string array; column 1 is the
%   namespace URI, column 2 the prefix. Row order preserves the inversion
%   order of nsmap (H11). Lookup: map(map(:,1) == uri, 2).
%
%   Representation note: Python `pfxmap` is a dict keyed by URI. URIs are not
%   valid MATLAB struct field names, so the dict is represented as an ordered
%   Nx2 string array (design.md: insertion-ordered arrays/structs; never
%   containers.Map). It is only ever used for lookup (from_clark_name).
%
%   Outputs: map - Nx2 string array, [URI, prefix] rows
%
%   Example:
%       map = mat2doc.oxml.pfxmap();
%       uri = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
%       map(map(:, 1) == uri, 2)    % "w"
%
%   Ported from python-docx v1.2.0: src/docx/oxml/ns.py::pfxmap

persistent cached
if isempty(cached)
    % Python: pfxmap = {value: key for key, value in nsmap.items()}
    % Dict-inversion semantics: a repeated URI keeps its first-insertion row
    % position and its LAST prefix wins (no duplicates exist in nsmap, but the
    % construction reproduces the semantics mechanically).
    nsmap_s = mat2doc.oxml.nsmap();
    prefixes = string(fieldnames(nsmap_s));  % insertion order (H11)
    m = strings(0, 2);
    for k = 1:numel(prefixes)
        uri = nsmap_s.(prefixes(k));
        row = find(m(:, 1) == uri, 1);
        if isempty(row)
            m(end + 1, :) = [uri, prefixes(k)]; %#ok<AGROW>
        else
            m(row, 2) = prefixes(k);  % later value overwrites, position kept
        end
    end
    cached = m;
end
map = cached;
end

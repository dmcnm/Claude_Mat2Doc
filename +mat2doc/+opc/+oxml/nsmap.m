function map = nsmap()
% NSMAP OPC-local prefix -> namespace-URI map for [Content_Types].xml and .rels.
%
%   map = MAT2DOC.OPC.OXML.NSMAP() returns a scalar struct whose fields are the
%   OPC-layer namespace prefixes and whose values are the corresponding URIs
%   (string scalars). Field order preserves the Python source dict's insertion
%   order exactly (H11): ct, pr, r.
%
%   This is a SEPARATE, OPC-LOCAL map (docx/opc/oxml.py:24-28), DISTINCT from the
%   WordprocessingML map mat2doc.oxml.nsmap: the main map has NO ct/pr/r bindings
%   for these OPC namespaces (it binds `r` to the OFFICE relationships URI too,
%   but has no ct/pr at all). The OPC CT_* classes and their qn() lookups
%   (Relationship_lst / defaults / overrides) resolve `ct:`/`pr:` ONLY through
%   this map -- never mat2doc.oxml.qn (which would KeyError on ct/pr). H8.
%
%     ct -> package content-types URI
%     pr -> package relationships URI (the .rels root namespace)
%     r  -> officeDocument relationships URI
%
%   Outputs: map - 1x1 struct, <prefix> -> namespace URI (string scalar)
%
%   Example:
%       m = mat2doc.opc.oxml.nsmap();
%       m.ct   % "http://schemas.openxmlformats.org/package/2006/content-types"
%
%   Ported from python-docx v1.2.0: src/docx/opc/oxml.py::nsmap (lines 24-28)

persistent cached
if isempty(cached)
    NS = mat2doc.opc.NAMESPACE;
    m = struct();
    % Field assignment order reproduces opc/oxml.py lines 25-27 verbatim (H11).
    m.ct = NS.OPC_CONTENT_TYPES;
    m.pr = NS.OPC_RELATIONSHIPS;
    m.r = NS.OFC_RELATIONSHIPS;
    cached = m;
end
map = cached;
end

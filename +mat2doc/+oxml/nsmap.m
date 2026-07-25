function map = nsmap()
% NSMAP Map namespace prefix to namespace name for the WordprocessingML namespaces.
%
%   map = MAT2DOC.OXML.NSMAP() returns a scalar struct whose field names are
%   the namespace prefixes and whose values are the corresponding namespace
%   URIs (string scalars). Field order preserves the Python source dict's
%   insertion order exactly (H11) -- iterate with fieldnames(map).
%
%   Outputs: map - 1x1 struct, <prefix> -> namespace URI (string scalar)
%
%   REPRESENTATION (H11): Python `nsmap` is a module-level dict; here it is a
%   scalar struct with insertion-ordered fields. It is indexed as a whole map
%   (nsmap[prefix] in Python -> map.(prefix) here). This is the PUBLIC docx
%   nsmap (unlike pptx, whose module-private `_nsmap` was rotated to nsmap_);
%   docx names it publicly, so this file is nsmap.m with no rotation and there
%   is no separate subset alias (the subset function is nspfxmap, see
%   nspfxmap.m). This is the FIXED map used by qn / nsdecls /
%   NamespacePrefixedTag / createElement / OxmlElement for the generation path.
%
%   DELTA FROM pptx (source of truth = docx ns.py): the docx map has 16
%   entries -- it ADDS dgm, w14 and rebinds sl to the schemaLibrary namespace,
%   and DROPS the many pptx-only prefixes (ct, ep, i, mo, mv, o, p, pd, pr, v,
%   ve, w10, wne). The OPC prefixes ct/pr/r used by [Content_Types].xml and
%   .rels live in a SEPARATE map in docx/opc/oxml.py (the OPC-layer WP), NOT
%   here.
%
%   Example:
%       map = mat2doc.oxml.nsmap();
%       map.w                  % "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
%       string(fieldnames(map))'   % all 16 prefixes in insertion order
%
%   Ported from python-docx v1.2.0: src/docx/oxml/ns.py::nsmap

persistent cached
if isempty(cached)
    m = struct();
    % Field assignment order below reproduces ns.py lines 7-24 verbatim (H11).
    m.a = "http://schemas.openxmlformats.org/drawingml/2006/main";
    m.c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    m.cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    m.dc = "http://purl.org/dc/elements/1.1/";
    m.dcmitype = "http://purl.org/dc/dcmitype/";
    m.dcterms = "http://purl.org/dc/terms/";
    m.dgm = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
    m.m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    m.pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    m.r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    m.sl = "http://schemas.openxmlformats.org/schemaLibrary/2006/main";
    m.w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    m.w14 = "http://schemas.microsoft.com/office/word/2010/wordml";
    m.wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    m.xml = "http://www.w3.org/XML/1998/namespace";
    m.xsi = "http://www.w3.org/2001/XMLSchema-instance";
    cached = m;
end
map = cached;
end

function uri = CONTENT_TYPES_URI()
% CONTENT_TYPES_URI The pack URI of the content-types stream.
%
%   uri = MAT2DOC.OPC.CONTENT_TYPES_URI() returns PackURI("/[Content_Types].xml").
%   Mirrors the module constant CONTENT_TYPES_URI = PackURI("/[Content_Types].xml").
%
%   Example:
%       disp(string(mat2doc.opc.CONTENT_TYPES_URI()))   % "/[Content_Types].xml"
%
%   Ported from python-docx v1.2.0: src/docx/opc/packuri.py::CONTENT_TYPES_URI
%   (line 109)

uri = mat2doc.opc.PackURI("/[Content_Types].xml");
end

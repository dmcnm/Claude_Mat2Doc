classdef NAMESPACE
% NAMESPACE Constant values for OPC XML namespaces.
%
%   Access as mat2doc.opc.NAMESPACE.OPC_RELATIONSHIPS etc. Mirrors the Python
%   class attributes (class NAMESPACE with class-level constants). These are the
%   OPC-layer namespace URIs used by [Content_Types].xml and .rels; the
%   OPC-local prefix map {ct, pr, r} in +opc\+oxml\nsmap.m binds prefixes to
%   three of these URIs.
%
%   Example:
%       disp(mat2doc.opc.NAMESPACE.OPC_CONTENT_TYPES)
%       % "http://schemas.openxmlformats.org/package/2006/content-types"
%
%   Ported from python-docx v1.2.0: src/docx/opc/constants.py::NAMESPACE
%   (lines 159-168)

    properties (Constant)
        DML_WORDPROCESSING_DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        OFC_RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        OPC_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships"
        OPC_CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
        WML_MAIN = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    end
end

classdef XsdString < mat2doc.oxml.simpletypes.BaseStringType
% XSDSTRING The xsd:string simple type (identity string transform).
%
%   Inherits from_xml / to_xml / convert_* / validate from BaseStringType
%   unchanged (Python `class XsdString(BaseStringType): pass`). Base for
%   ST_String / ST_RelationshipId and the inline-valid-values string
%   enumerations (ST_BrClear, ST_BrType, ST_TblLayoutType, ST_TblWidth).
%
%   Example:
%       disp(mat2doc.oxml.simpletypes.XsdString.to_xml("Heading1"))
%       % "Heading1" (identity string transform)
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::XsdString
%   (lines 154-155)
end

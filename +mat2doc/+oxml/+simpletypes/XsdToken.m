classdef XsdToken < mat2doc.oxml.simpletypes.BaseStringType
% XSDTOKEN The xsd:token simple type (identity string transform).
%
%   Inherits from_xml / to_xml / convert_* / validate from BaseStringType
%   unchanged (Python `class XsdToken(BaseStringType)` with only a docstring
%   describing xsd:token whitespace collapsing; the source implements NO
%   collapsing override, so the transform is plain identity, exactly as
%   XsdString).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::XsdToken
%   (lines 162-166)
end

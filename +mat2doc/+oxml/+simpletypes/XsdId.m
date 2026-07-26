classdef XsdId < mat2doc.oxml.simpletypes.BaseStringType
% XSDID The xsd:ID simple type (identity string transform).
%
%   Inherits from_xml / to_xml / convert_* / validate from BaseStringType
%   unchanged (Python `class XsdId(BaseStringType)` with only a docstring:
%   "String that must begin with a letter or underscore and cannot contain any
%   colons. Not fully validated because not used in external API"). Resolved
%   by the typed-attr engine via feval.
%
%   Example:
%       disp(mat2doc.oxml.simpletypes.XsdId.to_xml("_x1"))   % "_x1"  (identity)
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::XsdId
%   (lines 133-139)
end

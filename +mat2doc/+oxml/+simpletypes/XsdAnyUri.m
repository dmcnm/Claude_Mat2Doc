classdef XsdAnyUri < mat2doc.oxml.simpletypes.BaseStringType
% XSDANYURI The xsd:anyURI simple type (identity string transform).
%
%   Inherits from_xml / to_xml / convert_* / validate from BaseStringType
%   unchanged (Python `class XsdAnyUri(BaseStringType)` with only a docstring:
%   the schema regex is deliberately NOT validated -- "spending cycles on
%   validating wouldn't be worth it for the number of programming errors it
%   would catch"). Resolved by the typed-attr engine via feval.
%
%   Example:
%       disp(mat2doc.oxml.simpletypes.XsdAnyUri.to_xml("http://x/y"))
%       % "http://x/y"  (identity)
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::XsdAnyUri
%   (lines 104-109)
end

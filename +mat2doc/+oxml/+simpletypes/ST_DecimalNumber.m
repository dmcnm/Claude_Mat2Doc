classdef ST_DecimalNumber < mat2doc.oxml.simpletypes.XsdInt
% ST_DECIMALNUMBER Decimal-number simple type (xsd:int).
%
%   Inherits from_xml / to_xml / convert_* / validate unchanged (Python
%   `class ST_DecimalNumber(XsdInt): pass`). The inherited XsdInt.to_xml ->
%   XsdInt.validate (range -2147483648..2147483647) is exactly this type's
%   behavior (H10: safe to inherit when no convert/validate is overridden).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::ST_DecimalNumber
%   (lines 269-270)
end

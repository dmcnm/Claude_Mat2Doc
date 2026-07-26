classdef ST_DrawingElementId < mat2doc.oxml.simpletypes.XsdUnsignedInt
% ST_DRAWINGELEMENTID Drawing-element id simple type (xsd:unsignedInt).
%
%   Inherits from_xml / to_xml / convert_* / validate unchanged (Python
%   `class ST_DrawingElementId(XsdUnsignedInt): pass`). The inherited
%   XsdUnsignedInt.to_xml -> XsdUnsignedInt.validate (range 0..4294967295) is
%   exactly this type's behavior (H10: safe to inherit when no convert/validate
%   is overridden).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::ST_DrawingElementId
%   (lines 273-274)
end

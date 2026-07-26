classdef XsdInt < mat2doc.oxml.simpletypes.BaseIntType
% XSDINT The xsd:int simple type: signed 32-bit integer.
%
%   Overrides validate with the signed 32-bit range. Because it overrides
%   validate, it re-declares to_xml so the BaseSimpleType.to_xml template
%   routes to THIS validate (H10, no classmethod late binding). from_xml /
%   convert_from_xml / convert_to_xml are inherited from BaseIntType.
%
%   Example:
%       disp(mat2doc.oxml.simpletypes.XsdInt.to_xml(42))   % "42"
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::XsdInt
%   (lines 142-145)

    methods (Static)
        function s = to_xml(value)
            % TO_XML BaseSimpleType.to_xml (lines 29-33), bound to THIS validate.
            mat2doc.oxml.simpletypes.XsdInt.validate(value);
            s = mat2doc.oxml.simpletypes.XsdInt.convert_to_xml(value);
        end

        function validate(value)
            % VALIDATE lines 143-145: int in range -2147483648..2147483647.
            mat2doc.oxml.simpletypes.BaseSimpleType.validate_int_in_range( ...
                value, -2147483648, 2147483647);
        end
    end
end

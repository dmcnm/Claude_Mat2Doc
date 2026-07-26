classdef XsdUnsignedInt < mat2doc.oxml.simpletypes.BaseIntType
% XSDUNSIGNEDINT The xsd:unsignedInt simple type: 0..4294967295.
%
%   Overrides validate; re-declares to_xml so the template routes to THIS
%   validate (H10). from_xml / convert_* inherited from BaseIntType.
%
%   Example:
%       disp(mat2doc.oxml.simpletypes.XsdUnsignedInt.to_xml(7))   % "7"
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::XsdUnsignedInt
%   (lines 169-172)

    methods (Static)
        function s = to_xml(value)
            % TO_XML BaseSimpleType.to_xml (lines 29-33), bound to THIS validate.
            mat2doc.oxml.simpletypes.XsdUnsignedInt.validate(value);
            s = mat2doc.oxml.simpletypes.XsdUnsignedInt.convert_to_xml(value);
        end

        function validate(value)
            % VALIDATE lines 170-172: int in range 0..4294967295.
            mat2doc.oxml.simpletypes.BaseSimpleType.validate_int_in_range( ...
                value, 0, 4294967295);
        end
    end
end

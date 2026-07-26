classdef ST_CoordinateUnqualified < mat2doc.oxml.simpletypes.XsdLong
% ST_COORDINATEUNQUALIFIED EMU coordinate range: -27273042329600..27273042316900.
%
%   Overrides validate with the coordinate range (well within 2^53, so exact
%   in a double); re-declares to_xml so the template routes to THIS validate
%   (H10). from_xml / convert_* inherited from XsdLong / BaseIntType (int
%   parse; str format).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::ST_CoordinateUnqualified
%   (lines 211-214)

    methods (Static)
        function s = to_xml(value)
            % TO_XML BaseSimpleType.to_xml (lines 29-33), bound to THIS validate.
            mat2doc.oxml.simpletypes.ST_CoordinateUnqualified.validate(value);
            s = mat2doc.oxml.simpletypes.ST_CoordinateUnqualified.convert_to_xml(value);
        end

        function validate(value)
            % VALIDATE lines 212-214: int in range -27273042329600..27273042316900.
            mat2doc.oxml.simpletypes.BaseSimpleType.validate_int_in_range( ...
                value, -27273042329600, 27273042316900);
        end
    end
end

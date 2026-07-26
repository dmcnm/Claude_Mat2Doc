classdef ST_PositiveCoordinate < mat2doc.oxml.simpletypes.XsdLong
% ST_POSITIVECOORDINATE Non-negative EMU coordinate: 0..27273042316900.
%
%   convert_from_xml parses the long and wraps it in Emu; validate enforces
%   0..27273042316900. Overrides convert_from_xml + validate, so re-declares
%   from_xml + to_xml (H10). convert_to_xml (str) inherited from BaseIntType.
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::ST_PositiveCoordinate
%   (lines 347-354)

    methods (Static)
        function v = from_xml(xml_value)
            % FROM_XML BaseSimpleType.from_xml (lines 25-27): convert_from_xml.
            v = mat2doc.oxml.simpletypes.ST_PositiveCoordinate.convert_from_xml(xml_value);
        end

        function s = to_xml(value)
            % TO_XML BaseSimpleType.to_xml (lines 29-33), bound to THIS validate.
            mat2doc.oxml.simpletypes.ST_PositiveCoordinate.validate(value);
            s = mat2doc.oxml.simpletypes.ST_PositiveCoordinate.convert_to_xml(value);
        end

        function emu_value = convert_from_xml(str_value)
            % CONVERT_FROM_XML lines 349-350: Emu(int(str_value)) -- the super
            %   (XsdLong/BaseIntType) parse wrapped in Emu (H6).
            emu_value = mat2doc.shared.Emu(str_value);
        end

        function validate(value)
            % VALIDATE lines 352-354: int in range 0..27273042316900.
            mat2doc.oxml.simpletypes.BaseSimpleType.validate_int_in_range( ...
                value, 0, 27273042316900);
        end
    end
end

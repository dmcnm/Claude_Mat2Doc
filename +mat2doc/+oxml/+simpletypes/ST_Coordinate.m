classdef ST_Coordinate < mat2doc.oxml.simpletypes.BaseIntType
% ST_COORDINATE Coordinate: a bare EMU integer, or a measure with a unit suffix.
%
%   Extends BaseIntType. convert_from_xml detects a unit suffix by the presence
%   of 'i', 'm' or 'p' (every unit -- in, cm, mm, pt, pc, pi -- contains one;
%   a bare EMU integer contains none) and defers to ST_UniversalMeasure,
%   otherwise wraps Emu(int(str_value)). Both branches return an Emu (Length).
%   validate delegates to ST_CoordinateUnqualified.validate (the EMU coordinate
%   range). Overrides convert_from_xml + validate, so re-declares from_xml +
%   to_xml (H10); convert_to_xml (str) is inherited from BaseIntType.
%
%   Example:
%       double(mat2doc.oxml.simpletypes.ST_Coordinate.from_xml("914400"))  % 914400
%       double(mat2doc.oxml.simpletypes.ST_Coordinate.from_xml("1in"))     % 914400
%       mat2doc.oxml.simpletypes.ST_Coordinate.to_xml(mat2doc.shared.Emu(914400)) % "914400"
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::ST_Coordinate
%   (lines 199-208)

    methods (Static)
        function v = from_xml(xml_value)
            % FROM_XML BaseSimpleType.from_xml (lines 25-27): convert_from_xml.
            v = mat2doc.oxml.simpletypes.ST_Coordinate.convert_from_xml(xml_value);
        end

        function s = to_xml(value)
            % TO_XML BaseSimpleType.to_xml (lines 29-33), bound to THIS validate
            %   + the inherited BaseIntType.convert_to_xml (str).
            mat2doc.oxml.simpletypes.ST_Coordinate.validate(value);
            s = mat2doc.oxml.simpletypes.ST_Coordinate.convert_to_xml(value);
        end

        function v = convert_from_xml(str_value)
            % CONVERT_FROM_XML lines 201-204: unit suffix ('i'/'m'/'p') ->
            %   ST_UniversalMeasure; else Emu(int(str_value)) (H6).
            sv = string(str_value);
            if contains(sv, "i") || contains(sv, "m") || contains(sv, "p")
                v = mat2doc.oxml.simpletypes.ST_UniversalMeasure.convert_from_xml(str_value);
            else
                v = mat2doc.shared.Emu(str_value);
            end
        end

        function validate(value)
            % VALIDATE lines 206-208: ST_CoordinateUnqualified.validate(value).
            mat2doc.oxml.simpletypes.ST_CoordinateUnqualified.validate(value);
        end
    end
end

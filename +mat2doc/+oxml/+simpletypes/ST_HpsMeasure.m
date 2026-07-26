classdef ST_HpsMeasure < mat2doc.oxml.simpletypes.XsdUnsignedLong
% ST_HPSMEASURE Half-point measure, e.g. 24 represents 12.0 points -> Length.
%
%   convert_from_xml: if the string contains 'm', 'n' or 'p' (a unit suffix:
%   mm/cm both contain 'm', in contains 'n', pt/pc/pi contain 'p') defer to
%   ST_UniversalMeasure; else the bare value is a half-point count, so the
%   length is Pt(int(str_value) / 2.0). convert_to_xml: half_points =
%   int(Emu(value).pt * 2), serialized as str. Overrides convert_from_xml +
%   convert_to_xml, so re-declares from_xml + to_xml (H10). validate is
%   inherited from XsdUnsignedLong (0..2^64-1, D-STYPE-3 dead upper edge).
%
%   NOTE the unit-detection letter set is 'm'/'n'/'p' here (NOT the 'i'/'m'/'p'
%   used by ST_Coordinate / the twips measures) -- ported verbatim from the
%   source (simpletypes.py 316).
%
%   int() truncates toward zero -> fix() (H6); the /2.0 and Emu.pt divisions
%   are true float divisions; str() via pyStr "int" (H14).
%
%   Example:
%       double(mat2doc.oxml.simpletypes.ST_HpsMeasure.from_xml("24"))   % 152400 (12 pt)
%       mat2doc.oxml.simpletypes.ST_HpsMeasure.to_xml(mat2doc.shared.Pt(12)) % "24"
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::ST_HpsMeasure
%   (lines 311-324)

    methods (Static)
        function v = from_xml(xml_value)
            % FROM_XML BaseSimpleType.from_xml (lines 25-27): convert_from_xml.
            v = mat2doc.oxml.simpletypes.ST_HpsMeasure.convert_from_xml(xml_value);
        end

        function s = to_xml(value)
            % TO_XML BaseSimpleType.to_xml (lines 29-33): validate (inherited
            %   XsdUnsignedLong); convert_to_xml (THIS class).
            mat2doc.oxml.simpletypes.ST_HpsMeasure.validate(value);
            s = mat2doc.oxml.simpletypes.ST_HpsMeasure.convert_to_xml(value);
        end

        function v = convert_from_xml(str_value)
            % CONVERT_FROM_XML lines 315-318: unit suffix ('m'/'n'/'p') ->
            %   ST_UniversalMeasure; else Pt(int(str_value) / 2.0).
            sv = string(str_value);
            if contains(sv, "m") || contains(sv, "n") || contains(sv, "p")
                v = mat2doc.oxml.simpletypes.ST_UniversalMeasure.convert_from_xml(str_value);
            else
                v = mat2doc.shared.Pt(intFromXml(str_value) / 2.0);
            end
        end

        function s = convert_to_xml(value)
            % CONVERT_TO_XML lines 320-324: emu = Emu(value); half_points =
            %   int(emu.pt * 2); str(half_points). int() -> fix() (H6).
            emu = mat2doc.shared.Emu(value);
            half_points = fix(emu.pt * 2);
            s = mat2doc.shared.pyStr(half_points, "int");
        end
    end
end

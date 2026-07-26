classdef ST_TwipsMeasure < mat2doc.oxml.simpletypes.XsdUnsignedLong
% ST_TWIPSMEASURE Twips measure (1/20 point) -> Length.
%
%   convert_from_xml: if the string contains 'i', 'm' or 'p' (a unit suffix)
%   defer to ST_UniversalMeasure; else the bare value is a twip count, so the
%   length is Twips(int(str_value)). convert_to_xml: Emu(value).twips,
%   serialized as str. Overrides convert_from_xml + convert_to_xml, so
%   re-declares from_xml + to_xml (H10). validate inherited from
%   XsdUnsignedLong (0..2^64-1, D-STYPE-3 dead upper edge).
%
%   int(str_value) is applied FIRST (via intFromXml) yielding a plain int,
%   THEN Twips(int) -- matching Python `Twips(int(str_value))`. A raw string is
%   never handed to Twips (whose string mode rejects, D-002/D-003); the
%   pre-int()ed numeric is the live path. convert_to_xml uses Length.twips
%   (int(round(emu/635)), Python round-half-to-even, H6/H14).
%
%   Example:
%       double(mat2doc.oxml.simpletypes.ST_TwipsMeasure.from_xml("1440"))  % 914400 (1 in)
%       mat2doc.oxml.simpletypes.ST_TwipsMeasure.to_xml(mat2doc.shared.Twips(1440)) % "1440"
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::ST_TwipsMeasure
%   (lines 397-408)

    methods (Static)
        function v = from_xml(xml_value)
            % FROM_XML BaseSimpleType.from_xml (lines 25-27): convert_from_xml.
            v = mat2doc.oxml.simpletypes.ST_TwipsMeasure.convert_from_xml(xml_value);
        end

        function s = to_xml(value)
            % TO_XML BaseSimpleType.to_xml (lines 29-33): validate (inherited
            %   XsdUnsignedLong); convert_to_xml (THIS class).
            mat2doc.oxml.simpletypes.ST_TwipsMeasure.validate(value);
            s = mat2doc.oxml.simpletypes.ST_TwipsMeasure.convert_to_xml(value);
        end

        function v = convert_from_xml(str_value)
            % CONVERT_FROM_XML lines 399-402: unit suffix ('i'/'m'/'p') ->
            %   ST_UniversalMeasure; else Twips(int(str_value)).
            sv = string(str_value);
            if contains(sv, "i") || contains(sv, "m") || contains(sv, "p")
                v = mat2doc.oxml.simpletypes.ST_UniversalMeasure.convert_from_xml(str_value);
            else
                v = mat2doc.shared.Twips(intFromXml(str_value));
            end
        end

        function s = convert_to_xml(value)
            % CONVERT_TO_XML lines 404-408: emu = Emu(value); twips = emu.twips;
            %   str(twips). Length.twips is int(round(emu/635)) (H6/H14).
            emu = mat2doc.shared.Emu(value);
            twips = emu.twips;
            s = mat2doc.shared.pyStr(twips, "int");
        end
    end
end

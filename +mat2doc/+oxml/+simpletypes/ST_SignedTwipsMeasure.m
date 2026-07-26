classdef ST_SignedTwipsMeasure < mat2doc.oxml.simpletypes.XsdInt
% ST_SIGNEDTWIPSMEASURE Signed twips measure (1/20 point) -> Length.
%
%   convert_from_xml: if the string contains 'i', 'm' or 'p' (a unit suffix)
%   defer to ST_UniversalMeasure; else the bare value is a (possibly
%   fractional/signed) twip count, so the length is
%   Twips(int(round(float(str_value)))). convert_to_xml: Emu(value).twips,
%   serialized as str. Overrides convert_from_xml + convert_to_xml, so
%   re-declares from_xml + to_xml (H10). validate inherited from XsdInt
%   (signed 32-bit range).
%
%   The bare branch: float(str_value) parses a decimal (str2double, the same
%   substitution the ST_UniversalMeasure float parse uses -- D-STYPE-2 / D-002,
%   no new D-number); round() is Python round-half-to-even (pyRound, H6);
%   int() of the integral rounded value is exact; Twips(that_int) applies the
%   635-EMU multiplier. int(str) is applied FIRST so a raw string never reaches
%   Twips (whose string mode rejects). convert_to_xml uses Length.twips.
%
%   Example:
%       double(mat2doc.oxml.simpletypes.ST_SignedTwipsMeasure.from_xml("-360"))  % -228600
%       mat2doc.oxml.simpletypes.ST_SignedTwipsMeasure.to_xml(mat2doc.shared.Twips(-360)) % "-360"
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::ST_SignedTwipsMeasure
%   (lines 361-372)

    methods (Static)
        function v = from_xml(xml_value)
            % FROM_XML BaseSimpleType.from_xml (lines 25-27): convert_from_xml.
            v = mat2doc.oxml.simpletypes.ST_SignedTwipsMeasure.convert_from_xml(xml_value);
        end

        function s = to_xml(value)
            % TO_XML BaseSimpleType.to_xml (lines 29-33): validate (inherited
            %   XsdInt); convert_to_xml (THIS class).
            mat2doc.oxml.simpletypes.ST_SignedTwipsMeasure.validate(value);
            s = mat2doc.oxml.simpletypes.ST_SignedTwipsMeasure.convert_to_xml(value);
        end

        function v = convert_from_xml(str_value)
            % CONVERT_FROM_XML lines 363-366: unit suffix ('i'/'m'/'p') ->
            %   ST_UniversalMeasure; else Twips(int(round(float(str_value)))).
            sv = string(str_value);
            if contains(sv, "i") || contains(sv, "m") || contains(sv, "p")
                v = mat2doc.oxml.simpletypes.ST_UniversalMeasure.convert_from_xml(str_value);
            else
                v = mat2doc.shared.Twips(pyRound(str2double(sv)));
            end
        end

        function s = convert_to_xml(value)
            % CONVERT_TO_XML lines 368-372: emu = Emu(value); twips = emu.twips;
            %   str(twips). Length.twips is int(round(emu/635)) (H6/H14).
            emu = mat2doc.shared.Emu(value);
            twips = emu.twips;
            s = mat2doc.shared.pyStr(twips, "int");
        end
    end
end

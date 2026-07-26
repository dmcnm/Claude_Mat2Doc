classdef ST_HexColor < mat2doc.oxml.simpletypes.BaseStringType
% ST_HEXCOLOR `w:color/@val` value: an RGBColor, or the literal "auto".
%
%   docx SPLIT (vs the python-pptx ST_HexColorRGB precedent -- verified
%   distinct): python-pptx had a single ST_HexColorRGB that took/emitted a
%   six-hex-digit STRING. python-docx splits the concept in two:
%     * ST_HexColor (THIS class) -- convert_from_xml returns an RGBColor
%       OBJECT for a hex value, or the string "auto" (ST_HexColorAuto.AUTO)
%       for the literal "auto"; convert_to_xml formats an RGBColor via
%       "%02X%02X%02X"; validate requires an RGBColor.
%     * ST_HexColorAuto -- a separate one-member XsdStringEnumeration ("auto")
%       supplying the AUTO constant this class returns.
%   So the port mirrors the docx pair exactly, NOT the pptx single class.
%
%   D-STYPE-4 (adopted): the hex acceptance boundary. python-docx defers hex
%   parsing to RGBColor.from_string, whose parseHexByte_ accepts exactly the
%   six-hex-digit contract [0-9A-Fa-f]{6} (P1-1 shared/RGBColor.m) -- the same
%   RGB contract, identical to Python int(_, 16) for every real colour; the
%   extra int(str,16) acceptances (sign, 0x prefix, underscores, surrounding
%   space) are non-RGB programming errors, folded under the adopted D-STYPE-4
%   ruling (no new D-number). NOTE: this class does NOT range/length-check the
%   string itself -- it just calls RGBColor.from_string, so an over-long or
%   under-long string raises whatever RGBColor.from_string raises (faithful:
%   python-docx has no length guard here either -- contrast the pptx
%   ST_HexColorRGB which DID length-guard; do not port that guard).
%
%   Overrides convert_from_xml + convert_to_xml + validate, so re-declares
%   from_xml + to_xml (H10).
%
%   Example:
%       c = mat2doc.oxml.simpletypes.ST_HexColor.from_xml("3C2F80");  % RGBColor
%       mat2doc.oxml.simpletypes.ST_HexColor.from_xml("auto")         % "auto"
%       mat2doc.oxml.simpletypes.ST_HexColor.to_xml(c)                % "3C2F80"
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::ST_HexColor
%   (lines 277-300)

    methods (Static)
        function v = from_xml(xml_value)
            % FROM_XML BaseSimpleType.from_xml (lines 25-27): convert_from_xml.
            v = mat2doc.oxml.simpletypes.ST_HexColor.convert_from_xml(xml_value);
        end

        function s = to_xml(value)
            % TO_XML BaseSimpleType.to_xml (lines 29-33), bound to THIS
            %   validate + convert_to_xml.
            mat2doc.oxml.simpletypes.ST_HexColor.validate(value);
            s = mat2doc.oxml.simpletypes.ST_HexColor.convert_to_xml(value);
        end

        function v = convert_from_xml(str_value)
            % CONVERT_FROM_XML lines 279-284: "auto" -> ST_HexColorAuto.AUTO
            %   (the string "auto"); else RGBColor.from_string(str_value).
            if string(str_value) == "auto"
                v = mat2doc.oxml.simpletypes.ST_HexColorAuto.AUTO;
            else
                v = mat2doc.shared.RGBColor.from_string(str_value);
            end
        end

        function s = convert_to_xml(value)
            % CONVERT_TO_XML lines 286-292: "%02X%02X%02X" % value -- the
            %   RGBColor 3-tuple formatted as fixed-width UPPERCASE hex (H14
            %   does not apply: this is %X hex, not a decimal serialization).
            s = string(sprintf("%02X%02X%02X", value.r, value.g, value.b));
        end

        function validate(value)
            % VALIDATE lines 294-300: must be an RGBColor object, else
            %   ValueError "rgb color value must be RGBColor object, got %s %s"
            %   % (type(value), value). type(value) -> MATLAB class token
            %   (D-005); value -> best-effort str().
            if ~isa(value, "mat2doc.shared.RGBColor")
                error("mat2doc:ValueError", ...
                    "rgb color value must be RGBColor object, got %s %s", ...
                    class(value), ...
                    mat2doc.oxml.simpletypes.ST_HexColor.valueRepr_(value));
            end
        end
    end

    methods (Static, Access = private)
        function r = valueRepr_(value)
            % Best-effort str(value) for the ValueError message. pyStr covers
            % the realistic mistaken inputs (string/char/logical/numeric/
            % Length); anything else falls back to the class token (dead path,
            % D-005).
            try
                r = mat2doc.shared.pyStr(value);
            catch
                r = string(class(value));
            end
        end
    end
end

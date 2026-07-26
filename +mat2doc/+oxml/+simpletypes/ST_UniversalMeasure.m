classdef ST_UniversalMeasure < mat2doc.oxml.simpletypes.BaseSimpleType
% ST_UNIVERSALMEASURE Measurement with a unit suffix -> Emu (Length).
%
%   Parses a number followed by a two-letter unit (mm, cm, in, pt, pc, pi)
%   into an Emu using the fixed EMU-per-unit multipliers. Read-only in the
%   source: only convert_from_xml is defined (used by ST_Coordinate,
%   ST_HpsMeasure, ST_TwipsMeasure, ST_SignedTwipsMeasure), so this port
%   provides convert_from_xml plus a from_xml template routing to it (H10).
%   Python has no convert_to_xml / validate here, so to_xml is intentionally
%   undefined (calling it would fail, as it does in Python).
%
%   D-STYPE-2 RE-HOME (adopted): python-pptx housed float parsing in a
%   BaseFloatType base; python-docx has NO BaseFloatType. The ONLY float parse
%   in the docx simpletypes tier is this class's `quantity = float(float_part)`
%   (simpletypes.py 415). The float handling is therefore ported HERE per the
%   D-STYPE-2 re-home ruling -- no BaseFloatType class is created for docx.
%   float() is realized with str2double (D-002 note: str2double accepts a
%   slightly wider lexical set than CPython float(); the inputs are well-formed
%   XML measure literals, and this is the same substitution the Mat2Ppt
%   ST_UniversalMeasure precedent used, folded under D-STYPE-2 / D-002 with no
%   new D-number).
%
%   quantity = float(number) via str2double; multiplier lookup via switch
%   (a fixed dict; order is irrelevant, so no H11 concern); Emu wraps
%   int(round(quantity * multiplier)) -- round is Python round-half-to-even
%   (pyRound, H6), int() -> fix() is a no-op on the integral rounded value.
%   An unknown unit raises mat2doc:KeyError (Python KeyError from the dict
%   subscript).
%
%   Example:
%       double(mat2doc.oxml.simpletypes.ST_UniversalMeasure.convert_from_xml("1in"))
%       % 914400
%       double(mat2doc.oxml.simpletypes.ST_UniversalMeasure.convert_from_xml("2.5cm"))
%       % 900000
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::ST_UniversalMeasure
%   (lines 411-424)

    methods (Static)
        function v = from_xml(xml_value)
            % FROM_XML BaseSimpleType.from_xml (lines 25-27): convert_from_xml.
            v = mat2doc.oxml.simpletypes.ST_UniversalMeasure.convert_from_xml(xml_value);
        end

        function emu_value = convert_from_xml(str_value)
            % CONVERT_FROM_XML lines 413-424: split trailing 2-char unit, parse
            %   quantity (float, D-STYPE-2 re-home), multiply, round-half-to-even,
            %   wrap in Emu.
            sv = char(str_value);
            float_part = sv(1:end - 2);       % str_value[:-2]
            units_part = string(sv(end - 1:end));  % str_value[-2:]
            quantity = str2double(float_part);
            switch units_part
                case "mm"
                    multiplier = 36000;
                case "cm"
                    multiplier = 360000;
                case "in"
                    multiplier = 914400;
                case "pt"
                    multiplier = 12700;
                case "pc"
                    multiplier = 152400;
                case "pi"
                    multiplier = 152400;
                otherwise
                    error("mat2doc:KeyError", "'%s'", units_part);
            end
            emu_value = mat2doc.shared.Emu(pyRound(quantity * multiplier));
        end
    end
end

classdef BaseIntType < mat2doc.oxml.simpletypes.BaseSimpleType
% BASEINTTYPE Simple-type base for integer-valued XML attributes.
%
%   Ports BaseSimpleType.from_xml / to_xml (simpletypes.py 25-33) at this
%   branch base because MATLAB static methods have no Python classmethod late
%   binding (see BaseSimpleType header, H10). A subclass that OVERRIDES
%   convert_from_xml must re-declare from_xml; one that overrides validate or
%   convert_to_xml must re-declare to_xml, so the routed convert/validate is
%   the subclass's own.
%
%   convert_from_xml parses the XML integer literal exactly as Python
%   int(str_value) via the package-private intFromXml helper (H6, D-002 ASCII
%   grammar). convert_to_xml formats via pyStr "int" -- Python str(value) with
%   no decimal point (H14). validate is validate_int (D-STYPE-1).
%
%   Example:
%       T = "mat2doc.oxml.simpletypes.XsdInt";
%       disp(feval(T + ".from_xml", "42"))    % 42  (double)
%       disp(feval(T + ".to_xml",   42))      % "42"
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::BaseIntType
%   (lines 66-77) + BaseSimpleType.from_xml/to_xml (lines 25-33)

    methods (Static)
        function v = from_xml(xml_value)
            % FROM_XML BaseSimpleType.from_xml (lines 25-27): convert_from_xml.
            v = mat2doc.oxml.simpletypes.BaseIntType.convert_from_xml(xml_value);
        end

        function s = to_xml(value)
            % TO_XML BaseSimpleType.to_xml (lines 29-33): validate; convert_to_xml.
            mat2doc.oxml.simpletypes.BaseIntType.validate(value);
            s = mat2doc.oxml.simpletypes.BaseIntType.convert_to_xml(value);
        end

        function n = convert_from_xml(str_value)
            % CONVERT_FROM_XML lines 67-69: int(str_value) (H6, via intFromXml).
            n = intFromXml(str_value);
        end

        function s = convert_to_xml(value)
            % CONVERT_TO_XML lines 71-73: str(value). pyStr "int" (H14); a
            %   Length value formats as Python int digits regardless of kind.
            s = mat2doc.shared.pyStr(value, "int");
        end

        function validate(value)
            % VALIDATE lines 75-77: validate_int(value).
            mat2doc.oxml.simpletypes.BaseSimpleType.validate_int(value);
        end
    end
end

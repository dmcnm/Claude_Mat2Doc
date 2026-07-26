classdef XsdBoolean < mat2doc.oxml.simpletypes.BaseSimpleType
% XSDBOOLEAN The xsd:boolean simple type: '1'/'0'/'true'/'false' <-> logical.
%
%   Extends BaseSimpleType directly (not a string/int branch), so it declares
%   its own from_xml / to_xml routing to its own convert / validate (H10, no
%   classmethod late binding). from_xml returns a logical; convert_to_xml maps
%   true->"1", false->"0". Base for ST_OnOff (which widens the accepted XML
%   token set to also allow 'on'/'off').
%
%   convert_from_xml raises the docx.exceptions.InvalidXmlError (NOT
%   ValueError) on an unrecognized token, via the canonical raiser
%   mat2doc.exc.InvalidXmlError -- matching python-docx which imports
%   InvalidXmlError from docx.exceptions here (simpletypes.py 15, raise at 116).
%
%   Example:
%       disp(mat2doc.oxml.simpletypes.XsdBoolean.from_xml("true"))  % 1 (logical)
%       disp(mat2doc.oxml.simpletypes.XsdBoolean.to_xml(false))     % "0"
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::XsdBoolean
%   (lines 112-130)

    methods (Static)
        function v = from_xml(xml_value)
            % FROM_XML BaseSimpleType.from_xml (lines 25-27): convert_from_xml.
            v = mat2doc.oxml.simpletypes.XsdBoolean.convert_from_xml(xml_value);
        end

        function s = to_xml(value)
            % TO_XML BaseSimpleType.to_xml (lines 29-33): validate; convert_to_xml.
            mat2doc.oxml.simpletypes.XsdBoolean.validate(value);
            s = mat2doc.oxml.simpletypes.XsdBoolean.convert_to_xml(value);
        end

        function v = convert_from_xml(str_value)
            % CONVERT_FROM_XML lines 113-119: one of '1','0','true','false',
            %   else InvalidXmlError; True for '1'/'true'. Message verbatim.
            sv = string(str_value);
            if ~any(sv == ["1", "0", "true", "false"])
                mat2doc.exc.InvalidXmlError( ...
                    "value must be one of '1', '0', 'true' or 'false', got '" + sv + "'");
            end
            v = any(sv == ["1", "true"]);
        end

        function s = convert_to_xml(value)
            % CONVERT_TO_XML lines 121-123: {True:'1', False:'0'}[value].
            if value
                s = "1";
            else
                s = "0";
            end
        end

        function validate(value)
            % VALIDATE lines 125-130: `if value not in (True, False)` raise
            %   TypeError. Python `in` uses ==, so it accepts True/False AND any
            %   numeric equal to 1 or 0 (1==True, 0==False); a non-numeric
            %   compares unequal and raises. Replicated: logical scalar, or a
            %   real numeric scalar equal to 0 or 1, passes. Message verbatim;
            %   the got-value renders via pyStr (str(value)).
            ok = (islogical(value) && isscalar(value)) || ...
                (isnumeric(value) && isscalar(value) && isreal(value) && ...
                    (value == 1 || value == 0));
            if ~ok
                error("mat2doc:TypeError", ...
                    "only True or False (and possibly None) may be assigned, got '%s'", ...
                    mat2doc.oxml.simpletypes.XsdBoolean.valueRepr_(value));
            end
        end
    end

    methods (Static, Access = private)
        function r = valueRepr_(value)
            % Best-effort str(value) for the validate TypeError message. pyStr
            % covers the realistic mistaken inputs (string/char/logical/
            % numeric/Length); anything else -- including [] (the None
            % analogue) -- falls back to the class token so the TypeError is
            % still raised with the faithful identifier + template (D-005;
            % Gate-2 fix: a bare pyStr call here leaked
            % mat2doc:pyStr:unsupportedType for [] / non-numeric inputs
            % instead of the TypeError).
            try
                r = mat2doc.shared.pyStr(value);
            catch
                r = string(class(value));
            end
        end
    end
end

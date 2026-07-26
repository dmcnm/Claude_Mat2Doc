classdef BaseStringType < mat2doc.oxml.simpletypes.BaseSimpleType
% BASESTRINGTYPE Simple-type base for string-valued XML attributes.
%
%   Ports BaseSimpleType.from_xml / to_xml (simpletypes.py 25-33) at this
%   branch base -- the level that owns the convert/validate they call --
%   because MATLAB static methods do not have Python's classmethod late
%   binding (see BaseSimpleType header). XsdString / XsdToken / XsdAnyUri /
%   XsdId extend this class with no overrides, so their from_xml / to_xml
%   resolve here and use these identity convert + validate_string, exactly
%   what Python's classmethod dispatch would produce.
%
%   Example:
%       T = "mat2doc.oxml.simpletypes.BaseStringType";
%       disp(feval(T + ".to_xml",   "rId7"))   % "rId7" (validate + identity)
%       disp(feval(T + ".from_xml", "rId7"))   % "rId7" (identity)
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::BaseStringType
%   (lines 80-91) + BaseSimpleType.from_xml/to_xml (lines 25-33)

    methods (Static)
        function v = from_xml(xml_value)
            % FROM_XML BaseSimpleType.from_xml (lines 25-27): convert_from_xml.
            v = mat2doc.oxml.simpletypes.BaseStringType.convert_from_xml(xml_value);
        end

        function s = to_xml(value)
            % TO_XML BaseSimpleType.to_xml (lines 29-33): validate; convert_to_xml.
            mat2doc.oxml.simpletypes.BaseStringType.validate(value);
            s = mat2doc.oxml.simpletypes.BaseStringType.convert_to_xml(value);
        end

        function v = convert_from_xml(str_value)
            % CONVERT_FROM_XML lines 81-83: identity (the string as-is).
            v = str_value;
        end

        function s = convert_to_xml(value)
            % CONVERT_TO_XML lines 85-87: identity (the value as-is).
            s = value;
        end

        function validate(value)
            % VALIDATE lines 89-91: validate_string(value).
            mat2doc.oxml.simpletypes.BaseSimpleType.validate_string(value);
        end
    end
end

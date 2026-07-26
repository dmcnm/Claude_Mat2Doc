classdef ST_BrType < mat2doc.oxml.simpletypes.XsdString
% ST_BRTYPE Valid values for `w:br/@w:type`: page, column, textWrapping.
%
%   XsdString subclass overriding validate with an inline valid-values tuple
%   raising ValueError (message identical to the enumeration base), delegated
%   to BaseStringEnumerationType.validate_enum. Re-declares to_xml so the
%   template routes to THIS validate (H10). from_xml / convert_* inherited
%   (identity).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::ST_BrType
%   (lines 190-196)

    methods (Static)
        function s = to_xml(value)
            % TO_XML BaseSimpleType.to_xml (lines 29-33), bound to THIS validate.
            mat2doc.oxml.simpletypes.ST_BrType.validate(value);
            s = mat2doc.oxml.simpletypes.ST_BrType.convert_to_xml(value);
        end

        function validate(value)
            % VALIDATE lines 191-196: validate_string; must be one of the tuple.
            mat2doc.oxml.simpletypes.BaseStringEnumerationType.validate_enum( ...
                value, ["page", "column", "textWrapping"]);
        end
    end
end

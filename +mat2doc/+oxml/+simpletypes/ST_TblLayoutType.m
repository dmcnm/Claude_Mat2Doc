classdef ST_TblLayoutType < mat2doc.oxml.simpletypes.XsdString
% ST_TBLLAYOUTTYPE Valid values for `w:tblLayout/@w:type`: fixed, autofit.
%
%   XsdString subclass overriding validate with an inline valid-values tuple
%   raising ValueError (message identical to the enumeration base), delegated
%   to BaseStringEnumerationType.validate_enum. Re-declares to_xml so the
%   template routes to THIS validate (H10). from_xml / convert_* inherited
%   (identity).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::ST_TblLayoutType
%   (lines 379-385)

    methods (Static)
        function s = to_xml(value)
            % TO_XML BaseSimpleType.to_xml (lines 29-33), bound to THIS validate.
            mat2doc.oxml.simpletypes.ST_TblLayoutType.validate(value);
            s = mat2doc.oxml.simpletypes.ST_TblLayoutType.convert_to_xml(value);
        end

        function validate(value)
            % VALIDATE lines 380-385: validate_string; must be one of the tuple.
            mat2doc.oxml.simpletypes.BaseStringEnumerationType.validate_enum( ...
                value, ["fixed", "autofit"]);
        end
    end
end

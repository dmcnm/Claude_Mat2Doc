classdef ST_TblWidth < mat2doc.oxml.simpletypes.XsdString
% ST_TBLWIDTH Valid values for `w:tblW/@w:type`: auto, dxa, nil, pct.
%
%   XsdString subclass overriding validate with an inline valid-values tuple
%   raising ValueError (message identical to the enumeration base), delegated
%   to BaseStringEnumerationType.validate_enum. Re-declares to_xml so the
%   template routes to THIS validate (H10). from_xml / convert_* inherited
%   (identity).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::ST_TblWidth
%   (lines 388-394)

    methods (Static)
        function s = to_xml(value)
            % TO_XML BaseSimpleType.to_xml (lines 29-33), bound to THIS validate.
            mat2doc.oxml.simpletypes.ST_TblWidth.validate(value);
            s = mat2doc.oxml.simpletypes.ST_TblWidth.convert_to_xml(value);
        end

        function validate(value)
            % VALIDATE lines 389-394: validate_string; must be one of the tuple.
            mat2doc.oxml.simpletypes.BaseStringEnumerationType.validate_enum( ...
                value, ["auto", "dxa", "nil", "pct"]);
        end
    end
end

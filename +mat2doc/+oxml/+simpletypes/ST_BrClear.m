classdef ST_BrClear < mat2doc.oxml.simpletypes.XsdString
% ST_BRCLEAR Valid values for `w:br/@w:clear`: none, left, right, all.
%
%   XsdString subclass that overrides validate with an inline valid-values
%   tuple raising ValueError (NOT the _members mechanism -- it is a plain
%   XsdString, not a BaseStringEnumerationType). The ValueError message is
%   identical to the enumeration base ("must be one of %s, got '%s'"), so
%   validate delegates to BaseStringEnumerationType.validate_enum with the
%   member list. Overriding validate, it re-declares to_xml so the template
%   routes to THIS validate (H10). from_xml / convert_* inherited (identity).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::ST_BrClear
%   (lines 181-187)

    methods (Static)
        function s = to_xml(value)
            % TO_XML BaseSimpleType.to_xml (lines 29-33), bound to THIS validate.
            mat2doc.oxml.simpletypes.ST_BrClear.validate(value);
            s = mat2doc.oxml.simpletypes.ST_BrClear.convert_to_xml(value);
        end

        function validate(value)
            % VALIDATE lines 182-187: validate_string; must be one of the tuple.
            mat2doc.oxml.simpletypes.BaseStringEnumerationType.validate_enum( ...
                value, ["none", "left", "right", "all"]);
        end
    end
end

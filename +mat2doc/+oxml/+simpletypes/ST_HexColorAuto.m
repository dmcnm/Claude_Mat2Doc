classdef ST_HexColorAuto < mat2doc.oxml.simpletypes.XsdStringEnumeration
% ST_HEXCOLORAUTO Value for `w:color/@val="auto"` (the single member "auto").
%
%   XsdStringEnumeration with one member, AUTO = "auto" (Python
%   `_members = (AUTO,)`, simpletypes.py 303-308). Referenced by
%   ST_HexColor.convert_from_xml (returns ST_HexColorAuto.AUTO for the literal
%   "auto"). validate delegates to BaseStringEnumerationType.validate_enum
%   with the singleton member list; to_xml re-declared so the template routes
%   to THIS validate (H10). from_xml / convert_* inherited (identity).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::ST_HexColorAuto
%   (lines 303-308)

    properties (Constant)
        AUTO = "auto"          % Python class attr AUTO = "auto"
        members_ = "auto"      % Python _members = (AUTO,); trailing-underscore rotation
    end

    methods (Static)
        function s = to_xml(value)
            % TO_XML BaseSimpleType.to_xml (lines 29-33), bound to THIS validate.
            mat2doc.oxml.simpletypes.ST_HexColorAuto.validate(value);
            s = mat2doc.oxml.simpletypes.ST_HexColorAuto.convert_to_xml(value);
        end

        function validate(value)
            % VALIDATE BaseStringEnumerationType.validate (lines 97-101) against members_.
            mat2doc.oxml.simpletypes.BaseStringEnumerationType.validate_enum( ...
                value, mat2doc.oxml.simpletypes.ST_HexColorAuto.members_);
        end
    end
end

classdef ST_Merge < mat2doc.oxml.simpletypes.XsdStringEnumeration
% ST_MERGE Valid values for `<w:vMerge/xMerge val="">`: continue, restart.
%
%   XsdStringEnumeration with members CONTINUE = "continue", RESTART =
%   "restart" (Python `_members = (CONTINUE, RESTART)`, simpletypes.py
%   327-333). validate delegates to BaseStringEnumerationType.validate_enum;
%   to_xml re-declared so the template routes to THIS validate (H10).
%   from_xml / convert_* inherited (identity).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::ST_Merge
%   (lines 327-333)

    properties (Constant)
        CONTINUE = "continue"                % Python class attr
        RESTART = "restart"                  % Python class attr
        members_ = ["continue", "restart"]   % Python _members = (CONTINUE, RESTART)
    end

    methods (Static)
        function s = to_xml(value)
            % TO_XML BaseSimpleType.to_xml (lines 29-33), bound to THIS validate.
            mat2doc.oxml.simpletypes.ST_Merge.validate(value);
            s = mat2doc.oxml.simpletypes.ST_Merge.convert_to_xml(value);
        end

        function validate(value)
            % VALIDATE BaseStringEnumerationType.validate (lines 97-101) against members_.
            mat2doc.oxml.simpletypes.BaseStringEnumerationType.validate_enum( ...
                value, mat2doc.oxml.simpletypes.ST_Merge.members_);
        end
    end
end

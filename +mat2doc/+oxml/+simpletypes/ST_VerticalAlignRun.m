classdef ST_VerticalAlignRun < mat2doc.oxml.simpletypes.XsdStringEnumeration
% ST_VERTICALALIGNRUN Valid values for `w:vertAlign/@val`.
%
%   XsdStringEnumeration with members BASELINE = "baseline", SUPERSCRIPT =
%   "superscript", SUBSCRIPT = "subscript" (Python `_members = (BASELINE,
%   SUPERSCRIPT, SUBSCRIPT)`, simpletypes.py 427-434). validate delegates to
%   BaseStringEnumerationType.validate_enum; to_xml re-declared so the template
%   routes to THIS validate (H10). from_xml / convert_* inherited (identity).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::ST_VerticalAlignRun
%   (lines 427-434)

    properties (Constant)
        BASELINE = "baseline"          % Python class attr
        SUPERSCRIPT = "superscript"    % Python class attr
        SUBSCRIPT = "subscript"        % Python class attr
        % Python _members = (BASELINE, SUPERSCRIPT, SUBSCRIPT)
        members_ = ["baseline", "superscript", "subscript"]
    end

    methods (Static)
        function s = to_xml(value)
            % TO_XML BaseSimpleType.to_xml (lines 29-33), bound to THIS validate.
            mat2doc.oxml.simpletypes.ST_VerticalAlignRun.validate(value);
            s = mat2doc.oxml.simpletypes.ST_VerticalAlignRun.convert_to_xml(value);
        end

        function validate(value)
            % VALIDATE BaseStringEnumerationType.validate (lines 97-101) against members_.
            mat2doc.oxml.simpletypes.BaseStringEnumerationType.validate_enum( ...
                value, mat2doc.oxml.simpletypes.ST_VerticalAlignRun.members_);
        end
    end
end

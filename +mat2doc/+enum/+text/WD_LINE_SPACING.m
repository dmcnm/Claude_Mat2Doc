classdef WD_LINE_SPACING < mat2doc.enum.base.BaseXmlEnum
% WD_LINE_SPACING Specifies a line spacing format to be applied to a paragraph.
%
%   Example:
%       % paragraph.line_spacing_rule = mat2doc.enum.text.WD_LINE_SPACING.EXACTLY;
%
%   The SINGLE, ONE_POINT_FIVE and DOUBLE members carry the placeholder
%   xml_value "UNMAPPED" (a real, truthy string - NOT None and NOT ""). This
%   matches python-docx exactly: from_xml("UNMAPPED") returns SINGLE (the first
%   member with that xml_value) and to_xml(SINGLE) returns "UNMAPPED" (the
%   `if not xml_value` guard passes because "UNMAPPED" is truthy). These tokens
%   never appear in real WordprocessingML; they are reproduced verbatim for
%   fidelity, no special handling.
%
%   MS API name: WdLineSpacing
%   http://msdn.microsoft.com/en-us/library/office/ff844910.aspx
%
%   Ported from python-docx v1.2.0: src/docx/enum/text.py::WD_LINE_SPACING

    methods
        function obj = WD_LINE_SPACING(ms_api_value, xml_value, docstr)
            obj@mat2doc.enum.base.BaseXmlEnum(ms_api_value, xml_value, docstr);
        end
    end

    methods (Static)
        function member = from_xml(xml_value)
            member = mat2doc.enum.base.BaseXmlEnum.from_xml_( ...
                "mat2doc.enum.text.WD_LINE_SPACING", xml_value);
        end
        function s = to_xml(value)
            s = mat2doc.enum.base.BaseXmlEnum.to_xml_( ...
                "mat2doc.enum.text.WD_LINE_SPACING", value);
        end
    end

    enumeration
        % Single spaced (default).
        SINGLE         (0, "UNMAPPED", "Single spaced (default).")
        % Space-and-a-half line spacing.
        ONE_POINT_FIVE (1, "UNMAPPED", "Space-and-a-half line spacing.")
        % Double spaced.
        DOUBLE         (2, "UNMAPPED", "Double spaced.")
        % Minimum line spacing is specified amount. Amount is specified separately.
        AT_LEAST       (3, "atLeast", ...
            "Minimum line spacing is specified amount. Amount is specified separately.")
        % Line spacing is exactly specified amount. Amount is specified separately.
        EXACTLY        (4, "exact", ...
            "Line spacing is exactly specified amount. Amount is specified separately.")
        % Line spacing is specified as multiple of line heights.
        MULTIPLE       (5, "auto", ...
            "Line spacing is specified as multiple of line heights. Changing font size" + ...
            " will change line spacing proportionately.")
    end
end

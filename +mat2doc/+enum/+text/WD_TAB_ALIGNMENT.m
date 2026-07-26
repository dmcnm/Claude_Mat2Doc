classdef WD_TAB_ALIGNMENT < mat2doc.enum.base.BaseXmlEnum
% WD_TAB_ALIGNMENT Specifies the tab stop alignment to apply.
%
%   Example:
%       m = mat2doc.enum.text.WD_TAB_ALIGNMENT.from_xml("center");   % CENTER
%       double(m.value)                                              % 1
%       mat2doc.enum.text.WD_TAB_ALIGNMENT.to_xml(m)                 % "center"
%
%   MS API name: WdTabAlignment
%   https://msdn.microsoft.com/EN-US/library/office/ff195609.aspx
%
%   Ported from python-docx v1.2.0: src/docx/enum/text.py::WD_TAB_ALIGNMENT

    methods
        function obj = WD_TAB_ALIGNMENT(ms_api_value, xml_value, docstr)
            obj@mat2doc.enum.base.BaseXmlEnum(ms_api_value, xml_value, docstr);
        end
    end

    methods (Static)
        function member = from_xml(xml_value)
            member = mat2doc.enum.base.BaseXmlEnum.from_xml_( ...
                "mat2doc.enum.text.WD_TAB_ALIGNMENT", xml_value);
        end
        function s = to_xml(value)
            s = mat2doc.enum.base.BaseXmlEnum.to_xml_( ...
                "mat2doc.enum.text.WD_TAB_ALIGNMENT", value);
        end
    end

    enumeration
        % Left-aligned.
        LEFT    (0,   "left",    "Left-aligned.")
        % Center-aligned.
        CENTER  (1,   "center",  "Center-aligned.")
        % Right-aligned.
        RIGHT   (2,   "right",   "Right-aligned.")
        % Decimal-aligned.
        DECIMAL (3,   "decimal", "Decimal-aligned.")
        % Bar-aligned.
        BAR     (4,   "bar",     "Bar-aligned.")
        % List-aligned. (deprecated)
        LIST    (6,   "list",    "List-aligned. (deprecated)")
        % Clear an inherited tab stop.
        CLEAR   (101, "clear",   "Clear an inherited tab stop.")
        % Right-aligned.  (deprecated)
        END     (102, "end",     "Right-aligned.  (deprecated)")
        % Left-aligned.  (deprecated)
        NUM     (103, "num",     "Left-aligned.  (deprecated)")
        % Left-aligned.  (deprecated)
        START   (104, "start",   "Left-aligned.  (deprecated)")
    end
end

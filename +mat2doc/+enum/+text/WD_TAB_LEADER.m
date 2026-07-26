classdef WD_TAB_LEADER < mat2doc.enum.base.BaseXmlEnum
% WD_TAB_LEADER Specifies the character to use as the leader with formatted tabs.
%
%   Example:
%       m = mat2doc.enum.text.WD_TAB_LEADER.from_xml("dot");   % DOTS
%       double(m.value)                                        % 1
%       mat2doc.enum.text.WD_TAB_LEADER.to_xml(m)              % "dot"
%
%   MS API name: WdTabLeader
%   https://msdn.microsoft.com/en-us/library/office/ff845050.aspx
%
%   Ported from python-docx v1.2.0: src/docx/enum/text.py::WD_TAB_LEADER

    methods
        function obj = WD_TAB_LEADER(ms_api_value, xml_value, docstr)
            obj@mat2doc.enum.base.BaseXmlEnum(ms_api_value, xml_value, docstr);
        end
    end

    methods (Static)
        function member = from_xml(xml_value)
            member = mat2doc.enum.base.BaseXmlEnum.from_xml_( ...
                "mat2doc.enum.text.WD_TAB_LEADER", xml_value);
        end
        function s = to_xml(value)
            s = mat2doc.enum.base.BaseXmlEnum.to_xml_( ...
                "mat2doc.enum.text.WD_TAB_LEADER", value);
        end
    end

    enumeration
        % Spaces. Default.
        SPACES     (0, "none",       "Spaces. Default.")
        % Dots.
        DOTS       (1, "dot",        "Dots.")
        % Dashes.
        DASHES     (2, "hyphen",     "Dashes.")
        % Double lines.
        LINES      (3, "underscore", "Double lines.")
        % A heavy line.
        HEAVY      (4, "heavy",      "A heavy line.")
        % A vertically-centered dot.
        MIDDLE_DOT (5, "middleDot",  "A vertically-centered dot.")
    end
end

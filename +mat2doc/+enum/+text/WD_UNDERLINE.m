classdef WD_UNDERLINE < mat2doc.enum.base.BaseXmlEnum
% WD_UNDERLINE Specifies the style of underline applied to a run of characters.
%
%   The INHERITED member (value -1) carries xml_value None (a <missing> string
%   here; H3). It is reachable via from_xml(None) -> INHERITED (docx
%   BaseXmlEnum.from_xml has NO None short-circuit; None == None matches).
%   to_xml(INHERITED) raises "WD_UNDERLINE.INHERITED has no XML representation".
%
%   Example:
%       mat2doc.enum.text.WD_UNDERLINE.from_xml([])              % INHERITED (None)
%       mat2doc.enum.text.WD_UNDERLINE.from_xml("single")        % SINGLE
%       mat2doc.enum.text.WD_UNDERLINE.to_xml(mat2doc.enum.text.WD_UNDERLINE.WAVY) % "wave"
%
%   MS API name: WdUnderline
%   http://msdn.microsoft.com/en-us/library/office/ff822388.aspx
%
%   Ported from python-docx v1.2.0: src/docx/enum/text.py::WD_UNDERLINE

    methods
        function obj = WD_UNDERLINE(ms_api_value, xml_value, docstr)
            obj@mat2doc.enum.base.BaseXmlEnum(ms_api_value, xml_value, docstr);
        end
    end

    methods (Static)
        function member = from_xml(xml_value)
            member = mat2doc.enum.base.BaseXmlEnum.from_xml_( ...
                "mat2doc.enum.text.WD_UNDERLINE", xml_value);
        end
        function s = to_xml(value)
            s = mat2doc.enum.base.BaseXmlEnum.to_xml_( ...
                "mat2doc.enum.text.WD_UNDERLINE", value);
        end
    end

    enumeration
        % Inherit underline setting from containing paragraph. (xml_value None)
        INHERITED          (-1, string(missing), ...
            "Inherit underline setting from containing paragraph.")
        % No underline. (see full docstring in python-docx source)
        NONE               (0,  "none", ...
            "No underline." + newline + newline + ...
            "This setting overrides any inherited underline value, so can" + ...
            " be used to remove underline from a run that inherits underlining from its" + ...
            " containing paragraph. Note this is not the same as assigning |None| to" + ...
            " Run.underline. |None| is a valid assignment value, but causes the run to" + ...
            " inherit its underline value. Assigning `WD_UNDERLINE.NONE` causes" + ...
            " underlining to be unconditionally turned off.")
        % A single line. (write-only; see full docstring in python-docx source)
        SINGLE             (1,  "single", ...
            "A single line." + newline + newline + ...
            "Note that this setting is write-only in the sense that" + ...
            " |True| (rather than `WD_UNDERLINE.SINGLE`) is returned for a run having" + ...
            " this setting.")
        % Underline individual words only.
        WORDS              (2,  "words",           "Underline individual words only.")
        % A double line.
        DOUBLE             (3,  "double",          "A double line.")
        % Dots.
        DOTTED             (4,  "dotted",          "Dots.")
        % A single thick line.
        THICK              (6,  "thick",           "A single thick line.")
        % Dashes.
        DASH               (7,  "dash",            "Dashes.")
        % Alternating dots and dashes.
        DOT_DASH           (9,  "dotDash",         "Alternating dots and dashes.")
        % An alternating dot-dot-dash pattern.
        DOT_DOT_DASH       (10, "dotDotDash",      "An alternating dot-dot-dash pattern.")
        % A single wavy line.
        WAVY               (11, "wave",            "A single wavy line.")
        % Heavy dots.
        DOTTED_HEAVY       (20, "dottedHeavy",     "Heavy dots.")
        % Heavy dashes.
        DASH_HEAVY         (23, "dashedHeavy",     "Heavy dashes.")
        % Alternating heavy dots and heavy dashes.
        DOT_DASH_HEAVY     (25, "dashDotHeavy", ...
            "Alternating heavy dots and heavy dashes.")
        % An alternating heavy dot-dot-dash pattern.
        DOT_DOT_DASH_HEAVY (26, "dashDotDotHeavy", ...
            "An alternating heavy dot-dot-dash pattern.")
        % A heavy wavy line.
        WAVY_HEAVY         (27, "wavyHeavy",       "A heavy wavy line.")
        % Long dashes.
        DASH_LONG          (39, "dashLong",        "Long dashes.")
        % A double wavy line.
        WAVY_DOUBLE        (43, "wavyDouble",      "A double wavy line.")
        % Long heavy dashes.
        DASH_LONG_HEAVY    (55, "dashLongHeavy",   "Long heavy dashes.")
    end
end

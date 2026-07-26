classdef WD_COLOR_INDEX < mat2doc.enum.base.BaseXmlEnum
% WD_COLOR_INDEX Specifies a standard preset color to apply.
%
%   Used for font highlighting and perhaps other applications.
%
%   Alias: WD_COLOR (mat2doc.enum.text.WD_COLOR).
%
%   The INHERITED member (value -1) carries xml_value None (a <missing> string
%   here; H3). It is reachable via from_xml(None) -> INHERITED (the docx
%   BaseXmlEnum.from_xml has NO None short-circuit; it scans for the member whose
%   xml_value equals the query, and None == None matches). to_xml(INHERITED)
%   raises "WD_COLOR_INDEX.INHERITED has no XML representation" (the `if not
%   xml_value` guard is true for a <missing> xml_value).
%
%   Example:
%       mat2doc.enum.text.WD_COLOR_INDEX.from_xml([])                % INHERITED (None)
%       double(mat2doc.enum.text.WD_COLOR_INDEX.BRIGHT_GREEN.value)  % 4
%       mat2doc.enum.text.WD_COLOR_INDEX.to_xml(mat2doc.enum.text.WD_COLOR_INDEX.BRIGHT_GREEN) % "green"
%
%   MS API name: WdColorIndex
%   https://msdn.microsoft.com/EN-US/library/office/ff195343.aspx
%
%   Ported from python-docx v1.2.0: src/docx/enum/text.py::WD_COLOR_INDEX

    methods
        function obj = WD_COLOR_INDEX(ms_api_value, xml_value, docstr)
            obj@mat2doc.enum.base.BaseXmlEnum(ms_api_value, xml_value, docstr);
        end
    end

    methods (Static)
        function member = from_xml(xml_value)
            member = mat2doc.enum.base.BaseXmlEnum.from_xml_( ...
                "mat2doc.enum.text.WD_COLOR_INDEX", xml_value);
        end
        function s = to_xml(value)
            s = mat2doc.enum.base.BaseXmlEnum.to_xml_( ...
                "mat2doc.enum.text.WD_COLOR_INDEX", value);
        end
    end

    enumeration
        % Color is inherited from the style hierarchy. (xml_value None)
        INHERITED    (-1, string(missing), "Color is inherited from the style hierarchy.")
        % Automatic color. Default; usually black.
        AUTO         (0,  "default",       "Automatic color. Default; usually black.")
        % Black color.
        BLACK        (1,  "black",         "Black color.")
        % Blue color
        BLUE         (2,  "blue",          "Blue color")
        % Bright green color.
        BRIGHT_GREEN (4,  "green",         "Bright green color.")
        % Dark blue color.
        DARK_BLUE    (9,  "darkBlue",      "Dark blue color.")
        % Dark red color.
        DARK_RED     (13, "darkRed",       "Dark red color.")
        % Dark yellow color.
        DARK_YELLOW  (14, "darkYellow",    "Dark yellow color.")
        % 25% shade of gray color.
        GRAY_25      (16, "lightGray",     "25% shade of gray color.")
        % 50% shade of gray color.
        GRAY_50      (15, "darkGray",      "50% shade of gray color.")
        % Green color.
        GREEN        (11, "darkGreen",     "Green color.")
        % Pink color.
        PINK         (5,  "magenta",       "Pink color.")
        % Red color.
        RED          (6,  "red",           "Red color.")
        % Teal color.
        TEAL         (10, "darkCyan",      "Teal color.")
        % Turquoise color.
        TURQUOISE    (3,  "cyan",          "Turquoise color.")
        % Violet color.
        VIOLET       (12, "darkMagenta",   "Violet color.")
        % White color.
        WHITE        (8,  "white",         "White color.")
        % Yellow color.
        YELLOW       (7,  "yellow",        "Yellow color.")
    end
end

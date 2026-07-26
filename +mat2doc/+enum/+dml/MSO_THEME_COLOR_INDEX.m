classdef MSO_THEME_COLOR_INDEX < mat2doc.enum.base.BaseXmlEnum
% MSO_THEME_COLOR_INDEX Indicates an Office theme color.
%
%   One of those shown in the color gallery on the formatting ribbon.
%
%   Alias: MSO_THEME_COLOR (mat2doc.enum.dml.MSO_THEME_COLOR).
%
%   NOTE: the docx member set differs from the pptx enum of the same name - do
%   NOT copy the Mat2Ppt version. docx uses full-word xml_values (accent1,
%   background1, dark1, hyperlink, light1, text1, ...), NOT_THEME_COLOR carries
%   the placeholder xml_value "UNMAPPED" (a truthy string, NOT ""), and there is
%   NO MIXED member. to_xml(NOT_THEME_COLOR) therefore returns "UNMAPPED" (the
%   `if not xml_value` guard passes), matching python-docx exactly.
%
%   Example:
%       c = mat2doc.enum.dml.MSO_THEME_COLOR.ACCENT_1;
%       double(c.value)                                          % 5
%       mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.to_xml(c)         % "accent1"
%       mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.from_xml("dark1") % DARK_1
%
%   MS API name: MsoThemeColorIndex
%   http://msdn.microsoft.com/en-us/library/office/ff860782(v=office.15).aspx
%
%   Ported from python-docx v1.2.0: src/docx/enum/dml.py::MSO_THEME_COLOR_INDEX

    methods
        function obj = MSO_THEME_COLOR_INDEX(ms_api_value, xml_value, docstr)
            obj@mat2doc.enum.base.BaseXmlEnum(ms_api_value, xml_value, docstr);
        end
    end

    methods (Static)
        function member = from_xml(xml_value)
            member = mat2doc.enum.base.BaseXmlEnum.from_xml_( ...
                "mat2doc.enum.dml.MSO_THEME_COLOR_INDEX", xml_value);
        end
        function s = to_xml(value)
            s = mat2doc.enum.base.BaseXmlEnum.to_xml_( ...
                "mat2doc.enum.dml.MSO_THEME_COLOR_INDEX", value);
        end
    end

    enumeration
        % Indicates the color is not a theme color.
        NOT_THEME_COLOR    (0,  "UNMAPPED",          "Indicates the color is not a theme color.")
        % Specifies the Accent 1 theme color.
        ACCENT_1           (5,  "accent1",           "Specifies the Accent 1 theme color.")
        % Specifies the Accent 2 theme color.
        ACCENT_2           (6,  "accent2",           "Specifies the Accent 2 theme color.")
        % Specifies the Accent 3 theme color.
        ACCENT_3           (7,  "accent3",           "Specifies the Accent 3 theme color.")
        % Specifies the Accent 4 theme color.
        ACCENT_4           (8,  "accent4",           "Specifies the Accent 4 theme color.")
        % Specifies the Accent 5 theme color.
        ACCENT_5           (9,  "accent5",           "Specifies the Accent 5 theme color.")
        % Specifies the Accent 6 theme color.
        ACCENT_6           (10, "accent6",           "Specifies the Accent 6 theme color.")
        % Specifies the Background 1 theme color.
        BACKGROUND_1       (14, "background1",       "Specifies the Background 1 theme color.")
        % Specifies the Background 2 theme color.
        BACKGROUND_2       (16, "background2",       "Specifies the Background 2 theme color.")
        % Specifies the Dark 1 theme color.
        DARK_1             (1,  "dark1",             "Specifies the Dark 1 theme color.")
        % Specifies the Dark 2 theme color.
        DARK_2             (3,  "dark2",             "Specifies the Dark 2 theme color.")
        % Specifies the theme color for a clicked hyperlink.
        FOLLOWED_HYPERLINK (12, "followedHyperlink", ...
            "Specifies the theme color for a clicked hyperlink.")
        % Specifies the theme color for a hyperlink.
        HYPERLINK          (11, "hyperlink",         "Specifies the theme color for a hyperlink.")
        % Specifies the Light 1 theme color.
        LIGHT_1            (2,  "light1",            "Specifies the Light 1 theme color.")
        % Specifies the Light 2 theme color.
        LIGHT_2            (4,  "light2",            "Specifies the Light 2 theme color.")
        % Specifies the Text 1 theme color.
        TEXT_1             (13, "text1",             "Specifies the Text 1 theme color.")
        % Specifies the Text 2 theme color.
        TEXT_2             (15, "text2",             "Specifies the Text 2 theme color.")
    end
end

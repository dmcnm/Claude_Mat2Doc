classdef MSO_THEME_COLOR
% MSO_THEME_COLOR Alias of MSO_THEME_COLOR_INDEX.
%
%   Mirrors the python-docx module-level assignment
%   ``MSO_THEME_COLOR = MSO_THEME_COLOR_INDEX`` (dml.py line 103). MSO_THEME_COLOR
%   is the spelling the python-docx API examples use, so consuming code
%   references mat2doc.enum.dml.MSO_THEME_COLOR. MATLAB has no class aliasing, so
%   this class re-exports the canonical enumeration's members as Constant
%   properties and forwards the static methods. The members ARE
%   mat2doc.enum.dml.MSO_THEME_COLOR_INDEX instances, so identity (==), isa, and
%   from_xml/to_xml behave exactly as if the two names referred to one
%   enumeration.
%
%   Example:
%       mat2doc.enum.dml.MSO_THEME_COLOR.ACCENT_1 == ...
%           mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.ACCENT_1   % true
%
%   Ported from python-docx v1.2.0: src/docx/enum/dml.py::MSO_THEME_COLOR
%   (alias of MSO_THEME_COLOR_INDEX)

    properties (Constant)
        NOT_THEME_COLOR    = mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.NOT_THEME_COLOR
        ACCENT_1           = mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.ACCENT_1
        ACCENT_2           = mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.ACCENT_2
        ACCENT_3           = mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.ACCENT_3
        ACCENT_4           = mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.ACCENT_4
        ACCENT_5           = mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.ACCENT_5
        ACCENT_6           = mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.ACCENT_6
        BACKGROUND_1       = mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.BACKGROUND_1
        BACKGROUND_2       = mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.BACKGROUND_2
        DARK_1             = mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.DARK_1
        DARK_2             = mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.DARK_2
        FOLLOWED_HYPERLINK = mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.FOLLOWED_HYPERLINK
        HYPERLINK          = mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.HYPERLINK
        LIGHT_1            = mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.LIGHT_1
        LIGHT_2            = mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.LIGHT_2
        TEXT_1             = mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.TEXT_1
        TEXT_2             = mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.TEXT_2
    end

    methods (Static)
        function member = from_xml(xml_value)
            member = mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.from_xml(xml_value);
        end
        function s = to_xml(value)
            s = mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.to_xml(value);
        end
    end
end

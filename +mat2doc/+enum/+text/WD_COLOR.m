classdef WD_COLOR
% WD_COLOR Alias of WD_COLOR_INDEX.
%
%   Mirrors the python-docx module-level assignment
%   ``WD_COLOR = WD_COLOR_INDEX`` (text.py line 156). MATLAB has no class
%   aliasing, so this class re-exports the canonical enumeration's members as
%   Constant properties and forwards the static methods. The members ARE
%   mat2doc.enum.text.WD_COLOR_INDEX instances, so identity (==), isa, and
%   from_xml/to_xml behave exactly as if the two names referred to one
%   enumeration.
%
%   Example:
%       mat2doc.enum.text.WD_COLOR.AUTO == ...
%           mat2doc.enum.text.WD_COLOR_INDEX.AUTO      % true
%
%   Ported from python-docx v1.2.0: src/docx/enum/text.py::WD_COLOR
%   (alias of WD_COLOR_INDEX)

    properties (Constant)
        INHERITED    = mat2doc.enum.text.WD_COLOR_INDEX.INHERITED
        AUTO         = mat2doc.enum.text.WD_COLOR_INDEX.AUTO
        BLACK        = mat2doc.enum.text.WD_COLOR_INDEX.BLACK
        BLUE         = mat2doc.enum.text.WD_COLOR_INDEX.BLUE
        BRIGHT_GREEN = mat2doc.enum.text.WD_COLOR_INDEX.BRIGHT_GREEN
        DARK_BLUE    = mat2doc.enum.text.WD_COLOR_INDEX.DARK_BLUE
        DARK_RED     = mat2doc.enum.text.WD_COLOR_INDEX.DARK_RED
        DARK_YELLOW  = mat2doc.enum.text.WD_COLOR_INDEX.DARK_YELLOW
        GRAY_25      = mat2doc.enum.text.WD_COLOR_INDEX.GRAY_25
        GRAY_50      = mat2doc.enum.text.WD_COLOR_INDEX.GRAY_50
        GREEN        = mat2doc.enum.text.WD_COLOR_INDEX.GREEN
        PINK         = mat2doc.enum.text.WD_COLOR_INDEX.PINK
        RED          = mat2doc.enum.text.WD_COLOR_INDEX.RED
        TEAL         = mat2doc.enum.text.WD_COLOR_INDEX.TEAL
        TURQUOISE    = mat2doc.enum.text.WD_COLOR_INDEX.TURQUOISE
        VIOLET       = mat2doc.enum.text.WD_COLOR_INDEX.VIOLET
        WHITE        = mat2doc.enum.text.WD_COLOR_INDEX.WHITE
        YELLOW       = mat2doc.enum.text.WD_COLOR_INDEX.YELLOW
    end

    methods (Static)
        function member = from_xml(xml_value)
            member = mat2doc.enum.text.WD_COLOR_INDEX.from_xml(xml_value);
        end
        function s = to_xml(value)
            s = mat2doc.enum.text.WD_COLOR_INDEX.to_xml(value);
        end
    end
end

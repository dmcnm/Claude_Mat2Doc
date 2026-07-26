classdef WD_BREAK
% WD_BREAK Alias of WD_BREAK_TYPE.
%
%   Mirrors the python-docx module-level assignment
%   ``WD_BREAK = WD_BREAK_TYPE`` (text.py line 89). MATLAB has no class aliasing,
%   so this class re-exports the canonical enumeration's members (including the
%   TEXT_WRAPPING member-alias) as Constant properties. The members ARE
%   mat2doc.enum.text.WD_BREAK_TYPE instances, so identity (==) and isa behave
%   exactly as if the two names referred to one enumeration. WD_BREAK_TYPE is a
%   plain enum with no XML mapping, so there are no static methods to forward.
%
%   Example:
%       mat2doc.enum.text.WD_BREAK.PAGE == ...
%           mat2doc.enum.text.WD_BREAK_TYPE.PAGE           % true
%       string(mat2doc.enum.text.WD_BREAK.TEXT_WRAPPING)   % "LINE_CLEAR_ALL"
%
%   Ported from python-docx v1.2.0: src/docx/enum/text.py::WD_BREAK
%   (alias of WD_BREAK_TYPE)

    properties (Constant)
        COLUMN             = mat2doc.enum.text.WD_BREAK_TYPE.COLUMN
        LINE               = mat2doc.enum.text.WD_BREAK_TYPE.LINE
        LINE_CLEAR_LEFT    = mat2doc.enum.text.WD_BREAK_TYPE.LINE_CLEAR_LEFT
        LINE_CLEAR_RIGHT   = mat2doc.enum.text.WD_BREAK_TYPE.LINE_CLEAR_RIGHT
        LINE_CLEAR_ALL     = mat2doc.enum.text.WD_BREAK_TYPE.LINE_CLEAR_ALL
        PAGE               = mat2doc.enum.text.WD_BREAK_TYPE.PAGE
        SECTION_CONTINUOUS = mat2doc.enum.text.WD_BREAK_TYPE.SECTION_CONTINUOUS
        SECTION_EVEN_PAGE  = mat2doc.enum.text.WD_BREAK_TYPE.SECTION_EVEN_PAGE
        SECTION_NEXT_PAGE  = mat2doc.enum.text.WD_BREAK_TYPE.SECTION_NEXT_PAGE
        SECTION_ODD_PAGE   = mat2doc.enum.text.WD_BREAK_TYPE.SECTION_ODD_PAGE
        TEXT_WRAPPING      = mat2doc.enum.text.WD_BREAK_TYPE.TEXT_WRAPPING
    end
end

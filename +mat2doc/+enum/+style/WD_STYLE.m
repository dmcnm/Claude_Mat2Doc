classdef WD_STYLE
% WD_STYLE Alias of WD_BUILTIN_STYLE.
%
%   Mirrors the python-docx module-level assignment
%   ``WD_STYLE = WD_BUILTIN_STYLE`` (style.py line 423). MATLAB has no class
%   aliasing, so this class re-exports the canonical enumeration's 132 members
%   as Constant properties. The members ARE mat2doc.enum.style.WD_BUILTIN_STYLE
%   instances, so identity (==) and isa behave exactly as if the two names
%   referred to one enumeration. WD_BUILTIN_STYLE is a BaseEnum with no XML
%   mapping, so there are no static methods to forward.
%
%   Example:
%       mat2doc.enum.style.WD_STYLE.BODY_TEXT == ...
%           mat2doc.enum.style.WD_BUILTIN_STYLE.BODY_TEXT   % true
%
%   Ported from python-docx v1.2.0: src/docx/enum/style.py::WD_STYLE
%   (alias of WD_BUILTIN_STYLE)

    properties (Constant)
        BLOCK_QUOTATION                 = mat2doc.enum.style.WD_BUILTIN_STYLE.BLOCK_QUOTATION
        BODY_TEXT                       = mat2doc.enum.style.WD_BUILTIN_STYLE.BODY_TEXT
        BODY_TEXT_2                     = mat2doc.enum.style.WD_BUILTIN_STYLE.BODY_TEXT_2
        BODY_TEXT_3                     = mat2doc.enum.style.WD_BUILTIN_STYLE.BODY_TEXT_3
        BODY_TEXT_FIRST_INDENT          = mat2doc.enum.style.WD_BUILTIN_STYLE.BODY_TEXT_FIRST_INDENT
        BODY_TEXT_FIRST_INDENT_2        = mat2doc.enum.style.WD_BUILTIN_STYLE.BODY_TEXT_FIRST_INDENT_2
        BODY_TEXT_INDENT                = mat2doc.enum.style.WD_BUILTIN_STYLE.BODY_TEXT_INDENT
        BODY_TEXT_INDENT_2              = mat2doc.enum.style.WD_BUILTIN_STYLE.BODY_TEXT_INDENT_2
        BODY_TEXT_INDENT_3              = mat2doc.enum.style.WD_BUILTIN_STYLE.BODY_TEXT_INDENT_3
        BOOK_TITLE                      = mat2doc.enum.style.WD_BUILTIN_STYLE.BOOK_TITLE
        CAPTION                         = mat2doc.enum.style.WD_BUILTIN_STYLE.CAPTION
        CLOSING                         = mat2doc.enum.style.WD_BUILTIN_STYLE.CLOSING
        COMMENT_REFERENCE               = mat2doc.enum.style.WD_BUILTIN_STYLE.COMMENT_REFERENCE
        COMMENT_TEXT                    = mat2doc.enum.style.WD_BUILTIN_STYLE.COMMENT_TEXT
        DATE                            = mat2doc.enum.style.WD_BUILTIN_STYLE.DATE
        DEFAULT_PARAGRAPH_FONT          = mat2doc.enum.style.WD_BUILTIN_STYLE.DEFAULT_PARAGRAPH_FONT
        EMPHASIS                        = mat2doc.enum.style.WD_BUILTIN_STYLE.EMPHASIS
        ENDNOTE_REFERENCE               = mat2doc.enum.style.WD_BUILTIN_STYLE.ENDNOTE_REFERENCE
        ENDNOTE_TEXT                    = mat2doc.enum.style.WD_BUILTIN_STYLE.ENDNOTE_TEXT
        ENVELOPE_ADDRESS                = mat2doc.enum.style.WD_BUILTIN_STYLE.ENVELOPE_ADDRESS
        ENVELOPE_RETURN                 = mat2doc.enum.style.WD_BUILTIN_STYLE.ENVELOPE_RETURN
        FOOTER                          = mat2doc.enum.style.WD_BUILTIN_STYLE.FOOTER
        FOOTNOTE_REFERENCE              = mat2doc.enum.style.WD_BUILTIN_STYLE.FOOTNOTE_REFERENCE
        FOOTNOTE_TEXT                   = mat2doc.enum.style.WD_BUILTIN_STYLE.FOOTNOTE_TEXT
        HEADER                          = mat2doc.enum.style.WD_BUILTIN_STYLE.HEADER
        HEADING_1                       = mat2doc.enum.style.WD_BUILTIN_STYLE.HEADING_1
        HEADING_2                       = mat2doc.enum.style.WD_BUILTIN_STYLE.HEADING_2
        HEADING_3                       = mat2doc.enum.style.WD_BUILTIN_STYLE.HEADING_3
        HEADING_4                       = mat2doc.enum.style.WD_BUILTIN_STYLE.HEADING_4
        HEADING_5                       = mat2doc.enum.style.WD_BUILTIN_STYLE.HEADING_5
        HEADING_6                       = mat2doc.enum.style.WD_BUILTIN_STYLE.HEADING_6
        HEADING_7                       = mat2doc.enum.style.WD_BUILTIN_STYLE.HEADING_7
        HEADING_8                       = mat2doc.enum.style.WD_BUILTIN_STYLE.HEADING_8
        HEADING_9                       = mat2doc.enum.style.WD_BUILTIN_STYLE.HEADING_9
        HTML_ACRONYM                    = mat2doc.enum.style.WD_BUILTIN_STYLE.HTML_ACRONYM
        HTML_ADDRESS                    = mat2doc.enum.style.WD_BUILTIN_STYLE.HTML_ADDRESS
        HTML_CITE                       = mat2doc.enum.style.WD_BUILTIN_STYLE.HTML_CITE
        HTML_CODE                       = mat2doc.enum.style.WD_BUILTIN_STYLE.HTML_CODE
        HTML_DFN                        = mat2doc.enum.style.WD_BUILTIN_STYLE.HTML_DFN
        HTML_KBD                        = mat2doc.enum.style.WD_BUILTIN_STYLE.HTML_KBD
        HTML_NORMAL                     = mat2doc.enum.style.WD_BUILTIN_STYLE.HTML_NORMAL
        HTML_PRE                        = mat2doc.enum.style.WD_BUILTIN_STYLE.HTML_PRE
        HTML_SAMP                       = mat2doc.enum.style.WD_BUILTIN_STYLE.HTML_SAMP
        HTML_TT                         = mat2doc.enum.style.WD_BUILTIN_STYLE.HTML_TT
        HTML_VAR                        = mat2doc.enum.style.WD_BUILTIN_STYLE.HTML_VAR
        HYPERLINK                       = mat2doc.enum.style.WD_BUILTIN_STYLE.HYPERLINK
        HYPERLINK_FOLLOWED              = mat2doc.enum.style.WD_BUILTIN_STYLE.HYPERLINK_FOLLOWED
        INDEX_1                         = mat2doc.enum.style.WD_BUILTIN_STYLE.INDEX_1
        INDEX_2                         = mat2doc.enum.style.WD_BUILTIN_STYLE.INDEX_2
        INDEX_3                         = mat2doc.enum.style.WD_BUILTIN_STYLE.INDEX_3
        INDEX_4                         = mat2doc.enum.style.WD_BUILTIN_STYLE.INDEX_4
        INDEX_5                         = mat2doc.enum.style.WD_BUILTIN_STYLE.INDEX_5
        INDEX_6                         = mat2doc.enum.style.WD_BUILTIN_STYLE.INDEX_6
        INDEX_7                         = mat2doc.enum.style.WD_BUILTIN_STYLE.INDEX_7
        INDEX_8                         = mat2doc.enum.style.WD_BUILTIN_STYLE.INDEX_8
        INDEX_9                         = mat2doc.enum.style.WD_BUILTIN_STYLE.INDEX_9
        INDEX_HEADING                   = mat2doc.enum.style.WD_BUILTIN_STYLE.INDEX_HEADING
        INTENSE_EMPHASIS                = mat2doc.enum.style.WD_BUILTIN_STYLE.INTENSE_EMPHASIS
        INTENSE_QUOTE                   = mat2doc.enum.style.WD_BUILTIN_STYLE.INTENSE_QUOTE
        INTENSE_REFERENCE               = mat2doc.enum.style.WD_BUILTIN_STYLE.INTENSE_REFERENCE
        LINE_NUMBER                     = mat2doc.enum.style.WD_BUILTIN_STYLE.LINE_NUMBER
        LIST                            = mat2doc.enum.style.WD_BUILTIN_STYLE.LIST
        LIST_2                          = mat2doc.enum.style.WD_BUILTIN_STYLE.LIST_2
        LIST_3                          = mat2doc.enum.style.WD_BUILTIN_STYLE.LIST_3
        LIST_4                          = mat2doc.enum.style.WD_BUILTIN_STYLE.LIST_4
        LIST_5                          = mat2doc.enum.style.WD_BUILTIN_STYLE.LIST_5
        LIST_BULLET                     = mat2doc.enum.style.WD_BUILTIN_STYLE.LIST_BULLET
        LIST_BULLET_2                   = mat2doc.enum.style.WD_BUILTIN_STYLE.LIST_BULLET_2
        LIST_BULLET_3                   = mat2doc.enum.style.WD_BUILTIN_STYLE.LIST_BULLET_3
        LIST_BULLET_4                   = mat2doc.enum.style.WD_BUILTIN_STYLE.LIST_BULLET_4
        LIST_BULLET_5                   = mat2doc.enum.style.WD_BUILTIN_STYLE.LIST_BULLET_5
        LIST_CONTINUE                   = mat2doc.enum.style.WD_BUILTIN_STYLE.LIST_CONTINUE
        LIST_CONTINUE_2                 = mat2doc.enum.style.WD_BUILTIN_STYLE.LIST_CONTINUE_2
        LIST_CONTINUE_3                 = mat2doc.enum.style.WD_BUILTIN_STYLE.LIST_CONTINUE_3
        LIST_CONTINUE_4                 = mat2doc.enum.style.WD_BUILTIN_STYLE.LIST_CONTINUE_4
        LIST_CONTINUE_5                 = mat2doc.enum.style.WD_BUILTIN_STYLE.LIST_CONTINUE_5
        LIST_NUMBER                     = mat2doc.enum.style.WD_BUILTIN_STYLE.LIST_NUMBER
        LIST_NUMBER_2                   = mat2doc.enum.style.WD_BUILTIN_STYLE.LIST_NUMBER_2
        LIST_NUMBER_3                   = mat2doc.enum.style.WD_BUILTIN_STYLE.LIST_NUMBER_3
        LIST_NUMBER_4                   = mat2doc.enum.style.WD_BUILTIN_STYLE.LIST_NUMBER_4
        LIST_NUMBER_5                   = mat2doc.enum.style.WD_BUILTIN_STYLE.LIST_NUMBER_5
        LIST_PARAGRAPH                  = mat2doc.enum.style.WD_BUILTIN_STYLE.LIST_PARAGRAPH
        MACRO_TEXT                      = mat2doc.enum.style.WD_BUILTIN_STYLE.MACRO_TEXT
        MESSAGE_HEADER                  = mat2doc.enum.style.WD_BUILTIN_STYLE.MESSAGE_HEADER
        NAV_PANE                        = mat2doc.enum.style.WD_BUILTIN_STYLE.NAV_PANE
        NORMAL                          = mat2doc.enum.style.WD_BUILTIN_STYLE.NORMAL
        NORMAL_INDENT                   = mat2doc.enum.style.WD_BUILTIN_STYLE.NORMAL_INDENT
        NORMAL_OBJECT                   = mat2doc.enum.style.WD_BUILTIN_STYLE.NORMAL_OBJECT
        NORMAL_TABLE                    = mat2doc.enum.style.WD_BUILTIN_STYLE.NORMAL_TABLE
        NOTE_HEADING                    = mat2doc.enum.style.WD_BUILTIN_STYLE.NOTE_HEADING
        PAGE_NUMBER                     = mat2doc.enum.style.WD_BUILTIN_STYLE.PAGE_NUMBER
        PLAIN_TEXT                      = mat2doc.enum.style.WD_BUILTIN_STYLE.PLAIN_TEXT
        QUOTE                           = mat2doc.enum.style.WD_BUILTIN_STYLE.QUOTE
        SALUTATION                      = mat2doc.enum.style.WD_BUILTIN_STYLE.SALUTATION
        SIGNATURE                       = mat2doc.enum.style.WD_BUILTIN_STYLE.SIGNATURE
        STRONG                          = mat2doc.enum.style.WD_BUILTIN_STYLE.STRONG
        SUBTITLE                        = mat2doc.enum.style.WD_BUILTIN_STYLE.SUBTITLE
        SUBTLE_EMPHASIS                 = mat2doc.enum.style.WD_BUILTIN_STYLE.SUBTLE_EMPHASIS
        SUBTLE_REFERENCE                = mat2doc.enum.style.WD_BUILTIN_STYLE.SUBTLE_REFERENCE
        TABLE_COLORFUL_GRID             = mat2doc.enum.style.WD_BUILTIN_STYLE.TABLE_COLORFUL_GRID
        TABLE_COLORFUL_LIST             = mat2doc.enum.style.WD_BUILTIN_STYLE.TABLE_COLORFUL_LIST
        TABLE_COLORFUL_SHADING          = mat2doc.enum.style.WD_BUILTIN_STYLE.TABLE_COLORFUL_SHADING
        TABLE_DARK_LIST                 = mat2doc.enum.style.WD_BUILTIN_STYLE.TABLE_DARK_LIST
        TABLE_LIGHT_GRID                = mat2doc.enum.style.WD_BUILTIN_STYLE.TABLE_LIGHT_GRID
        TABLE_LIGHT_GRID_ACCENT_1       = mat2doc.enum.style.WD_BUILTIN_STYLE.TABLE_LIGHT_GRID_ACCENT_1
        TABLE_LIGHT_LIST                = mat2doc.enum.style.WD_BUILTIN_STYLE.TABLE_LIGHT_LIST
        TABLE_LIGHT_LIST_ACCENT_1       = mat2doc.enum.style.WD_BUILTIN_STYLE.TABLE_LIGHT_LIST_ACCENT_1
        TABLE_LIGHT_SHADING             = mat2doc.enum.style.WD_BUILTIN_STYLE.TABLE_LIGHT_SHADING
        TABLE_LIGHT_SHADING_ACCENT_1    = mat2doc.enum.style.WD_BUILTIN_STYLE.TABLE_LIGHT_SHADING_ACCENT_1
        TABLE_MEDIUM_GRID_1             = mat2doc.enum.style.WD_BUILTIN_STYLE.TABLE_MEDIUM_GRID_1
        TABLE_MEDIUM_GRID_2             = mat2doc.enum.style.WD_BUILTIN_STYLE.TABLE_MEDIUM_GRID_2
        TABLE_MEDIUM_GRID_3             = mat2doc.enum.style.WD_BUILTIN_STYLE.TABLE_MEDIUM_GRID_3
        TABLE_MEDIUM_LIST_1             = mat2doc.enum.style.WD_BUILTIN_STYLE.TABLE_MEDIUM_LIST_1
        TABLE_MEDIUM_LIST_1_ACCENT_1    = mat2doc.enum.style.WD_BUILTIN_STYLE.TABLE_MEDIUM_LIST_1_ACCENT_1
        TABLE_MEDIUM_LIST_2             = mat2doc.enum.style.WD_BUILTIN_STYLE.TABLE_MEDIUM_LIST_2
        TABLE_MEDIUM_SHADING_1          = mat2doc.enum.style.WD_BUILTIN_STYLE.TABLE_MEDIUM_SHADING_1
        TABLE_MEDIUM_SHADING_1_ACCENT_1 = mat2doc.enum.style.WD_BUILTIN_STYLE.TABLE_MEDIUM_SHADING_1_ACCENT_1
        TABLE_MEDIUM_SHADING_2          = mat2doc.enum.style.WD_BUILTIN_STYLE.TABLE_MEDIUM_SHADING_2
        TABLE_MEDIUM_SHADING_2_ACCENT_1 = mat2doc.enum.style.WD_BUILTIN_STYLE.TABLE_MEDIUM_SHADING_2_ACCENT_1
        TABLE_OF_AUTHORITIES            = mat2doc.enum.style.WD_BUILTIN_STYLE.TABLE_OF_AUTHORITIES
        TABLE_OF_FIGURES                = mat2doc.enum.style.WD_BUILTIN_STYLE.TABLE_OF_FIGURES
        TITLE                           = mat2doc.enum.style.WD_BUILTIN_STYLE.TITLE
        TOAHEADING                      = mat2doc.enum.style.WD_BUILTIN_STYLE.TOAHEADING
        TOC_1                           = mat2doc.enum.style.WD_BUILTIN_STYLE.TOC_1
        TOC_2                           = mat2doc.enum.style.WD_BUILTIN_STYLE.TOC_2
        TOC_3                           = mat2doc.enum.style.WD_BUILTIN_STYLE.TOC_3
        TOC_4                           = mat2doc.enum.style.WD_BUILTIN_STYLE.TOC_4
        TOC_5                           = mat2doc.enum.style.WD_BUILTIN_STYLE.TOC_5
        TOC_6                           = mat2doc.enum.style.WD_BUILTIN_STYLE.TOC_6
        TOC_7                           = mat2doc.enum.style.WD_BUILTIN_STYLE.TOC_7
        TOC_8                           = mat2doc.enum.style.WD_BUILTIN_STYLE.TOC_8
        TOC_9                           = mat2doc.enum.style.WD_BUILTIN_STYLE.TOC_9
    end
end

classdef TableStyle_ < mat2doc.styles.ParagraphStyle
% TABLESTYLE_ A table style.
%
%   A table style provides character and paragraph formatting for its contents as
%   well as special table formatting properties. python-docx v1.2.0 adds NO
%   members of its own beyond ParagraphStyle except a __repr__ override
%   ("_TableStyle('<name>') id: <id>", style.py 246-247) -- MATLAB object display
%   is a separate mechanism and is not ported (no output-visible effect). So this
%   class is a faithful thin subclass: it inherits the entire ParagraphStyle /
%   CharacterStyle / BaseStyle surface unchanged.
%
%   FLAG-3 (naming): Python's private `_TableStyle` maps to the trailing-
%   underscore convention TableStyle_ (leading-underscore rotation used across
%   the toolbox: _Cell -> Cell_, _TableStyle -> TableStyle_).
%
%   Ported from python-docx v1.2.0: src/docx/styles/style.py::_TableStyle

    methods
        function obj = TableStyle_(style_elm)
            % TABLESTYLE_ Wrap a `w:style` (inherits BaseStyle.__init__).
            %   Ported from python-docx v1.2.0: styles/style.py::_TableStyle
            obj@mat2doc.styles.ParagraphStyle(style_elm);
        end
    end
end

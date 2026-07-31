classdef NumberingStyle_ < mat2doc.styles.BaseStyle
% NUMBERINGSTYLE_ A numbering style.
%
%   Not yet implemented (python-docx v1.2.0, style.py 250-254): the class body is
%   empty -- a numbering style exposes only the inherited BaseStyle surface. This
%   port is a faithful thin subclass adding nothing.
%
%   FLAG-3 (naming): Python's private `_NumberingStyle` maps to the trailing-
%   underscore convention NumberingStyle_ (leading-underscore rotation:
%   _NumberingStyle -> NumberingStyle_).
%
%   Ported from python-docx v1.2.0: src/docx/styles/style.py::_NumberingStyle

    methods
        function obj = NumberingStyle_(style_elm)
            % NUMBERINGSTYLE_ Wrap a `w:style` (inherits BaseStyle.__init__).
            %   Ported from python-docx v1.2.0: styles/style.py::_NumberingStyle
            obj@mat2doc.styles.BaseStyle(style_elm);
        end
    end
end

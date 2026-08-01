classdef PngParser_ < handle
% PNGPARSER_ Parses a PNG image stream to extract properties from its chunks.
%
%   UNDERSCORE ROTATION (design.md section 2): python-docx _PngParser (a
%   module-private class) -> PngParser_.
%
%   NONE IDIOM (H3): python-docx tests `if pHYs is None`; Mat2Doc has no shared
%   isNone helper, so the None-guard is the inline isequal(x, []) test (ratified
%   none-idiom decision 2026-07-26).
%
%   Ported from python-docx v1.2.0: src/docx/image/png.py::_PngParser
%   (lines 35-89)

    properties (Access = private)
        chunks_          % a Chunks_ instance
    end

    methods
        function obj = PngParser_(chunks)
            % Ported from _PngParser.__init__ (png.py 38-40).
            obj.chunks_ = chunks;
        end

        function v = px_width(obj)
            % px_width @property (png.py 49-54): IHDR.px_width.
            IHDR = obj.chunks_.IHDR;
            v = IHDR.px_width;
        end

        function v = px_height(obj)
            % px_height @property (png.py 55-60): IHDR.px_height.
            IHDR = obj.chunks_.IHDR;
            v = IHDR.px_height;
        end

        function v = horz_dpi(obj)
            % horz_dpi @property (png.py 61-71): from pHYs, or 72 when absent.
            pHYs = obj.chunks_.pHYs;
            if isequal(pHYs, [])                              % None (H3)
                v = 72;
                return
            end
            v = mat2doc.image.PngParser_.dpi_(pHYs.units_specifier, pHYs.horz_px_per_unit);
        end

        function v = vert_dpi(obj)
            % vert_dpi @property (png.py 72-81): from pHYs, or 72 when absent.
            pHYs = obj.chunks_.pHYs;
            if isequal(pHYs, [])                              % None (H3)
                v = 72;
                return
            end
            v = mat2doc.image.PngParser_.dpi_(pHYs.units_specifier, pHYs.vert_px_per_unit);
        end
    end

    methods (Static)
        function obj = parse(stream)
            % parse (png.py 42-47): build the parser from `stream`'s chunks.
            chunks = mat2doc.image.Chunks_.from_stream(stream);
            obj = mat2doc.image.PngParser_(chunks);
        end

        function d = dpi_(units_specifier, px_per_unit)
            % _dpi (png.py 83-89): dpi from units_specifier + px_per_unit.
            %   Only units_specifier == 1 (meters) yields a real dpi; H4:
            %   px_per_unit truthy means non-zero. int(round(...)) is
            %   round-half-to-even (pyRound) then truncate-toward-zero (fix)
            %   (H6/H14). This is the docx png.py formula (the same 0.0254
            %   metre-to-inch constant), NOT any PIL variant.
            if units_specifier == 1 && px_per_unit ~= 0       % H4 truthiness
                d = fix(pyRound(px_per_unit * 0.0254));       % int(round(..))
                return
            end
            d = 72;
        end
    end
end

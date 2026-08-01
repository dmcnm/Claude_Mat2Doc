classdef Bmp < mat2doc.image.BaseImageHeader
% BMP Image header parser for BMP images.
%
%   Ported from python-docx v1.2.0: src/docx/image/bmp.py::Bmp (lines 6-43)

    methods
        function obj = Bmp(px_width, px_height, horz_dpi, vert_dpi)
            % Transparent pass-through to BaseImageHeader.
            obj@mat2doc.image.BaseImageHeader(px_width, px_height, horz_dpi, vert_dpi);
        end

        function ct = content_type(obj) %#ok<MANU>
            % content_type @property (bmp.py 26-30): unconditionally image/bmp.
            ct = mat2doc.image.MIME_TYPE.BMP;
        end

        function e = default_ext(obj) %#ok<MANU>
            % default_ext @property (bmp.py 32-35): always "bmp".
            e = "bmp";
        end
    end

    methods (Static)
        function obj = from_stream(stream)
            % from_stream (bmp.py 9-24): parse the BITMAPINFOHEADER. All fields
            %   are LITTLE-ENDIAN unsigned longs. Offsets mirror python-docx
            %   exactly (px_width read as an unsigned long, matching docx).
            stream_rdr = mat2doc.image.StreamReader( ...
                stream, mat2doc.image.StreamReader.LITTLE_ENDIAN);

            px_width = stream_rdr.read_long(18);              % 0x12
            px_height = stream_rdr.read_long(22);             % 0x16

            horz_px_per_meter = stream_rdr.read_long(38);     % 0x26
            vert_px_per_meter = stream_rdr.read_long(42);     % 0x2A

            horz_dpi = mat2doc.image.Bmp.dpi_(horz_px_per_meter);
            vert_dpi = mat2doc.image.Bmp.dpi_(vert_px_per_meter);

            obj = mat2doc.image.Bmp(px_width, px_height, horz_dpi, vert_dpi);
        end

        function d = dpi_(px_per_meter)
            % _dpi (bmp.py 37-43): integer pixels-per-inch from px_per_meter,
            %   defaulting to 96 when px_per_meter is zero.
            %
            %   docx formula (NOT the Mat2Ppt PIL oracle): a zero px_per_meter
            %   returns 96 (the BMP default), otherwise int(round(px_per_meter *
            %   0.0254)) with the 0.0254 metre-to-inch constant. This reverts the
            %   Mat2Ppt D-bmp-dpi PIL-equivalence variant (which used /39.3701 and
            %   mapped 0 -> 72); Mat2Doc's value oracle is python-docx, so the
            %   docx math is exact and carries NO deviation.
            %
            %   H6: int(round(...)) is round-half-to-even (pyRound) then
            %   truncate-toward-zero (fix); pyRound already yields an integral
            %   double, so fix is exact.
            if px_per_meter == 0
                d = 96;
                return
            end
            d = fix(pyRound(px_per_meter * 0.0254));          % int(round(..)) H6/H14
        end
    end
end

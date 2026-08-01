classdef Png < mat2doc.image.BaseImageHeader
% PNG Image header parser for PNG images.
%
%   Ported from python-docx v1.2.0: src/docx/image/png.py::Png (lines 7-32)

    methods
        function obj = Png(px_width, px_height, horz_dpi, vert_dpi)
            % Transparent pass-through to BaseImageHeader (MATLAB does not
            %   inherit constructors).
            obj@mat2doc.image.BaseImageHeader(px_width, px_height, horz_dpi, vert_dpi);
        end

        function ct = content_type(obj) %#ok<MANU>
            % content_type @property (png.py 10-14): unconditionally image/png.
            ct = mat2doc.image.MIME_TYPE.PNG;
        end

        function e = default_ext(obj) %#ok<MANU>
            % default_ext @property (png.py 16-19): always "png".
            e = "png";
        end
    end

    methods (Static)
        function obj = from_stream(stream)
            % from_stream (png.py 21-32): header properties parsed from `stream`.
            parser = mat2doc.image.PngParser_.parse(stream);

            px_width = parser.px_width;
            px_height = parser.px_height;
            horz_dpi = parser.horz_dpi;
            vert_dpi = parser.vert_dpi;

            obj = mat2doc.image.Png(px_width, px_height, horz_dpi, vert_dpi);
        end
    end
end

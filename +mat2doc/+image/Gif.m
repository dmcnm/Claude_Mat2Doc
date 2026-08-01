classdef Gif < mat2doc.image.BaseImageHeader
% GIF Image header parser for GIF images.
%
%   The GIF format carries no resolution (DPI) information, so both horizontal
%   and vertical DPI default to 72.
%
%   Ported from python-docx v1.2.0: src/docx/image/gif.py::Gif (lines 7-38)

    methods
        function obj = Gif(px_width, px_height, horz_dpi, vert_dpi)
            % Transparent pass-through to BaseImageHeader.
            obj@mat2doc.image.BaseImageHeader(px_width, px_height, horz_dpi, vert_dpi);
        end

        function ct = content_type(obj) %#ok<MANU>
            % content_type @property (gif.py 21-25): unconditionally image/gif.
            ct = mat2doc.image.MIME_TYPE.GIF;
        end

        function e = default_ext(obj) %#ok<MANU>
            % default_ext @property (gif.py 27-30): always "gif".
            e = "gif";
        end
    end

    methods (Static)
        function obj = from_stream(stream)
            % from_stream (gif.py 14-19): px dimensions from `stream`; dpi 72/72.
            %   GIF carries no resolution, so dpi is UNCONDITIONALLY 72/72 (no
            %   computation) -- matching python-docx exactly.
            [px_width, px_height] = mat2doc.image.Gif.dimensions_from_stream_(stream);
            obj = mat2doc.image.Gif(px_width, px_height, 72, 72);
        end

        function [px_width, px_height] = dimensions_from_stream_(stream)
            % _dimensions_from_stream (gif.py 32-38): the Logical Screen
            %   Descriptor holds width/height as two LITTLE-ENDIAN unsigned
            %   shorts ("<HH") at byte offset 6. Read directly off the stream,
            %   exactly as python-docx does (this parser does NOT use
            %   StreamReader).
            stream.seek(6);
            bytes_ = stream.read(4);
            % struct.unpack("<HH", bytes_): little-endian -> low byte first.
            px_width = double(bytes_(1)) + double(bytes_(2)) * 256;   % IDX 1-based
            px_height = double(bytes_(3)) + double(bytes_(4)) * 256;
        end
    end
end

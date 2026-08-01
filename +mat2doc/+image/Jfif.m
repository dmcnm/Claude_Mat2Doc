classdef Jfif < mat2doc.image.Jpeg
% JFIF Image header parser for the JFIF JPEG image format.
%
%   Un-stubbed at P7-2b (was a mat2doc:notYetPorted stub at P7-1a). Parses the
%   JFIF marker stream and takes px dimensions from the SOFn marker and dpi from
%   the APP0 marker (density units + x/y density). This is the plain python-docx
%   behavior: the APP0 dpi is used UNCONDITIONALLY (no PIL-style fallback to the
%   Exif/APP1 dpi for aspect-only density units -- the Mat2Ppt CLASS-J variant is
%   reverted here).
%
%   Ported from python-docx v1.2.0: src/docx/image/jpeg.py::Jfif (lines 47-61)

    methods
        function obj = Jfif(px_width, px_height, horz_dpi, vert_dpi)
            % Transparent pass-through to Jpeg/BaseImageHeader.
            obj@mat2doc.image.Jpeg(px_width, px_height, horz_dpi, vert_dpi);
        end
    end

    methods (Static)
        function obj = from_stream(stream)
            % from_stream (jpeg.py 50-61): dims from the SOF marker; dpi from the
            %   APP0 marker (density units + x/y density).
            markers = mat2doc.image.JfifMarkers_.from_stream(stream);

            px_width = markers.sof.px_width;
            px_height = markers.sof.px_height;
            horz_dpi = markers.app0.horz_dpi;
            vert_dpi = markers.app0.vert_dpi;

            obj = mat2doc.image.Jfif(px_width, px_height, horz_dpi, vert_dpi);
        end
    end
end

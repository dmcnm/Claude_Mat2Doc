classdef Exif < mat2doc.image.Jpeg
% EXIF Image header parser for the Exif JPEG image format.
%
%   Un-stubbed at P7-2b (was a mat2doc:notYetPorted stub at P7-1a). Identical to
%   Jfif except the dpi comes from the APP1 (Exif) marker rather than APP0. The
%   APP1 dpi is read via the embedded-TIFF path (App1Marker_ -> Tiff.from_stream),
%   the python-docx behavior -- NOT the Mat2Ppt PIL-oracle CLASS-E variant, which
%   is reverted here per the boundary audit (C1: Tiff ported first at P7-2a so the
%   Exif dpi comes from the real Tiff parser).
%
%   Ported from python-docx v1.2.0: src/docx/image/jpeg.py::Exif (lines 29-44)

    methods
        function obj = Exif(px_width, px_height, horz_dpi, vert_dpi)
            % Transparent pass-through to Jpeg/BaseImageHeader.
            obj@mat2doc.image.Jpeg(px_width, px_height, horz_dpi, vert_dpi);
        end
    end

    methods (Static)
        function obj = from_stream(stream)
            % from_stream (jpeg.py 32-44): dims from SOF, dpi from APP1 (Exif).
            markers = mat2doc.image.JfifMarkers_.from_stream(stream);

            px_width = markers.sof.px_width;
            px_height = markers.sof.px_height;
            horz_dpi = markers.app1.horz_dpi;
            vert_dpi = markers.app1.vert_dpi;

            obj = mat2doc.image.Exif(px_width, px_height, horz_dpi, vert_dpi);
        end
    end
end

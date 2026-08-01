classdef Tiff < mat2doc.image.BaseImageHeader
% TIFF Image header parser for TIFF images (both big- and little-endian).
%
%   Un-stubbed at P7-2a (was a mat2doc:notYetPorted stub at P7-1a). Parses the
%   main Image File Directory (IFD) to extract px_width / px_height and the
%   XResolution / YResolution / ResolutionUnit dpi. Also used as the embedded
%   parser for the Exif (APP1) segment of an Exif JPEG (P7-2b jpeg's Exif marker
%   calls Tiff.from_stream on the exif segment).
%
%   Ported from python-docx v1.2.0: src/docx/image/tiff.py::Tiff (lines 6-34)

    methods
        function obj = Tiff(px_width, px_height, horz_dpi, vert_dpi)
            % Transparent pass-through to BaseImageHeader (MATLAB does not
            %   inherit constructors).
            obj@mat2doc.image.BaseImageHeader(px_width, px_height, horz_dpi, vert_dpi);
        end

        function ct = content_type(obj) %#ok<MANU>
            % content_type @property (tiff.py 12-16): unconditionally image/tiff.
            ct = mat2doc.image.MIME_TYPE.TIFF;
        end

        function e = default_ext(obj) %#ok<MANU>
            % default_ext @property (tiff.py 18-21): always "tiff".
            e = "tiff";
        end
    end

    methods (Static)
        function obj = from_stream(stream)
            % from_stream (tiff.py 23-34): header properties parsed from `stream`.
            parser = mat2doc.image.TiffParser_.parse(stream);

            px_width = parser.px_width;
            px_height = parser.px_height;
            horz_dpi = parser.horz_dpi;
            vert_dpi = parser.vert_dpi;

            obj = mat2doc.image.Tiff(px_width, px_height, horz_dpi, vert_dpi);
        end
    end
end

classdef Jpeg < mat2doc.image.BaseImageHeader
% JPEG Base class for the JFIF and Exif JPEG header parsers.
%
%   content_type is unconditionally image/jpeg and default_ext is always "jpg"
%   for both JPEG sub-formats; the concrete from_stream lives on the Jfif / Exif
%   subclasses.
%
%   Ported from python-docx v1.2.0: src/docx/image/jpeg.py::Jpeg (lines 14-26)

    methods
        function obj = Jpeg(px_width, px_height, horz_dpi, vert_dpi)
            % Transparent pass-through to BaseImageHeader (MATLAB does not
            %   inherit constructors).
            obj@mat2doc.image.BaseImageHeader(px_width, px_height, horz_dpi, vert_dpi);
        end

        function ct = content_type(obj) %#ok<MANU>
            % content_type @property (jpeg.py 17-21): unconditionally image/jpeg.
            ct = mat2doc.image.MIME_TYPE.JPEG;
        end

        function e = default_ext(obj) %#ok<MANU>
            % default_ext @property (jpeg.py 23-26): always "jpg".
            e = "jpg";
        end
    end
end

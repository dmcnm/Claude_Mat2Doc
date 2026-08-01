classdef MIME_TYPE
% MIME_TYPE Image content types.
%
%   Access as mat2doc.image.MIME_TYPE.PNG etc. String values byte-identical to
%   python-docx constants.py. These are the content types the docx image parsers
%   assign to a recognized image stream.
%
%   Example:
%       disp(mat2doc.image.MIME_TYPE.PNG)   % "image/png"
%
%   Ported from python-docx v1.2.0: src/docx/image/constants.py::MIME_TYPE
%   (lines 100-107)

    properties (Constant)
        BMP = "image/bmp"
        GIF = "image/gif"
        JPEG = "image/jpeg"
        PNG = "image/png"
        TIFF = "image/tiff"
    end
end

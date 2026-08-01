classdef s0083_DoubleHeader < mat2doc.image.BaseImageHeader
% S0083_DOUBLEHEADER Gate-3 test-double image header for s0083.
%
%   A concrete BaseImageHeader whose px_width/px_height/horz_dpi/vert_dpi and
%   content_type/default_ext are SEEDED at construction from the python-docx
%   oracle's OWN parse (frozen in references/s0083/probe.json). It exists solely
%   because the P7-1a FORMAT PARSERS (Png/Jfif/Exif/Gif/Tiff/Bmp) are stubbed
%   until P7-1b/P7-2, so a real Image.from_blob(real_png) cannot yet produce the
%   header on the MATLAB side. Seeding lets the Image CORE's derived surface
%   (width/height Length math, ext, filename, sha1 end-to-end, content_type
%   delegation) be proven value-identical to the oracle NOW; the raw header
%   PARSE (does MATLAB read 860 px out of the PNG bytes?) is re-proven end-to-end
%   at P7-1b (VERIFY V-2).
%
%   Overrides content_type() and default_ext() -- which BaseImageHeader defines
%   as abstract METHODS raising NotImplementedError -- to return the seeds.
%
%   Gate-3 validation scaffold (mso-validator); NOT toolbox code.

    properties (Access = private)
        content_type_
        default_ext_
    end

    methods
        function obj = s0083_DoubleHeader(px_w, px_h, hdpi, vdpi, ct, ext)
            obj@mat2doc.image.BaseImageHeader(px_w, px_h, hdpi, vdpi);
            obj.content_type_ = ct;
            obj.default_ext_ = ext;
        end

        function v = content_type(obj)
            v = obj.content_type_;
        end

        function v = default_ext(obj)
            v = obj.default_ext_;
        end
    end
end

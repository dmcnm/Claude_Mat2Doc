classdef TIFF_TAG
% TIFF_TAG Tag codes for TIFF Image File Directory (IFD) entries.
%
%   The five tag codes the TIFF parser looks up (image dimensions and
%   resolution). Held as doubles carrying the exact integer tag code; IFD tag
%   codes are read as unsigned shorts, so the parser's `in`/get lookups compare
%   these against read_short results. tag_names is a debug-only lookup.
%
%   Ported from python-docx v1.2.0: src/docx/image/constants.py::TIFF_TAG
%   (lines 139-172)

    properties (Constant)
        IMAGE_WIDTH     = 256    % 0x0100
        IMAGE_LENGTH    = 257    % 0x0101
        X_RESOLUTION    = 282    % 0x011A
        Y_RESOLUTION    = 283    % 0x011B
        RESOLUTION_UNIT = 296    % 0x0128
    end

    methods (Static)
        function nm = tag_name(tag_code)
            % tag_names lookup (constants.py 148-172). Debug-only (unexercised on
            %   any live path); raises KeyError for an unmapped code.
            m = mat2doc.image.TIFF_TAG.tag_names_();
            if isKey(m, tag_code)
                nm = m(tag_code);
            else
                error("mat2doc:KeyError", "%s", string(tag_code));
            end
        end
    end

    methods (Static, Access = private)
        function m = tag_names_()
            % Lookup-only dictionary (H11: order-inert, key access only).
            keys_ = [254 256 257 258 259 262 270 271 272 273 274 277 279 ...
                     282 283 284 296 305 306 531 34665 34853 50341];
            vals_ = ["NewSubfileType" "ImageWidth" "ImageLength" "BitsPerSample" ...
                     "Compression" "PhotometricInterpretation" "ImageDescription" ...
                     "Make" "Model" "StripOffsets" "Orientation" "SamplesPerPixel" ...
                     "StripByteCounts" "XResolution" "YResolution" ...
                     "PlanarConfiguration" "ResolutionUnit" "Software" "DateTime" ...
                     "YCbCrPositioning" "ExifTag" "GPS IFD" "PrintImageMatching"];
            m = dictionary(keys_, vals_);
        end
    end
end

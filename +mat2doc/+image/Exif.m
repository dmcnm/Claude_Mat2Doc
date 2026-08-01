classdef Exif < mat2doc.image.BaseImageHeader
% EXIF Image header parser for Exif (JPEG) images. STUB (not yet ported).
%
%   The concrete Exif BaseImageHeader subclass (whose dpi comes from the embedded
%   TIFF parse -- reverting the Mat2Ppt PIL-oracle CLASS-E variant per the
%   boundary audit) is owned by the P7-2 jpeg WP. This stub exists so
%   ImageHeaderFactory_'s signature table (which resolves
%   @mat2doc.image.Exif.from_stream at call time) dispatches to a named target;
%   invoking it raises mat2doc:notYetPorted.
%
%   Target symbol: python-docx v1.2.0 src/docx/image/jpeg.py::Exif (owning WP: P7-2 jpeg)

    methods (Static)
        function obj = from_stream(stream) %#ok<INUSD,STOUT>
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.image.Exif.from_stream is not yet ported (owning WP: P7-2 jpeg).");
        end
    end
end

classdef Bmp < mat2doc.image.BaseImageHeader
% BMP Image header parser for BMP images. STUB (not yet ported).
%
%   The concrete BMP BaseImageHeader subclass is owned by P7-1b. Its dpi math
%   must follow the python-docx bmp.py::Bmp._dpi formula (int(round(ppm*0.0254)),
%   ppm==0 -> 96), NOT the Mat2Ppt PIL-oracle variant -- see the P6->P7 boundary
%   audit. This stub exists so ImageHeaderFactory_'s signature table (which
%   resolves @mat2doc.image.Bmp.from_stream at call time) dispatches to a named
%   target; invoking it raises mat2doc:notYetPorted.
%
%   Target symbol: python-docx v1.2.0 src/docx/image/bmp.py::Bmp (owning WP: P7-1b)

    methods (Static)
        function obj = from_stream(stream) %#ok<INUSD,STOUT>
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.image.Bmp.from_stream is not yet ported (owning WP: P7-1b).");
        end
    end
end

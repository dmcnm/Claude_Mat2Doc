classdef Tiff < mat2doc.image.BaseImageHeader
% TIFF Image header parser for TIFF images. STUB (not yet ported).
%
%   The concrete TIFF BaseImageHeader subclass and its IFD machinery are owned by
%   the P7-2 tiff WP. Its dpi math must follow python-docx tiff.py::_TiffParser
%   ._dpi (per-axis, tag-absent -> 72), NOT the Mat2Ppt PIL-oracle variant -- see
%   boundary audit. This stub exists so ImageHeaderFactory_'s signature table
%   (which resolves @mat2doc.image.Tiff.from_stream at call time) dispatches to a
%   named target; invoking it raises mat2doc:notYetPorted.
%
%   Target symbol: python-docx v1.2.0 src/docx/image/tiff.py::Tiff (owning WP: P7-2 tiff)

    methods (Static)
        function obj = from_stream(stream) %#ok<INUSD,STOUT>
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.image.Tiff.from_stream is not yet ported (owning WP: P7-2 tiff).");
        end
    end
end

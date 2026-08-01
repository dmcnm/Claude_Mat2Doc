classdef Png < mat2doc.image.BaseImageHeader
% PNG Image header parser for PNG images. STUB (not yet ported).
%
%   The concrete PNG BaseImageHeader subclass and its chunk machinery
%   (PngParser_, Chunk_, ChunkFactory_, ChunkParser_, Chunks_, IHDRChunk_,
%   pHYsChunk_) are owned by P7-1b. This stub exists so ImageHeaderFactory_'s
%   signature table (which resolves @mat2doc.image.Png.from_stream at call time)
%   dispatches to a named target; invoking it raises mat2doc:notYetPorted.
%
%   Target symbol: python-docx v1.2.0 src/docx/image/png.py::Png (owning WP: P7-1b)

    methods (Static)
        function obj = from_stream(stream) %#ok<INUSD,STOUT>
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.image.Png.from_stream is not yet ported (owning WP: P7-1b).");
        end
    end
end

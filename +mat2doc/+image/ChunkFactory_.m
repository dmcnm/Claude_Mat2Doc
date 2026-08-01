function chunk = ChunkFactory_(chunk_type, stream_rdr, offset)
% CHUNKFACTORY_ Return a chunk object appropriate to `chunk_type`.
%
%   Mirrors python-docx's chunk_cls_map dispatch: IHDR -> IHDRChunk_, pHYs ->
%   pHYsChunk_, everything else -> the default Chunk_ (which only records its
%   type). Each class supplies its own from_offset static (MATLAB does not
%   dispatch inherited statics), matching the Python per-class from_offset.
%
%   UNDERSCORE ROTATION (design.md section 2): python-docx _ChunkFactory (a
%   module-private function) -> ChunkFactory_.
%
%   Ported from python-docx v1.2.0: src/docx/image/png.py::_ChunkFactory
%   (lines 168-176)

PCT = mat2doc.image.PNG_CHUNK_TYPE;
switch chunk_type
    case PCT.IHDR
        chunk = mat2doc.image.IHDRChunk_.from_offset(chunk_type, stream_rdr, offset);
    case PCT.pHYs
        chunk = mat2doc.image.pHYsChunk_.from_offset(chunk_type, stream_rdr, offset);
    otherwise
        chunk = mat2doc.image.Chunk_.from_offset(chunk_type, stream_rdr, offset);
end
end

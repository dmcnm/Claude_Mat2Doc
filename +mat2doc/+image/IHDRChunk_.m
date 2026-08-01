classdef IHDRChunk_ < mat2doc.image.Chunk_
% IHDRCHUNK_ IHDR chunk; contains the image dimensions.
%
%   UNDERSCORE ROTATION (design.md section 2): python-docx _IHDRChunk ->
%   IHDRChunk_.
%
%   Ported from python-docx v1.2.0: src/docx/image/png.py::_IHDRChunk
%   (lines 200-222)

    properties (Access = private)
        px_width_
        px_height_
    end

    methods
        function obj = IHDRChunk_(chunk_type, px_width, px_height)
            % Ported from _IHDRChunk.__init__ (png.py 203-206).
            obj@mat2doc.image.Chunk_(chunk_type);
            obj.px_width_ = px_width;
            obj.px_height_ = px_height;
        end

        function v = px_width(obj)
            % px_width @property (png.py 216-218).
            v = obj.px_width_;
        end

        function v = px_height(obj)
            % px_height @property (png.py 220-222).
            v = obj.px_height_;
        end
    end

    methods (Static)
        function obj = from_offset(chunk_type, stream_rdr, offset)
            % from_offset (png.py 208-214): dimensions from the IHDR data at
            %   `offset` (BIG_ENDIAN longs: px_width at +0, px_height at +4).
            px_width = stream_rdr.read_long(offset);
            px_height = stream_rdr.read_long(offset, 4);
            obj = mat2doc.image.IHDRChunk_(chunk_type, px_width, px_height);
        end
    end
end

classdef Chunk_ < handle
% CHUNK_ Base class for specific PNG chunk types; also the default chunk type.
%
%   UNDERSCORE ROTATION (design.md section 2): python-docx _Chunk -> Chunk_.
%
%   Ported from python-docx v1.2.0: src/docx/image/png.py::_Chunk (lines 179-197)

    properties (Access = protected)
        chunk_type_
    end

    methods
        function obj = Chunk_(chunk_type)
            % Ported from _Chunk.__init__ (png.py 185-187).
            obj.chunk_type_ = chunk_type;
        end

        function v = type_name(obj)
            % type_name @property (png.py 194-197): chunk type name, e.g.
            %   "IHDR", "pHYs".
            v = obj.chunk_type_;
        end
    end

    methods (Static)
        function obj = from_offset(chunk_type, stream_rdr, offset) %#ok<INUSD>
            % from_offset (png.py 189-192): a default chunk that only knows its
            %   type (stream_rdr / offset unused).
            obj = mat2doc.image.Chunk_(chunk_type);
        end
    end
end

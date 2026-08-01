classdef pHYsChunk_ < mat2doc.image.Chunk_
% PHYSCHUNK_ pHYs chunk; contains the image dpi (resolution) information.
%
%   UNDERSCORE ROTATION (design.md section 2): python-docx _pHYsChunk ->
%   pHYsChunk_.
%
%   Ported from python-docx v1.2.0: src/docx/image/png.py::_pHYsChunk
%   (lines 225-253)

    properties (Access = private)
        horz_px_per_unit_
        vert_px_per_unit_
        units_specifier_
    end

    methods
        function obj = pHYsChunk_(chunk_type, horz_px_per_unit, vert_px_per_unit, units_specifier)
            % Ported from _pHYsChunk.__init__ (png.py 228-232). Argument order
            %   matches Python exactly.
            obj@mat2doc.image.Chunk_(chunk_type);
            obj.horz_px_per_unit_ = horz_px_per_unit;
            obj.vert_px_per_unit_ = vert_px_per_unit;
            obj.units_specifier_ = units_specifier;
        end

        function v = horz_px_per_unit(obj)
            % horz_px_per_unit @property (png.py 243-245).
            v = obj.horz_px_per_unit_;
        end

        function v = vert_px_per_unit(obj)
            % vert_px_per_unit @property (png.py 247-249).
            v = obj.vert_px_per_unit_;
        end

        function v = units_specifier(obj)
            % units_specifier @property (png.py 251-253).
            v = obj.units_specifier_;
        end
    end

    methods (Static)
        function obj = from_offset(chunk_type, stream_rdr, offset)
            % from_offset (png.py 234-241): resolution from the pHYs data at
            %   `offset` (BIG_ENDIAN: horz long +0, vert long +4, units byte +8).
            horz_px_per_unit = stream_rdr.read_long(offset);
            vert_px_per_unit = stream_rdr.read_long(offset, 4);
            units_specifier = stream_rdr.read_byte(offset, 8);
            obj = mat2doc.image.pHYsChunk_(chunk_type, horz_px_per_unit, vert_px_per_unit, units_specifier);
        end
    end
end

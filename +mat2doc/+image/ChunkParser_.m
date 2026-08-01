classdef ChunkParser_ < handle
% CHUNKPARSER_ Extracts chunks from a PNG image stream.
%
%   UNDERSCORE ROTATION (design.md section 2): python-docx _ChunkParser ->
%   ChunkParser_.
%
%   Ported from python-docx v1.2.0: src/docx/image/png.py::_ChunkParser
%   (lines 129-165)

    properties (Access = private)
        stream_rdr_      % a StreamReader (BIG_ENDIAN)
    end

    methods
        function obj = ChunkParser_(stream_rdr)
            % Ported from _ChunkParser.__init__ (png.py 132-134).
            obj.stream_rdr_ = stream_rdr;
        end

        function chunks = iter_chunks(obj)
            % iter_chunks (png.py 143-148): a chunk object per chunk, in stream
            %   order. H9: the Python generator is realized as a precomputed
            %   1xN cell array (no mutation occurs during iteration, so dropping
            %   laziness is unobservable).
            offsets = obj.iter_chunk_offsets_();          % Kx2 cell {type, off}
            chunks = cell(1, size(offsets, 1));
            for i = 1:size(offsets, 1)
                chunk_type = offsets{i, 1};
                offset = offsets{i, 2};
                chunks{i} = mat2doc.image.ChunkFactory_(chunk_type, obj.stream_rdr_, offset);
            end
        end
    end

    methods (Static)
        function obj = from_stream(stream)
            % from_stream (png.py 136-141): wrap `stream` in a BIG_ENDIAN reader.
            stream_rdr = mat2doc.image.StreamReader( ...
                stream, mat2doc.image.StreamReader.BIG_ENDIAN);
            obj = mat2doc.image.ChunkParser_(stream_rdr);
        end
    end

    methods (Access = private)
        function offsets = iter_chunk_offsets_(obj)
            % _iter_chunk_offsets (png.py 150-165): a (chunk_type, chunk_offset)
            %   pair per chunk; iteration stops after IEND. H9: precomputed into
            %   a Kx2 cell {chunk_type(string), data_offset(double)}.
            offsets = cell(0, 2);
            chunk_offset = 8;
            while true
                chunk_data_len = obj.stream_rdr_.read_long(chunk_offset);
                chunk_type = obj.stream_rdr_.read_str(4, chunk_offset, 4);
                data_offset = chunk_offset + 8;
                offsets(end + 1, :) = {chunk_type, data_offset}; %#ok<AGROW>
                if chunk_type == "IEND"
                    break
                end
                % incr offset for chunk len long, chunk type, chunk data, and CRC
                chunk_offset = chunk_offset + 4 + 4 + chunk_data_len + 4;
            end
        end
    end
end

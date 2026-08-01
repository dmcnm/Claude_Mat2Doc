classdef Chunks_ < handle
% CHUNKS_ Collection of the chunks parsed from a PNG image stream.
%
%   UNDERSCORE ROTATION (design.md section 2): python-docx _Chunks -> Chunks_.
%
%   NONE IDIOM (H3): _find_first returns None when no chunk matches; Mat2Doc has
%   no shared isNone helper, so the None-guard is the inline isequal(x, []) test
%   (ratified none-idiom decision 2026-07-26).
%
%   Ported from python-docx v1.2.0: src/docx/image/png.py::_Chunks (lines 92-126)

    properties (Access = private)
        chunks_          % 1xN cell array of chunk objects (the parsed list)
    end

    methods
        function obj = Chunks_(chunk_iterable)
            % Ported from _Chunks.__init__ (png.py 95-97): list(chunk_iterable).
            %   `chunk_iterable` is already a 1xN cell array (see from_stream).
            obj.chunks_ = chunk_iterable;
        end

        function v = IHDR(obj)
            % IHDR @property (png.py 106-113): first IHDR chunk, else
            %   InvalidImageStreamError.
            match = @(chunk) chunk.type_name == mat2doc.image.PNG_CHUNK_TYPE.IHDR;
            v = obj.find_first_(match);
            if isequal(v, [])                                 % None (H3)
                error("mat2doc:InvalidImageStreamError", "%s", ...
                    "no IHDR chunk in PNG image");
            end
        end

        function v = pHYs(obj)
            % pHYs @property (png.py 115-119): first pHYs chunk, or None ([]).
            match = @(chunk) chunk.type_name == mat2doc.image.PNG_CHUNK_TYPE.pHYs;
            v = obj.find_first_(match);
        end
    end

    methods (Static)
        function obj = from_stream(stream)
            % from_stream (png.py 99-104): parse the chunks in `stream`.
            chunk_parser = mat2doc.image.ChunkParser_.from_stream(stream);
            chunks = chunk_parser.iter_chunks();   % 1xN cell (H9: precomputed)
            obj = mat2doc.image.Chunks_(chunks);
        end
    end

    methods (Access = private)
        function v = find_first_(obj, match)
            % _find_first (png.py 121-126): first chunk in stream order for
            %   which `match` returns true; None ([]) if none.
            for i = 1:numel(obj.chunks_)
                chunk = obj.chunks_{i};
                if match(chunk)
                    v = chunk;
                    return
                end
            end
            v = [];        % None (H3)
        end
    end
end

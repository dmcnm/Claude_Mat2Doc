classdef BytesIO < handle
% BYTESIO Minimal seekable in-memory byte stream (the io.BytesIO analogue).
%
%   stream = MAT2DOC.IMAGE.BYTESIO(blob) wraps a uint8 byte vector `blob` in a
%   seekable/readable cursor object, substituting Python's io.BytesIO for the
%   +image header parsers. Only the surface those parsers exercise is provided:
%   seek(pos), read(n) / read() (read-all), and tell(). The cursor position is
%   0-based, mirroring Python file positions exactly (the parsers compute
%   absolute byte offsets, e.g. StreamReader base_offset + base + offset).
%
%   read(n) returns up to n bytes (fewer at EOF, like Python's stream.read),
%   never raising on a short read; the StreamReader layer is what raises
%   UnexpectedEndOfFileError when a fixed-width field runs past EOF. read() with
%   no count returns all remaining bytes (used by Image.from_file's file-like
%   branch). Bytes are returned as a 1xN uint8 row (empty 1x0 uint8 at EOF).
%
%   Mat2Doc infrastructure (no python-docx counterpart; substitutes io.BytesIO).
%   Mandated by design.md section 8 (binary byte I/O).

    properties (Access = private)
        buf_          % 1xN uint8 buffer
        pos_          % 0-based cursor position
    end

    methods
        function obj = BytesIO(blob)
            obj.buf_ = uint8(blob(:))';     % normalize to a 1xN uint8 row
            obj.pos_ = 0;
        end

        function seek(obj, pos)
            % seek to an absolute 0-based byte position (may be past EOF, as
            % Python allows; a subsequent read then returns empty).
            obj.pos_ = pos;
        end

        function out = read(obj, n)
            % read up to `n` bytes from the current position (all remaining
            % bytes when `n` is omitted); advance the cursor by the count read.
            arguments
                obj
                n = []      % None -> read all remaining
            end
            navail = numel(obj.buf_) - obj.pos_;
            if isempty(n)
                k = navail;
            else
                k = min(n, navail);
            end
            if k <= 0
                out = uint8([]);
                out = reshape(out, 1, 0);   % 1x0 uint8
                return
            end
            out = obj.buf_(obj.pos_ + 1 : obj.pos_ + k);    % IDX 0-based -> 1-based
            obj.pos_ = obj.pos_ + k;
        end

        function p = tell(obj)
            % return the current 0-based cursor position.
            p = obj.pos_;
        end
    end
end

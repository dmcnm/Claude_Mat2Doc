classdef MarkerFinder_ < handle
% MARKERFINDER_ Finds the next JFIF marker in a stream.
%
%   UNDERSCORE ROTATION (design.md section 2): python-docx _MarkerFinder ->
%   MarkerFinder_.
%
%   All +/-1 arithmetic below is BYTE-ADDRESS arithmetic (Python's own): `tell()
%   - 1` recovers the offset of the byte just read (tell points past it), and
%   `position + 1` steps to the byte after a marker code. These are NOT H1
%   0/1-index shifts (the StreamReader/BytesIO cursor is 0-based, matching
%   Python file positions exactly).
%
%   Ported from python-docx v1.2.0: src/docx/image/jpeg.py::_MarkerFinder
%   (lines 155-225)

    properties (Access = private)
        stream_       % a StreamReader
    end

    methods
        function obj = MarkerFinder_(stream)
            % Ported from _MarkerFinder.__init__ (jpeg.py 158-160).
            obj.stream_ = stream;
        end

        function [marker_code, segment_offset] = next(obj, start)
            % next (jpeg.py 167-187): locate the first marker after offset `start`.
            %   segment_offset points just past the 2-byte marker code (the start
            %   of the marker segment).
            position = start;
            while true
                % skip over any non-0xFF bytes
                position = obj.offset_of_next_ff_byte_(position);
                % skip over any 0xFF padding bytes
                [position, byte_] = obj.next_non_ff_byte_(position + 1);
                % 'FF 00' is not a marker; start over if found
                if byte_ == uint8(0)
                    continue
                end
                % this is a marker; gather return values and break out of scan
                marker_code = byte_;
                segment_offset = position + 1;   % byte address
                break
            end
        end
    end

    methods (Static)
        function obj = from_stream(stream)
            % from_stream (jpeg.py 162-165).
            obj = mat2doc.image.MarkerFinder_(stream);
        end
    end

    methods (Access = private)
        function [offset_of_non_ff_byte, byte_] = next_non_ff_byte_(obj, start)
            % _next_non_ff_byte (jpeg.py 189-201): next byte at/after `start` that
            %   is not 0xFF, with its offset.
            obj.stream_.seek(start);
            byte_ = obj.read_byte_();
            while byte_ == uint8(255)
                byte_ = obj.read_byte_();
            end
            offset_of_non_ff_byte = obj.stream_.tell() - 1;   % byte address
        end

        function offset_of_ff_byte = offset_of_next_ff_byte_(obj, start)
            % _offset_of_next_ff_byte (jpeg.py 203-215): offset of the next 0xFF
            %   byte at/after `start` (returns `start` if it already points at one).
            obj.stream_.seek(start);
            byte_ = obj.read_byte_();
            while byte_ ~= uint8(255)
                byte_ = obj.read_byte_();
            end
            offset_of_ff_byte = obj.stream_.tell() - 1;       % byte address
        end

        function byte_ = read_byte_(obj)
            % _read_byte (jpeg.py 217-225): next byte, or raise at EOF. Python
            %   raises a bare Exception("unexpected end of file") here
            %   (# pragma: no cover -- a corrupt/truncated stream that a valid
            %   JPEG never reaches, since scanning stops at SOS). Python's own
            %   identifier is not a docx image exception, so the id is
            %   port-authored; the message is verbatim.
            byte_ = obj.stream_.read(1);
            if isempty(byte_)                  % Python `if not byte_:` (H4)
                error("mat2doc:Exception", "%s", "unexpected end of file");
            end
        end
    end
end

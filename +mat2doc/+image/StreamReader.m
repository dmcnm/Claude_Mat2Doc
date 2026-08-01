classdef StreamReader < handle
% STREAMREADER Structured binary reads over a seekable byte stream.
%
%   rdr = MAT2DOC.IMAGE.STREAMREADER(stream, byte_order[, base_offset]) wraps a
%   BytesIO-like `stream` to provide fixed-width integer / string reads at
%   computed byte offsets. Byte-order is configurable (BIG_ENDIAN ">" or
%   LITTLE_ENDIAN "<"); `base_offset` is added to every base value to locate the
%   actual read position (used by the JPEG/TIFF parsers at P7-2; 0 for PNG/BMP).
%
%   All integer reads are UNSIGNED (Python struct "B"/"L"/"H"): read_byte (1
%   byte), read_short (2 bytes), read_long (4 bytes). Values are held in doubles
%   (all < 2^32, exact). read_str decodes `char_count` bytes as UTF-8 (H2).
%   A read that runs past EOF raises mat2doc:UnexpectedEndOfFileError.
%
%   Ported from python-docx v1.2.0: src/docx/image/helpers.py::StreamReader
%   (lines 9-86)

    properties (Constant)
        BIG_ENDIAN = ">"       % helpers.py BIG_ENDIAN
        LITTLE_ENDIAN = "<"    % helpers.py LITTLE_ENDIAN
    end

    properties (Access = private)
        stream_          % underlying seekable byte stream (BytesIO)
        byte_order_      % ">" or "<"
        base_offset_     % added to every base value
    end

    methods
        function obj = StreamReader(stream, byte_order, base_offset)
            % Ported from StreamReader.__init__ (helpers.py 16-20).
            arguments
                stream
                byte_order
                base_offset = 0
            end
            obj.stream_ = stream;
            % self._byte_order = LITTLE_ENDIAN if byte_order == LITTLE_ENDIAN
            %                    else BIG_ENDIAN
            if byte_order == mat2doc.image.StreamReader.LITTLE_ENDIAN
                obj.byte_order_ = mat2doc.image.StreamReader.LITTLE_ENDIAN;
            else
                obj.byte_order_ = mat2doc.image.StreamReader.BIG_ENDIAN;
            end
            obj.base_offset_ = base_offset;
        end

        function out = read(obj, count)
            % read (helpers.py 22-24): pass-through read() call.
            out = obj.stream_.read(count);
        end

        function v = read_byte(obj, base, offset)
            % read_byte (helpers.py 26-33): unsigned byte ("B") at
            %   base_offset + base + offset.
            arguments
                obj
                base
                offset = 0
            end
            b = obj.read_bytes_(1, base, offset);
            v = double(b);                          % single byte, endian-agnostic
        end

        function v = read_long(obj, base, offset)
            % read_long (helpers.py 35-44): unsigned 32-bit ("<L"/">L") at
            %   base_offset + base + offset, per this instance's byte order.
            arguments
                obj
                base
                offset = 0
            end
            b = obj.read_bytes_(4, base, offset);
            v = obj.unpack_uint_(b);
        end

        function v = read_short(obj, base, offset)
            % read_short (helpers.py 46-50): unsigned 16-bit ("<H"/">H").
            arguments
                obj
                base
                offset = 0
            end
            b = obj.read_bytes_(2, base, offset);
            v = obj.unpack_uint_(b);
        end

        function s = read_str(obj, char_count, base, offset)
            % read_str (helpers.py 52-63): `char_count` bytes decoded as UTF-8.
            arguments
                obj
                char_count
                base
                offset = 0
            end
            b = obj.read_bytes_(char_count, base, offset);
            s = string(native2unicode(uint8(b), "UTF-8"));   % H2: bytes->text
        end

        function seek(obj, base, offset)
            % seek (helpers.py 65-67): position = base_offset + base + offset.
            arguments
                obj
                base
                offset = 0
            end
            location = obj.base_offset_ + base + offset;
            obj.stream_.seek(location);
        end

        function p = tell(obj)
            % tell (helpers.py 69-71): pass-through tell() call.
            p = obj.stream_.tell();
        end
    end

    methods (Access = private)
        function bytes_ = read_bytes_(obj, byte_count, base, offset)
            % _read_bytes (helpers.py 73-78): seek then read exactly
            %   byte_count bytes; a short read is EOF.
            obj.seek(base, offset);
            bytes_ = obj.stream_.read(byte_count);
            if numel(bytes_) < byte_count
                error("mat2doc:UnexpectedEndOfFileError", "%s", ...
                    "EOF was unexpectedly encountered while reading an image stream");
            end
        end

        function v = unpack_uint_(obj, b)
            % struct.unpack of an unsigned integer from `b` (1xN uint8) per this
            %   instance's byte order. Little-endian: least-significant byte
            %   first; big-endian: most-significant byte first.
            b = double(b);
            n = numel(b);
            if obj.byte_order_ == mat2doc.image.StreamReader.LITTLE_ENDIAN
                weights = 256 .^ (0:n-1);
            else
                weights = 256 .^ (n-1:-1:0);
            end
            v = sum(b .* weights);
        end
    end
end

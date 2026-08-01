classdef AsciiIfdEntry_ < mat2doc.image.IfdEntry_
% ASCIIIFDENTRY_ IFD entry holding a NULL-terminated ASCII string.
%
%   UNDERSCORE ROTATION (design.md section 2): python-docx _AsciiIfdEntry ->
%   AsciiIfdEntry_.
%
%   Ported from python-docx v1.2.0: src/docx/image/tiff.py::_AsciiIfdEntry
%   (lines 229-239)

    methods
        function obj = AsciiIfdEntry_(tag_code, value)
            % Transparent pass-through to IfdEntry_.
            obj@mat2doc.image.IfdEntry_(tag_code, value);
        end
    end

    methods (Static)
        function obj = from_stream(stream_rdr, offset)
            % Common from_stream (tiff.py 196-208), dispatching to this class's
            %   parse_value_ and constructing this class.
            tag_code = stream_rdr.read_short(offset, 0);
            value_count = stream_rdr.read_long(offset, 4);
            value_offset = stream_rdr.read_long(offset, 8);
            value = mat2doc.image.AsciiIfdEntry_.parse_value_( ...
                stream_rdr, offset, value_count, value_offset);
            obj = mat2doc.image.AsciiIfdEntry_(tag_code, value);
        end

        function v = parse_value_(stream_rdr, offset, value_count, value_offset) %#ok<INUSD>
            % _parse_value (tiff.py 232-239): the ASCII string at `value_offset`.
            %   value_count includes the terminating NUL, so read value_count-1
            %   chars (byte-address arithmetic, Python's own; not an H1 shift).
            v = stream_rdr.read_str(value_count - 1, value_offset);
        end
    end
end

classdef ShortIfdEntry_ < mat2doc.image.IfdEntry_
% SHORTIFDENTRY_ IFD entry expressed as a short (2-byte) integer.
%
%   UNDERSCORE ROTATION (design.md section 2): python-docx _ShortIfdEntry ->
%   ShortIfdEntry_.
%
%   Ported from python-docx v1.2.0: src/docx/image/tiff.py::_ShortIfdEntry
%   (lines 242-255)

    methods
        function obj = ShortIfdEntry_(tag_code, value)
            % Transparent pass-through to IfdEntry_.
            obj@mat2doc.image.IfdEntry_(tag_code, value);
        end
    end

    methods (Static)
        function obj = from_stream(stream_rdr, offset)
            % Common from_stream (tiff.py 196-208).
            tag_code = stream_rdr.read_short(offset, 0);
            value_count = stream_rdr.read_long(offset, 4);
            value_offset = stream_rdr.read_long(offset, 8);
            value = mat2doc.image.ShortIfdEntry_.parse_value_( ...
                stream_rdr, offset, value_count, value_offset);
            obj = mat2doc.image.ShortIfdEntry_(tag_code, value);
        end

        function v = parse_value_(stream_rdr, offset, value_count, value_offset) %#ok<INUSD>
            % _parse_value (tiff.py 245-255): a single short int read inline from
            %   the entry's value field (offset+8). Multi-value shorts are not
            %   implemented upstream (# pragma: no cover); ported verbatim.
            if value_count == 1
                v = stream_rdr.read_short(offset, 8);
            else
                v = "Multi-value short integer NOT IMPLEMENTED";
            end
        end
    end
end

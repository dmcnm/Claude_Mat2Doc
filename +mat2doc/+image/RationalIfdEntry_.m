classdef RationalIfdEntry_ < mat2doc.image.IfdEntry_
% RATIONALIFDENTRY_ IFD entry expressed as a numerator/denominator pair.
%
%   UNDERSCORE ROTATION (design.md section 2): python-docx _RationalIfdEntry ->
%   RationalIfdEntry_.
%
%   Ported from python-docx v1.2.0: src/docx/image/tiff.py::_RationalIfdEntry
%   (lines 274-289)

    methods
        function obj = RationalIfdEntry_(tag_code, value)
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
            value = mat2doc.image.RationalIfdEntry_.parse_value_( ...
                stream_rdr, offset, value_count, value_offset);
            obj = mat2doc.image.RationalIfdEntry_(tag_code, value);
        end

        function v = parse_value_(stream_rdr, offset, value_count, value_offset) %#ok<INUSD>
            % _parse_value (tiff.py 277-289): a single rational at `value_offset`
            %   (numerator long +0, denominator long +4), returned as the float
            %   quotient (Python true division; kept as a double, H6/H14).
            %   Multi-value rationals are not implemented upstream
            %   (# pragma: no cover); ported verbatim.
            %
            %   D-tiff-den0 RE-LITIGATION (docx oracle, verified P7-2a): python
            %   `numerator / denominator` with denominator == 0 raises
            %   ZeroDivisionError, which propagates out through IfdEntries_.
            %   from_stream / TiffParser_.parse / Tiff.from_stream to
            %   Image.from_file (docx tiff.py:287). Confirmed empirically:
            %   docx.image.image.Image.from_file on a TIFF whose X_RESOLUTION
            %   rational has den==0 -> ZeroDivisionError("division by zero").
            %   MATLAB's `n/0` silently yields Inf, so a guard is required to
            %   MATCH docx (an ERROR-PATH match, NOT a value divergence -> no
            %   D-number). The pptx D-tiff-den0 (which routed Inf through the
            %   int_dpi [1,2048] clamp) does NOT transfer -- docx has no clamp.
            if value_count == 1
                numerator = stream_rdr.read_long(value_offset);
                denominator = stream_rdr.read_long(value_offset, 4);
                if denominator == 0
                    error("mat2doc:ZeroDivisionError", "%s", "division by zero");
                end
                v = numerator / denominator;                 % true division
            else
                v = "Multi-value Rational NOT IMPLEMENTED";
            end
        end
    end
end

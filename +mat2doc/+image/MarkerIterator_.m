classdef MarkerIterator_ < handle
% MARKERITERATOR_ Explicit cursor for _MarkerParser.iter_markers (H9).
%
%   MATLAB has no generators, so the Python generator's loop state -- the marker
%   finder, the running `start` offset, and the last-seen marker code (the while
%   condition) -- is held explicitly here and advanced one marker at a time. This
%   preserves the generator's laziness, which IS observable: the sole consumer
%   (JfifMarkers_.from_stream) breaks at the SOS marker, so bytes after SOS are
%   never parsed as markers (eager evaluation to EOI could raise on a truncated
%   stream where Python never reads).
%
%   Mirrors the generator body of python-docx _MarkerParser.iter_markers
%   (src/docx/image/jpeg.py lines 142-152); no standalone Python counterpart
%   class (infrastructure for the generator).

    properties (Access = private)
        stream_          % StreamReader passed to _MarkerFactory
        marker_finder_   % a MarkerFinder_
        start_           % running scan offset (Python `start`)
        marker_code_     % last-yielded marker code ([] = Python None sentinel)
    end

    methods
        function obj = MarkerIterator_(stream_reader, marker_finder)
            obj.stream_ = stream_reader;
            obj.marker_finder_ = marker_finder;
            obj.start_ = 0;               % start = 0
            obj.marker_code_ = [];        % marker_code = None (H3)
        end

        function tf = has_more(obj)
            % The generator's `while marker_code != EOI` guard, evaluated BEFORE
            %   each fetch. None (init) is not EOI, so the first fetch always runs.
            EOI = mat2doc.image.JPEG_MARKER_CODE.EOI;
            tf = isempty(obj.marker_code_) || obj.marker_code_ ~= EOI;
        end

        function marker = next_marker(obj)
            % One generator step (jpeg.py 149-152): find the next marker, build it,
            %   yield it, then advance `start` past this marker's segment.
            [marker_code, segment_offset] = obj.marker_finder_.next(obj.start_);
            marker = mat2doc.image.MarkerFactory_(marker_code, obj.stream_, segment_offset);
            obj.start_ = segment_offset + marker.segment_length;   % byte address
            obj.marker_code_ = marker_code;
        end
    end
end

classdef MarkerParser_ < handle
% MARKERPARSER_ Parses a JFIF stream and iterates over its markers.
%
%   UNDERSCORE ROTATION (design.md section 2): python-docx _MarkerParser ->
%   MarkerParser_.
%
%   Ported from python-docx v1.2.0: src/docx/image/jpeg.py::_MarkerParser
%   (lines 128-152)

    properties (Access = private)
        stream_       % a StreamReader (BIG_ENDIAN)
    end

    methods
        function obj = MarkerParser_(stream_reader)
            % Ported from _MarkerParser.__init__ (jpeg.py 132-134).
            obj.stream_ = stream_reader;
        end

        function it = iter_markers(obj)
            % iter_markers (jpeg.py 142-152): a Python generator yielding a
            %   marker per JFIF marker in stream order. Realized as an explicit
            %   MarkerIterator_ cursor (H9: laziness is observable -- the consumer
            %   breaks at SOS, so scan data past SOS must never be parsed).
            marker_finder = mat2doc.image.MarkerFinder_.from_stream(obj.stream_);
            it = mat2doc.image.MarkerIterator_(obj.stream_, marker_finder);
        end
    end

    methods (Static)
        function obj = from_stream(stream)
            % from_stream (jpeg.py 136-140): wrap `stream` in a BIG_ENDIAN reader.
            stream_reader = mat2doc.image.StreamReader( ...
                stream, mat2doc.image.StreamReader.BIG_ENDIAN);
            obj = mat2doc.image.MarkerParser_(stream_reader);
        end
    end
end

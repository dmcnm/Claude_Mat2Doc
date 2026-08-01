classdef SofMarker_ < mat2doc.image.Marker_
% SOFMARKER_ Represents a JFIF start-of-frame (SOFn) marker segment.
%
%   Carries the image pixel dimensions (px_width, px_height).
%
%   UNDERSCORE ROTATION (design.md section 2): python-docx _SofMarker ->
%   SofMarker_.
%
%   Ported from python-docx v1.2.0: src/docx/image/jpeg.py::_SofMarker
%   (lines 394-425)

    properties (Access = private)
        px_width_
        px_height_
    end

    properties (Dependent)
        px_width      % image width in pixels
        px_height     % image height in pixels
    end

    methods
        function obj = SofMarker_(marker_code, offset, segment_length, px_width, px_height)
            % Ported from _SofMarker.__init__ (jpeg.py 397-400).
            obj@mat2doc.image.Marker_(marker_code, offset, segment_length);
            obj.px_width_ = px_width;
            obj.px_height_ = px_height;
        end

        function v = get.px_height(obj)
            % px_height @property (jpeg.py 417-420).
            v = obj.px_height_;
        end

        function v = get.px_width(obj)
            % px_width @property (jpeg.py 422-425).
            v = obj.px_width_;
        end
    end

    methods (Static)
        function obj = from_stream(stream, marker_code, offset)
            % from_stream (jpeg.py 402-415). SOFn segment layout (byte addresses,
            %   verbatim -- NOT H1 shifts):
            %     segment length   +0  short
            %     data precision   +2  byte
            %     vertical lines   +3  short  px_height
            %     horizontal lines +5  short  px_width
            segment_length = stream.read_short(offset);
            px_height = stream.read_short(offset, 3);
            px_width = stream.read_short(offset, 5);
            obj = mat2doc.image.SofMarker_( ...
                marker_code, offset, segment_length, px_width, px_height);
        end
    end
end

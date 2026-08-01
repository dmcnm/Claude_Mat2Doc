classdef Marker_ < handle
% MARKER_ Base class for JFIF marker classes (also the generic marker type).
%
%   Represents a marker and its segment in a JPEG byte stream. The concrete
%   subclasses (App0Marker_, App1Marker_, SofMarker_) add format-specific fields;
%   this base is used directly for every other marker.
%
%   UNDERSCORE ROTATION (design.md section 2): python-docx _Marker -> Marker_.
%
%   Ported from python-docx v1.2.0: src/docx/image/jpeg.py::_Marker
%   (lines 242-281)

    properties (Access = protected)
        marker_code_
        offset_
        segment_length_
    end

    properties (Dependent)
        marker_code       % single-byte code identifying the marker type (uint8)
        name              % marker name (debug-only; # pragma: no cover upstream)
        offset            % marker offset (debug-only; # pragma: no cover upstream)
        segment_length    % length in bytes of this marker's segment
    end

    methods
        function obj = Marker_(marker_code, offset, segment_length)
            % Ported from _Marker.__init__ (jpeg.py 248-252).
            obj.marker_code_ = marker_code;
            obj.offset_ = offset;
            obj.segment_length_ = segment_length;
        end

        function v = get.marker_code(obj)
            % marker_code @property (jpeg.py 264-268).
            v = obj.marker_code_;
        end

        function v = get.name(obj)
            % name @property (jpeg.py 270-272): marker_names lookup. Debug-only.
            v = mat2doc.image.JPEG_MARKER_CODE.marker_name(obj.marker_code_);
        end

        function v = get.offset(obj)
            % offset @property (jpeg.py 274-276). Debug-only.
            v = obj.offset_;
        end

        function v = get.segment_length(obj)
            % segment_length @property (jpeg.py 278-281).
            v = obj.segment_length_;
        end
    end

    methods (Static)
        function obj = from_stream(stream, marker_code, offset)
            % from_stream (jpeg.py 254-262): a generic marker. Standalone markers
            %   have no segment (length 0); otherwise the segment length is the
            %   leading short of the segment.
            if mat2doc.image.JPEG_MARKER_CODE.is_standalone(marker_code)
                segment_length = 0;
            else
                segment_length = stream.read_short(offset);
            end
            obj = mat2doc.image.Marker_(marker_code, offset, segment_length);
        end
    end
end

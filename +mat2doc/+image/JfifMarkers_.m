classdef JfifMarkers_ < handle
% JFIFMARKERS_ Sequence of markers in a JPEG file (up to the first SOS).
%
%   UNDERSCORE ROTATION (design.md section 2): python-docx _JfifMarkers ->
%   JfifMarkers_.
%
%   H9 (laziness IS observable here): the Python generator (_MarkerParser.
%   iter_markers) is consumed with an early break at the SOS marker, so the
%   entropy-coded scan data AFTER SOS is never parsed as markers. Eagerly
%   precomputing to EOI would parse that scan data (and could raise on a
%   truncated stream where Python never reads). The port therefore preserves
%   laziness via a MarkerIterator_ cursor and breaks at SOS exactly as Python.
%
%   __str__ (jpeg.py 72-89) is a debug tabular dump (# pragma: no cover upstream);
%   not ported (unexercised on any live path).
%
%   Ported from python-docx v1.2.0: src/docx/image/jpeg.py::_JfifMarkers
%   (lines 64-125)

    properties (Access = private)
        markers_      % 1xN cell of Marker_ (and subclasses)
    end

    methods
        function obj = JfifMarkers_(markers)
            % Ported from _JfifMarkers.__init__ (jpeg.py 68-70): list(markers).
            obj.markers_ = markers;
        end

        function m = app0(obj)
            % app0 @property (jpeg.py 103-109): first APP0 marker; KeyError if none.
            APP0 = mat2doc.image.JPEG_MARKER_CODE.APP0;
            for i = 1:numel(obj.markers_)
                mk = obj.markers_{i};
                if mk.marker_code == APP0
                    m = mk;
                    return
                end
            end
            error("mat2doc:KeyError", "%s", "no APP0 marker in image");
        end

        function m = app1(obj)
            % app1 @property (jpeg.py 111-117): first APP1 marker; KeyError if none.
            APP1 = mat2doc.image.JPEG_MARKER_CODE.APP1;
            for i = 1:numel(obj.markers_)
                mk = obj.markers_{i};
                if mk.marker_code == APP1
                    m = mk;
                    return
                end
            end
            error("mat2doc:KeyError", "%s", "no APP1 marker in image");
        end

        function m = sof(obj)
            % sof @property (jpeg.py 119-125): first SOFn marker; KeyError if none.
            SOF = mat2doc.image.JPEG_MARKER_CODE.SOF_MARKER_CODES;
            for i = 1:numel(obj.markers_)
                mk = obj.markers_{i};
                if any(mk.marker_code == SOF)
                    m = mk;
                    return
                end
            end
            error("mat2doc:KeyError", "%s", ...
                "no start of frame (SOFn) marker in image");
        end
    end

    methods (Static)
        function obj = from_stream(stream)
            % from_stream (jpeg.py 91-101): collect markers up to and including
            %   the first SOS (early break -- laziness preserved via the cursor).
            marker_parser = mat2doc.image.MarkerParser_.from_stream(stream);
            SOS = mat2doc.image.JPEG_MARKER_CODE.SOS;
            markers = {};
            it = marker_parser.iter_markers();          % MarkerIterator_ cursor
            while it.has_more()
                marker = it.next_marker();
                markers{end + 1} = marker;              %#ok<AGROW>
                if marker.marker_code == SOS
                    break
                end
            end
            obj = mat2doc.image.JfifMarkers_(markers);
        end
    end
end

classdef App1Marker_ < mat2doc.image.Marker_
% APP1MARKER_ Represents a JFIF APP1 (Exif) marker segment (carries Exif dpi).
%
%   The dpi is extracted by parsing the embedded TIFF (Exif) stream with Tiff;
%   a non-Exif APP1 segment defaults to 72/72. This is the plain python-docx
%   behavior: the dpi comes straight from the embedded Tiff parser's
%   horz_dpi / vert_dpi (P7-2a). The Mat2Ppt CLASS-E PIL-oracle variant
%   (_read_dpi_from_exif: XResolution for both axes, only ResolutionUnit==3
%   converts) is REVERTED here -- docx uses the real Tiff parser (boundary
%   audit C1: Tiff ported first so this path could land faithfully).
%
%   UNDERSCORE ROTATION (design.md section 2): python-docx _App1Marker ->
%   App1Marker_.
%
%   Ported from python-docx v1.2.0: src/docx/image/jpeg.py::_App1Marker
%   (lines 336-391)

    properties (Access = private)
        horz_dpi_
        vert_dpi_
    end

    properties (Dependent)
        horz_dpi      % horizontal dots-per-inch (defaults to 72 if unspecified)
        vert_dpi      % vertical dots-per-inch (defaults to 72 if unspecified)
    end

    methods
        function obj = App1Marker_(marker_code, offset, length, horz_dpi, vert_dpi)
            % Ported from _App1Marker.__init__ (jpeg.py 339-342).
            obj@mat2doc.image.Marker_(marker_code, offset, length);
            obj.horz_dpi_ = horz_dpi;
            obj.vert_dpi_ = vert_dpi;
        end

        function v = get.horz_dpi(obj)
            % horz_dpi @property (jpeg.py 362-366).
            v = obj.horz_dpi_;
        end

        function v = get.vert_dpi(obj)
            % vert_dpi @property (jpeg.py 368-372).
            v = obj.vert_dpi_;
        end
    end

    methods (Static)
        function obj = from_stream(stream, marker_code, offset)
            % from_stream (jpeg.py 344-360): a non-Exif APP1 segment -> 72/72;
            %   otherwise the dpi is the embedded Tiff parser's horz/vert_dpi.
            segment_length = stream.read_short(offset);
            if mat2doc.image.App1Marker_.is_non_Exif_APP1_segment_(stream, offset)
                obj = mat2doc.image.App1Marker_(marker_code, offset, segment_length, 72, 72);
                return
            end
            tiff = mat2doc.image.App1Marker_.tiff_from_exif_segment_( ...
                stream, offset, segment_length);
            obj = mat2doc.image.App1Marker_( ...
                marker_code, offset, segment_length, tiff.horz_dpi, tiff.vert_dpi);
        end
    end

    methods (Static, Access = private)
        function tf = is_non_Exif_APP1_segment_(stream, offset)
            % _is_non_Exif_APP1_segment (jpeg.py 374-381): True unless the
            %   'Exif\x00\x00' signature is present at offset+2 in the segment.
            stream.seek(offset + 2);                  % byte address
            exif_signature = stream.read(6);
            tf = ~isequal(exif_signature, uint8([69 120 105 102 0 0]));  % "Exif\x00\x00"
        end

        function tiff = tiff_from_exif_segment_(stream, offset, segment_length)
            % _tiff_from_exif_segment (jpeg.py 383-391): wrap the Exif APP1 segment
            %   in its own byte stream and feed it to Tiff.from_stream. The segment
            %   body starts at offset+8 (past the 2-byte segment length + the
            %   6-byte 'Exif\x00\x00' signature) and is segment_length-8 bytes long;
            %   that slice is the embedded TIFF (its own 'II'/'MM' + IFD offset).
            stream.seek(offset + 8);                  % byte address
            segment_bytes = stream.read(segment_length - 8);
            substream = mat2doc.image.BytesIO(segment_bytes);
            tiff = mat2doc.image.Tiff.from_stream(substream);
        end
    end
end

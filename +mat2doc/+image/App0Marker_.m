classdef App0Marker_ < mat2doc.image.Marker_
% APP0MARKER_ Represents a JFIF APP0 marker segment (carries the JFIF dpi).
%
%   UNDERSCORE ROTATION (design.md section 2): python-docx _App0Marker ->
%   App0Marker_.
%
%   The dpi comes from the APP0 density_units + x/y density via dpi_ (unit 1 ->
%   density, unit 2 -> int(round(density*2.54)), anything else -> 72). This is
%   the python-docx formula VERBATIM (jpeg.py 305-313); the Mat2Ppt CLASS-J
%   aspect-only (unit 0) fallback to the Exif dpi is NOT ported (docx uses the
%   APP0 dpi unconditionally -- see Jfif.from_stream).
%
%   Ported from python-docx v1.2.0: src/docx/image/jpeg.py::_App0Marker
%   (lines 284-333)

    properties (Access = private)
        density_units_
        x_density_
        y_density_
    end

    properties (Dependent)
        horz_dpi      % horizontal dots-per-inch (defaults to 72 if unspecified)
        vert_dpi      % vertical dots-per-inch (defaults to 72 if unspecified)
    end

    methods
        function obj = App0Marker_(marker_code, offset, length, density_units, x_density, y_density)
            % Ported from _App0Marker.__init__ (jpeg.py 287-291). Argument order
            %   matches Python exactly.
            obj@mat2doc.image.Marker_(marker_code, offset, length);
            obj.density_units_ = density_units;
            obj.x_density_ = x_density;
            obj.y_density_ = y_density;
        end

        function v = get.horz_dpi(obj)
            % horz_dpi @property (jpeg.py 293-297): _dpi(x_density).
            v = obj.dpi_(obj.x_density_);
        end

        function v = get.vert_dpi(obj)
            % vert_dpi @property (jpeg.py 299-303): _dpi(y_density).
            v = obj.dpi_(obj.y_density_);
        end
    end

    methods (Static)
        function obj = from_stream(stream, marker_code, offset)
            % from_stream (jpeg.py 315-333). APP0 segment layout (byte addresses,
            %   verbatim -- NOT H1 shifts):
            %     segment length   +0  short
            %     density units    +9  byte   1=inches, 2=cm
            %     horz density    +10  short
            %     vert density    +12  short
            segment_length = stream.read_short(offset);
            density_units = stream.read_byte(offset, 9);
            x_density = stream.read_short(offset, 10);
            y_density = stream.read_short(offset, 12);
            obj = mat2doc.image.App0Marker_( ...
                marker_code, offset, segment_length, density_units, x_density, y_density);
        end
    end

    methods (Access = private)
        function dpi = dpi_(obj, density)
            % _dpi (jpeg.py 305-313): dpi corresponding to a density value.
            %     - density_units == 1 (dots/inch): dpi = density.
            %     - density_units == 2 (dots/cm):   dpi = int(round(density*2.54)).
            %     - else (aspect-ratio-only / unknown): 72.
            %   int(round(.)) -> fix(pyRound(.)) (H6): round-half-to-even then
            %   truncate toward zero (density and 2.54 are non-negative, so fix
            %   and floor agree here -- fix is the faithful int() mapping).
            if obj.density_units_ == 1
                dpi = density;
            elseif obj.density_units_ == 2
                dpi = fix(pyRound(density * 2.54));       % int(round(.)) H6
            else
                dpi = 72;
            end
        end
    end
end

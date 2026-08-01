classdef TiffParser_ < handle
% TIFFPARSER_ Parses a TIFF image stream to extract properties from its main IFD.
%
%   UNDERSCORE ROTATION (design.md section 2): python-docx _TiffParser ->
%   TiffParser_.
%
%   Ported from python-docx v1.2.0: src/docx/image/tiff.py::_TiffParser
%   (lines 37-115)

    properties (Access = private)
        ifd_entries_      % an IfdEntries_ (mapping of tag -> value)
    end

    methods
        function obj = TiffParser_(ifd_entries)
            % Ported from _TiffParser.__init__ (tiff.py 41-43).
            obj.ifd_entries_ = ifd_entries;
        end

        function v = horz_dpi(obj)
            % horz_dpi @property (tiff.py 54-58): dpi from XResolution/ResolutionUnit.
            v = obj.dpi_(mat2doc.image.TIFF_TAG.X_RESOLUTION);
        end

        function v = vert_dpi(obj)
            % vert_dpi @property (tiff.py 60-64): dpi from YResolution/ResolutionUnit.
            %   (The docx docstring copies XResolution here; the code reads
            %   Y_RESOLUTION -- ported verbatim.)
            v = obj.dpi_(mat2doc.image.TIFF_TAG.Y_RESOLUTION);
        end

        function v = px_height(obj)
            % px_height @property (tiff.py 66-71): ImageLength tag, or None ([])
            %   when absent (the expected case for a TIFF embedded in an Exif
            %   image, where dimensions come from the JPEG SOF marker instead).
            v = obj.ifd_entries_.get(mat2doc.image.TIFF_TAG.IMAGE_LENGTH);
        end

        function v = px_width(obj)
            % px_width @property (tiff.py 73-78): ImageWidth tag, or None ([]).
            v = obj.ifd_entries_.get(mat2doc.image.TIFF_TAG.IMAGE_WIDTH);
        end
    end

    methods (Static)
        function obj = parse(stream)
            % parse (tiff.py 45-52): build a parser from the TIFF in `stream`.
            stream_rdr = mat2doc.image.TiffParser_.make_stream_reader_(stream);
            ifd0_offset = stream_rdr.read_long(4);
            ifd_entries = mat2doc.image.IfdEntries_.from_stream(stream_rdr, ifd0_offset);
            obj = mat2doc.image.TiffParser_(ifd_entries);
        end
    end

    methods (Access = private)
        function d = dpi_(obj, resolution_tag)
            % _dpi (tiff.py 88-108): the dpi value for `resolution_tag`
            %   (X_RESOLUTION or Y_RESOLUTION), computed from that tag plus the
            %   RESOLUTION_UNIT tag. This is the python-docx formula VERBATIM --
            %   NOT the Mat2Ppt PIL-oracle variant. No int_dpi [1,2048] clamp
            %   (that is a pptx seam; docx has none), no PIL joint-axes.
            ie = obj.ifd_entries_;

            if ~ie.contains_(resolution_tag)          % `resolution_tag not in ifd_entries`
                d = 72;
                return
            end

            % resolution unit defaults to inches (2)
            resolution_unit = ie.get(mat2doc.image.TIFF_TAG.RESOLUTION_UNIT, 2);

            % Gate-2 F-1: docx does `resolution_unit == 1` / `== 2` directly
            %   (tiff.py 103/106). In Python a non-scalar or non-numeric
            %   resolution_unit (the multi-value SHORT placeholder string, when
            %   RESOLUTION_UNIT has count>1) compares `== 1`/`== 2` as False
            %   SILENTLY and falls through to units_per_inch = 2.54 (returning a
            %   dpi). MATLAB `string/array == 1` THROWS, so each comparison is
            %   guarded with isnumeric && isscalar to evaluate as False exactly
            %   where Python does -- error only where Python errors, no error
            %   where Python does not (the same principle as the den0 guard).
            %   Faithful mapping: scalar 1 -> 72; scalar 2 -> x1; everything else
            %   (scalar 3, an invalid unit, or a non-scalar/non-numeric
            %   multi-value) -> x2.54.
            if isnumeric(resolution_unit) && isscalar(resolution_unit) ...
                    && resolution_unit == 1            % aspect ratio only
                d = 72;
                return
            end
            % resolution_unit == 2 for inches, 3 for centimeters
            if isnumeric(resolution_unit) && isscalar(resolution_unit) ...
                    && resolution_unit == 2
                units_per_inch = 1;
            else
                units_per_inch = 2.54;
            end
            dots_per_unit = ie.getitem_(resolution_tag);      % ifd_entries[resolution_tag]
            % int(round(...)): round-half-to-even (pyRound) then truncate toward
            %   zero (fix) (H6/H14).
            d = fix(pyRound(dots_per_unit * units_per_inch));
        end
    end

    methods (Static, Access = private)
        function endian = detect_endian_(stream)
            % _detect_endian (tiff.py 80-86): 'MM' -> BIG_ENDIAN else LITTLE_ENDIAN.
            %   `stream` here is the raw byte stream (BytesIO), read before it is
            %   wrapped in a StreamReader.
            stream.seek(0);
            endian_str = stream.read(2);
            if isequal(endian_str, uint8([77 77]))   % b"MM"
                endian = mat2doc.image.StreamReader.BIG_ENDIAN;
            else
                endian = mat2doc.image.StreamReader.LITTLE_ENDIAN;
            end
        end

        function rdr = make_stream_reader_(stream)
            % _make_stream_reader (tiff.py 110-115): wrap `stream` with the endian
            %   determined by its 'MM'/'II' header indicator.
            endian = mat2doc.image.TiffParser_.detect_endian_(stream);
            rdr = mat2doc.image.StreamReader(stream, endian);
        end
    end
end

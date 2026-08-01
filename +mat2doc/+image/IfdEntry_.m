classdef IfdEntry_ < handle
% IFDENTRY_ Base class for TIFF IFD entry classes (also the default entry type).
%
%   Holds a single directory entry's tag code and parsed value. Subclasses differ
%   only in _parse_value (value-type-specific). Because MATLAB statics do not
%   dispatch polymorphically, each subclass carries its own from_stream (the
%   common read of tag/value_count/value_offset) and its own parse_value_; the
%   constructor and the tag/value getters are inherited.
%
%   UNDERSCORE ROTATION (design.md section 2): python-docx _IfdEntry ->
%   IfdEntry_; _parse_value -> parse_value_.
%
%   Ported from python-docx v1.2.0: src/docx/image/tiff.py::_IfdEntry
%   (lines 185-226)

    properties (Access = protected)
        tag_code_
        value_
    end

    methods
        function obj = IfdEntry_(tag_code, value)
            % Ported from _IfdEntry.__init__ (tiff.py 191-194).
            obj.tag_code_ = tag_code;
            obj.value_ = value;
        end

        function v = tag(obj)
            % tag @property (tiff.py 218-221): the entry's short-int tag code.
            v = obj.tag_code_;
        end

        function v = value(obj)
            % value @property (tiff.py 223-226): the tag value (type per subclass).
            v = obj.value_;
        end
    end

    methods (Static)
        function obj = from_stream(stream_rdr, offset)
            % from_stream (tiff.py 196-208): read the common tag/value_count/
            %   value_offset triple and parse the value. Common to all subclasses
            %   in Python via `cls._parse_value`; realized here per-class.
            tag_code = stream_rdr.read_short(offset, 0);
            value_count = stream_rdr.read_long(offset, 4);
            value_offset = stream_rdr.read_long(offset, 8);
            value = mat2doc.image.IfdEntry_.parse_value_( ...
                stream_rdr, offset, value_count, value_offset);
            obj = mat2doc.image.IfdEntry_(tag_code, value);
        end

        function v = parse_value_(stream_rdr, offset, value_count, value_offset) %#ok<INUSD>
            % _parse_value (tiff.py 210-216): base returns a placeholder for an
            %   unimplemented field type (# pragma: no cover upstream).
            v = "UNIMPLEMENTED FIELD TYPE";
        end
    end
end

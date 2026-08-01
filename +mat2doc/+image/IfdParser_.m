classdef IfdParser_ < handle
% IFDPARSER_ Extracts directory entries from a TIFF Image File Directory (IFD).
%
%   UNDERSCORE ROTATION (design.md section 2): python-docx _IfdParser ->
%   IfdParser_.
%
%   Ported from python-docx v1.2.0: src/docx/image/tiff.py::_IfdParser
%   (lines 148-168)

    properties (Access = private)
        stream_rdr_
        offset_
    end

    methods
        function obj = IfdParser_(stream_rdr, offset)
            % Ported from _IfdParser.__init__ (tiff.py 152-155).
            obj.stream_rdr_ = stream_rdr;
            obj.offset_ = offset;
        end

        function entries = iter_entries(obj)
            % iter_entries (tiff.py 157-163): an IfdEntry_ per directory entry.
            %   H9: the Python generator is realized as a precomputed 1xN cell
            %   (no mutation during iteration in the original). Each entry sits at
            %   offset + 2 + idx*12 (idx 0-based, matching the byte layout; the
            %   +2 skips the 2-byte entry-count, 12 = IFD entry width).
            n = obj.entry_count_();
            entries = cell(1, n);
            for idx = 0:n-1
                dir_entry_offset = obj.offset_ + 2 + (idx * 12);   % byte address
                entries{idx + 1} = mat2doc.image.IfdEntryFactory_( ...
                    obj.stream_rdr_, dir_entry_offset);            % IDX 0->1-based
            end
        end
    end

    methods (Access = private)
        function n = entry_count_(obj)
            % _entry_count @property (tiff.py 165-168): entry count short at the
            %   top of the IFD header.
            n = obj.stream_rdr_.read_short(obj.offset_);
        end
    end
end

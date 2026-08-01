function ifd_entry = IfdEntryFactory_(stream_rdr, offset)
% IFDENTRYFACTORY_ Return an IfdEntry_ subclass for the entry at `offset`.
%
%   Reads the entry's field-type short (at offset+2) and dispatches to the
%   value-typed IfdEntry_ subclass (ASCII/SHORT/LONG/RATIONAL), or the base
%   IfdEntry_ for any other field type. Mirrors python-docx's ifd_entry_classes
%   dict + .get(field_type, _IfdEntry) default (H10). Each class supplies its own
%   from_stream static (MATLAB does not dispatch inherited statics).
%
%   UNDERSCORE ROTATION (design.md section 2): python-docx _IfdEntryFactory (a
%   module-private function) -> IfdEntryFactory_.
%
%   Ported from python-docx v1.2.0: src/docx/image/tiff.py::_IfdEntryFactory
%   (lines 171-182)

FLD = mat2doc.image.TIFF_FLD;
field_type = stream_rdr.read_short(offset, 2);
switch field_type
    case FLD.ASCII
        ifd_entry = mat2doc.image.AsciiIfdEntry_.from_stream(stream_rdr, offset);
    case FLD.SHORT
        ifd_entry = mat2doc.image.ShortIfdEntry_.from_stream(stream_rdr, offset);
    case FLD.LONG
        ifd_entry = mat2doc.image.LongIfdEntry_.from_stream(stream_rdr, offset);
    case FLD.RATIONAL
        ifd_entry = mat2doc.image.RationalIfdEntry_.from_stream(stream_rdr, offset);
    otherwise
        ifd_entry = mat2doc.image.IfdEntry_.from_stream(stream_rdr, offset);
end
end

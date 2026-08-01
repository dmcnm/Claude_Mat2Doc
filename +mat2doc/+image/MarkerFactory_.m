function marker = MarkerFactory_(marker_code, stream, offset)
% MARKERFACTORY_ Return the Marker_ subclass appropriate for `marker_code`.
%
%   Dispatches APP0 -> App0Marker_, APP1 -> App1Marker_, any SOFn -> SofMarker_,
%   and everything else -> the base Marker_ (H10). Each class supplies its own
%   from_stream static (MATLAB does not dispatch inherited statics).
%
%   UNDERSCORE ROTATION (design.md section 2): python-docx _MarkerFactory (a
%   module-private function) -> MarkerFactory_.
%
%   Ported from python-docx v1.2.0: src/docx/image/jpeg.py::_MarkerFactory
%   (lines 228-239)

JMC = mat2doc.image.JPEG_MARKER_CODE;
if marker_code == JMC.APP0
    marker = mat2doc.image.App0Marker_.from_stream(stream, marker_code, offset);
elseif marker_code == JMC.APP1
    marker = mat2doc.image.App1Marker_.from_stream(stream, marker_code, offset);
elseif any(marker_code == JMC.SOF_MARKER_CODES)
    marker = mat2doc.image.SofMarker_.from_stream(stream, marker_code, offset);
else
    marker = mat2doc.image.Marker_.from_stream(stream, marker_code, offset);
end
end

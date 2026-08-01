function header = ImageHeaderFactory_(stream)
% IMAGEHEADERFACTORY_ Return a BaseImageHeader subclass that can parse `stream`.
%
%   Reads the first 32 bytes of `stream` and matches them, in order, against the
%   python-docx SIGNATURES table (docx/image/__init__.py lines 13-23). The first
%   signature whose bytes equal the header slice at its offset selects the parser
%   class; its from_stream is invoked. If nothing matches,
%   mat2doc:UnrecognizedImageError is raised.
%
%   The eight signatures dispatch to the format parsers: PNG / GIF / BMP (P7-1b)
%   and JPEG (JFIF / Exif) and TIFF (MM / II) (P7-2). Those parser classes are
%   currently clean notYetPorted stubs; the factory itself lands whole here at
%   P7-1a because the dispatch targets are function handles resolved at call time.
%
%   UNDERSCORE ROTATION (design.md section 2): python-docx _ImageHeaderFactory
%   (a module-private function) -> ImageHeaderFactory_.
%
%   Ported from python-docx v1.2.0: src/docx/image/image.py::_ImageHeaderFactory
%   (lines 168-182); SIGNATURES from src/docx/image/__init__.py (lines 13-23).

% -- read_32 (image.py 172-174): first 32 bytes from position 0 --
stream.seek(0);
header32 = stream.read(32);        % up to 32 bytes (fewer for a short stream)

sigs = signatures_();
for i = 1:size(sigs, 1)
    cls = sigs{i, 1};              % @<Class>.from_stream handle
    offset = sigs{i, 2};          % 0-based byte offset of the signature
    signature_bytes = sigs{i, 3}; % 1xK uint8

    endIdx = offset + numel(signature_bytes);        % Python end = offset+len
    % Python `header[offset:end]` yields fewer bytes than the signature when the
    % header is short, and `==` then fails on the length mismatch. Guard the
    % MATLAB slice the same way (a short header simply cannot match).
    if numel(header32) >= endIdx
        found_bytes = header32(offset + 1 : endIdx);  % IDX 0-based slice -> 1-based
        if isequal(found_bytes, signature_bytes)
            header = cls(stream);                     % cls.from_stream(stream)
            return
        end
    end
end

% raise UnrecognizedImageError (image.py 182). Python raises it with no message
% (str(exc) == ''); this message text is port-authored (never asserted upstream).
error("mat2doc:UnrecognizedImageError", "%s", ...
    "The provided image stream could not be recognized.");
end


function sigs = signatures_()
% SIGNATURES table (docx/image/__init__.py 13-23): {from_stream handle, offset,
%   signature bytes}, in the EXACT upstream order (order is behavior: the first
%   match wins). All eight route to the format parsers (PNG/GIF/BMP at P7-1b,
%   JFIF/Exif/TIFF at P7-2).
sigs = {
    @mat2doc.image.Png.from_stream,  0, uint8([137 80 78 71 13 10 26 10]);  % \x89PNG\r\n\x1a\n
    @mat2doc.image.Jfif.from_stream, 6, uint8([74 70 73 70]);               % "JFIF"
    @mat2doc.image.Exif.from_stream, 6, uint8([69 120 105 102]);            % "Exif"
    @mat2doc.image.Gif.from_stream,  0, uint8([71 73 70 56 55 97]);         % "GIF87a"
    @mat2doc.image.Gif.from_stream,  0, uint8([71 73 70 56 57 97]);         % "GIF89a"
    @mat2doc.image.Tiff.from_stream, 0, uint8([77 77 0 42]);                % "MM\x00*" big-endian
    @mat2doc.image.Tiff.from_stream, 0, uint8([73 73 42 0]);                % "II*\x00" little-endian
    @mat2doc.image.Bmp.from_stream,  0, uint8([66 77]);                     % "BM"
};
end

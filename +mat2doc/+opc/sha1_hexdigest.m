function hexdigest = sha1_hexdigest(blob)
% SHA1_HEXDIGEST 40-character lowercase SHA1 hex digest of a byte blob.
%
%   hexdigest = MAT2DOC.OPC.SHA1_HEXDIGEST(blob) returns the SHA-1 hash of the
%   uint8 `blob` as a 40-character lowercase hexadecimal string, replicating
%   Python `hashlib.sha1(blob).hexdigest()`. Used by the image value object
%   (docx/image/image.py::Image.sha1) and by ImagePart (docx/parts/image.py,
%   P7-4) to identify and de-duplicate image bytestreams by content.
%
%   Java boundary (design.md section 4): the blob is handed to
%   java.security.MessageDigest as a signed Java byte[] via the SOLE audited
%   converters mat2doc.opc.bytesToJava / bytesFromJava -- a raw uint8->Java
%   conversion would saturate values > 127 and corrupt the digest, exactly the
%   failure mode the two converters exist to prevent.
%
%   Inputs:  blob      - 1xN uint8 (may be empty; SHA-1 of "" is the standard
%                        da39a3ee5e6b4b0d3255bfef95601890afd80709)
%   Outputs: hexdigest - 1x1 string, 40 lowercase hex characters
%
%   Example:
%       h = mat2doc.opc.sha1_hexdigest(uint8('abc'));
%       disp(h)    % "a9993e364706816aba3e25717850c26c9cd0d89d"
%
%   Mat2Doc infrastructure (shared helper): replicates hashlib.sha1(...).hexdigest()
%   for python-docx v1.2.0 src/docx/image/image.py + src/docx/parts/image.py sha1.

arguments
    blob (1,:) uint8
end
md = java.security.MessageDigest.getInstance("SHA-1");
% Feed the bytes only when non-empty: an empty blob leaves the digest in its
% initial state, and MessageDigest.digest() then returns the SHA-1 of "".
% (Passing an empty MATLAB int8 array to update() surfaces as a null Java
% buffer -- "No input buffer given" -- so it must be skipped.)
if ~isempty(blob)
    md.update(mat2doc.opc.bytesToJava(blob));
end
digest = mat2doc.opc.bytesFromJava(md.digest());   % 1x20 uint8
hexdigest = string(sprintf('%02x', digest));       % 40 lowercase hex chars
end

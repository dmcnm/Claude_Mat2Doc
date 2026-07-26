function blob = bytesFromJava(jbytes)
% BYTESFROMJAVA Convert a Java signed byte[] result to a uint8 blob.
%
%   blob = MAT2DOC.OPC.BYTESFROMJAVA(jbytes) returns the uint8 row vector for
%   the signed bytes jbytes returned by a Java call (e.g.
%   ByteArrayOutputStream.toByteArray, which materializes in MATLAB as an int8
%   column). Uses typecast (NOT cast) so the bit pattern is preserved
%   (int8 -128..127 -> uint8 128..255). Twin of bytesToJava; the two are the SOLE
%   audited Java boundary of the OPC layer (design.md section 4).
%
%   RE-PORT NOTE: re-ported into +mat2doc\+opc (design.md section 7 forbids
%   sharing code between the Mat2Ppt and Mat2Doc toolboxes). Logic identical to
%   the Mat2Ppt bytesFromJava (Spike S3).
%
%   Inputs:  jbytes - int8 array (any shape) from a Java byte[]
%   Outputs: blob   - 1xN uint8
%
%   Example:
%       u = mat2doc.opc.bytesFromJava(int8([0 -128 -1]));
%       disp(u)    % 0  128  255  (uint8, bit pattern preserved)
%
%   Ported from python-docx v1.2.0: src/docx/opc/phys_pkg.py (the java.util.zip
%   Java-boundary helper is design-realization per design.md section 4 / Spike S3)

blob = typecast(int8(jbytes(:))', 'uint8');
end

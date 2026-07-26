function jbytes = bytesToJava(blob)
% BYTESTOJAVA Convert a uint8 blob to a signed int8 array for the Java zip API.
%
%   jbytes = MAT2DOC.OPC.BYTESTOJAVA(blob) returns typecast(blob, 'int8'), the
%   signed representation MATLAB hands to java.io/java.util.zip as a Java
%   `byte[]` (Java bytes are signed; a direct uint8->java conversion SATURATES
%   values > 127, corrupting binary parts). This helper and its twin
%   bytesFromJava are the SOLE audited Java boundary of the OPC layer
%   (design.md section 4 "Java boundary"): nothing else touches Java arrays.
%
%   RE-PORT NOTE: re-ported into +mat2doc\+opc (design.md section 7 forbids
%   sharing code between the Mat2Ppt and Mat2Doc toolboxes; the Mat2Ppt twin is
%   NOT referenced). Logic identical to the Mat2Ppt bytesToJava (Spike S3).
%
%   Inputs:  blob   - 1xN uint8 (a part blob or whole-zip bytes)
%   Outputs: jbytes - 1xN int8, ready for ByteArrayInputStream / ZipOutputStream
%
%   Example:
%       j = mat2doc.opc.bytesToJava(uint8([0 127 128 255]));
%       disp(class(j))    % "int8"  (128 -> -128, 255 -> -1: bit pattern kept)
%
%   Ported from python-docx v1.2.0: src/docx/opc/phys_pkg.py (the java.util.zip
%   Java-boundary helper is design-realization per design.md section 4 / Spike S3)

arguments
    blob (1,:) uint8
end
jbytes = typecast(blob, 'int8');
end

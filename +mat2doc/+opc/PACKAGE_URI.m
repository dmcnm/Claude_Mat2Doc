function uri = PACKAGE_URI()
% PACKAGE_URI The pack URI for the package pseudo-partname "/".
%
%   uri = MAT2DOC.OPC.PACKAGE_URI() returns PackURI("/"). Mirrors the module
%   constant PACKAGE_URI = PackURI("/"). Returned fresh (PackURI is an immutable
%   value type, so this is indistinguishable from a shared constant).
%
%   Example:
%       disp(string(mat2doc.opc.PACKAGE_URI()))   % "/"
%
%   Ported from python-docx v1.2.0: src/docx/opc/packuri.py::PACKAGE_URI (line 108)

uri = mat2doc.opc.PackURI("/");
end

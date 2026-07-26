classdef PhysPkgWriter < handle
% PHYSPKGWRITER Factory for physical package writer objects.
%
%   mat2doc.opc.PhysPkgWriter is PUBLIC in python-docx. Its Python `__new__`
%   ALWAYS returns a _ZipPkgWriter (phys_pkg.py 31-32) -- there is no directory
%   writer; a directory package can be read but not written. MATLAB realizes the
%   dispatch as a static `factory` method returning ZipPkgWriter_.
%
%   UNDERSCORE ROTATION (design.md section 2): `_ZipPkgWriter` -> ZipPkgWriter_.
%   This base keeps the docx PUBLIC name PhysPkgWriter.
%
%   Example:
%       w = mat2doc.opc.PhysPkgWriter.factory("out.docx");
%       disp(class(w))     % "mat2doc.opc.ZipPkgWriter_"
%
%   Ported from python-docx v1.2.0: src/docx/opc/phys_pkg.py::PhysPkgWriter
%   (lines 28-32)

    methods (Static)
        function writer = factory(pkg_file)
            % FACTORY Return the PhysPkgWriter subtype for pkg_file
            %   (phys_pkg.py 31-32): always ZipPkgWriter_.
            writer = mat2doc.opc.ZipPkgWriter_(pkg_file);
        end
    end
end

classdef PhysPkgReader < handle
% PHYSPKGREADER Factory for physical package reader objects.
%
%   mat2doc.opc.PhysPkgReader is PUBLIC in python-docx (class PhysPkgReader, NO
%   leading underscore -- contrast the private pptx `_PhysPkgReader`). Its Python
%   `__new__` dispatches on the pkg_file argument and returns a subtype INSTANCE;
%   MATLAB constructors cannot return a different class, so the dispatch is a
%   static `factory` method (the Mat2Ppt precedent), and the concrete readers
%   subclass this base.
%
%   Dispatch (phys_pkg.py 13-25), faithful to docx (which tests str FIRST):
%       * pkg_file is a path (string/char):
%             a directory  -> DirPkgReader_
%             a zip file   -> ZipPkgReader_
%             otherwise    -> raise mat2doc:PackageNotFoundError
%       * pkg_file is a stream (our currency: uint8 whole-zip bytes) -> ZipPkgReader_
%
%   UNDERSCORE ROTATION (design.md section 2): the concrete readers `_DirPkgReader`
%   / `_ZipPkgReader` -> DirPkgReader_ / ZipPkgReader_. This base keeps the docx
%   PUBLIC name PhysPkgReader (no rotation).
%
%   Example:
%       tpl = fullfile(fileparts(fileparts(which( ...
%           'mat2doc.opc.PhysPkgReader'))), "templates", "default.docx");
%       reader = mat2doc.opc.PhysPkgReader.factory(tpl);
%       disp(class(reader))     % "mat2doc.opc.ZipPkgReader_"
%
%   Ported from python-docx v1.2.0: src/docx/opc/phys_pkg.py::PhysPkgReader
%   (lines 10-25)

    methods (Static)
        function reader = factory(pkg_file)
            % FACTORY Return the PhysPkgReader subtype for pkg_file
            %   (phys_pkg.py 13-25).
            % Python: `if isinstance(pkg_file, str): ... else: _ZipPkgReader`.
            if isstring(pkg_file) || ischar(pkg_file)
                path = string(pkg_file);
                if isfolder(path)
                    reader = mat2doc.opc.DirPkgReader_(path);
                elseif mat2doc.opc.PhysPkgReader.isZipFile_(path)
                    reader = mat2doc.opc.ZipPkgReader_(path);
                else
                    error("mat2doc:PackageNotFoundError", ...
                        "Package not found at '%s'", path);
                end
            else
                % assume it's a stream and pass it to the Zip reader (uint8 bytes)
                reader = mat2doc.opc.ZipPkgReader_(pkg_file);
            end
        end
    end

    methods (Static, Access = private)
        function tf = isZipFile_(path)
            % zipfile.is_zipfile analogue: a readable file whose first two bytes
            %   are the local-file-header signature "PK". (OOXML packages are
            %   never empty archives, so the PK local-header check is sufficient;
            %   an empty/other file falls through to PackageNotFoundError.)
            tf = false;
            if ~isfile(path)
                return
            end
            fid = fopen(path, "rb");
            if fid < 0
                return
            end
            sig = fread(fid, 2, "*uint8")';
            fclose(fid);
            tf = numel(sig) == 2 && sig(1) == uint8('P') && sig(2) == uint8('K');
        end
    end
end

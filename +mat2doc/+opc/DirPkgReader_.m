classdef DirPkgReader_ < mat2doc.opc.PhysPkgReader
% DIRPKGREADER_ PhysPkgReader for an OPC package extracted into a directory.
%
%   The plain-file-read reader for a package unzipped into a directory tree.
%   `path` is the directory containing the expanded package; blob_for resolves a
%   PackURI to <path>/<membername> and reads it. Read-only: there is NO directory
%   writer (mirroring docx, phys_pkg.py has no _DirPkgWriter).
%
%   UNDERSCORE ROTATION (design.md section 2): `_DirPkgReader` -> DirPkgReader_.
%
%   ERROR-ID fidelity: docx `_DirPkgReader.blob_for` opens the member file with
%   open(...,"rb"); a missing member raises IOError/FileNotFoundError. Ported as
%   mat2doc:IOError, so `rels_xml_for` (which catches IOError -> None) matches the
%   Python except clause exactly (contrast the zip reader, which raises KeyError).
%
%   Example:
%       d = fullfile(tempdir, "pkg_expanded"); % a directory of an expanded docx
%       reader = mat2doc.opc.DirPkgReader_(d);
%       cts = reader.content_types_xml;        % uint8 [Content_Types].xml bytes
%
%   Ported from python-docx v1.2.0: src/docx/opc/phys_pkg.py::_DirPkgReader
%   (lines 35-68)

    properties (Access = private)
        path_ (1,1) string
    end

    properties (Dependent, SetAccess = private)
        content_types_xml   % [Content_Types].xml blob (phys_pkg.py 56-59)
    end

    methods
        function obj = DirPkgReader_(path)
            % phys_pkg.py 39-42: self._path = os.path.abspath(path).
            arguments
                path (1,1) string
            end
            f = dir(path);
            if isempty(f)
                obj.path_ = path;   % keep as given if not resolvable
            else
                obj.path_ = string(f(1).folder);
            end
        end

        function blob = blob_for(obj, pack_uri)
            % BLOB_FOR Contents of the member file for pack_uri (phys_pkg.py
            %   44-49): open(join(path, membername), "rb").read(). Missing member
            %   -> mat2doc:IOError (Python open() -> IOError/FileNotFoundError).
            p = fullfile(obj.path_, pack_uri.membername);
            fid = fopen(p, "rb");
            if fid < 0
                error("mat2doc:IOError", "cannot open member '%s'", ...
                    string(pack_uri.membername));
            end
            blob = fread(fid, inf, "*uint8")';
            fclose(fid);
        end

        function close(obj) %#ok<MANU>
            % CLOSE Interface consistency; a directory file system needs no
            %   closing (phys_pkg.py 51-54): pass.
        end

        function value = get.content_types_xml(obj)
            % @property content_types_xml (phys_pkg.py 56-59): blob_for(CT_URI).
            value = obj.blob_for(mat2doc.opc.CONTENT_TYPES_URI());
        end

        function rels_xml = rels_xml_for(obj, source_uri)
            % RELS_XML_FOR rels item XML for source_uri, or [] (None) if the item
            %   has no rels item (phys_pkg.py 61-68): try blob_for(rels_uri)
            %   except IOError -> None.
            try
                rels_xml = obj.blob_for(source_uri.rels_uri);
            catch ME
                if ME.identifier == "mat2doc:IOError"
                    rels_xml = [];   % None (H3)
                else
                    rethrow(ME);
                end
            end
        end
    end
end

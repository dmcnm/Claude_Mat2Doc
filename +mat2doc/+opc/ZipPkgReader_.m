classdef ZipPkgReader_ < mat2doc.opc.PhysPkgReader
% ZIPPKGREADER_ PhysPkgReader for a zip-file (.docx) OPC package.
%
%   Provides random-access read of package members by PackURI. Source is either a
%   path string (a .docx file) or a uint8 whole-zip byte vector (the in-memory
%   stream currency). On construction the archive is enumerated once into an
%   ordered {membername -> uint8 blob} map via an in-memory java.util.zip pass
%   (whole bytes -> ByteArrayInputStream -> ZipInputStream, each entry drained
%   through the InterruptibleStreamCopier -- MATLAB cannot observe mutations of a
%   byte[] passed into read(byte[],off,len), so the buffered copy is delegated to
%   Java). The signed byte[] <-> uint8 boundary goes ONLY through bytesToJava.
%
%   Realization note: docx `_ZipPkgReader.__init__` opens a ZipFile and reads
%   members lazily on each blob_for; this port preloads the member map (design.md
%   section 4 -- the whole-file -> {member -> blob} map). blob_for is then an
%   observable-equivalent lookup: it returns the member bytes, or raises
%   mat2doc:KeyError when absent (matching zipfile's `read` KeyError), which
%   rels_xml_for catches -> [].
%
%   UNDERSCORE ROTATION (design.md section 2): `_ZipPkgReader` -> ZipPkgReader_.
%
%   Example:
%       tpl = fullfile(fileparts(fileparts(which( ...
%           'mat2doc.opc.ZipPkgReader_'))), "templates", "default.docx");
%       z   = mat2doc.opc.ZipPkgReader_(tpl);
%       cts = z.content_types_xml;              % uint8 [Content_Types].xml bytes
%
%   Ported from python-docx v1.2.0: src/docx/opc/phys_pkg.py::_ZipPkgReader
%   (lines 71-101; the java.util.zip realization is design-realization per
%   design.md section 4 / Spike S3)

    properties (Access = private)
        source_                            % path string OR uint8 whole-zip bytes
        entries_ = struct('membername', {}, 'blob', {})  % ordered {membername, blob}
        loaded_ (1,1) logical = false      % computed-flag sentinel (design.md section 2)
    end

    properties (Dependent, SetAccess = private)
        content_types_xml   % [Content_Types].xml blob (phys_pkg.py 89-92)
    end

    methods
        function obj = ZipPkgReader_(pkg_file)
            % phys_pkg.py 74-76: self._zipf = ZipFile(pkg_file, "r"). docx opens
            %   eagerly; the member map is loaded now so a corrupt source is
            %   observed at construction as in docx.
            obj.source_ = pkg_file;
            obj.ensureLoaded_();
        end

        function blob = blob_for(obj, pack_uri)
            % BLOB_FOR blob corresponding to pack_uri (phys_pkg.py 78-83):
            %   self._zipf.read(pack_uri.membername). mat2doc:KeyError when no
            %   matching member is present in the archive.
            i = obj.indexOf_(pack_uri.membername);
            if i == 0
                error("mat2doc:KeyError", "no member '%s' in package", ...
                    string(pack_uri.membername));
            end
            blob = obj.entries_(i).blob;
        end

        function close(obj) %#ok<MANU>
            % CLOSE Release the archive (phys_pkg.py 85-87). The member map is
            %   in-memory and holds no OS handle, so this is a no-op.
        end

        function value = get.content_types_xml(obj)
            % @property content_types_xml (phys_pkg.py 89-92): blob_for(CT_URI).
            value = obj.blob_for(mat2doc.opc.CONTENT_TYPES_URI());
        end

        function rels_xml = rels_xml_for(obj, source_uri)
            % RELS_XML_FOR rels item XML for source_uri, or [] (None) if no rels
            %   item is present (phys_pkg.py 94-101): try blob_for(rels_uri)
            %   except KeyError -> None.
            try
                rels_xml = obj.blob_for(source_uri.rels_uri);
            catch ME
                if ME.identifier == "mat2doc:KeyError"
                    rels_xml = [];   % None (H3)
                else
                    rethrow(ME);
                end
            end
        end
    end

    methods (Access = private)
        function ensureLoaded_(obj)
            if obj.loaded_
                return
            end
            if isa(obj.source_, "uint8")
                zipBytes = obj.source_;
            else
                fid = fopen(obj.source_, "rb");
                if fid < 0
                    error("mat2doc:PackageNotFoundError", ...
                        "cannot open package '%s'", string(obj.source_));
                end
                zipBytes = fread(fid, inf, "*uint8")';
                fclose(fid);
            end
            obj.entries_ = mat2doc.opc.ZipPkgReader_.readZipEntries_(zipBytes);
            obj.loaded_ = true;
        end

        function i = indexOf_(obj, membername)
            i = 0;
            target = string(membername);
            for k = 1:numel(obj.entries_)
                if obj.entries_(k).membername == target
                    i = k;
                    return
                end
            end
        end
    end

    methods (Static, Access = private)
        function entries = readZipEntries_(zipBytes)
            % Enumerate all entries of an in-memory zip in order -> struct array
            %   {membername (the raw zip name), blob (uint8)} (Spike S3).
            bais = java.io.ByteArrayInputStream(mat2doc.opc.bytesToJava(zipBytes));
            zis = java.util.zip.ZipInputStream(bais);
            entries = struct('membername', {}, 'blob', {});
            cleanup = onCleanup(@() zis.close()); %#ok<NASGU>  auto-close on exit/error
            while true
                ze = zis.getNextEntry();
                if isempty(ze)      % Java null => no more entries
                    break
                end
                name = string(ze.getName());
                blob = mat2doc.opc.ZipPkgReader_.drainEntry_(zis);
                entries(end + 1) = struct('membername', name, 'blob', blob); %#ok<AGROW>
                zis.closeEntry();
            end
        end

        function data = drainEntry_(in)
            % Drain the current zip entry to uint8. MATLAB cannot observe
            %   mutations of a byte[] passed into read(byte[],off,len), so the
            %   buffered copy is delegated to the Java-side InterruptibleStream
            %   Copier (loops read() until -1 internally) -- Spike S3.
            baos = java.io.ByteArrayOutputStream;
            copier = com.mathworks.mlwidgets.io.InterruptibleStreamCopier ...
                .getInterruptibleStreamCopier;
            copier.copyStream(in, baos);
            data = mat2doc.opc.bytesFromJava(baos.toByteArray());
        end
    end
end

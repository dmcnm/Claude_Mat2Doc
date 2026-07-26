classdef ZipPkgWriter_ < mat2doc.opc.PhysPkgWriter
% ZIPPKGWRITER_ PhysPkgWriter for a zip-file (.docx) OPC package.
%
%   STATEFUL writer mirroring docx `_ZipPkgWriter` (phys_pkg.py 104-119): the
%   constructor opens the archive, `write(pack_uri, blob)` appends one DEFLATED
%   entry per call (membername = pack_uri.membername), and `close()` finalizes it.
%   Realized over an in-memory java.util.zip pass (ByteArrayOutputStream ->
%   ZipOutputStream); on close the bytes are flushed to the pkg_file path (a
%   string/char) and are also retrievable via `to_bytes()` for the stream case.
%   The signed byte[] <-> uint8 boundary goes ONLY through bytesToJava.
%
%   D-zip-time (design.md section 4; ADOPTED deviation, NO new D-number): Python
%   `ZipFile.writestr(membername, blob)` stamps each entry with the wall-clock
%   time (time.localtime()), so whole-file zip bytes are not reproducible. Every
%   entry time here is pinned to 1980-01-01 00:00 (the DOS-time floor) so
%   MATLAB<->MATLAB output is byte-reproducible. Invisible to Office; whole-file
%   zip bytes are OUT OF SCOPE for equivalence (part-level after extraction).
%
%   BYTE-PARITY with Mat2Ppt: the Java calls (ZipOutputStream over a
%   ByteArrayOutputStream, setMethod DEFLATED, per-entry ZipEntry + setTime(1980)
%   + putNextEntry + write + closeEntry, then finish/close) are identical to the
%   Mat2Ppt ZipPkgWriter_.write_blob_map; given the same entries in the same order
%   this writer produces byte-identical output. The only difference is structural:
%   docx drives the entries incrementally (stateful write/close) rather than as a
%   batch, which does not change the emitted bytes.
%
%   UNDERSCORE ROTATION (design.md section 2): `_ZipPkgWriter` -> ZipPkgWriter_.
%
%   Example:
%       w = mat2doc.opc.ZipPkgWriter_("out.docx");
%       w.write(mat2doc.opc.PackURI("/hello.txt"), uint8('hi'));
%       w.close();     % out.docx written; w.to_bytes() also holds the zip bytes
%
%   Ported from python-docx v1.2.0: src/docx/opc/phys_pkg.py::_ZipPkgWriter
%   (lines 104-119; the java.util.zip realization is design-realization per
%   design.md section 4 / Spike S3)

    properties (Access = private)
        pkg_file_          % path string/char, [] (bytes only), or stream target
        baos_              % java.io.ByteArrayOutputStream
        zos_               % java.util.zip.ZipOutputStream
        fixedMs_ (1,1) double   % pinned entry time (1980-01-01), D-zip-time
        bytes_ = uint8([])      % whole-zip bytes, filled on close()
        closed_ (1,1) logical = false
    end

    methods
        function obj = ZipPkgWriter_(pkg_file)
            % phys_pkg.py 107-109: self._zipf = ZipFile(pkg_file, "w",
            %   compression=ZIP_DEFLATED). Realized as an in-memory
            %   ByteArrayOutputStream -> ZipOutputStream (DEFLATED).
            obj.pkg_file_ = pkg_file;
            obj.baos_ = java.io.ByteArrayOutputStream;
            obj.zos_ = java.util.zip.ZipOutputStream(obj.baos_);
            obj.zos_.setMethod(java.util.zip.ZipOutputStream.DEFLATED);
            % Fixed entry time: 1980-01-01 00:00 local (the DOS-time floor),
            % safely at/above the floor in any timezone -> deterministic output
            % across runs on the same machine (D-zip-time / Spike S3).
            cal = java.util.GregorianCalendar(1980, 0, 1, 0, 0, 0);
            obj.fixedMs_ = cal.getTimeInMillis();
        end

        function write(obj, pack_uri, blob)
            % WRITE Write blob to the package with the membername for pack_uri
            %   (phys_pkg.py 116-119): self._zipf.writestr(membername, blob).
            %   The entry time is pinned to 1980 (D-zip-time) instead of Python's
            %   wall-clock time.
            arguments
                obj (1,1) mat2doc.opc.ZipPkgWriter_
                pack_uri (1,1) mat2doc.opc.PackURI
                blob
            end
            membername = char(pack_uri.membername);
            ze = java.util.zip.ZipEntry(membername);
            ze.setTime(obj.fixedMs_);
            obj.zos_.putNextEntry(ze);
            u = uint8(blob);
            if ~isempty(u)
                jb = mat2doc.opc.bytesToJava(u);
                obj.zos_.write(jb, 0, numel(u));
            end
            obj.zos_.closeEntry();
        end

        function close(obj)
            % CLOSE Finalize the archive (phys_pkg.py 111-114). Flush pending
            %   writes, capture the whole-zip bytes, and (if pkg_file is a path)
            %   write them to disk.
            if obj.closed_
                return
            end
            obj.zos_.finish();
            obj.zos_.close();
            obj.bytes_ = mat2doc.opc.bytesFromJava(obj.baos_.toByteArray());
            obj.closed_ = true;
            if isstring(obj.pkg_file_) || ischar(obj.pkg_file_)
                fid = fopen(obj.pkg_file_, "wb");
                if fid < 0
                    error("mat2doc:IOError", "cannot open '%s' for writing", ...
                        string(obj.pkg_file_));
                end
                fwrite(fid, obj.bytes_, "uint8");
                fclose(fid);
            end
        end

        function bytes = to_bytes(obj)
            % TO_BYTES The whole-zip uint8 bytes captured at close() (the stream
            %   currency for a non-path pkg_file). Empty until close() runs.
            bytes = obj.bytes_;
        end
    end
end

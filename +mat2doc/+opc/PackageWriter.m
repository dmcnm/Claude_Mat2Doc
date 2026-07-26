classdef PackageWriter
% PACKAGEWRITER Writes a zip-format OPC package (.docx) to a file.
%
%   Its single API method, write, is static, so this class is not instantiated.
%   pkg_file is a path (string/char) or the in-memory stream currency.
%
%   M1 ZIP-ENTRY ORDER (design.md section 4; the M1 byte-risk). write emits, in
%   this exact order:
%       0. [Content_Types].xml   (write_content_types_stream_)
%       1. /_rels/.rels          (write_pkg_rels_)
%       2..N. each part in `parts` order, and immediately after each part its
%             `.rels` item IFF the part has any relationships   (write_parts_)
%   The `parts` sequence is supplied by the caller (OpcPackage.iter_parts, a DFS
%   of the relationship graph -- P1-6b/P1-8 scope). This writer FAITHFULLY
%   preserves whatever order it is handed; write_parts_ does NOT reorder. The
%   frozen M1 target order (validation\mat2doc\references\m1_skeleton_target.md)
%   IS this content-types/pkg-rels/part+part.rels sequence. NOTE: the full
%   integrated default.docx order proof lands at P1-8 (when iter_parts + real
%   Part objects exist); this WP proves the writer MECHANISM preserves a given
%   order (stub multi-part probe, audit section 6).
%
%   BYTES CURRENCY (P1-5 audit ADVISORY, dispositioned): every value handed to
%   phys_writer.write is uint8 bytes -- cti.blob (serialize_part_xml), part.blob,
%   part.rels.xml (= Relationships.xml -> CT_Relationships.xml_file_bytes ->
%   uint8), pkg_rels.xml (same). ZipPkgWriter_.write does uint8(blob), exact for a
%   uint8 currency; NO char/string ever reaches the physical writer (that would
%   latin-1-corrupt non-ASCII), so the advisory is satisfied at every boundary.
%
%   Ported from python-docx v1.2.0: src/docx/opc/pkgwriter.py::PackageWriter
%   (lines 22-59)

    methods (Static)
        function write(pkg_file, pkg_rels, parts)
            % WRITE Write the physical package to pkg_file containing pkg_rels and
            %   parts, plus a content-types stream derived from the parts
            %   (pkgwriter.py 30-38).
            phys_writer = mat2doc.opc.PhysPkgWriter.factory(pkg_file);
            mat2doc.opc.PackageWriter.write_content_types_stream_(phys_writer, parts);
            mat2doc.opc.PackageWriter.write_pkg_rels_(phys_writer, pkg_rels);
            mat2doc.opc.PackageWriter.write_parts_(phys_writer, parts);
            phys_writer.close();
        end
    end

    methods (Static, Access = private)
        function write_content_types_stream_(phys_writer, parts)
            % _write_content_types_stream (pkgwriter.py 40-45).
            cti = mat2doc.opc.ContentTypesItem_.from_parts(parts);
            phys_writer.write(mat2doc.opc.CONTENT_TYPES_URI(), cti.blob);
        end

        function write_parts_(phys_writer, parts)
            % _write_parts (pkgwriter.py 47-54): write each part's blob, then --
            %   iff it has any relationships -- its .rels item, in `parts` order.
            for k = 1:numel(parts)
                part = parts(k);
                phys_writer.write(part.partname, part.blob);
                if part.rels.len > 0   % Python: if len(part.rels): (H4 truthiness)
                    phys_writer.write(part.partname.rels_uri, part.rels.xml);
                end
            end
        end

        function write_pkg_rels_(phys_writer, pkg_rels)
            % _write_pkg_rels (pkgwriter.py 56-59): write '/_rels/.rels'.
            phys_writer.write(mat2doc.opc.PACKAGE_URI().rels_uri, pkg_rels.xml);
        end
    end
end

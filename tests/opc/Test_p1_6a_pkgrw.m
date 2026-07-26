classdef Test_p1_6a_pkgrw < matlab.unittest.TestCase
% TEST_P1_6A_PKGRW  Gate-4 permanent unit tests for Mat2Doc P1-6a
%   (opc pkgreader + pkgwriter).
%
%   Freezes the P1-6a guarantees ported in +mat2doc\+opc\ from python-docx v1.2.0
%   src/docx/opc/{pkgreader,pkgwriter}.py: the OPC serialized-READ path
%   (PackageReader + _ContentTypeMap / _SerializedPart / _SerializedRelationship(s))
%   and the physical zip-WRITE path (PackageWriter + _ContentTypesItem). This is
%   the M1 byte-critical #2 WP -- PackageWriter's traversal IS the frozen M1
%   zip-entry order; PackageReader is the load path.
%
%   Provenance (Gate-1..3, all 2026-07-26):
%     * Audit    : validation\mat2doc\audit_P1-6a_pkgrw.md (Porter Gate-1 +
%                  Fable Gate-2 APPROVE -- zero defects, zero inline fixes,
%                  zero new D-numbers)
%     * Validate : validation\mat2doc\validate_P1-6a_pkgrw.md (Gate-3 PASS, all
%                  5 bar items; 0 FAIL; 0 new D-numbers; regression 265/265)
%     * Scenarios: validation\mat2doc\scenarios\s0005_writer_dfs_order.{py,m}
%                  (+ S0005StubPart.m); s0006_reader_roundtrip.{py,m};
%                  s0007_ctmap_v3_probe.{py,m}
%     * Frozen refs (python-docx 1.2.0 oracle, frozen ONCE):
%         references\s0005\ (8-part pkgcompare freeze + manifest.json; the 3 L1
%           XML parts copied byte-for-byte into tests\opc\data\ as
%           s0005_content_types.bin (773 B), s0005_pkg_rels.bin (298 B),
%           s0005_doc_rels.bin (283 B) so this suite is self-contained)
%         references\s0006\probe.json  -> copied to data\s0006_probe.json
%           (13 sparts partname|content_type|reltype|size|sha256 + 13 srels
%           source|rId|reltype|target|is_external, DFS order)
%         references\s0007\probe.json  -> copied to data\s0007_probe.json
%           (_ContentTypeMap precedence + V-3 target_partname cache)
%     * Template : +mat2doc\templates\default.docx (sha256 2094b5bd..40d35d),
%                  the real reader round-trip target (ships in the toolbox, so
%                  the reader lanes are self-contained relative to the worktree).
%
%   Coverage taxonomy
%   -----------------
%   * Regression (byte-level L1, the M1-critical writer pin) -- rebuild the s0005
%     5-part mixed-case stub via mat2doc.opc.PackageWriter.write, unzip, and pin
%     the zip-entry SEQUENCE (8 entries, exact) + [Content_Types].xml (773 B) +
%     _rels/.rels (298 B) + word/_rels/document.xml.rels (283 B) BYTE-identical
%     and SHA-256-identical to the frozen s0005 refs. Goes RED on any writer
%     reorder or serializer perturbation. The Gate-3 ladder demanded L1 for every
%     XML part (validate_P1-6a lane 1); these are byte-identical assertions.
%   * Equivalence -- reader round-trip against the REAL default.docx: iter_sparts
%     13/13 and iter_srels 13/13 equal the frozen s0006 probe (DFS order,
%     blob SHA-256 per part, doc-order rels NOT rId-sorted, visited-dedup).
%   * Regression + Edge/error path -- _ContentTypeMap precedence (Override exact,
%     Override case-insensitive, Default-ext, Default-ext CI, xml Default,
%     missing->mat2doc:KeyError verbatim, non-PackURI->mat2doc:KeyError
%     identifier) and V-3 target_partname (posix ..-join, cache stability,
%     External->mat2doc:ValueError verbatim, None->empty).
%   * H2 bytes currency -- every writer-boundary value and all 13 reader blobs
%     are uint8 (no char leak).
%
%   Deviations exercised (adopt-only, ZERO new D-numbers -- Gate-3): D-zip-time
%   (the s0005 package is written by the P1-5 ZipPkgWriter_ 1980-stamp mechanism;
%   pkgcompare/this suite compare part bytes after unzip, never the zip envelope)
%   and D-serializer-nsdecl / D-001 (the [Content_Types].xml + .rels bytes transit
%   the P1-4 serialize_part_xml; re-verified byte-neutral by the L1 pins here).
%
%   Determinism: no network, no hard-coded absolute paths -- fixtures and the
%   template resolve relative to this file via fileparts(mfilename('fullpath')).
%   The one written artifact is a tempname .docx, deleted via onCleanup; the
%   PackageWriter opens it in BINARY mode ("wb"), so no CRLF corruption (the
%   Gate-3 CRLF fixture finding).

    properties (Constant)
        % The frozen s0005 writer zip-DFS entry order (manifest.json parts[] order;
        % validate_P1-6a lane 1). Part-then-its-own-.rels interleave; no-rels parts
        % (styles/jpeg/png/core) emit NO .rels item.
        S0005_ORDER = [ ...
            "[Content_Types].xml", ...
            "_rels/.rels", ...
            "word/document.xml", ...
            "word/_rels/document.xml.rels", ...
            "word/styles.xml", ...
            "docProps/thumbnail.jpeg", ...
            "word/media/IMG.PNG", ...
            "docProps/Core.XML"];
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied from tests\opc\Test_p1_5_opc_packuri.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\opc
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
            % Also add tests\opc so the StubPart_p1_6a helper class resolves under
            % a cold suite run regardless of the runner's cd behavior.
            testCase.applyFixture(PathFixture(here));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. Writer zip-DFS order + [Content_Types].xml/.rels L1          %
        %    (the M1-critical byte pin)                                    %
        % =============================================================== %

        function test_writer_zip_dfs_order_8_entries(testCase)
            % Regression (M1 zip-DFS order, s0005 L0): the PackageWriter emits the
            % 8 zip entries in EXACTLY the frozen s0005 sequence -- content-types,
            % pkg-rels, then each part followed immediately by its own .rels iff it
            % has rels. Goes RED on any writer reorder. (validate_P1-6a lane 1.)
            [~, names] = testCase.buildS0005Package();
            testCase.verifyEqual(names, testCase.S0005_ORDER, ...
                'writer zip-entry sequence must equal the frozen s0005 DFS order');
        end

        function test_writer_content_types_L1_773(testCase)
            % Regression (L1, THE M1 byte pin): the written [Content_Types].xml is
            % BYTE-identical (773 B) + SHA-256-identical to the frozen s0005 ref.
            % Proves sort-at-emit (defaults-by-ext then overrides-by-partname),
            % Default/Override classification, serializer/nsdecl inheritance.
            [blobs, names] = testCase.buildS0005Package();
            got = testCase.entryBlob(blobs, names, "[Content_Types].xml");
            want = testCase.loadFixture('s0005_content_types');
            testCase.verifyEqual(numel(got), 773, ...
                '[Content_Types].xml must be exactly 773 B');
            verifyByteIdentical(testCase, got, want, '[Content_Types].xml');
            testCase.verifyEqual(sha256hex(got), ...
                "3c99ff59835177c9a514724e6bbbe9a36c44f49ca476d688422065f7bf22c2a9", ...
                '[Content_Types].xml SHA-256 must equal the frozen s0005 reference');
        end

        function test_writer_pkg_rels_L1_298(testCase)
            % Regression (L1): the written _rels/.rels is BYTE-identical (298 B) +
            % SHA-256-identical to the frozen s0005 ref (insertion-order .rels
            % serialization in-situ).
            [blobs, names] = testCase.buildS0005Package();
            got = testCase.entryBlob(blobs, names, "_rels/.rels");
            want = testCase.loadFixture('s0005_pkg_rels');
            testCase.verifyEqual(numel(got), 298, '_rels/.rels must be exactly 298 B');
            verifyByteIdentical(testCase, got, want, '_rels/.rels');
            testCase.verifyEqual(sha256hex(got), ...
                "77494e8e16bbf29213e494792349e3d49d2397a4007b2816a058337d804e79a1", ...
                '_rels/.rels SHA-256 must equal the frozen s0005 reference');
        end

        function test_writer_document_rels_L1_283(testCase)
            % Regression (L1): the written word/_rels/document.xml.rels is
            % BYTE-identical (283 B) + SHA-256-identical to the frozen s0005 ref.
            [blobs, names] = testCase.buildS0005Package();
            got = testCase.entryBlob(blobs, names, "word/_rels/document.xml.rels");
            want = testCase.loadFixture('s0005_doc_rels');
            testCase.verifyEqual(numel(got), 283, ...
                'word/_rels/document.xml.rels must be exactly 283 B');
            verifyByteIdentical(testCase, got, want, 'word/_rels/document.xml.rels');
            testCase.verifyEqual(sha256hex(got), ...
                "f2764a1203a064b26382ba924f8f77fc6cc634f35955ebe41d57af43a68eff12", ...
                'document.xml.rels SHA-256 must equal the frozen s0005 reference');
        end

        function test_writer_interleave_and_no_rels_emits_none(testCase)
            % Regression (H4 part-then-.rels interleave + no-rels-emits-none):
            % the ONLY .rels item is word/_rels/document.xml.rels, and it appears
            % IMMEDIATELY after word/document.xml. The four no-rels parts
            % (styles/thumbnail.jpeg/IMG.PNG/Core.XML) emit NO .rels item
            % (write_parts_ writes `.rels` iff part.rels.len > 0 -- pkgwriter.py 47-54).
            [~, names] = testCase.buildS0005Package();
            relsItems = names(endsWith(names, ".rels"));
            % Two .rels total: the package _rels/.rels and document.xml.rels only.
            testCase.verifyEqual(sort(relsItems), ...
                sort(["_rels/.rels", "word/_rels/document.xml.rels"]), ...
                'exactly the pkg-rels and document.xml.rels items may appear');
            iDoc = find(names == "word/document.xml");
            testCase.verifyEqual(names(iDoc + 1), "word/_rels/document.xml.rels", ...
                'document.xml.rels must immediately follow word/document.xml');
            % No .rels for any of the four no-rels parts.
            for nm = ["word/_rels/styles.xml.rels", ...
                      "docProps/_rels/thumbnail.jpeg.rels", ...
                      "word/media/_rels/IMG.PNG.rels", ...
                      "docProps/_rels/Core.XML.rels"]
                testCase.verifyFalse(any(names == nm), ...
                    sprintf('no-rels part must emit no .rels item: %s', nm));
            end
        end

        function test_writer_preserves_caller_order_no_sort(testCase)
            % Regression (write_parts_ does NOT reorder): the caller-supplied part
            % order (document, styles, thumbnail, IMG.PNG, Core.XML) is NOT
            % alphabetically sorted -- a sort would place docProps/* before word/*.
            % The emitted part order matching the GIVEN order proves the writer
            % preserves the caller sequence with no sort/reorder (dedup is
            % iter_parts scope -- P1-8; the reader's visited-dedup is pinned below).
            [~, names] = testCase.buildS0005Package();
            partOrder = names(~endsWith(names, ".rels") & ...
                names ~= "[Content_Types].xml");
            testCase.verifyEqual(partOrder, [ ...
                "word/document.xml", "word/styles.xml", ...
                "docProps/thumbnail.jpeg", "word/media/IMG.PNG", ...
                "docProps/Core.XML"], ...
                'parts must appear in the caller-given (unsorted) order');
            % Sanity: this order is genuinely NOT what a sort would produce.
            testCase.verifyNotEqual(partOrder, sort(partOrder), ...
                'the caller order must differ from sorted order (guards the pin)');
        end

        % =============================================================== %
        % 2. Reader round-trip vs real default.docx (s0006)               %
        % =============================================================== %

        function test_reader_sparts_count_dfs_order_dedup(testCase)
            % Equivalence (s0006): iter_sparts yields exactly 13 parts, in the
            % frozen DFS emission order, with all 13 partnames UNIQUE (the threaded
            % visited-list dedup -- a part reached by two rel paths is emitted once).
            rdr = testCase.reader();
            sparts = rdr.iter_sparts();
            exp = testCase.probe6().sparts;
            testCase.verifyEqual(numel(sparts), 13, 'iter_sparts must yield 13 parts');
            got = arrayfun(@(s) string(s.partname), sparts);
            want = arrayfun(@(e) string(e.partname), exp(:)');
            testCase.verifyEqual(got, want, ...
                'iter_sparts partnames must match the frozen s0006 DFS order');
            testCase.verifyEqual(numel(unique(got)), 13, ...
                'all 13 reader partnames must be unique (visited-dedup)');
        end

        function test_reader_sparts_ct_reltype_size_sha(testCase)
            % Equivalence (s0006): for each of the 13 serialized parts the
            % content_type, reltype, blob size, and blob SHA-256 equal the frozen
            % python-docx reader oracle byte-for-byte.
            rdr = testCase.reader();
            sparts = rdr.iter_sparts();
            exp = testCase.probe6().sparts;
            testCase.assertEqual(numel(sparts), numel(exp), 'spart count');
            for k = 1:numel(sparts)
                s = sparts(k);  e = exp(k);
                pn = char(e.partname);
                testCase.verifyEqual(string(s.content_type), string(e.content_type), ...
                    sprintf('%s content_type', pn));
                testCase.verifyEqual(string(s.reltype), string(e.reltype), ...
                    sprintf('%s reltype', pn));
                testCase.verifyEqual(numel(s.blob), double(e.size), ...
                    sprintf('%s blob size', pn));
                testCase.verifyEqual(sha256hex(s.blob), lower(string(e.sha256)), ...
                    sprintf('%s blob SHA-256 must match the frozen s0006 reference', pn));
            end
        end

        function test_reader_srels_count_order(testCase)
            % Equivalence (s0006): iter_srels yields exactly 13 rows, in emission
            % order (package rels first, then per-part rels), each
            % source_uri | rId | reltype | target | is_external matching the frozen
            % oracle. Internal target = joined target_partname; external = target_ref.
            rdr = testCase.reader();
            srels = rdr.iter_srels();
            exp = testCase.probe6().srels;
            testCase.verifyEqual(numel(srels), 13, 'iter_srels must yield 13 rows');
            for k = 1:numel(srels)
                r = srels(k).srel;  e = exp(k);
                tag = sprintf('srel %d', k);
                testCase.verifyEqual(string(srels(k).source_uri), string(e.source_uri), ...
                    [tag ' source_uri']);
                testCase.verifyEqual(string(r.rId), string(e.rId), [tag ' rId']);
                testCase.verifyEqual(string(r.reltype), string(e.reltype), [tag ' reltype']);
                testCase.verifyEqual(logical(r.is_external), logical(e.is_external), ...
                    [tag ' is_external']);
                if r.is_external
                    tgt = string(r.target_ref);
                else
                    tgt = string(r.target_partname);
                end
                testCase.verifyEqual(tgt, string(e.target), [tag ' target']);
            end
        end

        function test_reader_pkg_rels_doc_order_not_rid_sorted(testCase)
            % Equivalence + Regression (H11, no rId sort): the 4 package-level rels
            % (source "/") emit in DOCUMENT order rId3,rId4,rId1,rId2 -- NOT
            % rId-sorted. This pins the reader's insertion-order rel emission.
            rdr = testCase.reader();
            srels = rdr.iter_srels();
            isPkg = arrayfun(@(t) string(t.source_uri) == "/", srels);
            pkgIds = arrayfun(@(t) string(t.srel.rId), srels(isPkg));
            testCase.verifyEqual(pkgIds, ["rId3" "rId4" "rId1" "rId2"], ...
                'package rels must be in document order rId3,rId4,rId1,rId2 (not sorted)');
        end

        % =============================================================== %
        % 3. _ContentTypeMap precedence (s0007)                           %
        % =============================================================== %

        function test_ctmap_precedence(testCase)
            % Regression + Equivalence (s0007): Override exact beats Default;
            % Override AND Default lookups are case-INSENSITIVE on the reader side
            % (both maps CaseInsensitiveDict, H15); xml Default resolves. Values are
            % the frozen s0007 precedence oracle over the real default.docx.
            ctm = testCase.ctmap();
            p = testCase.probe7().precedence;
            P = @(s) mat2doc.opc.PackURI(s);
            testCase.verifyEqual(string(ctm.getitem(P("/word/document.xml"))), ...
                string(p.override_exact), 'Override exact partname');
            testCase.verifyEqual(string(ctm.getitem(P("/WORD/DOCUMENT.XML"))), ...
                string(p.override_ci), 'Override case-insensitive (/WORD/DOCUMENT.XML)');
            testCase.verifyEqual(string(ctm.getitem(P("/docProps/thumbnail.jpeg"))), ...
                string(p.default_ext), 'Default by extension (jpeg)');
            testCase.verifyEqual(string(ctm.getitem(P("/docProps/THUMBNAIL.JPEG"))), ...
                string(p.default_ext_ci), 'Default extension case-insensitive');
            testCase.verifyEqual(string(ctm.getitem(P("/customXml/item1.xml"))), ...
                string(p.xml_default), 'xml Default');
        end

        function test_ctmap_missing_keyerror_verbatim(testCase)
            % Edge / error path (s0007): a partname absent from both maps raises
            % the IDENTIFIER mat2doc:KeyError, with the message VERBATIM-identical
            % to python-docx (pkgreader.py 95-105).
            ctm = testCase.ctmap();
            P = @(s) mat2doc.opc.PackURI(s);
            f = @() ctm.getitem(P("/nope/nothing.zzz"));
            testCase.verifyError(f, "mat2doc:KeyError", ...
                'missing partname must raise mat2doc:KeyError');
            wantMsg = string(testCase.probe7().precedence.missing_msg);
            got = "";
            try, f(); catch ME, got = string(ME.message); end
            testCase.verifyEqual(got, wantMsg, ...
                'KeyError message must be verbatim-identical to the frozen s0007 oracle');
        end

        function test_ctmap_nonpackuri_keyerror(testCase)
            % Edge / error path (s0007, VERIFY-3): a non-PackURI key raises the
            % IDENTIFIER mat2doc:KeyError (isa guard, pkgreader.py 95-96). The
            % message text differs from Python only in the type spelling
            % (accepted unreachable-diagnostic divergence, Gate-2/Gate-3), so ONLY
            % the identifier is pinned here.
            ctm = testCase.ctmap();
            testCase.verifyError(@() ctm.getitem('/word/document.xml'), ...
                "mat2doc:KeyError", ...
                'non-PackURI key must raise mat2doc:KeyError (identifier pin)');
        end

        % =============================================================== %
        % 4. V-3 target_partname cache (s0007)                            %
        % =============================================================== %

        function test_v3_load_rows(testCase)
            % Regression (s0007 V-3): load_from_xml over a fixed 2-row Relationships
            % XML yields 2 rels carrying rId / reltype / is_external as frozen.
            arr = testCase.v3array();
            v = testCase.probe7().v3;
            testCase.verifyEqual(numel(arr), double(v.count), 'V-3 rel count');
            testCase.verifyEqual(arrayfun(@(r) string(r.rId), arr), ...
                string(v.rId(:)'), 'V-3 rIds');
            testCase.verifyEqual(arrayfun(@(r) string(r.reltype), arr), ...
                string(v.reltype(:)'), 'V-3 reltypes');
            testCase.verifyEqual(arrayfun(@(r) logical(r.is_external), arr), ...
                logical(v.is_external(:)'), 'V-3 is_external flags');
        end

        function test_v3_target_partname_posix_join(testCase)
            % Regression (s0007 V-3): the internal rel's target_partname joins the
            % /word baseURI with ../docProps/core.xml via posix `..` traversal ->
            % /docProps/core.xml (frozen oracle).
            arr = testCase.v3array();
            testCase.verifyEqual(string(arr(1).target_partname), ...
                string(testCase.probe7().v3.target_partname), ...
                'target_partname posix ..-join must equal /docProps/core.xml');
        end

        function test_v3_cache_stable(testCase)
            % Regression (s0007 V-3): target_partname is lazily computed and cached
            % (hasattr guard); repeat access returns the SAME value (cache-stable).
            arr = testCase.v3array();
            tp1 = arr(1).target_partname;
            tp2 = arr(1).target_partname;
            testCase.verifyEqual(string(tp1), string(tp2), ...
                'target_partname must be stable across repeat access (lazy cache)');
            testCase.verifyTrue(testCase.probe7().v3.cache_stable, ...
                'frozen oracle records cache_stable == true');
        end

        function test_v3_external_target_ref_and_flag(testCase)
            % Regression (s0007 V-3): the external rel reports the URI verbatim as
            % target_ref and is_external == true.
            arr = testCase.v3array();
            testCase.verifyTrue(logical(arr(2).is_external), 'row 2 is external');
            testCase.verifyEqual(string(arr(2).target_ref), ...
                string(testCase.probe7().v3.target_ref_external), ...
                'external target_ref must be the URI verbatim');
        end

        function test_v3_external_target_partname_valueerror_verbatim(testCase)
            % Edge / error path (s0007 V-3): target_partname on an External rel
            % raises the IDENTIFIER mat2doc:ValueError with the message
            % VERBATIM-identical to python-docx (incl. the quoted "External").
            arr = testCase.v3array();
            f = @() arr(2).target_partname;
            testCase.verifyError(f, "mat2doc:ValueError", ...
                'External target_partname must raise mat2doc:ValueError');
            got = "";
            try, f(); catch ME, got = string(ME.message); end
            testCase.verifyEqual(got, string(testCase.probe7().v3.external_msg), ...
                'ValueError message must be verbatim-identical to the frozen oracle');
        end

        function test_v3_load_none_empty(testCase)
            % Edge (s0007 V-3): load_from_xml with [] (Python None -- no .rels item)
            % returns an EMPTY collection (count 0), NOT an error.
            empt = mat2doc.opc.SerializedRelationships_.load_from_xml("/word", []);
            testCase.verifyEqual(numel(empt.to_array()), ...
                double(testCase.probe7().v3.empty_on_none_count), ...
                'load_from_xml(None) must yield an empty collection (count 0)');
        end

        % =============================================================== %
        % 5. bytes currency (H2)                                          %
        % =============================================================== %

        function test_bytes_currency_writer_boundary_uint8(testCase)
            % Regression (H2): every value handed to phys_writer.write is uint8 --
            % cti.blob (serialize_part_xml), pkg_rels.xml, and part.rels.xml. A char
            % leak would latin-1-corrupt non-ASCII (and break the s0005 L1 pins).
            [parts, pkg_rels, docPart] = testCase.buildS0005Inputs();
            cti = mat2doc.opc.ContentTypesItem_.from_parts(parts);
            testCase.verifyClass(cti.blob, 'uint8', 'cti.blob must be uint8');
            testCase.verifyClass(pkg_rels.xml, 'uint8', 'pkg_rels.xml must be uint8');
            testCase.verifyClass(docPart.rels.xml, 'uint8', 'part.rels.xml must be uint8');
        end

        function test_bytes_currency_reader_blobs_uint8(testCase)
            % Regression (H2): all 13 reader blobs are uint8 -- no char/string leak
            % out of the read path.
            rdr = testCase.reader();
            sparts = rdr.iter_sparts();
            for k = 1:numel(sparts)
                testCase.verifyClass(sparts(k).blob, 'uint8', ...
                    sprintf('reader blob %d must be uint8', k));
            end
        end

    end

    % ===================== instance helpers ============================ %
    methods (Access = private)

        function [parts, pkg_rels, docPart] = buildS0005Inputs(testCase) %#ok<MANU>
            % The IDENTICAL call sequence as scenarios\s0005_writer_dfs_order.m
            % (the Gate-3 twin): a 5-part mixed-case stub, pkg-rels rooted at the
            % document, and a single document->styles rel.
            CT = mat2doc.opc.CONTENT_TYPE;
            RT = mat2doc.opc.RELATIONSHIP_TYPE;
            P  = @(s) mat2doc.opc.PackURI(s);
            docPart = StubPart_p1_6a(P("/word/document.xml"), CT.WML_DOCUMENT_MAIN, uint8('<doc/>'));
            styPart = StubPart_p1_6a(P("/word/styles.xml"), CT.WML_STYLES, uint8('<sty/>'));
            jpgPart = StubPart_p1_6a(P("/docProps/thumbnail.jpeg"), "image/jpeg", uint8([255 216 255]));
            imgPart = StubPart_p1_6a(P("/word/media/IMG.PNG"), "image/png", uint8([137 80 78 71]));
            mixPart = StubPart_p1_6a(P("/docProps/Core.XML"), CT.OPC_CORE_PROPERTIES, uint8('<core/>'));
            docPart.rels.add_relationship(RT.STYLES, styPart, "rId1", false);
            pkg_rels = mat2doc.opc.Relationships("/");
            pkg_rels.add_relationship(RT.OFFICE_DOCUMENT, docPart, "rId1", false);
            parts = [docPart styPart jpgPart imgPart mixPart];
        end

        function [blobs, names] = buildS0005Package(testCase)
            % Write the s0005 stub package via mat2doc.opc.PackageWriter.write to a
            % BINARY-mode temp .docx (the writer opens it "wb" -- no CRLF), read the
            % bytes back, and enumerate the zip entries in stream (write) order.
            [parts, pkg_rels] = testCase.buildS0005Inputs();
            tmp = [tempname '.docx'];
            cleanup = onCleanup(@() deleteIfExists(tmp)); %#ok<NASGU>
            mat2doc.opc.PackageWriter.write(tmp, pkg_rels, parts);
            f = fopen(tmp, 'r', 'n');            % binary read
            assert(f >= 0, 'could not reopen written package %s', tmp);
            zipBytes = fread(f, Inf, '*uint8')';
            fclose(f);
            [blobs, names] = zipEntryList(zipBytes);
        end

        function rdr = reader(testCase)
            rdr = mat2doc.opc.PackageReader.from_file(char(testCase.templatePath()));
        end

        function ctm = ctmap(testCase)
            % Build a _ContentTypeMap from the real default.docx [Content_Types].xml
            % via the live PhysPkgReader flow (as scenarios\s0007_ctmap_v3_probe.m).
            tpl = testCase.templatePath();
            f = fopen(tpl, 'r', 'n');
            assert(f >= 0, 'template missing: %s', tpl);
            tplb = fread(f, Inf, '*uint8')';
            fclose(f);
            pr = mat2doc.opc.PhysPkgReader.factory(tplb);
            ctm = mat2doc.opc.ContentTypeMap_.from_xml(pr.content_types_xml);
            pr.close();
        end

        function arr = v3array(testCase) %#ok<MANU>
            % Load the fixed 2-row Relationships XML (identical to
            % scenarios\s0007_ctmap_v3_probe.m) and return the rel array.
            relxml = uint8(['<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' ...
                '<Relationship Id="rId1" Type="http://x/y" Target="../docProps/core.xml"/>' ...
                '<Relationship Id="rId2" Type="http://x/z" Target="http://example.com/x" TargetMode="External"/>' ...
                '</Relationships>']);
            sr = mat2doc.opc.SerializedRelationships_.load_from_xml("/word", relxml);
            arr = sr.to_array();
        end

        function p = templatePath(~)
            here = fileparts(mfilename('fullpath'));       % tests\opc
            root = fileparts(fileparts(here));             % worktree root
            p = fullfile(root, '+mat2doc', 'templates', 'default.docx');
        end

        function b = loadFixture(~, base)
            here = fileparts(mfilename('fullpath'));       % tests\opc
            p = fullfile(here, 'data', [base '.bin']);
            f = fopen(p, 'r', 'n');
            assert(f >= 0, 'byte fixture missing: %s', p);
            b = fread(f, Inf, '*uint8')';
            fclose(f);
        end

        function s = probe6(testCase)
            s = jsondecode(testCase.readTextFixture('s0006_probe.json'));
        end

        function s = probe7(testCase)
            s = jsondecode(testCase.readTextFixture('s0007_probe.json'));
        end

        function txt = readTextFixture(~, name)
            here = fileparts(mfilename('fullpath'));       % tests\opc
            p = fullfile(here, 'data', name);
            f = fopen(p, 'r', 'n');                        % binary read (no CRLF xlate)
            assert(f >= 0, 'json fixture missing: %s', p);
            raw = fread(f, Inf, '*uint8')';
            fclose(f);
            txt = native2unicode(raw, 'UTF-8');
        end

        function blob = entryBlob(~, blobs, names, membername)
            i = find(names == string(membername), 1);
            assert(~isempty(i), 'zip entry not found: %s', membername);
            blob = blobs{i};
        end
    end
end

% ===================== file-local helpers ============================== %

function [blobs, names] = zipEntryList(zipBytes)
    % Enumerate an in-memory zip in stream (write) order -> ordered {names, blobs}.
    % Uses java.util.zip.ZipInputStream (reads local file headers in physical
    % order) so `names` is the true zip-entry write sequence. Mirrors the
    % ZipPkgReader_.readZipEntries_ mechanism; kept file-local so this order pin
    % is independent of the reader under test.
    bais = java.io.ByteArrayInputStream(int8(typecast(uint8(zipBytes(:)'), 'int8')));
    zis  = java.util.zip.ZipInputStream(bais);
    cleanup = onCleanup(@() zis.close()); %#ok<NASGU>
    copier = com.mathworks.mlwidgets.io.InterruptibleStreamCopier.getInterruptibleStreamCopier;
    names = strings(1, 0);
    blobs = {};
    while true
        ze = zis.getNextEntry();
        if isempty(ze)          % Java null -> no more entries
            break
        end
        names(end + 1) = string(ze.getName()); %#ok<AGROW>
        baos = java.io.ByteArrayOutputStream;
        copier.copyStream(zis, baos);
        blobs{end + 1} = typecast(int8(baos.toByteArray()), 'uint8')'; %#ok<AGROW>
        zis.closeEntry();
    end
end

function deleteIfExists(p)
    if isfile(p)
        delete(p);
    end
end

function verifyByteIdentical(testCase, got, want, label)
    % Byte-level assertion (L1). On mismatch report sizes and first diff offset.
    got = uint8(got(:)');  want = uint8(want(:)');
    if ~isequal(got, want)
        n = min(numel(got), numel(want));
        d = find(got(1:n) ~= want(1:n), 1);
        if isempty(d), d = n + 1; end
        gv = 0; wv = 0;
        if d <= numel(got), gv = double(got(d)); end
        if d <= numel(want), wv = double(want(d)); end
        diag = sprintf(['%s: bytes differ (got %d B, want %d B); first diff @%d ' ...
            '(got 0x%02X, want 0x%02X)'], char(label), numel(got), numel(want), ...
            d, gv, wv);
    else
        diag = char(label);
    end
    testCase.verifyEqual(got, want, diag);
end

function h = sha256hex(bytes)
    % Lowercase hex SHA-256 of a uint8 vector (base MATLAB / Java MessageDigest).
    md  = java.security.MessageDigest.getInstance('SHA-256');
    dig = md.digest(typecast(uint8(bytes(:)'), 'int8'));
    b   = typecast(dig, 'uint8');
    h   = lower(string(sprintf('%02x', b)));
end

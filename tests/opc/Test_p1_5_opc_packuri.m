classdef Test_p1_5_opc_packuri < matlab.unittest.TestCase
% TEST_P1_5_OPC_PACKURI  Gate-4 permanent unit tests for Mat2Doc P1-5
%   (opc packuri + phys_pkg + rel).
%
%   Freezes the P1-5 guarantees ported in +mat2doc\+opc\ from python-docx v1.2.0
%   src/docx/opc/{packuri,phys_pkg,rel}.py: the PackURI value machinery, the
%   Relationships/Relationship_ collection + .rels serialization, the D-zip-time
%   zip writer, and the java.util.zip byte boundary.
%
%   Provenance (Gate-1..3, all 2026-07-25):
%     * Audit   : validation\mat2doc\audit_P1-5_opc_packuri.md (Porter Gate-1 +
%                 Opus Gate-2 APPROVE; FLAG-1/FLAG-2 both RATIFY)
%     * Validate : validation\mat2doc\validate_P1-5_opc_packuri.md (Gate-3 PASS,
%                 all 5 bar items; 0 FAIL; 0 new D-numbers; regression 232/232)
%     * Scenarios: validation\mat2doc\scenarios\s0003_packuri_rel_probe.{py,m}
%                 (PackURI/rId probe) and s0004_rels_l1.{py,m} + s0004_FakePart.m
%                 (the 3 .rels L1 fixtures)
%     * Frozen refs: references\s0003\probe.json (probe_diff MATCH vs live
%                 python-docx 1.2.0) and references\s0004\parts\... (SHA-256 per
%                 part, manifest.json). The 3 .rels are copied byte-for-byte into
%                 tests\opc\data\ (rels{1,2,3}.bin) so this suite is self-contained.
%
%   Coverage taxonomy
%   -----------------
%   * Equivalence + Regression -- the PackURI member battery values are the frozen
%     s0003 oracle values (probe_diff MATCH); the 3 .rels byte-fixtures are the
%     frozen s0004 oracle bytes (byte-identical AND SHA-256-identical).
%   * Regression -- the docx-specific idx `[1-9][0-9]*` distinctions
%     (image01.png->[], image10.png->10, image1.jpeg->1, media.png->[]) that
%     separate the docx regex from pptx's `[0-9][0-9]*`; the rels2 insertion-order
%     (rId3,rId1,rId2 -- no rId sort, H11) and rels3 internal-target + H5 identity;
%     hard-coded rId gap-reuse / dedup / target_ref; the D-zip-time 1980 stamps.
%   * Edge / error path -- ctor no-leading-slash and empty-string both raise the
%     IDENTIFIER mat2doc:ValueError; target_part on an external rel raises
%     mat2doc:ValueError; non-ASCII (é + CJK + emoji) + all-256-byte blob through
%     the zip round-trip and the Java boundary.
%
%   Deviations exercised (adopt-only, ZERO new D-numbers -- Gate-3):
%     D-zip-time (fixed 1980-01-01 entry stamp) and D-001 (own OOXML-subset
%     serializer, via CT_Relationships.xml_file_bytes). Every pin re-verifies them
%     byte-neutral. The three .rels are L1 (the Gate-3 ladder demanded L1).
%
%   D-zip-time parity note (item 3): Gate-3 proved the Mat2Doc stateful writer
%   byte-identical to the shipped mat2ppt.opc.ZipPkgWriter_. Reaching Mat2Ppt from
%   the COLD mat2doc test path would require a non-self-contained cross-repo path
%   (Mat2Ppt checked out as a sibling), breaking Gate-4 self-containment, so this
%   class pins the FALLBACK the brief allows: run-to-run byte reproducibility + a
%   frozen whole-zip expected-bytes fixture (data\dzip_expected.bin -- the
%   Gate-3-proven Mat2Ppt-parity output, captured on this machine) + the fixed-1980
%   DOS stamp on every header. The 1980-stamp + reproducibility + round-trip pins
%   are timezone-robust; the whole-zip fixture is a same-machine hard pin (the
%   D-zip-time envelope's UT extra field encodes the fixed 1980 local instant, so
%   whole-zip bytes are deterministic per machine -- the signed same-machine scope
%   of D-zip-time; the merge check runs on the same machine).
%
%   Determinism: no network, no absolute paths -- fixtures resolve relative to this
%   file via fileparts(mfilename('fullpath')).

    properties (TestParameter)
        % --- PackURI member battery: one case per partname. Expected values are
        %     the FROZEN s0003 oracle (references\s0003\probe.json, probe_diff
        %     MATCH vs live python-docx 1.2.0). idxNone==true means Python None
        %     (MATLAB []). Fields: {uri, baseURI, ext, filename, idxNone, idxVal,
        %     membername, rels_uri, relref_root, relref_word, relref_media}.
        puCase = struct( ...
            'root', {{ ...
                "/", "/", "", "", true, 0, "", "/_rels/.rels", ...
                "", "..", "../.."}}, ...
            'content_types', {{ ...
                "/[Content_Types].xml", "/", "xml", "[Content_Types].xml", true, 0, ...
                "[Content_Types].xml", "/_rels/[Content_Types].xml.rels", ...
                "[Content_Types].xml", "../[Content_Types].xml", "../../[Content_Types].xml"}}, ...
            'docprops_app', {{ ...
                "/docProps/app.xml", "/docProps", "xml", "app.xml", true, 0, ...
                "docProps/app.xml", "/docProps/_rels/app.xml.rels", ...
                "docProps/app.xml", "../docProps/app.xml", "../../docProps/app.xml"}}, ...
            'document_rels', {{ ...
                "/word/_rels/document.xml.rels", "/word/_rels", "rels", ...
                "document.xml.rels", true, 0, "word/_rels/document.xml.rels", ...
                "/word/_rels/_rels/document.xml.rels.rels", ...
                "word/_rels/document.xml.rels", "_rels/document.xml.rels", ...
                "../_rels/document.xml.rels"}}, ...
            'document', {{ ...
                "/word/document.xml", "/word", "xml", "document.xml", true, 0, ...
                "word/document.xml", "/word/_rels/document.xml.rels", ...
                "word/document.xml", "document.xml", "../document.xml"}}, ...
            'embedding_docx', {{ ...
                "/word/embeddings/Microsoft_Word_Document.docx", "/word/embeddings", ...
                "docx", "Microsoft_Word_Document.docx", true, 0, ...
                "word/embeddings/Microsoft_Word_Document.docx", ...
                "/word/embeddings/_rels/Microsoft_Word_Document.docx.rels", ...
                "word/embeddings/Microsoft_Word_Document.docx", ...
                "embeddings/Microsoft_Word_Document.docx", ...
                "../embeddings/Microsoft_Word_Document.docx"}}, ...
            'image01_png', {{ ...
                "/word/media/image01.png", "/word/media", "png", "image01.png", ...
                true, 0, "word/media/image01.png", ...
                "/word/media/_rels/image01.png.rels", "word/media/image01.png", ...
                "media/image01.png", "image01.png"}}, ...
            'image1_jpeg', {{ ...
                "/word/media/image1.jpeg", "/word/media", "jpeg", "image1.jpeg", ...
                false, 1, "word/media/image1.jpeg", ...
                "/word/media/_rels/image1.jpeg.rels", "word/media/image1.jpeg", ...
                "media/image1.jpeg", "image1.jpeg"}}, ...
            'image10_png', {{ ...
                "/word/media/image10.png", "/word/media", "png", "image10.png", ...
                false, 10, "word/media/image10.png", ...
                "/word/media/_rels/image10.png.rels", "word/media/image10.png", ...
                "media/image10.png", "image10.png"}}, ...
            'media_png', {{ ...
                "/word/media/media.png", "/word/media", "png", "media.png", ...
                true, 0, "word/media/media.png", ...
                "/word/media/_rels/media.png.rels", "word/media/media.png", ...
                "media/media.png", "media.png"}}, ...
            'theme1', {{ ...
                "/word/theme/theme1.xml", "/word/theme", "xml", "theme1.xml", ...
                false, 1, "word/theme/theme1.xml", ...
                "/word/theme/_rels/theme1.xml.rels", "word/theme/theme1.xml", ...
                "theme/theme1.xml", "../theme/theme1.xml"}});

        % --- .rels L1 byte-pins: {fixture-basename, expected-size, sha256}
        %     from references\s0004\manifest.json (frozen python-docx oracle).
        relsCase = struct( ...
            'rels1_two_ext', {{ 'rels1', 474, ...
                'f7912867b28478e009c178d8476ff60e8b05f3f896ace215a1a6707198432fb6'}}, ...
            'rels2_insertion_order', {{ 'rels2', 629, ...
                '9cdedcc8804449ea49ca49e2558fad36001ed1e57f326024a0fca10b47232123'}}, ...
            'rels3_internal_h5', {{ 'rels3', 579, ...
                'f7a0c446c45feac13883a11aeb9c4fa43a573f4d7c05e00d745bb4ee0b090851'}});
    end

    properties (Constant)
        HYP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
        IMG = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied from tests\opc\Test_p1_4_opc_oxml.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\opc
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
            % Also add tests\opc so the FakePart_p1_5 helper class resolves under
            % a cold suite run regardless of the runner's cd behavior.
            testCase.applyFixture(PathFixture(here));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. PackURI value battery (Equivalence + Regression, s0003)      %
        % =============================================================== %

        function test_packuri_members(testCase, puCase)
            % Equivalence + Regression: every PackURI member equals the frozen
            % s0003 oracle value (probe_diff MATCH vs python-docx 1.2.0). Covers
            % baseURI / ext (dot-stripped) / filename / idx / membername /
            % rels_uri / relative_ref x3 / from_rel_ref-companion relref values.
            c = puCase;
            uri = c{1};
            p = mat2doc.opc.PackURI(uri);
            testCase.verifyEqual(string(p.baseURI), c{2}, 'baseURI');
            testCase.verifyEqual(string(p.ext), c{3}, 'ext (dot-stripped)');
            testCase.verifyEqual(string(p.filename), c{4}, 'filename');
            if c{5}   % idxNone
                testCase.verifyEmpty(p.idx, ...
                    'idx must be [] (Python None) for this partname');
            else
                testCase.verifyEqual(p.idx, c{6}, 'idx');
            end
            testCase.verifyEqual(string(p.membername), c{7}, 'membername');
            testCase.verifyEqual(string(p.rels_uri), c{8}, 'rels_uri');
            testCase.verifyEqual(string(p.relative_ref("/")), c{9}, ...
                'relative_ref("/")');
            testCase.verifyEqual(string(p.relative_ref("/word")), c{10}, ...
                'relative_ref("/word")');
            testCase.verifyEqual(string(p.relative_ref("/word/media")), c{11}, ...
                'relative_ref("/word/media")');
        end

        function test_idx_docx_leadingzero_regex(testCase)
            % Regression (H12, THE docx-vs-pptx guard): docx packuri.py:18 uses
            % `([a-zA-Z]+)([1-9][0-9]*)?` -- a leading-ZERO numeric suffix does
            % NOT match, so image01.png -> [] (None), NOT 1. pptx uses
            % `[0-9][0-9]*` and would yield 1 here. These four pins go RED if the
            % regex is ever relaxed to the pptx form. (Frozen s0003 idx column.)
            testCase.verifyEmpty(mat2doc.opc.PackURI("/word/media/image01.png").idx, ...
                'image01.png: leading-zero suffix must NOT match [1-9][0-9]* -> []');
            testCase.verifyEqual(mat2doc.opc.PackURI("/word/media/image10.png").idx, 10, ...
                'image10.png: idx 10');
            testCase.verifyEqual(mat2doc.opc.PackURI("/word/media/image1.jpeg").idx, 1, ...
                'image1.jpeg: idx 1');
            testCase.verifyEmpty(mat2doc.opc.PackURI("/word/media/media.png").idx, ...
                'media.png: no trailing integer -> []');
        end

        function test_from_rel_ref(testCase)
            % Equivalence + Regression (s0003 from_rel_ref, incl. `../`
            % normalization): abspath(join(baseURI, relative_ref)).
            frr = @(b, r) string(mat2doc.opc.PackURI.from_rel_ref(b, r));
            testCase.verifyEqual(frr("/word", "media/image1.png"), ...
                "/word/media/image1.png", 'from_rel_ref a');
            testCase.verifyEqual(frr("/word/slides", "../media/image1.png"), ...
                "/word/media/image1.png", 'from_rel_ref b (../ normalization)');
            testCase.verifyEqual(frr("/", "word/document.xml"), ...
                "/word/document.xml", 'from_rel_ref c');
            testCase.verifyEqual(frr("/word/_rels", "../theme/theme1.xml"), ...
                "/word/theme/theme1.xml", 'from_rel_ref d');
        end

        function test_uri_constants(testCase)
            % Equivalence + Regression (s0003): the module constants.
            pu = mat2doc.opc.PACKAGE_URI();
            testCase.verifyEqual(string(pu), "/", 'PACKAGE_URI');
            testCase.verifyEqual(string(pu.rels_uri), "/_rels/.rels", ...
                'PACKAGE_URI.rels_uri');
            testCase.verifyEqual(string(mat2doc.opc.CONTENT_TYPES_URI()), ...
                "/[Content_Types].xml", 'CONTENT_TYPES_URI');
        end

        function test_ctor_no_leading_slash_raises(testCase)
            % Edge / error path (s0003 ctor_noslash): a partname without a leading
            % slash raises the IDENTIFIER mat2doc:ValueError (packuri.py:20-24).
            testCase.verifyError(@() mat2doc.opc.PackURI("word/doc.xml"), ...
                "mat2doc:ValueError", ...
                'no-leading-slash partname must raise mat2doc:ValueError');
        end

        function test_ctor_empty_raises(testCase)
            % Edge / error path: the empty string raises mat2doc:ValueError (the
            % guard for Python str[0] IndexError; documented in-code as defensive).
            testCase.verifyError(@() mat2doc.opc.PackURI(""), ...
                "mat2doc:ValueError", ...
                'empty partname must raise mat2doc:ValueError');
        end

        % =============================================================== %
        % 2. Relationships.xml L1 byte-pins (Equivalence + Regression)    %
        % =============================================================== %

        function test_rels_L1_byte_identical(testCase, relsCase)
            % Equivalence + Regression (L1): rebuild each .rels via the frozen
            % s0004 call sequence and assert its xml (== Python Relationships.xml)
            % BYTE-IDENTICAL and SHA-256-identical to the frozen python-docx part.
            % Byte-level L1: the Gate-3 ladder demanded L1 for every .rels.
            base = relsCase{1};  expsz = relsCase{2};  expsha = relsCase{3};
            r = testCase.buildRels(base);
            got = uint8(r.xml);  got = got(:)';
            want = testCase.loadFixture(base);
            testCase.verifyEqual(numel(got), expsz, ...
                sprintf('%s must be exactly %d B', base, expsz));
            verifyByteIdentical(testCase, got, want, base);
            testCase.verifyEqual(sha256hex(got), string(expsha), ...
                sprintf('%s SHA-256 must equal the frozen s0004 reference', base));
        end

        function test_rels2_insertion_order_no_rId_sort(testCase)
            % Regression (H11, no rId sort): rels2 inserts rId3 BEFORE rId1, then
            % gap-fills rId2. The emitted <Relationship Id=...> order MUST be the
            % INSERTION order rId3,rId1,rId2 -- NOT numerically sorted (which the
            % pptx writer would produce). Parse the Id attributes off the bytes.
            r = testCase.buildRels('rels2');
            ids = idOrder(char(r.xml));
            testCase.verifyEqual(ids, ["rId3" "rId1" "rId2"], ...
                'rels2 must emit Ids in insertion order rId3,rId1,rId2 (no sort)');
        end

        function test_rels3_h5_distinct_rids_same_partname(testCase)
            % Regression (H5 identity): two DISTINCT FakePart objects with the SAME
            % partname must get DISTINCT rIds (rId1, rId2) -- the match is on part
            % identity, not partname equality -- and re-adding the SAME object
            % returns its existing rId. The internal Target serializes relative to
            % the /word baseURI (media/image1.png, no TargetMode).
            pA = FakePart_p1_5(mat2doc.opc.PackURI("/word/media/image1.png"));
            pB = FakePart_p1_5(mat2doc.opc.PackURI("/word/media/image1.png"));
            r = mat2doc.opc.Relationships("/word");
            relA1 = r.get_or_add(testCase.IMG, pA);
            relA2 = r.get_or_add(testCase.IMG, pA);   % identity -> same rId1
            relB  = r.get_or_add(testCase.IMG, pB);   % distinct object -> rId2
            testCase.verifyEqual(string(relA1.rId), "rId1", 'first internal -> rId1');
            testCase.verifyEqual(string(relA2.rId), "rId1", ...
                'same Part object -> SAME rId1 (H5 identity)');
            testCase.verifyEqual(string(relB.rId), "rId2", ...
                'different Part object, same partname -> DISTINCT rId2 (H5)');
            testCase.verifyEqual(string(relA1.target_ref), "media/image1.png", ...
                'internal target_ref relative to /word baseURI');
        end

        % =============================================================== %
        % 3. D-zip-time reproducibility + 1980 stamp + parity fixture     %
        % =============================================================== %

        function test_dzip_run_to_run_reproducible(testCase)
            % Regression (D-zip-time): two independent writers over the identical
            % ordered entries produce BYTE-IDENTICAL whole-zip output. This is the
            % core D-zip-time guarantee -- run-to-run reproducibility -- and is
            % timezone-robust.
            b1 = testCase.buildDzip();
            b2 = testCase.buildDzip();
            testCase.verifyEqual(b1, b2, ...
                'D-zip-time: whole-zip bytes must be run-to-run reproducible');
        end

        function test_dzip_expected_bytes_parity(testCase)
            % Regression (D-zip-time / Mat2Ppt parity, captured): the whole-zip
            % bytes match the frozen data\dzip_expected.bin -- the Gate-3-proven
            % Mat2Ppt-parity output captured on this machine. Goes RED on any
            % serializer/compression/order perturbation. Same-machine hard pin
            % (see the class D-zip-time parity note); the reproducibility + 1980
            % stamp + round-trip pins are the timezone-robust core.
            got = testCase.buildDzip();
            want = testCase.loadFixture('dzip_expected');
            verifyByteIdentical(testCase, got, want, 'dzip_expected.bin');
        end

        function test_dzip_all_headers_1980(testCase)
            % Regression (D-zip-time): every central-directory header carries the
            % fixed DOS stamp 1980-01-01 00:00 (time 0x0000, date 0x0021) -- the
            % GregorianCalendar(1980,0,1) setTime mechanism. Timezone-robust (the
            % DOS stamp is the fixed LOCAL 1980 instant in any timezone).
            bytes = testCase.buildDzip();
            [times, dates] = centralDirStamps(bytes);
            testCase.verifyEqual(numel(times), 3, ...
                'expected 3 central-directory headers (3 entries)');
            testCase.verifyTrue(all(times == 0), ...
                'every entry DOS mod-time must be 0x0000 (00:00)');
            testCase.verifyTrue(all(dates == hex2dec('0021')), ...
                'every entry DOS mod-date must be 0x0021 (1980-01-01)');
        end

        function test_dzip_roundtrip_nonascii_and_all_bytes(testCase)
            % Edge (non-ASCII + full byte range) + Regression: the é+CJK+emoji
            % UTF-8 blob concatenated with all 256 byte values round-trips
            % byte-identically through write -> zip -> ZipPkgReader_.blob_for.
            bytes = testCase.buildDzip();
            z = mat2doc.opc.ZipPkgReader_(bytes);
            got = z.blob_for(mat2doc.opc.PackURI("/word/document.xml"));
            testCase.verifyEqual(uint8(got(:)'), testCase.docBlob(), ...
                'document.xml blob (é+CJK+emoji+0..255) must round-trip byte-clean');
        end

        % =============================================================== %
        % 4. rId generation: gap-reuse, dedup, external target_ref        %
        % =============================================================== %

        function test_next_rId_gap_reuse(testCase)
            % Regression (H11 gap-reuse, s0003 gap_reuse): with {rId3, rId1}
            % present, the next external rel reuses the lowest gap -> rId2, and the
            % key order is insertion order rId3,rId1,rId2.
            r = mat2doc.opc.Relationships("/word");
            r.add_relationship(testCase.HYP, "http://x.example/", "rId3", true);
            r.add_relationship(testCase.HYP, "http://y.example/", "rId1", true);
            nid = r.get_or_add_ext_rel(testCase.IMG, "http://z.example/");
            testCase.verifyEqual(string(nid), "rId2", ...
                'next rId over {rId3,rId1} must reuse the gap -> rId2');
            testCase.verifyEqual(r.keys(), ["rId3" "rId1" "rId2"], ...
                'keys must be insertion order rId3,rId1,rId2');
        end

        function test_ext_rel_dedup(testCase)
            % Regression (s0003 ext_dedup): get_or_add_ext_rel dedups by
            % (reltype,target_ref) -> rId1, rId2, rId1; collection length 2.
            r = mat2doc.opc.Relationships("/word");
            r1 = r.get_or_add_ext_rel(testCase.HYP, "http://b.example/");
            r2 = r.get_or_add_ext_rel(testCase.HYP, "http://a.example/");
            r3 = r.get_or_add_ext_rel(testCase.HYP, "http://b.example/");
            testCase.verifyEqual(string(r1), "rId1", 'first ext -> rId1');
            testCase.verifyEqual(string(r2), "rId2", 'second ext -> rId2');
            testCase.verifyEqual(string(r3), "rId1", ...
                'duplicate (reltype,target) must dedup -> rId1');
            testCase.verifyEqual(r.len(), 2, 'collection length must be 2 after dedup');
        end

        function test_external_target_ref_and_flags(testCase)
            % Regression (s0003 external): an external rel reports the URI verbatim
            % as target_ref, is_external true, and its reltype.
            r = mat2doc.opc.Relationships("/word");
            r.get_or_add_ext_rel(testCase.HYP, "http://b.example/");
            rel = r.values(); rel = rel(1);
            testCase.verifyEqual(string(rel.rId), "rId1", 'rId');
            testCase.verifyTrue(logical(rel.is_external), 'is_external');
            testCase.verifyEqual(string(rel.reltype), testCase.HYP, 'reltype');
            testCase.verifyEqual(string(rel.target_ref), "http://b.example/", ...
                'external target_ref must be the URI verbatim');
        end

        function test_target_part_on_external_raises(testCase)
            % Edge / error path (s0003 target_part_ext): target_part on an external
            % relationship raises the IDENTIFIER mat2doc:ValueError (rel.py 139-145).
            r = mat2doc.opc.Relationships("/word");
            r.get_or_add_ext_rel(testCase.HYP, "http://b.example/");
            rel = r.values(); rel = rel(1);
            testCase.verifyError(@() rel.target_part, "mat2doc:ValueError", ...
                'target_part on an external rel must raise mat2doc:ValueError');
        end

        function test_dict_surface(testCase)
            % Regression (s0003 dict_surface): the inherited Dict[str,_Relationship]
            % surface -- contains / len / getitem.
            r = mat2doc.opc.Relationships("/word");
            r.get_or_add_ext_rel(testCase.HYP, "http://b.example/");
            r.get_or_add_ext_rel(testCase.HYP, "http://a.example/");
            testCase.verifyTrue(r.contains("rId1"), 'contains rId1');
            testCase.verifyFalse(r.contains("rId9"), 'not contains rId9');
            testCase.verifyEqual(r.len(), 2, 'len 2');
            testCase.verifyEqual(string(r.getitem("rId1").rId), "rId1", ...
                'getitem("rId1").rId');
        end

        % =============================================================== %
        % 5. Java byte boundary (H2)                                      %
        % =============================================================== %

        function test_java_boundary_roundtrip_0_255(testCase)
            % Regression (H2): bytesFromJava(bytesToJava(0..255)) == 0..255.
            u = uint8(0:255);
            back = mat2doc.opc.bytesFromJava(mat2doc.opc.bytesToJava(u));
            testCase.verifyEqual(uint8(back(:)'), u, ...
                '0..255 must round-trip identity through the Java boundary');
        end

        function test_java_boundary_signed_typecast(testCase)
            % Regression (H2 sign fidelity): bytesToJava is a signed typecast
            % (bit pattern kept), so 128->-128 and 255->-1, class int8.
            j = mat2doc.opc.bytesToJava(uint8([0 127 128 255]));
            testCase.verifyClass(j, 'int8', 'bytesToJava must yield int8');
            testCase.verifyEqual(j, int8([0 127 -128 -1]), ...
                'signed typecast: 128->-128, 255->-1 (no saturation)');
        end

        function test_java_boundary_utf8_nonascii(testCase)
            % Edge (non-ASCII, H2): a UTF-8 blob (é + CJK + emoji) round-trips
            % byte-clean through the Java boundary with no sign corruption.
            u = uint8(unicode2native(char("é 文字 " + char(55357) + char(56898)), 'UTF-8'));
            back = mat2doc.opc.bytesFromJava(mat2doc.opc.bytesToJava(u));
            testCase.verifyEqual(uint8(back(:)'), u, ...
                'UTF-8 (é+CJK+emoji) must round-trip byte-clean through the boundary');
        end

    end

    % ===================== instance helpers ============================ %
    methods (Access = private)
        function b = loadFixture(~, base)
            here = fileparts(mfilename('fullpath'));   % tests\opc
            p = fullfile(here, 'data', [base '.bin']);
            f = fopen(p, 'r', 'n');
            assert(f >= 0, 'byte fixture missing: %s', p);
            b = fread(f, Inf, '*uint8')';
            fclose(f);
        end

        function r = buildRels(testCase, base)
            % Rebuild each frozen s0004 .rels via the IDENTICAL call sequence as
            % scenarios\s0004_rels_l1.m (the Gate-3 twin).
            HYP = testCase.HYP;  IMG = testCase.IMG;
            switch base
                case 'rels1'   % two external, monotonic rId1/rId2 (474 B)
                    r = mat2doc.opc.Relationships("/word");
                    r.get_or_add_ext_rel(HYP, "http://b.example/");   % rId1
                    r.get_or_add_ext_rel(HYP, "http://a.example/");   % rId2
                    r.get_or_add_ext_rel(HYP, "http://b.example/");   % dedup -> rId1
                case 'rels2'   % insertion order rId3 before rId1, gap -> rId2 (629 B)
                    r = mat2doc.opc.Relationships("/word");
                    r.add_relationship(HYP, "http://x.example/", "rId3", true);
                    r.add_relationship(HYP, "http://y.example/", "rId1", true);
                    r.get_or_add_ext_rel(IMG, "http://z.example/");   % gap -> rId2
                case 'rels3'   % two internal same-partname diff-obj + 1 ext (579 B)
                    pA = FakePart_p1_5(mat2doc.opc.PackURI("/word/media/image1.png"));
                    pB = FakePart_p1_5(mat2doc.opc.PackURI("/word/media/image1.png"));
                    r = mat2doc.opc.Relationships("/word");
                    r.get_or_add(IMG, pA);                            % rId1 internal
                    r.get_or_add(IMG, pA);                            % identity -> rId1
                    r.get_or_add(IMG, pB);                            % diff obj -> rId2
                    r.get_or_add_ext_rel(HYP, "http://q.example/");   % rId3 external
                otherwise
                    error('Test_p1_5_opc_packuri:unknownRels', ...
                        'unknown rels base %s', base);
            end
        end

        function bytes = buildDzip(testCase)
            % Build the 3-entry D-zip package (the Gate-3 lane-3 package):
            % [Content_Types].xml, _rels/.rels, and a word/document.xml carrying
            % the é+CJK+emoji UTF-8 blob + all 256 byte values. In-memory ([] path).
            w = mat2doc.opc.ZipPkgWriter_([]);
            w.write(mat2doc.opc.PackURI("/[Content_Types].xml"), testCase.ctBlob());
            w.write(mat2doc.opc.PackURI("/_rels/.rels"), testCase.relsBlob());
            w.write(mat2doc.opc.PackURI("/word/document.xml"), testCase.docBlob());
            w.close();
            bytes = uint8(w.to_bytes());  bytes = bytes(:)';
        end

        function b = ctBlob(~)
            b = uint8(unicode2native( ...
                '<?xml version="1.0"?><Types/>', 'UTF-8'));
        end

        function b = relsBlob(~)
            b = uint8(unicode2native( ...
                '<?xml version="1.0"?><Relationships/>', 'UTF-8'));
        end

        function b = docBlob(~)
            % é + CJK + emoji (surrogate pair U+1F642) as UTF-8, then all 256 byte
            % values -- exercises non-ASCII text AND the full binary byte range.
            txt = uint8(unicode2native( ...
                char("é 文字 " + char(55357) + char(56898)), 'UTF-8'));
            b = [txt, uint8(0:255)];
        end
    end
end

% ===================== file-local helpers ============================== %

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

function ids = idOrder(xmltext)
    % Extract the <Relationship Id="..."> values in document order off the bytes.
    tok = regexp(xmltext, '<Relationship\s+Id="([^"]*)"', 'tokens');
    ids = strings(1, numel(tok));
    for k = 1:numel(tok), ids(k) = string(tok{k}{1}); end
end

function [times, dates] = centralDirStamps(bytes)
    % Walk the zip central directory (anchored at the EOCD record) and return the
    % DOS mod-time and mod-date (uint16) of every central-directory header. Robust
    % against false signature matches in compressed payload (only the central dir,
    % which follows all local data, is scanned).
    b = double(bytes(:)');
    n = numel(b);
    EOCD = [80 75 5 6];        % 0x06054b50 little-endian byte order
    CDH  = [80 75 1 2];        % 0x02014b50 central-directory file header
    eocdPos = 0;
    for i = n-21:-1:1          % EOCD is >= 22 bytes; scan back from the tail
        if i >= 1 && isequal(b(i:i+3), EOCD)
            eocdPos = i;
            break
        end
    end
    assert(eocdPos > 0, 'zip: EOCD record not found');
    total = b(eocdPos+10) + 256*b(eocdPos+11);        % total entries (this disk)
    cdOff = b(eocdPos+16) + 256*b(eocdPos+17) + ...
            65536*b(eocdPos+18) + 16777216*b(eocdPos+19);   % central-dir offset
    times = zeros(1, total);
    dates = zeros(1, total);
    p = cdOff + 1;             % 1-based
    for e = 1:total
        assert(isequal(b(p:p+3), CDH), 'zip: central-dir header %d not found', e);
        times(e) = b(p+12) + 256*b(p+13);   % DOS mod time
        dates(e) = b(p+14) + 256*b(p+15);   % DOS mod date
        nameLen  = b(p+28) + 256*b(p+29);
        extraLen = b(p+30) + 256*b(p+31);
        commLen  = b(p+32) + 256*b(p+33);
        p = p + 46 + nameLen + extraLen + commLen;
    end
end

function h = sha256hex(bytes)
    % Lowercase hex SHA-256 of a uint8 row vector (base MATLAB / Java
    % MessageDigest). typecast(uint8->int8) yields Java's signed byte[] view.
    md  = java.security.MessageDigest.getInstance('SHA-256');
    dig = md.digest(typecast(uint8(bytes(:)'), 'int8'));
    b   = typecast(dig, 'uint8');
    h   = lower(string(sprintf('%02x', b)));
end

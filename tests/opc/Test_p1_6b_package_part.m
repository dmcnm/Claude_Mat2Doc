classdef Test_p1_6b_package_part < matlab.unittest.TestCase
% TEST_P1_6B_PACKAGE_PART  Gate-4 permanent unit tests for Mat2Doc P1-6b
%   (opc package + part + PartFactory) -- the M1-defining open->save layer.
%
%   Freezes the P1-6b guarantees ported in +mat2doc\+opc\ from python-docx v1.2.0
%   src/docx/opc/package.py (OpcPackage/Unmarshaller) + src/docx/opc/part.py
%   (Part/XmlPart/PartFactory) + the docx/__init__.py 37-51 registration block.
%   This is the M1 byte-critical #3 (and headline) WP: it joins the P1-6a reader to
%   the P1-6a writer, so mat2doc.opc.OpcPackage.open(default.docx).save(tmp) is the
%   first time the full 17-part round-trip runs end-to-end through the real logical
%   model.
%
%   Provenance (Gate-1..3, all 2026-07-26):
%     * Audit    : validation\mat2doc\audit_P1-6b_package_part.md
%                  (Porter Gate-1 + Opus Gate-2 adversarial APPROVE; 63 PASS /
%                  1 FAIL, the FAIL being Finding 1 ruled M1-UNREACHABLE, task #60)
%     * Validate : validation\mat2doc\validate_P1-6b_package_part.md (Gate-3 PASS --
%                  the 17-part M1 open->save sweep is GREEN, three-way:
%                  MATLAB == python-docx OpcPackage.open->save == the frozen s0001
%                  Document().save() reference; 16 XML L1 + thumbnail bin; 0 FAIL;
%                  0 new D-numbers; regression 286/286)
%     * Scenario : validation\mat2doc\scenarios\s0008_m1_opensave_sweep.{py,m}
%                  (identical call sequence: OpcPackage.open(tpl).save(out))
%     * Frozen refs (python-docx 1.2.0 oracle, frozen ONCE):
%         references\s0001\  -- the frozen M1 acceptance set (Document().save(),
%           frozen 2026-07-25). Its manifest.json (17-part name|size|sha256, in the
%           frozen M1 zip-entry order + content-types + all three .rels) is embedded
%           below as the M1_MANIFEST Constant so this suite is self-contained.
%         references\s0008\  -- a fresh OpcPackage.open(default).save() oracle;
%           pkgcompare(s0008, s0001) is OVERALL PASS 17/17, so the headline
%           three-way reduces to the single frozen s0001 target embedded here.
%       The 4 regenerated container parts + the reserialized core.xml are ALSO
%       copied byte-for-byte into tests\opc\data\ (content_types.bin 1738 B,
%       rels_main.bin 734 B, rels_document.bin 1227 B, rels_customxml.bin 295 B --
%       all pre-existing P1-4 fixtures, SHA-verified == the s0001 manifest; plus
%       s0001_core.bin 721 B copied from references\s0001\parts\docProps\core.xml)
%       so the L1 container/core pins carry literal byte-diff diagnostics.
%     * Template : +mat2doc\templates\default.docx (sha256 2094b5bd..40d35d), the
%                  real open->save target -- ships in the toolbox, so the sweep is
%                  self-contained relative to the worktree.
%
%   Coverage taxonomy
%   -----------------
%   * Regression (THE 17-PART M1 OPEN->SAVE SWEEP, the headline byte pin) --
%     OpcPackage.open(default.docx).save(tmp), unzip, and pin (a) the 17-part
%     inventory + EXACT frozen M1 zip-entry order (the manifest `#` sequence), and
%     (b) every part's size + SHA-256 identical to the frozen s0001 manifest.
%     SHA-256 equality IS byte-identity (collision-resistant), so these are L1
%     byte-level assertions -- the Gate-3 ladder demanded L1 for all 16 XML parts
%     and bin for the thumbnail (validate_P1-6b lane 1). Goes RED if ANY part's
%     bytes OR the zip-entry order drift. The 4 regenerated container parts and the
%     reserialized core.xml carry an ADDITIONAL literal byte-identical pin (with
%     first-diff-offset diagnostics) against the co-located data\ fixtures.
%   * Regression (L0 structural) -- [Content_Types].xml 3 Default + 11 Override;
%     _rels/.rels 4 rows in DOCUMENT order rId3,rId4,rId1,rId2 (NOT rId-sorted, H11).
%   * Regression (idempotency / D-zip-time) -- second save == first save whole-zip
%     byte-identical; reopen(candidate)->save == candidate whole-zip byte-identical.
%   * Equivalence / VERIFY-1b -- iter_parts returns a NON-cell heterogeneous
%     mat2doc.opc.Part array (13 reachable parts); XmlPart bucket
%     {core,document,styles,settings,numbering}, base-Part bucket = the other 8.
%   * Regression (H10 dispatch matrix) -- 8 registered content types -> XmlPart,
%     8 unregistered M1 types -> base Part; selector (PNG,IMAGE)->Part /
%     (STYLES,OFFICE_DOCUMENT)->""; create(WML_DOCUMENT_MAIN)->XmlPart /
%     create(WML_WEB_SETTINGS)->Part.
%   * H5 identity -- synthetic diamond (pkg->A, pkg->B, A->C, B->C): iter_parts
%     yields C once, iter_rels yields BOTH A->C and B->C, target_part is the SAME
%     handle C from both.
%   * H9 currency -- XmlPart element parsed-once (same handle twice), blob uint8 +
%     re-serialized-on-demand stable + single-quote <?xml declaration; base
%     Part.blob verbatim; blob None -> empty uint8 (H4).
%   * H2 bytes currency -- every reachable part's blob is uint8 (no char leak at
%     the writer boundary).
%
%   Deviations exercised (adopt-only, ZERO new D-numbers -- Gate-3): D-001 (own
%   OOXML parser, via mat2doc.oxml.parse_xml in XmlPart.load), D-serializer-nsdecl
%   (lxml-convention serializer, via mat2doc.opc.oxml.serialize_part_xml in
%   XmlPart.blob), and D-zip-time (the MATLAB<->MATLAB whole-zip byte-stability of
%   save; pkgcompare/this suite compare part bytes after unzip, never the zip
%   envelope). The 17/17 L1 result proves ZERO output-visible divergence.
%
%   Finding 1 (task #60) NOT exercised: Part.drop_rel on an XmlPart hits the
%   M1-unreachable XmlElement.xpath gap (deferred to the P2 DocumentPart un-stub).
%   This suite never calls drop_rel on an XmlPart. drop_rel is not exercised at all.
%
%   Determinism: no network, no hard-coded absolute paths -- the template and byte
%   fixtures resolve relative to this file via fileparts(mfilename('fullpath')).
%   Written artifacts are tempname .docx, deleted via onCleanup; every file write is
%   the writer's BINARY mode ("wb") and every read here is binary ('r','n') -- no
%   CRLF translation, no 'wt'.

    properties (Constant)
        % The frozen s0001 M1 manifest (python-docx Document().save(), frozen
        % 2026-07-25), embedded so the sweep is self-contained. Column 1 is the
        % part name in the FROZEN M1 ZIP-ENTRY ORDER (== references\s0001\
        % manifest.json parts[] order == validate_P1-6b lane-1 `#` column); column
        % 2 is the exact byte size; column 3 is the lowercase SHA-256. SHA-256
        % equality == byte-identity, so this is the L1 pin for all 16 XML parts and
        % the bin pin for docProps/thumbnail.jpeg.
        M1_MANIFEST = [ ...
            "[Content_Types].xml",             "1738",   "66c84fb7a6aa3c4ead49f895e4a7044df1fb57de1ed76d09b2686e91f5bed5b4"; ...
            "_rels/.rels",                     "734",    "a9ae57efe9186f07d48303bcac5d54c7e359bf3f503939c03ad2954e5c59a5c4"; ...
            "docProps/core.xml",               "721",    "d14be8284e406d14dc2a576b19f15e7f637b955f9e27a3d9ed871ec99606f7ba"; ...
            "docProps/app.xml",                "1132",   "be664981c3141cddfc59362beb287ebf20d0773660e2dd6faac5968a5930a081"; ...
            "word/document.xml",               "1548",   "0e4dd503bc095219cd24e8312b372c588ca56bb834aacac53eafb7bf47836327"; ...
            "word/_rels/document.xml.rels",    "1227",   "1e7f0eb144a98e199249314f61ff32a8de2a27e56d8e9ee6b524e1c6b235d377"; ...
            "word/styles.xml",                 "349458", "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"; ...
            "word/stylesWithEffects.xml",      "438131", "463ae0928cf0d84775dbf8cf18d6c3029f6707c81bf590f6d6dd8757a5e93f15"; ...
            "word/settings.xml",               "2535",   "51a0d348fe85965c66e4748a03c3c0d055d78455514f03cb121c334af7d73689"; ...
            "word/webSettings.xml",            "438",    "349d36de7434d09f86987ff671d8814964a0588c1e630c06e562cda7e75e9f95"; ...
            "word/fontTable.xml",              "2811",   "79385fb7f60247507ecaffc292e9ebd52ea0657b8634f629ba6fccc54011d6bb"; ...
            "word/theme/theme1.xml",           "10939",  "e3a8ab7db9ca7afca56f5f2820a56e8b660016c647773555b060b0a02ac76941"; ...
            "customXml/item1.xml",             "262",    "a86086ffc5d8e83ebd6c71a55d1d2efaa31b137977f5f3a752366e1023612144"; ...
            "customXml/_rels/item1.xml.rels",  "295",    "1ca6c9a64edcebe24ee703a54403611b322d96da33371779e742d2d3f7ed7a6c"; ...
            "customXml/itemProps1.xml",        "354",    "c542307b13ec29a8b546217bb37936ab4822e044b265d2952985ec3d6afed24e"; ...
            "word/numbering.xml",              "5513",   "70976f19cbcd896e51890859fe6ecb3467a5a7ad0c040160fedc5a1993cb09ce"; ...
            "docProps/thumbnail.jpeg",         "8324",   "96367138dc44ce09bf2c8f0f8e49348a1478d2c5c0af69bbc2bbc38b63cdcead"];

        % The 13 M1-reachable parts (iter_parts subset) whose logical part class is
        % XmlPart (registered content types present in default.docx). The other 8
        % reachable parts fall to the base verbatim Part; the remaining 4 of the 17
        % ([Content_Types].xml, the three .rels) are writer artefacts, not graph
        % parts. See validate_P1-6b lane 2.
        XMLPART_PARTNAMES = [ ...
            "/docProps/core.xml", ...
            "/word/document.xml", ...
            "/word/styles.xml", ...
            "/word/settings.xml", ...
            "/word/numbering.xml"];
    end

    properties (Access = private)
        % Lazy cache of one open->save candidate (the sweep is rebuilt once; styles
        % 349 KB + stylesWithEffects 438 KB make re-serialization non-trivial).
        candBuilt_ (1,1) logical = false
        candNames_
        candBlobs_
        candZip_
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied from tests\opc\Test_p1_6a_pkgrw.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\opc
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
            testCase.applyFixture(PathFixture(here));  % StubPart helpers, if used

            % Build the M1 open->save candidate ONCE, in-line AFTER the PathFixtures
            % are active (so +mat2doc resolves; multiple TestClassSetup methods have
            % no guaranteed order, hence the single method). The framework copies
            % this post-setup instance into each Test method, so this single
            % OpcPackage.open(default.docx).save() is shared by every case that
            % inspects the 17 parts. A port defect here fails the whole class loudly.
            tpl = char(testCase.templatePath());
            pkg = mat2doc.opc.OpcPackage.open(tpl);
            zipBytes = testCase.saveToTemp(pkg);
            [b, n] = zipEntryList(zipBytes);
            testCase.candZip_   = zipBytes;
            testCase.candBlobs_ = b;
            testCase.candNames_ = n;
            testCase.candBuilt_ = true;
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. THE 17-PART M1 OPEN->SAVE SWEEP (the headline byte pin)       %
        % =============================================================== %

        function test_m1_sweep_inventory_and_zip_order(testCase)
            % Regression (M1 L0, the milestone-defining order pin): OpcPackage.open
            % (default.docx).save() emits EXACTLY the 17 frozen parts, in EXACTLY
            % the frozen M1 zip-entry order (the manifest `#` sequence). Goes RED on
            % any missing/extra part or any writer reorder. (validate_P1-6b lane 1.)
            [~, names] = testCase.candidate();
            want = testCase.M1_MANIFEST(:, 1)';
            testCase.verifyEqual(numel(names), 17, ...
                'open->save must emit exactly 17 zip entries');
            testCase.verifyEqual(names, want, ...
                'zip-entry sequence must equal the frozen M1 (s0001) order');
        end

        function test_m1_sweep_all_17_parts_byte_identical(testCase)
            % Regression (THE headline pin, L1/bin): every one of the 17 parts is
            % byte-identical (size + SHA-256) to the frozen s0001 python-docx oracle
            % -- the 16 XML parts at L1 and docProps/thumbnail.jpeg as bin. SHA-256
            % equality is a byte-level (L1) assertion. Goes RED on any single-byte
            % drift in any part. This IS the M1 acceptance sweep passing through the
            % P1-6a writer (validate_P1-6b lane 1; Gate-2 PROBE B).
            [blobs, names] = testCase.candidate();
            for k = 1:size(testCase.M1_MANIFEST, 1)
                nm       = testCase.M1_MANIFEST(k, 1);
                wantSize = str2double(testCase.M1_MANIFEST(k, 2));
                wantSha  = testCase.M1_MANIFEST(k, 3);
                got = testCase.entryBlob(blobs, names, nm);
                testCase.verifyEqual(numel(got), wantSize, ...
                    sprintf('part %s must be exactly %d B', nm, wantSize));
                testCase.verifyEqual(sha256hex(got), wantSha, ...
                    sprintf(['part %s SHA-256 must equal the frozen s0001 oracle ' ...
                    '(byte-identical L1/bin)'], nm));
            end
        end

        function test_m1_sweep_container_parts_literal_byte_identical(testCase)
            % Regression (L1, literal byte-diff diagnostics): the 4 REGENERATED
            % container parts are byte-identical to the co-located data\ fixtures
            % (themselves SHA-verified == the s0001 manifest). This adds first-diff
            % offset diagnostics on top of the SHA pin above, on the highest-value
            % writer-regenerated parts. (validate_P1-6b lane 1; the ladder demanded
            % L1 for these.)
            [blobs, names] = testCase.candidate();
            map = { ...
                "[Content_Types].xml",            'content_types'; ...
                "_rels/.rels",                    'rels_main'; ...
                "word/_rels/document.xml.rels",   'rels_document'; ...
                "customXml/_rels/item1.xml.rels", 'rels_customxml'};
            for k = 1:size(map, 1)
                got  = testCase.entryBlob(blobs, names, map{k, 1});
                want = testCase.loadFixture(map{k, 2});
                verifyByteIdentical(testCase, got, want, char(map{k, 1}));
            end
        end

        function test_m1_sweep_core_reserialized_literal_byte_identical(testCase)
            % Regression (L1, literal): the RESERIALIZED core.xml (routed through
            % XmlPart parse+serialize -- the whitespace-collapse path) is
            % byte-identical to the frozen s0001 reference. Literal byte-diff on the
            % reserialize path (contrast the passthrough parts).
            [blobs, names] = testCase.candidate();
            got  = testCase.entryBlob(blobs, names, "docProps/core.xml");
            want = testCase.loadFixture('s0001_core');
            verifyByteIdentical(testCase, got, want, 'docProps/core.xml');
        end

        function test_m1_content_types_3default_11override(testCase)
            % Regression (L0 structural, s0001): the regenerated [Content_Types].xml
            % declares exactly 3 <Default> and 11 <Override> entries (the M1
            % content-type map). Structural count over the pinned bytes -- RED on
            % any Default/Override add/drop. (validate_P1-6b lane 1 L0.)
            [blobs, names] = testCase.candidate();
            ct = char(testCase.entryBlob(blobs, names, "[Content_Types].xml"));
            nDefault  = numel(regexp(ct, '<Default\s',  'start'));
            nOverride = numel(regexp(ct, '<Override\s', 'start'));
            testCase.verifyEqual(nDefault, 3,  '[Content_Types].xml must have 3 Defaults');
            testCase.verifyEqual(nOverride, 11, '[Content_Types].xml must have 11 Overrides');
        end

        function test_m1_pkg_rels_doc_order_not_rid_sorted(testCase)
            % Regression (L0, H11, s0001): the package _rels/.rels lists its 4 rows
            % in DOCUMENT order rId3,rId4,rId1,rId2 -- NOT rId-sorted. Pins the
            % insertion-order rel emission the reader+writer preserve.
            [blobs, names] = testCase.candidate();
            rels = char(testCase.entryBlob(blobs, names, "_rels/.rels"));
            ids = regexp(rels, 'Id="(rId\d+)"', 'tokens');
            ids = string(cellfun(@(c) c{1}, ids, 'UniformOutput', false));
            testCase.verifyEqual(ids, ["rId3" "rId4" "rId1" "rId2"], ...
                '_rels/.rels rIds must be in document order (not sorted)');
        end

        % =============================================================== %
        % 2. Idempotency / D-zip-time whole-zip stability                 %
        % =============================================================== %

        function test_idempotent_second_save_equals_first(testCase)
            % Regression (D-zip-time): saving the SAME opened package twice yields
            % two whole-zip byte-identical files (fixed 1980 entry timestamps, P1-5
            % ZipPkgWriter_; the logical graph is not mutated by save).
            tpl = char(testCase.templatePath());
            pkg = mat2doc.opc.OpcPackage.open(tpl);
            a = testCase.saveToTemp(pkg);
            b = testCase.saveToTemp(pkg);
            verifyByteIdentical(testCase, a, b, 'second save == first save (whole zip)');
        end

        function test_reopen_candidate_save_equals_candidate(testCase)
            % Regression (D-zip-time, MATLAB<->MATLAB byte-stable): reopening a
            % saved candidate and re-saving reproduces it whole-zip byte-for-byte.
            tpl = char(testCase.templatePath());
            a = testCase.saveToTemp(mat2doc.opc.OpcPackage.open(tpl));
            tmpA = [tempname '.docx'];
            cA = onCleanup(@() deleteIfExists(tmpA)); %#ok<NASGU>
            writeBytes(tmpA, a);
            c = testCase.saveToTemp(mat2doc.opc.OpcPackage.open(tmpA));
            verifyByteIdentical(testCase, c, a, 'reopen->save == candidate (whole zip)');
        end

        % =============================================================== %
        % 3. iter_parts object-array (VERIFY-1b)                          %
        % =============================================================== %

        function test_iter_parts_noncell_heterogeneous_13(testCase)
            % Equivalence / VERIFY-1b (s0008 lane 2): iter_parts returns a NON-cell
            % heterogeneous mat2doc.opc.Part array (the root type over the Part /
            % XmlPart mix), 13 reachable parts, that the P1-6a PackageWriter indexes.
            pkg = mat2doc.opc.OpcPackage.open(char(testCase.templatePath()));
            parts = pkg.iter_parts();
            testCase.verifyFalse(iscell(parts), 'iter_parts must NOT be a cell');
            testCase.verifyEqual(class(parts), 'mat2doc.opc.Part', ...
                'iter_parts must be the heterogeneous mat2doc.opc.Part root array');
            testCase.verifyEqual(numel(parts), 13, ...
                'default.docx has 13 reachable parts');
        end

        function test_iter_parts_xmlpart_vs_part_buckets(testCase)
            % Equivalence (s0008 lane 2 / lane 3 split): exactly the 5 registered
            % content parts {core,document,styles,settings,numbering} are XmlPart
            % (reserialize bucket); the other 8 reachable parts are the base Part
            % (verbatim bucket). Read off the live object graph.
            pkg = mat2doc.opc.OpcPackage.open(char(testCase.templatePath()));
            parts = pkg.iter_parts();
            gotXml = strings(1, 0);
            for k = 1:numel(parts)
                if isa(parts(k), 'mat2doc.opc.XmlPart')
                    gotXml(end + 1) = string(parts(k).partname()); %#ok<AGROW>
                end
            end
            testCase.verifyEqual(sort(gotXml), sort(testCase.XMLPART_PARTNAMES), ...
                'exactly {core,document,styles,settings,numbering} must be XmlPart');
            testCase.verifyEqual(numel(parts) - numel(gotXml), 8, ...
                'the other 8 reachable parts must be base Part (verbatim bucket)');
        end

        % =============================================================== %
        % 4. PartFactory dispatch matrix (H10, Finding-1 split)           %
        % =============================================================== %

        function test_partfactory_8_registered_to_xmlpart(testCase)
            % Regression (H10, docx/__init__.py 44-51): all 8 registered content
            % types map to mat2doc.opc.XmlPart (row-for-row, count 8=8).
            CT = mat2doc.opc.CONTENT_TYPE;
            reg = [CT.OPC_CORE_PROPERTIES, CT.WML_COMMENTS, CT.WML_DOCUMENT_MAIN, ...
                   CT.WML_FOOTER, CT.WML_HEADER, CT.WML_NUMBERING, ...
                   CT.WML_SETTINGS, CT.WML_STYLES];
            testCase.verifyEqual(numel(reg), 8, 'exactly 8 registered content types');
            for ct = reg
                testCase.verifyEqual( ...
                    mat2doc.opc.PartFactory.part_cls_for_(ct), "mat2doc.opc.XmlPart", ...
                    sprintf('%s must dispatch to XmlPart', ct));
            end
        end

        function test_partfactory_8_unregistered_to_part(testCase)
            % Regression (H10): all 8 unregistered M1 content types fall to the base
            % mat2doc.opc.Part (verbatim passthrough), incl. the raw
            % ms-word.stylesWithEffects string.
            CT = mat2doc.opc.CONTENT_TYPE;
            unreg = [CT.WML_WEB_SETTINGS, CT.WML_FONT_TABLE, CT.OFC_THEME, ...
                     CT.OFC_EXTENDED_PROPERTIES, CT.XML, ...
                     CT.OFC_CUSTOM_XML_PROPERTIES, CT.JPEG, ...
                     "application/vnd.ms-word.stylesWithEffects+xml"];
            testCase.verifyEqual(numel(unreg), 8, 'exactly 8 unregistered M1 types');
            for ct = unreg
                testCase.verifyEqual( ...
                    mat2doc.opc.PartFactory.part_cls_for_(ct), "mat2doc.opc.Part", ...
                    sprintf('%s must fall to base Part', ct));
            end
        end

        function test_partfactory_selector_cases(testCase)
            % Regression (H10, docx/__init__.py 37-40): the part_class_selector maps
            % (PNG, IMAGE) -> "mat2doc.opc.Part" (ImagePart stand-in) and every
            % non-IMAGE reltype -> "" (Python None -> fall to the content-type map).
            CT = mat2doc.opc.CONTENT_TYPE;
            RT = mat2doc.opc.RELATIONSHIP_TYPE;
            testCase.verifyEqual( ...
                mat2doc.opc.PartFactory.part_class_selector_(CT.PNG, RT.IMAGE), ...
                "mat2doc.opc.Part", '(PNG, IMAGE) -> base Part (ImagePart stand-in)');
            testCase.verifyEqual( ...
                mat2doc.opc.PartFactory.part_class_selector_(CT.WML_STYLES, RT.OFFICE_DOCUMENT), ...
                "", '(STYLES, OFFICE_DOCUMENT) -> "" (None -> content-type map)');
        end

        function test_partfactory_create_dispatch(testCase)
            % Regression (H10, create() end-to-end): create(WML_DOCUMENT_MAIN) ->
            % XmlPart (parsed); create(WML_WEB_SETTINGS) -> base Part (verbatim).
            % Both dispatch through cls_method_fn(PartClass,"load").
            CT = mat2doc.opc.CONTENT_TYPE;
            RT = mat2doc.opc.RELATIONSHIP_TYPE;
            docBlob = uint8(['<w:document xmlns:w="http://schemas.openxmlformats.org' ...
                '/wordprocessingml/2006/main"/>']);
            xp = mat2doc.opc.PartFactory.create( ...
                mat2doc.opc.PackURI("/word/document.xml"), CT.WML_DOCUMENT_MAIN, ...
                RT.OFFICE_DOCUMENT, docBlob, []);
            testCase.verifyEqual(class(xp), 'mat2doc.opc.XmlPart', ...
                'create(WML_DOCUMENT_MAIN) must return XmlPart');
            bp = mat2doc.opc.PartFactory.create( ...
                mat2doc.opc.PackURI("/word/webSettings.xml"), CT.WML_WEB_SETTINGS, ...
                RT.WEB_SETTINGS, uint8('<w:webSettings/>'), []);
            testCase.verifyEqual(class(bp), 'mat2doc.opc.Part', ...
                'create(WML_WEB_SETTINGS) must return base Part');
        end

        % =============================================================== %
        % 5. H5 identity (synthetic diamond) + H9 XmlPart currency        %
        % =============================================================== %

        function test_h5_diamond_iter_parts_dedup(testCase)
            % H5 (element identity): a diamond pkg->A, pkg->B, A->C, B->C (C shared)
            % -- iter_parts yields exactly 3 parts with C ONCE (the visited set gates
            % recursion, dedup by handle identity via the Sealed eq).
            [~, A, B, C] = testCase.buildDiamond();
            pkg = testCase.diamondPkg_;
            parts = pkg.iter_parts();
            testCase.verifyEqual(numel(parts), 3, 'diamond iter_parts yields 3 parts');
            names = arrayfun(@(p) string(p.partname()), parts);
            testCase.verifyEqual(numel(unique(names)), 3, 'C must be yielded once');
            % A and B present; C present exactly once.
            testCase.verifyTrue(any(parts == A) && any(parts == B) && any(parts == C), ...
                'A, B and C must all be yielded');
            testCase.verifyEqual(sum(parts == C), 1, 'C must appear exactly once');
        end

        function test_h5_diamond_iter_rels_both_and_identity(testCase)
            % H5: iter_rels yields BOTH A->C and B->C (4 rels total -- the visited
            % set gates recursion only, not the yield), and target_part is the SAME
            % handle C from both A.rels and B.rels.
            [~, A, B, C] = testCase.buildDiamond();
            pkg = testCase.diamondPkg_;
            rels = pkg.iter_rels();
            testCase.verifyEqual(numel(rels), 4, ...
                'iter_rels yields all 4 rels (both A->C and B->C)');
            tA = A.rels().values();
            tB = B.rels().values();
            testCase.verifyTrue(tA(1).target_part == C, 'A->C target is C');
            testCase.verifyTrue(tB(1).target_part == C, 'B->C target is C');
            testCase.verifyTrue(tA(1).target_part == tB(1).target_part, ...
                'both rels resolve to the SAME handle C (H5 identity)');
        end

        function test_h9_xmlpart_element_parsed_once(testCase)
            % H9 (currency): the opened document.xml part is an XmlPart whose
            % element() returns the SAME handle across calls (parsed once at load,
            % never re-parsed).
            xp = testCase.documentPart();
            testCase.verifyEqual(class(xp), 'mat2doc.opc.XmlPart', ...
                'document.xml part must be an XmlPart');
            e1 = xp.element();
            e2 = xp.element();
            testCase.verifyTrue(e1 == e2, 'element() must be the same handle (parsed once)');
        end

        function test_h9_xmlpart_blob_uint8_stable_decl(testCase)
            % H9: XmlPart.blob is uint8, re-serialized on demand yet STABLE across
            % calls, and starts with the single-quote lxml <?xml declaration
            % (D-serializer-nsdecl). The reserialize path underlying the M1 collapse.
            xp = testCase.documentPart();
            b1 = xp.blob();
            b2 = xp.blob();
            testCase.verifyClass(b1, 'uint8', 'XmlPart.blob must be uint8');
            testCase.verifyEqual(b1, b2, 'XmlPart.blob must be stable across calls');
            testCase.verifyTrue(startsWith(char(b1), "<?xml"), ...
                'XmlPart.blob must begin with the <?xml declaration');
        end

        function test_h9_base_part_blob_verbatim(testCase)
            % H9: a base Part returns its blob VERBATIM -- the opened
            % stylesWithEffects.xml part (unregistered -> base Part) is byte-for-byte
            % the frozen s0001 reference (passthrough, no re-serialization).
            pkg = mat2doc.opc.OpcPackage.open(char(testCase.templatePath()));
            p = testCase.partByName(pkg, "/word/stylesWithEffects.xml");
            testCase.verifyEqual(class(p), 'mat2doc.opc.Part', ...
                'stylesWithEffects.xml must be a base Part (verbatim)');
            b = p.blob();
            testCase.verifyEqual(numel(b), 438131, 'stylesWithEffects verbatim size');
            testCase.verifyEqual(sha256hex(b), ...
                "463ae0928cf0d84775dbf8cf18d6c3029f6707c81bf590f6d6dd8757a5e93f15", ...
                'base Part.blob must be the s0001 reference bytes VERBATIM');
        end

        function test_h4_base_part_blob_none_empty(testCase)
            % H4 (truthiness): a base Part built with blob = [] (Python None) returns
            % EMPTY uint8 (Python b""), not [] and not an error.
            p = mat2doc.opc.Part(mat2doc.opc.PackURI("/x.bin"), ...
                mat2doc.opc.CONTENT_TYPE.XML, [], []);
            b = p.blob();
            testCase.verifyClass(b, 'uint8', 'blob None -> uint8');
            testCase.verifyEmpty(b, 'blob None -> empty uint8 (b"")');
        end

        function test_h2_reachable_blobs_uint8(testCase)
            % H2 (bytes currency): every reachable part's blob is uint8 -- no char
            % leak at the writer boundary (a leak would latin-1-corrupt non-ASCII and
            % break the L1 pins).
            pkg = mat2doc.opc.OpcPackage.open(char(testCase.templatePath()));
            parts = pkg.iter_parts();
            for k = 1:numel(parts)
                testCase.verifyClass(parts(k).blob(), 'uint8', ...
                    sprintf('reachable part %d blob must be uint8', k));
            end
        end

    end

    % ===================== instance helpers ============================ %
    properties (Access = private)
        diamondPkg_        % synthetic H5 diamond package (built lazily)
        diamondBuilt_ (1,1) logical = false
        diamondA_
        diamondB_
        diamondC_
    end

    methods (Access = private)

        function [blobs, names] = candidate(testCase)
            % Return the cached open->save candidate (built once in TestClassSetup):
            % the zip entries of OpcPackage.open(default.docx).save(), in stream
            % (write) order.
            assert(testCase.candBuilt_, 'candidate not built in TestClassSetup');
            blobs = testCase.candBlobs_;
            names = testCase.candNames_;
        end

        function bytes = saveToTemp(~, pkg)
            % pkg.save to a BINARY-mode temp .docx (the writer opens "wb"), read the
            % bytes back, delete. Returns the whole-zip uint8 vector.
            tmp = [tempname '.docx'];
            cleanup = onCleanup(@() deleteIfExists(tmp)); %#ok<NASGU>
            pkg.save(tmp);
            bytes = readBytes(tmp);
        end

        function xp = documentPart(testCase)
            pkg = mat2doc.opc.OpcPackage.open(char(testCase.templatePath()));
            xp = testCase.partByName(pkg, "/word/document.xml");
        end

        function p = partByName(~, pkg, name)
            parts = pkg.iter_parts();
            p = [];
            for k = 1:numel(parts)
                if string(parts(k).partname()) == string(name)
                    p = parts(k);
                    return
                end
            end
            assert(~isempty(p), 'part not found in graph: %s', name);
        end

        function [pkg, A, B, C] = buildDiamond(testCase)
            % Synthetic diamond pkg->A, pkg->B, A->C, B->C (C shared). Base Parts;
            % well-known rIds via load_rel (the load-path add). Built once + cached.
            if ~testCase.diamondBuilt_
                CT = mat2doc.opc.CONTENT_TYPE;
                RT = mat2doc.opc.RELATIONSHIP_TYPE;
                P  = @(s) mat2doc.opc.PackURI(s);
                a = mat2doc.opc.Part(P("/a.xml"), CT.XML, uint8('<a/>'), []);
                b = mat2doc.opc.Part(P("/b.xml"), CT.XML, uint8('<b/>'), []);
                c = mat2doc.opc.Part(P("/c.xml"), CT.XML, uint8('<c/>'), []);
                p = mat2doc.opc.OpcPackage();
                p.load_rel(RT.CUSTOM_XML, a, "rId1", false);   % pkg -> A
                p.load_rel(RT.OFFICE_DOCUMENT, b, "rId2", false); % pkg -> B
                a.load_rel(RT.STYLES, c, "rId1", false);       % A -> C
                b.load_rel(RT.STYLES, c, "rId1", false);       % B -> C (shared)
                testCase.diamondPkg_ = p;
                testCase.diamondA_ = a;
                testCase.diamondB_ = b;
                testCase.diamondC_ = c;
                testCase.diamondBuilt_ = true;
            end
            pkg = testCase.diamondPkg_;
            A = testCase.diamondA_;
            B = testCase.diamondB_;
            C = testCase.diamondC_;
        end

        function p = templatePath(~)
            here = fileparts(mfilename('fullpath'));       % tests\opc
            root = fileparts(fileparts(here));             % worktree root
            p = fullfile(root, '+mat2doc', 'templates', 'default.docx');
        end

        function b = loadFixture(~, base)
            here = fileparts(mfilename('fullpath'));       % tests\opc
            p = fullfile(here, 'data', [base '.bin']);
            b = readBytes(p);
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
    % java.util.zip.ZipInputStream reads local file headers in physical order, so
    % `names` is the true zip-entry write sequence. Kept file-local so the order pin
    % is independent of the reader under test. (Copied from Test_p1_6a_pkgrw.m.)
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

function b = readBytes(p)
    f = fopen(p, 'r', 'n');            % binary read (no CRLF translation)
    assert(f >= 0, 'could not open for read: %s', p);
    b = fread(f, Inf, '*uint8')';
    fclose(f);
end

function writeBytes(p, bytes)
    f = fopen(p, 'w', 'n');            % binary write (no 'wt')
    assert(f >= 0, 'could not open for write: %s', p);
    fwrite(f, uint8(bytes), 'uint8');
    fclose(f);
end

function deleteIfExists(p)
    if isfile(p)
        delete(p);
    end
end

function verifyByteIdentical(testCase, got, want, label)
    % Byte-level (L1) assertion. On mismatch report sizes and first diff offset.
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

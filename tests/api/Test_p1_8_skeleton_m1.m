classdef Test_p1_8_skeleton_m1 < matlab.unittest.TestCase
% TEST_P1_8_SKELETON_M1  Gate-4 permanent unit tests for Mat2Doc P1-8
%   (walking-skeleton wiring -> the M1 milestone).
%
%   Freezes the P1-8 guarantees that wire the PUBLIC entry
%   `mat2doc.Document(...)` (docx/api.py) down through
%   mat2doc.package.Package (docx/package.py), mat2doc.parts.DocumentPart
%   (docx/parts/document.py, M1 slice), mat2doc.document.Document
%   (docx/document.py, M1 slice), plus the TWO byte-neutral PartFactory row
%   flips (WML_DOCUMENT_MAIN -> mat2doc.parts.DocumentPart,
%   OPC_CORE_PROPERTIES -> mat2doc.opc.parts.CorePropertiesPart).
%
%   This is the **M1 milestone acceptance class**: it is the FIRST time the full
%   17-part round-trip is driven from the api-level factory `mat2doc.Document()`
%   (Test_p1_6b pins the same 17/17 sweep one layer lower, through
%   OpcPackage.open->save; THIS class pins it through the public entry the user
%   actually calls). RED on ANY byte/order drift is the milestone regression pin.
%
%   Provenance (Gate-1..3, all 2026-07-26):
%     * Audit    : validation\mat2doc\audit_P1-8_skeleton_m1.md
%                  (Porter Gate-1 + Fable Gate-2 cross-model APPROVE; F1 inline
%                  fix in place -- Document.m strict 0x0-double None sentinel).
%     * Validate : validation\mat2doc\validate_P1-8_skeleton_m1.md (Gate-3 PASS
%                  -- ** M1 BYTE LEG GREEN **; three-way MATLAB
%                  `mat2doc.Document().save` == python-docx `docx.Document().save`
%                  == frozen s0001: L0 PASS + 16 XML L1 + thumbnail bin = 17/17,
%                  OVERALL PASS; F1 tri-state, content-type guard, class-identity,
%                  stub-safety and row-flip byte-neutrality all PASS; 0 new
%                  D-numbers).
%     * Scenario : validation\mat2doc\scenarios\s0012_m1_document_entry.{py,m}
%                  (identical call sequence: the api.py `Document(...)` entry).
%     * Frozen refs (python-docx 1.2.0 oracle, frozen ONCE):
%         references\s0001\ -- the frozen 17-part M1 acceptance set
%           (Document().save(), frozen 2026-07-25). Its 17-part
%           name|size|sha256 manifest (frozen zip-entry order + content-types +
%           the three .rels) is embedded below as M1_MANIFEST so this suite is
%           self-contained; SHA-256 equality IS byte-identity (L1).
%         references\s0012\mat2doc_candidate.docx -- the 37141 B MATLAB
%           public-entry candidate (`mat2doc.Document().save` from a foreign cwd;
%           sha 8ec5f26b...), copied byte-for-byte into tests\api\data\
%           s0012_candidate.docx (co-located `* binary` .gitattributes) as the
%           WHOLE-PACKAGE regression pin.
%     * Template : +mat2doc\templates\default.docx -- ships in the toolbox, so the
%           sweep is self-contained relative to the worktree.
%
%   Coverage taxonomy
%   -----------------
%   * Regression (THE M1 ACCEPTANCE SWEEP, headline milestone pin) --
%     `mat2doc.Document()` (no arg -> default template) then `.save(tmp)`, unzip,
%     and pin (a) the 17-part inventory + EXACT frozen zip-entry order,
%     (b) every part's size + SHA-256 == the frozen s0001 oracle (16 XML L1 +
%     thumbnail bin), (c) L0 structure ([Content_Types].xml 3 Default / 11
%     Override; the three .rels 4 / 8 / 1 rows; the package .rels in DOCUMENT
%     order rId3,rId4,rId1,rId2 not rId-sorted), and (d) the WHOLE-package zip ==
%     the frozen 37141 B s0012 candidate. Goes RED on ANY single-byte or order
%     drift through the public entry. (validate_P1-8 Bar 1.)
%   * Regression (F1 H3 public-entry None tri-state) -- `Document()`/`Document([])`
%     open the default template; `Document('')` and `Document("")` raise
%     `mat2doc:PackageNotFoundError` "Package not found at ''" (identifier + message
%     pinned). Locks the Gate-2 F1 fix so `''`/`""` can never again be mistaken for
%     None. (validate_P1-8 Bar 2.)
%   * Regression / Edge (content-type guard, error path) -- a docx whose main-part
%     content-type is flipped off WML_DOCUMENT_MAIN raises `mat2doc:ValueError`
%     with the VERBATIM message `file '%s' is not a Word file, content type is
%     '%s'` (identifier + full message pinned). (validate_P1-8 Bar 3.)
%   * Regression (class-identity, inherited-static trap pin) -- the full opened
%     object graph: Package.open(default) isa mat2doc.package.Package;
%     main_document_part isa mat2doc.parts.DocumentPart; the core-props part isa
%     mat2doc.opc.parts.CorePropertiesPart; `.document` isa
%     mat2doc.document.Document and is a FRESH handle each access. A future
%     base-class regression (a factory silently building the base class) goes RED
%     here. (validate_P1-8 Bar 4.)
%   * Regression / Edge (stub-safety) -- a representative stub set
%     (add_paragraph / paragraphs / styles / numbering_part / image_parts) each
%     raises `mat2doc:notYetPorted`; and a clean `Document().save()` fires ZERO of
%     them (a thrown stub would fail the save). (validate_P1-8 Bar 4 stub-safety.)
%
%   Deviations exercised (adopt-only, ZERO new D-numbers -- Gate-3): D-001 (own
%   OOXML parser/attr-order, via the parse/serialize path -- byte-proven 17/17),
%   D-serializer-nsdecl (unreachable at M1: no created/mutated elements),
%   D-zip-time (container timestamps only; the whole-zip pin is the
%   MATLAB<->MATLAB byte-stable save, itself byte-proven), D-coreprops-time
%   (unreachable: default.docx HAS core.xml so CorePropertiesPart.default()'s
%   wall-clock never runs). The 17/17 L1 result proves ZERO output divergence.
%
%   Determinism: no network, no hard-coded absolute paths -- the template and byte
%   fixture resolve relative to this file via fileparts(mfilename('fullpath')).
%   Written artifacts are tempname .docx, deleted via onCleanup; every file write
%   is binary ("wb"/'w','n') and every read is binary ('r','n') -- no CRLF
%   translation, no 'wt'.

    properties (Constant)
        % The frozen s0001 M1 manifest (python-docx Document().save(), frozen
        % 2026-07-25), embedded so the sweep is self-contained. Column 1 is the
        % part name in the FROZEN zip-entry order (== references\s0001\
        % manifest.json parts[] order); column 2 the exact byte size; column 3 the
        % lowercase SHA-256. SHA-256 equality == byte-identity, so this is the L1
        % pin for all 16 XML parts and the bin pin for docProps/thumbnail.jpeg.
        % Identical to Test_p1_6b_package_part.M1_MANIFEST (same frozen oracle);
        % here it is driven through the api-level mat2doc.Document() entry.
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

        % The frozen s0012 whole-package candidate (mat2doc.Document().save from a
        % foreign cwd), copied to tests\api\data\s0012_candidate.docx.
        SHA_S0012_CANDIDATE = "8ec5f26bc7e0c497f7ccb72cbf3183169da1861d665b7dd7437290c9af25588f"
        SIZE_S0012_CANDIDATE = 37141
    end

    properties (Access = private)
        % Lazy cache of ONE public-entry open->save candidate (built once in
        % TestClassSetup; styles 349 KB + stylesWithEffects 438 KB make
        % re-serialization non-trivial).
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
            % Idiom copied from tests\coreprops\Test_p1_7_coreprops.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\api
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));

            % Build the M1 public-entry candidate ONCE, AFTER the PathFixture is
            % active (so +mat2doc resolves). This single mat2doc.Document().save()
            % is the milestone acceptance artifact shared by every sweep case; a
            % port defect here fails the whole class loudly.
            d = mat2doc.Document();                    % no arg -> default template
            zipBytes = testCase.saveToTemp(d);
            [b, n] = zipEntryList(zipBytes);
            testCase.candZip_   = zipBytes;
            testCase.candBlobs_ = b;
            testCase.candNames_ = n;
            testCase.candBuilt_ = true;
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. THE M1 ACCEPTANCE SWEEP via mat2doc.Document() (milestone)    %
        % =============================================================== %

        function test_m1_public_entry_inventory_and_zip_order(testCase)
            % Regression (M1 L0, the milestone-defining order pin): the PUBLIC
            % entry mat2doc.Document().save() emits EXACTLY the 17 frozen parts, in
            % EXACTLY the frozen zip-entry order. RED on any missing/extra part or
            % writer reorder. (validate_P1-8 Bar 1 L0.)
            [~, names] = testCase.candidate();
            want = testCase.M1_MANIFEST(:, 1)';
            testCase.verifyEqual(numel(names), 17, ...
                'mat2doc.Document().save must emit exactly 17 zip entries');
            testCase.verifyEqual(names, want, ...
                'zip-entry sequence must equal the frozen M1 (s0001) order');
        end

        function test_m1_public_entry_all_17_parts_byte_identical(testCase)
            % Regression (THE headline milestone pin, L1/bin): every one of the 17
            % parts emitted through mat2doc.Document().save() is byte-identical
            % (size + SHA-256) to the frozen s0001 python-docx oracle -- the 16 XML
            % parts at L1 and docProps/thumbnail.jpeg as bin. SHA-256 equality is a
            % byte-level (L1) assertion. RED on any single-byte drift in any part.
            % THIS is the M1 acceptance sweep passing through the api-level entry.
            % (validate_P1-8 Bar 1; three-way MATLAB==python==s0001.)
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

        function test_m1_public_entry_whole_zip_equals_s0012_candidate(testCase)
            % Regression (whole-package pin, D-zip-time byte-stable): the whole zip
            % produced by mat2doc.Document().save() is byte-identical to the frozen
            % 37141 B s0012 candidate (itself a mat2doc.Document().save from a
            % foreign cwd). Pins the complete package envelope -- part bytes AND zip
            % framing (order + 1980 timestamps) -- with first-diff diagnostics.
            % (validate_P1-8 Bar 1; the s0012 milestone artifact for the COM leg.)
            testCase.candidate();                 % assert the candidate is built
            got  = testCase.candZip_;
            want = testCase.loadCandidateFixture();
            testCase.verifyEqual(numel(got), testCase.SIZE_S0012_CANDIDATE, ...
                sprintf('candidate must be exactly %d B', testCase.SIZE_S0012_CANDIDATE));
            testCase.verifyEqual(sha256hex(got), testCase.SHA_S0012_CANDIDATE, ...
                'candidate SHA-256 must equal the frozen s0012 whole-package oracle');
            verifyByteIdentical(testCase, got, want, ...
                'mat2doc.Document().save whole zip == frozen s0012 candidate');
        end

        function test_m1_content_types_3default_11override(testCase)
            % Regression (L0 structural, s0001): the regenerated [Content_Types].xml
            % declares exactly 3 <Default> and 11 <Override> entries (the M1
            % content-type map). RED on any Default/Override add/drop.
            % (validate_P1-8 Bar 1 L0.)
            [blobs, names] = testCase.candidate();
            ct = char(testCase.entryBlob(blobs, names, "[Content_Types].xml"));
            nDefault  = numel(regexp(ct, '<Default\s',  'start'));
            nOverride = numel(regexp(ct, '<Override\s', 'start'));
            testCase.verifyEqual(nDefault, 3,  '[Content_Types].xml must have 3 Defaults');
            testCase.verifyEqual(nOverride, 11, '[Content_Types].xml must have 11 Overrides');
        end

        function test_m1_rels_row_counts_and_pkg_order(testCase)
            % Regression (L0, H11, s0001): the three .rels parts carry 4 / 8 / 1
            % rows respectively, and the package _rels/.rels lists its 4 rows in
            % DOCUMENT order rId3,rId4,rId1,rId2 -- NOT rId-sorted. Pins the
            % insertion-order rel emission the reader+writer preserve.
            % (validate_P1-8 Bar 1 L0.)
            [blobs, names] = testCase.candidate();
            pkgRels = char(testCase.entryBlob(blobs, names, "_rels/.rels"));
            docRels = char(testCase.entryBlob(blobs, names, "word/_rels/document.xml.rels"));
            cxRels  = char(testCase.entryBlob(blobs, names, "customXml/_rels/item1.xml.rels"));
            testCase.verifyEqual(numel(regexp(pkgRels, '<Relationship\s', 'start')), 4, ...
                '_rels/.rels must have 4 rows');
            testCase.verifyEqual(numel(regexp(docRels, '<Relationship\s', 'start')), 8, ...
                'word/_rels/document.xml.rels must have 8 rows');
            testCase.verifyEqual(numel(regexp(cxRels, '<Relationship\s', 'start')), 1, ...
                'customXml/_rels/item1.xml.rels must have 1 row');
            ids = regexp(pkgRels, 'Id="(rId\d+)"', 'tokens');
            ids = string(cellfun(@(c) c{1}, ids, 'UniformOutput', false));
            testCase.verifyEqual(ids, ["rId3" "rId4" "rId1" "rId2"], ...
                '_rels/.rels rIds must be in document order (not sorted)');
        end

        % =============================================================== %
        % 2. F1 H3 public-entry None tri-state (Gate-2 fix regression pin) %
        % =============================================================== %

        function test_f1_noarg_and_none_open_default(testCase)
            % Regression (F1, H3): the None sentinels -- no-arg and the 0x0 double
            % [] -- both open the default template and return a
            % mat2doc.document.Document. (validate_P1-8 Bar 2.)
            d0 = mat2doc.Document();
            testCase.verifyClass(d0, 'mat2doc.document.Document', ...
                'mat2doc.Document() (no arg) must open the default template');
            d1 = mat2doc.Document([]);
            testCase.verifyClass(d1, 'mat2doc.document.Document', ...
                'mat2doc.Document([]) (None sentinel) must open the default template');
        end

        function test_f1_empty_char_and_string_raise_packagenotfound(testCase)
            % Regression (F1, H3, THE Gate-2 fix pin): '' (0x0 char) and "" (empty
            % string) are NOT None -- Python treats them as a real path and
            % Package.open('') raises PackageNotFoundError. The strict 0x0-double
            % None test (Document.m F1 fix) must let them fall through to
            % mat2doc:PackageNotFoundError "Package not found at ''" (identifier +
            % byte-verbatim message pinned; the local-part == the Python exception
            % class name). Pre-fix, these silently opened the template.
            % (validate_P1-8 Bar 2.)
            for arg = {'', ""}
                caught = [];
                try
                    mat2doc.Document(arg{1});
                catch ME
                    caught = ME;
                end
                testCase.assertNotEmpty(caught, ...
                    'empty path must raise (NOT silently open the template)');
                testCase.verifyEqual(caught.identifier, 'mat2doc:PackageNotFoundError', ...
                    'empty path must raise mat2doc:PackageNotFoundError (== Python class)');
                testCase.verifyEqual(caught.message, char("Package not found at ''"), ...
                    'message must be byte-verbatim to Python PackageNotFoundError');
            end
        end

        % =============================================================== %
        % 3. Content-type guard (error path)                              %
        % =============================================================== %

        function test_content_type_guard_valueerror_verbatim(testCase)
            % Regression / Edge (error path, validate_P1-8 Bar 3): a docx whose
            % main-part Override content-type is flipped off WML_DOCUMENT_MAIN (to
            % WML_DOCUMENT_GLOSSARY) makes mat2doc.Document raise mat2doc:ValueError
            % with the VERBATIM message `file '%s' is not a Word file, content type
            % is '%s'` -- identifier + full message (path + content-type) pinned.
            % The tampered docx is derived at runtime from the shipped default.docx
            % (self-contained: no new byte fixture), keeping the exact Gate-3
            % scenario.
            CT = mat2doc.opc.CONTENT_TYPE;
            tampered = testCase.tamperMainContentType();     % temp path, auto-cleaned
            expMsg = char(sprintf("file '%s' is not a Word file, content type is '%s'", ...
                tampered, CT.WML_DOCUMENT_GLOSSARY));
            caught = [];
            try
                mat2doc.Document(tampered);
            catch ME
                caught = ME;
            end
            testCase.assertNotEmpty(caught, 'tampered content-type must raise');
            testCase.verifyEqual(caught.identifier, 'mat2doc:ValueError', ...
                'content-type guard must raise mat2doc:ValueError');
            testCase.verifyEqual(caught.message, expMsg, ...
                'content-type guard message must be byte-verbatim (path + content-type)');
        end

        % =============================================================== %
        % 4. Class identity (inherited-static trap regression pin)        %
        % =============================================================== %

        function test_class_identity_full_graph(testCase)
            % Regression (class-identity, validate_P1-8 Bar 4): the full opened
            % object graph lands on the exact ported classes -- this is the pin
            % against the inherited-static trap (a future base-class regression
            % where the factory silently builds the base class goes RED here).
            %   Package.open(default)          isa mat2doc.package.Package
            %   main_document_part             isa mat2doc.parts.DocumentPart
            %   core-properties part           isa mat2doc.opc.parts.CorePropertiesPart
            %   document_part.document         isa mat2doc.document.Document
            %   document non-cached (fresh handle each access)
            pkg = mat2doc.package.Package.open(char(testCase.templatePath()));
            testCase.verifyEqual(class(pkg), 'mat2doc.package.Package', ...
                'Package.open must build a mat2doc.package.Package (own static)');
            dp = pkg.main_document_part();
            testCase.verifyEqual(class(dp), 'mat2doc.parts.DocumentPart', ...
                'main_document_part must be a mat2doc.parts.DocumentPart (P1-8 flip)');
            cpp = testCase.corePropsPart(pkg);
            testCase.verifyEqual(class(cpp), 'mat2doc.opc.parts.CorePropertiesPart', ...
                'core-properties part must be a CorePropertiesPart (P1-8 flip)');
            d1 = dp.document();
            testCase.verifyEqual(class(d1), 'mat2doc.document.Document', ...
                'document_part.document must be a mat2doc.document.Document');
            d2 = dp.document();
            % P2-1 VERIFY-M1-DOC-BASE retrofit: Document now derives
            % mat2doc.shared.ElementProxy, whose `==` is H5 ELEMENT identity
            % (document.py 28 `class Document(ElementProxy)`; shared.py 289-299).
            % `DocumentPart.document` is a plain @property (docx/parts/document.py
            % 58-59), NOT a lazyproperty -- so each access builds a FRESH proxy,
            % but both wrap the SAME w:document element, hence they compare EQUAL.
            % This is the Python-faithful result (`part.document == part.document`
            % is True in python-docx). The PRE-retrofit assertion here was
            % verifyFalse(d1==d2) (MATLAB instance identity), which is NOT the
            % Python contract; it is corrected to verifyTrue below. Non-caching
            % (distinct instances) is no longer observable through the now-
            % element-identity `==`/`isequal` -- exactly as in Python, where it is
            % only visible via `is`; the fresh-instance construction is pinned by
            % the class contract (DocumentPart.document builds `Document(...)`
            % each call) rather than by handle inequality.
            testCase.verifyTrue(d1 == d2, ...
                ['two document() accesses wrap the SAME element, so they are ' ...
                 'ElementProxy-equal (H5 element identity, Python-correct)']);
        end

        % =============================================================== %
        % 5. Stub-safety                                                  %
        % =============================================================== %

        function test_stub_set_raises_notyetported(testCase)
            % Regression / Edge (stub-safety, validate_P1-8 Bar 4): a representative
            % feature-stub set each raises mat2doc:notYetPorted (identifier pinned),
            % NOT a silent no-op. Covers the still-wired stubs:
            %   document.Document.paragraphs   (still stubbed -- P4-7b VERIFY-1 scope)
            %
            % REGISTRY-FLIP RE-PINS (the un-stub moved the behavior, so the pin moves
            % with it -- registry-flip stale-pins lesson):
            %   * P4-7a: Document.styles was un-stubbed and now RESOLVES (a Styles
            %     proxy over the real styles part) -> positive check below.
            %   * P4-7b: Document.add_paragraph was un-stubbed (the M2 milestone WP)
            %     and now RESOLVES (returns a Paragraph) -> REMOVED from the
            %     notYetPorted set, positive check below. document.Document.paragraphs
            %     STAYS stubbed (P4-7b left it in scope for the next content WP -- its
            %     deps are live but it was not un-stubbed; audit VERIFY-1), so it
            %     remains pinned here.
            %   * P7-4: Package.image_parts was un-stubbed (the picture milestone WP)
            %     and now RESOLVES (returns an ImageParts collection) -> REMOVED from
            %     the notYetPorted set, positive check below (validate_P7-4 s5/re-pin 1).
            %   * P8-1: DocumentPart.numbering_part was un-stubbed. default.docx ships
            %     a numbering part (related via RT.NUMBERING), so it now RESOLVES to a
            %     NumberingPart via part_related_by (the faithful NotImplementedError
            %     new() branch is never reached) -> REMOVED from the notYetPorted set,
            %     positive check below (validate_P8-1 §4 re-pin 3). The full numbering
            %     surface is pinned in tests\oxml\Test_p8_1_numbering.m.
            %   * P8-3 (FINAL port WP): document.Document.paragraphs was un-stubbed
            %     (the last content delegator) and now RESOLVES to a Paragraph array
            %     delegated from the body -> REMOVED from the notYetPorted set, positive
            %     check below (validate_P8-3 s0110). The notYetPorted battery is now
            %     EMPTY: C4 met -- ZERO live mat2doc:notYetPorted sites remain in the
            %     whole toolbox. This test now asserts only positive resolution.
            pkg = mat2doc.package.Package.open(char(testCase.templatePath()));
            dp  = pkg.main_document_part();
            d   = dp.document();

            % Document.paragraphs now RESOLVES (un-stubbed at P8-3, the final port WP)
            % -> a Paragraph array delegated from the body (Python: self._body.paragraphs).
            testCase.verifyClass(d.paragraphs(), 'mat2doc.text.Paragraph', ...
                'Document.paragraphs RESOLVES to a Paragraph array (P8-3 un-stub, C4 met)');

            % Document.styles now RESOLVES (un-stubbed at P4-7a) -> a Styles proxy
            % over the real styles part.
            testCase.verifyClass(d.styles(), 'mat2doc.styles.Styles', ...
                'Document.styles RESOLVES to a Styles over the real styles part (P4-7a)');

            % Document.add_paragraph now RESOLVES (un-stubbed at P4-7b, the M2
            % milestone WP) -> a Paragraph appended to the body.
            testCase.verifyClass(d.add_paragraph(), 'mat2doc.text.Paragraph', ...
                'Document.add_paragraph RESOLVES to a Paragraph (P4-7b un-stub, M2)');

            % Package.image_parts now RESOLVES (un-stubbed at P7-4, the picture
            % milestone WP) -> an ImageParts collection over the package's media parts
            % (registry-flip re-pin 1; validate_P7-4 s5).
            testCase.verifyClass(pkg.image_parts(), 'mat2doc.package.ImageParts', ...
                'Package.image_parts RESOLVES to an ImageParts collection (P7-4 un-stub)');

            % DocumentPart.numbering_part now RESOLVES (un-stubbed at P8-1) -> the
            % package's NumberingPart (default.docx ships numbering.xml, related via
            % RT.NUMBERING). The faithful NotImplementedError new() branch is only
            % reached on a package WITHOUT a numbering part.
            testCase.verifyClass(dp.numbering_part(), 'mat2doc.parts.NumberingPart', ...
                'DocumentPart.numbering_part RESOLVES to a NumberingPart (P8-1 un-stub)');
        end

        function test_clean_save_fires_zero_stubs(testCase)
            % Regression (stub-safety, validate_P1-8 Bar 4): a clean
            % mat2doc.Document().save() completes with ZERO mat2doc:notYetPorted
            % stubs on the open/save path. If any stub were on the path, save would
            % throw; the file being written IS the mechanical no-stub-fired proof.
            d = mat2doc.Document();
            tmp = [tempname '.docx'];
            cleanup = onCleanup(@() deleteIfExists(tmp)); %#ok<NASGU>
            caught = [];
            try
                d.save(tmp);
            catch ME
                caught = ME;
            end
            testCase.verifyEmpty(caught, ...
                'clean Document().save must NOT fire any mat2doc:notYetPorted stub');
            testCase.verifyTrue(isfile(tmp), ...
                'clean Document().save must write the output file');
        end

    end

    % ===================== instance helpers ============================ %
    methods (Access = private)

        function [blobs, names] = candidate(testCase)
            % Return the cached public-entry open->save candidate (built once in
            % TestClassSetup): the zip entries of mat2doc.Document().save(), in
            % stream (write) order.
            assert(testCase.candBuilt_, 'candidate not built in TestClassSetup');
            blobs = testCase.candBlobs_;
            names = testCase.candNames_;
        end

        function bytes = saveToTemp(~, d)
            % d.save to a BINARY-mode temp .docx (the writer opens "wb"), read the
            % bytes back, delete. Returns the whole-zip uint8 vector.
            tmp = [tempname '.docx'];
            cleanup = onCleanup(@() deleteIfExists(tmp)); %#ok<NASGU>
            d.save(tmp);
            bytes = readBytes(tmp);
        end

        function p = corePropsPart(~, pkg)
            % The docProps/core.xml part off the live graph (CT match).
            CT = mat2doc.opc.CONTENT_TYPE;
            parts = pkg.iter_parts();
            p = [];
            for k = 1:numel(parts)
                if string(parts(k).content_type()) == CT.OPC_CORE_PROPERTIES
                    p = parts(k);
                    return
                end
            end
            assert(~isempty(p), 'core-properties part not found in graph');
        end

        function tampered = tamperMainContentType(testCase)
            % Derive a tampered docx from the shipped default.docx: rewrite the
            % [Content_Types].xml Override content-type of the main document part
            % from WML_DOCUMENT_MAIN to WML_DOCUMENT_GLOSSARY, re-zip to a temp
            % file. Self-contained (no new byte fixture); reproduces the Gate-3
            % content-type-guard scenario. The temp file is deleted on class teardown
            % via the returned path's onCleanup captured by addTeardown.
            CT = mat2doc.opc.CONTENT_TYPE;
            src = readBytes(char(testCase.templatePath()));
            tampered = [tempname '.docx'];
            testCase.addTeardown(@() deleteIfExists(tampered));
            rewriteZipContentType(src, tampered, ...
                char(CT.WML_DOCUMENT_MAIN), char(CT.WML_DOCUMENT_GLOSSARY));
        end

        function b = loadCandidateFixture(~)
            here = fileparts(mfilename('fullpath'));   % tests\api
            p = fullfile(here, 'data', 's0012_candidate.docx');
            b = readBytes(p);
        end

        function p = templatePath(~)
            here = fileparts(mfilename('fullpath'));   % tests\api
            root = fileparts(fileparts(here));         % worktree root
            p = fullfile(root, '+mat2doc', 'templates', 'default.docx');
        end

        function blob = entryBlob(~, blobs, names, membername)
            i = find(names == string(membername), 1);
            assert(~isempty(i), 'zip entry not found: %s', membername);
            blob = blobs{i};
        end
    end
end

% ===================== file-local helpers ============================== %

function rewriteZipContentType(srcBytes, outPath, fromCT, toCT)
    % Copy every entry of the in-memory zip `srcBytes` to `outPath`, replacing the
    % substring `fromCT` with `toCT` inside the [Content_Types].xml member only.
    % Reads with java.util.zip.ZipInputStream (physical order) and writes with
    % java.util.zip.ZipOutputStream (standard zip; ZipPkgReader_ reads it back via
    % ZipInputStream). Deterministic; used only to build the content-type-guard
    % fixture from the shipped template.
    bais = java.io.ByteArrayInputStream(int8(typecast(uint8(srcBytes(:)'), 'int8')));
    zis  = java.util.zip.ZipInputStream(bais);
    cin  = onCleanup(@() zis.close()); %#ok<NASGU>
    copier = com.mathworks.mlwidgets.io.InterruptibleStreamCopier.getInterruptibleStreamCopier;
    baosFile = java.io.ByteArrayOutputStream;
    zos = java.util.zip.ZipOutputStream(baosFile);
    while true
        ze = zis.getNextEntry();
        if isempty(ze)
            break
        end
        name = char(ze.getName());
        baos = java.io.ByteArrayOutputStream;
        copier.copyStream(zis, baos);
        blob = typecast(int8(baos.toByteArray()), 'uint8')';
        zis.closeEntry();
        if strcmp(name, '[Content_Types].xml')
            txt = char(blob);
            txt = strrep(txt, fromCT, toCT);
            blob = uint8(txt);
        end
        zos.putNextEntry(java.util.zip.ZipEntry(name));
        zos.write(int8(typecast(uint8(blob(:)'), 'int8')), 0, numel(blob));
        zos.closeEntry();
    end
    zos.close();
    outBytes = typecast(int8(baosFile.toByteArray()), 'uint8')';
    writeBytes(outPath, outBytes);
end

function [blobs, names] = zipEntryList(zipBytes)
    % Enumerate an in-memory zip in stream (write) order -> ordered {names, blobs}.
    % java.util.zip.ZipInputStream reads local file headers in physical order, so
    % `names` is the true zip-entry write sequence. Kept file-local so the order pin
    % is independent of the reader under test. (Copied from Test_p1_6b_package_part.m.)
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

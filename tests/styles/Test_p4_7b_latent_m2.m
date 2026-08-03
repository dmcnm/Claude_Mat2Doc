classdef Test_p4_7b_latent_m2 < matlab.unittest.TestCase
% TEST_P4_7B_LATENT_M2  Gate-4 permanent unit tests for Mat2Doc P4-7b
%   (latent styles + the adder un-stub) -- the ★ M2 MILESTONE work package, the
%   FINAL P4 WP. Ports src/docx/styles/latent.py::{LatentStyles, _LatentStyle} ->
%   +mat2doc\+styles\{LatentStyles, LatentStyle_} (NEW) and un-stubs the adder
%   chain: Styles.latent_styles, Document.add_heading/add_paragraph,
%   BlockItemContainer.add_paragraph/add_paragraph_/paragraphs.
%
%   ============================================================================
%   ★★★  THE M2 MILESTONE  ★★★
%   This class is the M2 milestone acceptance + permanent regression guard. M2 is
%   the FIRST time real body content is authored through the public adder chain:
%       Document();
%       add_heading("Document Title", 0);   -> style "Title"     -> <w:pStyle w:val="Title"/>
%       add_heading("First Section", 1);    -> style "Heading 1" -> <w:pStyle w:val="Heading1"/>
%       add_heading("A Subsection", 2);     -> style "Heading 2" -> <w:pStyle w:val="Heading2"/>
%       add_paragraph("Body paragraph text.");                   -> plain <w:p>
%       save()
%   Gate-3 froze this as the permanent M2 oracle (references\s0033): the package
%   is 17/17 BYTE-IDENTICAL to python-docx 1.2.0, ONLY word/document.xml differs
%   from the M1 default (1865 B, SHA a71e5502...), and word/styles.xml is
%   byte-UNCHANGED from M1 (Title/Heading 1/Heading 2 pre-exist as REAL template
%   styles resolved BY NAME through the un-stubbed Paragraph.style chain -- the
%   latent-styles table is never touched by M2). The Word COM oracle PASSED
%   (com_verify_M2.md): real Word opens it silently, shows the 4 paragraphs with
%   the correct styles, and round-trips it. test_m2_milestone_package_17_of_17 is
%   the headline permanent pin -- it goes RED on ANY single-byte or zip-order
%   drift through the public authoring path. Modeled on the M1 pin in
%   Test_p1_8_skeleton_m1 (same 17-part manifest freeze idiom).
%   ============================================================================
%
%   Provenance (Gate-1..3 + COM, all 2026-07-30):
%     * Audit    : validation\mat2doc\audit_P4-7b_latent_m2.md (Porter Gate-1
%                  self-probe 52/52-match + Fable/mso-auditor Gate-2 APPROVE --
%                  M2 INDEPENDENTLY re-derived, 6 scenario pairs x 17 = 102/102
%                  byte-identical, 19/19 behavioral oracle, 3/3 verbatim error
%                  strings; ZERO defects; ZERO new D-numbers; H17 delete_ SOUND).
%     * Validate : validation\mat2doc\validate_P4-7b_latent_m2.md (Gate-3 PASS --
%                  s0033 M2 17/17 L1 byte-identical; s0034 latent-write 17/17 L1;
%                  probe_diff s0035 MATCH 53/53 exit 0; adversarial battery
%                  levels/para/latdel 17/17; delete_/GC 17/17; M1 17/17; ZERO new
%                  D-numbers).
%     * COM      : validation\mat2doc\com_verify_M2.md (Word 16.0 build 16.0.20228
%                  -- open-clean, 4 paragraphs correct text + styles by name,
%                  open-edit-save round-trip clean, PDF render correct: M2 PASS).
%     * Scenarios: validation\mat2doc\scenarios\s0033_m2_hello.{py,m} (THE M2
%                  byte proof -- exact call sequence replayed in buildM2Candidate
%                  below), s0034_latent_mutation.{py,m} (the latent WRITE-path
%                  byte proof, replayed in buildLatentMutation), s0035_p4_7b_probe.
%                  {py,m} (the 53-key full-surface probe, replayed VERBATIM by
%                  runProbes for the Equivalence leg).
%     * Frozen refs (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE at Gate-3):
%         references\s0033\ -- the 17-part M2 milestone package. Its 17-part
%           manifest (name|size|sha256, zip-entry order) is embedded below as
%           M2_MANIFEST so the sweep is self-contained; word/document.xml is also
%           copied byte-for-byte into tests\styles\data\s0033_document.xml
%           (co-located `* binary` .gitattributes) for a whole-bytes pin.
%         references\s0034\parts\word\styles.xml (SHA 3981d463..., 349578 B) --
%           copied byte-for-byte into tests\styles\data\s0034_styles.xml (binary
%           pinned) as the latent-mutation whole-bytes fixture; SHA/size embedded.
%         references\s0035\probe.json -- copied verbatim (self-contained) into
%           tests\styles\data\s0035_probe_oracle.json (the 53-key value oracle;
%           jsondecode is line-ending agnostic -> no `* binary` pin, s0031
%           precedent).
%
%   Coverage taxonomy
%   -----------------
%   * Nominal      -- LatentStyles read surface (len_/default_priority/load_count/
%                     the 4 default_to_* bools/to_array/getitem_ by name);
%                     LatentStyle_ get/set (name/priority/hidden/locked/quick_style/
%                     unhide_when_used); Styles.latent_styles resolution;
%                     add_heading level->style mapping; add_paragraph text/style;
%                     paragraphs list; the M2 authoring sequence.
%   * Edge         -- H3 raw tri-state (a FRESH lsdException returns []/None for
%                     hidden/locked/priority, CONTRAST default_to_* bool_prop
%                     effective-False); empty "" add_paragraph -> 0 runs; no-arg
%                     add_paragraph; non-ASCII (café/CJK/emoji surrogate pair);
%                     & < > escaping; single-element; error paths -- getitem_ miss
%                     -> mat2doc:KeyError verbatim, add_heading(-1)/(10) ->
%                     mat2doc:ValueError verbatim (each verifies the IDENTIFIER +
%                     the byte-verbatim message).
%   * Equivalence  -- test_equivalence_full_probe_vs_frozen_oracle replays the
%                     ENTIRE s0035 battery (runProbes, the .m twin's body verbatim)
%                     and compares every one of the 53 keys to the frozen
%                     python-docx 1.2.0 oracle (Gate-3: probe_diff MATCH exit 0).
%   * Regression   -- the M2 17-part manifest (size+SHA-256) + the M2 document.xml
%                     whole-bytes + the latent-mutation styles.xml whole-bytes +
%                     the delete_latent_style parent-side effect + the GC-safety styles.xml
%                     invariance + hard-coded escaped/UTF-8 run bytes.
%   * Upstream     -- the verbatim ValueError ("level must be in range 0-9, got
%                     %d") / KeyError ("no latent style with name '%s'") messages,
%                     the level->style_id mapping, and the H3 latent tri-state ARE
%                     the python-docx latent.py / document.py / blkcntnr.py
%                     contract; the frozen oracle IS lxml's output for this API
%                     sequence.
%
%   Byte-level (L1) note: every serialized-bytes assertion is a SHA-256 (+ size)
%   pin of the raw UTF-8 shipping bytes, plus (for document.xml and the latent
%   styles.xml) a whole-bytes comparison against the co-located frozen fixture.
%   SHA-256 equality == byte identity (L1). NO D-number granted any L2 relaxation
%   in this WP (Gate-3: ZERO new; delete_latent_style is a method-naming
%   resolution, byte-identical to python-docx _LatentStyle.delete()), so every
%   byte pin is L1. The escaped/UTF-8 substring pins assert byte-exact OOXML
%   escaping / UTF-8 encoding of the run text -- an independent oracle (the
%   escaping/encoding rules, not a MATLAB-derived hash). The probe leaf-count and
%   the latent-count cross-checks are the only looser-than-byte checks, commented
%   at their site.
%
%   Determinism: no network, no absolute paths. The co-located oracle + byte
%   fixtures resolve relative to this file via fileparts(mfilename('fullpath')).
%   Saves go to tempname .docx / tempname dirs deleted via onCleanup; every file
%   read is binary ('r','n'). The +mat2doc package resolves via the MANDATORY
%   PathFixture(worktree-root) in TestClassSetup (WP9-F4 lesson).

    properties (Constant)
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        % --- ★ THE M2 MILESTONE 17-part manifest (frozen references\s0033) ---
        % Column 1 = part name in the FROZEN zip-entry order (== s0033
        % manifest.json parts[] order); column 2 = exact byte size; column 3 =
        % lowercase SHA-256. SHA-256 equality == byte-identity (L1). IDENTICAL to
        % the M1 manifest (Test_p1_8_skeleton_m1.M1_MANIFEST) EXCEPT
        % word/document.xml, which M2 rewrites (1865 B, a71e5502...) carrying the
        % Title/Heading1/Heading2 pStyle refs + the body paragraph. word/styles.xml
        % is byte-UNCHANGED from M1 (styles resolved by name).
        M2_MANIFEST = [ ...
            "[Content_Types].xml",             "1738",   "66c84fb7a6aa3c4ead49f895e4a7044df1fb57de1ed76d09b2686e91f5bed5b4"; ...
            "_rels/.rels",                     "734",    "a9ae57efe9186f07d48303bcac5d54c7e359bf3f503939c03ad2954e5c59a5c4"; ...
            "docProps/core.xml",               "721",    "d14be8284e406d14dc2a576b19f15e7f637b955f9e27a3d9ed871ec99606f7ba"; ...
            "docProps/app.xml",                "1132",   "be664981c3141cddfc59362beb287ebf20d0773660e2dd6faac5968a5930a081"; ...
            "word/document.xml",               "1865",   "a71e550253b8c6c9f472c740b13f6e184b29bf7ac7e5694dcdfba00ecbef7c2c"; ...
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

        % --- ★ M2 milestone word/document.xml byte pin (s0033) ---
        SIZE_M2_DOC = 1865
        SHA_M2_DOC  = "a71e550253b8c6c9f472c740b13f6e184b29bf7ac7e5694dcdfba00ecbef7c2c"

        % --- M1 word/styles.xml (byte-UNCHANGED baseline; also the GC invariant) ---
        SIZE_STYLES_M1 = 349458
        SHA_STYLES_M1  = "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"

        % --- latent WRITE-path word/styles.xml byte pin (s0034) ---
        SIZE_STYLES_LATENT = 349578
        SHA_STYLES_LATENT  = "3981d4630241fa89bb5fd2485bb070010c3dcde01d1d35caf12fd9478d36ab53"

        % Fresh-Document latent counts (frozen s0035 oracle / Gate-2 behavioral).
        LATENT_LEN_FRESH = 137
    end

    properties (Access = private)
        % The M2 milestone candidate built ONCE in TestClassSetup (styles 349 KB +
        % stylesWithEffects 438 KB make re-serialization non-trivial; the M2
        % authoring sequence is the shared milestone artifact for every sweep case).
        m2Built_ (1,1) logical = false
        m2Names_
        m2Blobs_
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied from tests\styles\Test_p4_7a_styles_api.m. here is
            % tests\styles; the worktree root is two levels up.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\styles
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));

            % Build the M2 milestone candidate ONCE, AFTER the PathFixture is
            % active (so +mat2doc resolves). This single Document()...save() is the
            % milestone acceptance artifact shared by the package/document/styles
            % pins; a port defect on the authoring path fails the whole class loudly.
            [b, n] = buildM2Candidate();
            testCase.m2Blobs_ = b;
            testCase.m2Names_ = n;
            testCase.m2Built_ = true;
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. ★ THE M2 MILESTONE PACKAGE PIN (headline permanent guard)     %
        % =============================================================== %

        function test_m2_milestone_package_17_of_17(testCase)
            % ★★★ THE M2 MILESTONE PIN ★★★  Regression (L1/bin): the M2 authoring
            % sequence Document(); add_heading(t0,0); add_heading(t1,1);
            % add_heading(t2,2); add_paragraph(p); save() emits EXACTLY the 17
            % frozen parts, in EXACTLY the frozen zip-entry order, each byte-identical
            % (size + SHA-256) to the python-docx 1.2.0 M2 oracle (references\s0033).
            % 16 XML parts at L1 + docProps/thumbnail.jpeg as bin. SHA-256 equality
            % is a byte-level (L1) assertion. RED on ANY single-byte drift OR zip
            % reorder through the public authoring path. This is the PERMANENT M2
            % milestone regression guard (Gate-3 froze it; the Word COM oracle
            % PASSED). Modeled on the M1 pin in Test_p1_8_skeleton_m1.
            [blobs, names] = testCase.m2Candidate();
            want = testCase.M2_MANIFEST(:, 1)';
            testCase.verifyEqual(numel(names), 17, ...
                'M2 Document()...save() must emit exactly 17 zip entries');
            testCase.verifyEqual(names, want, ...
                'M2 zip-entry sequence must equal the frozen references\s0033 order');
            for k = 1:size(testCase.M2_MANIFEST, 1)
                nm       = testCase.M2_MANIFEST(k, 1);
                wantSize = str2double(testCase.M2_MANIFEST(k, 2));
                wantSha  = testCase.M2_MANIFEST(k, 3);
                got = entryBlob(blobs, names, nm);
                testCase.verifyEqual(numel(got), wantSize, ...
                    sprintf('M2 part %s must be exactly %d B', nm, wantSize));
                testCase.verifyEqual(sha256hex(got), wantSha, ...
                    sprintf(['M2 part %s SHA-256 must equal the frozen s0033 oracle ' ...
                    '(byte-identical L1/bin)'], nm));
            end
        end

        function test_m2_document_xml_whole_bytes(testCase)
            % Regression (M2, whole-bytes + SHA + size): word/document.xml emitted
            % by the M2 sequence is byte-identical to the frozen s0033 reference --
            % 1865 B, SHA a71e5502..., AND whole-bytes equal to the co-located
            % s0033_document.xml fixture (with first-diff diagnostics). This is the
            % part that carries the Title/Heading1/Heading2 pStyle refs + the body
            % paragraph -- the actual M2 payload.
            [blobs, names] = testCase.m2Candidate();
            doc = entryBlob(blobs, names, "word/document.xml");
            testCase.verifyEqual(numel(doc), testCase.SIZE_M2_DOC, ...
                sprintf('M2 word/document.xml must be exactly %d B', testCase.SIZE_M2_DOC));
            testCase.verifyEqual(sha256hex(doc), testCase.SHA_M2_DOC, ...
                'M2 word/document.xml SHA-256 == frozen s0033 oracle (byte-identical L1)');
            want = loadFixture('s0033_document.xml');
            verifyByteIdentical(testCase, doc, want, ...
                'M2 word/document.xml == frozen s0033 reference (whole-bytes)');

            % looser-than-byte cross-check (justified: names the payload): the three
            % pStyle refs the M2 add_heading calls resolved BY NAME are present.
            txt = char(doc);
            testCase.verifyTrue(contains(txt, '<w:pStyle w:val="Title"/>'), ...
                'M2 document.xml carries the Title pStyle (add_heading level 0)');
            testCase.verifyTrue(contains(txt, '<w:pStyle w:val="Heading1"/>'), ...
                'M2 document.xml carries the Heading1 pStyle (add_heading level 1)');
            testCase.verifyTrue(contains(txt, '<w:pStyle w:val="Heading2"/>'), ...
                'M2 document.xml carries the Heading2 pStyle (add_heading level 2)');
        end

        function test_m2_styles_xml_byte_unchanged_from_m1(testCase)
            % Regression (M2 changed-part discipline): word/styles.xml emitted by
            % the M2 sequence is byte-UNCHANGED from the M1 default (349458 B, SHA
            % 02d71a68...). Proves add_heading resolves Title/Heading 1/Heading 2 as
            % PRE-EXISTING template styles BY NAME (via CT_Styles.get_by_name) and
            % NEVER touches the latent-styles table -- exactly as the M2 milestone
            % requires. RED if a future change makes add_heading mint styles.
            [blobs, names] = testCase.m2Candidate();
            styles = entryBlob(blobs, names, "word/styles.xml");
            testCase.verifyEqual(numel(styles), testCase.SIZE_STYLES_M1, ...
                'M2 word/styles.xml size must be UNCHANGED from M1');
            testCase.verifyEqual(sha256hex(styles), testCase.SHA_STYLES_M1, ...
                'M2 word/styles.xml SHA-256 UNCHANGED from M1 -- add_heading resolves by name');
        end

        % =============================================================== %
        % 2. LatentStyles read surface (fresh Document)                    %
        % =============================================================== %

        function test_latent_styles_read(testCase)
            % Nominal + Edge (latent.py 7-107): the LatentStyles read surface over a
            % fresh Document -- len_/default_priority/load_count/the 4 default_to_*
            % bools (bool_prop effective-False/True)/to_array first name/getitem_ by
            % name with BabelFish name mapping/getitem_ miss -> KeyError verbatim.
            d = mat2doc.Document();
            ls = d.styles.latent_styles();
            testCase.verifyClass(ls, 'mat2doc.styles.LatentStyles', ...
                'd.styles.latent_styles -> LatentStyles');

            testCase.verifyEqual(ls.len_(), testCase.LATENT_LEN_FRESH, ...
                'len_ == 137 (default template lsdException count)');
            testCase.verifyEqual(ls.default_priority, 99, 'default_priority == 99');
            testCase.verifyEqual(ls.load_count, 276, 'load_count == 276');

            % the 4 default_to_* read bool_prop -> logical (effective value), NOT []
            % when absent. The default Word template's <w:latentStyles> carries
            % w:defSemiHidden="1", w:defLockedState="0", w:defQFormat="0",
            % w:defUnhideWhenUsed="1", so default_to_hidden and
            % default_to_unhide_when_used are TRUE while default_to_locked and
            % default_to_quick_style are FALSE. Values pinned against python-docx
            % 1.2.0 run DIRECTLY on this exact template this Gate-4 session
            % (default_to_hidden=True, default_to_unhide_when_used=True) and the
            % frozen s0035 oracle -- NOT hand-transcribed. default_to_unhide_when_used
            % is additionally the Gate-1 probe-typo case: the porter's probe hard-coded
            % false; the PORT is right (Python returns True), independently
            % re-confirmed by the Gate-2 auditor and here by a live python-docx run.
            testCase.verifyTrue(ls.default_to_hidden, ...
                'default_to_hidden -> TRUE (w:defSemiHidden="1"; python-docx True)');
            testCase.verifyFalse(ls.default_to_locked,      'default_to_locked -> false (w:defLockedState="0")');
            testCase.verifyFalse(ls.default_to_quick_style, 'default_to_quick_style -> false (w:defQFormat="0")');
            testCase.verifyTrue(ls.default_to_unhide_when_used, ...
                'default_to_unhide_when_used -> TRUE (w:defUnhideWhenUsed="1"; python-docx True; Gate-1 probe-typo case)');
            testCase.verifyTrue(islogical(ls.default_to_hidden), ...
                'default_to_* are logical (bool_prop effective value, not [])');

            % to_array: first name "Normal" (next(iter(ls)).name), homogeneous array
            arr = ls.to_array();
            testCase.verifyEqual(numel(arr), testCase.LATENT_LEN_FRESH, 'to_array length == len_');
            testCase.verifyClass(arr, 'mat2doc.styles.LatentStyle_', 'to_array -> LatentStyle_ array');
            testCase.verifyEqual(arr(1).name, "Normal", 'to_array(1).name == "Normal"');

            % getitem_ by name (BabelFish name mapping back through internal2ui)
            testCase.verifyEqual(ls.getitem_("Normal").name, "Normal", 'getitem_("Normal").name');
            h1 = ls.getitem_("Heading 1");
            testCase.verifyClass(h1, 'mat2doc.styles.LatentStyle_', 'getitem_ -> LatentStyle_');
            testCase.verifyEqual(h1.name, "Heading 1", 'getitem_("Heading 1").name (internal2ui)');
            testCase.verifyEqual(h1.priority, 9, 'Heading 1 latent priority == 9');
            testCase.verifyFalse(h1.hidden, 'Heading 1 latent hidden -> false');
            testCase.verifyTrue(isequal(h1.locked, []), 'Heading 1 latent locked -> [] (H3 raw tri-state None)');
            testCase.verifyTrue(h1.quick_style, 'Heading 1 latent quick_style -> true');
            testCase.verifyFalse(h1.unhide_when_used, 'Heading 1 latent unhide_when_used -> false');

            % getitem_ miss -> KeyError (identifier + verbatim message)
            ME = captureError(@() ls.getitem_("No Such Style Xyz"));
            testCase.verifyEqual(string(ME.identifier), "mat2doc:KeyError", ...
                'getitem_ miss -> mat2doc:KeyError id');
            testCase.verifyEqual(string(ME.message), "no latent style with name 'No Such Style Xyz'", ...
                'getitem_ miss -> verbatim Python KeyError message');
        end

        function test_latent_styles_setters(testCase)
            % Nominal (latent.py 33-107): the LatentStyles default setters write and
            % read back -- default_priority (defUIPriority), load_count (count), and
            % the 4 default_to_* (set_bool_prop, incl. the F-1 fixed path).
            d = mat2doc.Document();
            ls = d.styles.latent_styles();

            ls.default_priority = 42;
            testCase.verifyEqual(ls.default_priority, 42, 'default_priority set/get');
            ls.load_count = 300;
            testCase.verifyEqual(ls.load_count, 300, 'load_count set/get');

            ls.default_to_hidden = true;
            testCase.verifyTrue(ls.default_to_hidden, 'default_to_hidden set true');
            ls.default_to_locked = false;
            testCase.verifyFalse(ls.default_to_locked, 'default_to_locked set false (F-1 set_bool_prop path)');
            ls.default_to_quick_style = true;
            testCase.verifyTrue(ls.default_to_quick_style, 'default_to_quick_style set true');
            ls.default_to_unhide_when_used = false;
            testCase.verifyFalse(ls.default_to_unhide_when_used, 'default_to_unhide_when_used set false');
        end

        function test_add_latent_style_fresh_tristate(testCase)
            % Nominal + Edge (latent.py 26-31; H3): add_latent_style appends a fresh
            % lsdException (+1 len), stores the name via ui2internal passthrough, and
            % the fresh override reads RAW tri-state [] (None) for
            % hidden/locked/priority -- CONTRAST the LatentStyles default_to_* which
            % read effective-False. This is the H3 latent tri-state distinction.
            d = mat2doc.Document();
            ls = d.styles.latent_styles();
            before = ls.len_();
            fresh = ls.add_latent_style("Table Grid Probe");
            testCase.verifyClass(fresh, 'mat2doc.styles.LatentStyle_', 'add_latent_style -> LatentStyle_');
            testCase.verifyEqual(ls.len_(), before + 1, 'add_latent_style -> +1 len');
            testCase.verifyEqual(fresh.name, "Table Grid Probe", 'fresh name passthrough (ui2internal)');
            testCase.verifyTrue(isequal(fresh.hidden, []),   'fresh hidden -> [] (H3 raw tri-state None)');
            testCase.verifyTrue(isequal(fresh.locked, []),   'fresh locked -> [] (None)');
            testCase.verifyTrue(isequal(fresh.priority, []), 'fresh priority -> [] (None)');
        end

        % =============================================================== %
        % 3. LatentStyle_ setters + H17 delete_ + GC-safety                %
        % =============================================================== %

        function test_latent_style_setters(testCase)
            % Nominal (latent.py 130-198): every LatentStyle_ setter writes a value
            % readable back through its getter -- priority (uiPriority), hidden
            % (semiHidden), locked, quick_style (qFormat), unhide_when_used. name is
            % read-only (matches Python's getter-only property).
            d = mat2doc.Document();
            ls = d.styles.latent_styles();
            s = ls.add_latent_style("Table Grid Probe");
            s.priority = 7;         testCase.verifyEqual(s.priority, 7,      'priority set/get');
            s.hidden = true;        testCase.verifyTrue(s.hidden,            'hidden set true');
            s.locked = false;       testCase.verifyFalse(s.locked,           'locked set false');
            s.quick_style = true;   testCase.verifyTrue(s.quick_style,       'quick_style set true');
            s.unhide_when_used = true; testCase.verifyTrue(s.unhide_when_used, 'unhide_when_used set true');
            % name is read-only (no set.name) -- the getter reflects the added name.
            testCase.verifyEqual(s.name, "Table Grid Probe", 'name read-only getter');
        end

        function test_latent_style_delete_parent_side(testCase)
            % Regression (latent.py 119-128; PARENT-side effect ONLY): a fresh
            % LatentStyle_ delete_latent_style() detaches its w:lsdException so the
            % parent count drops back. Gate-4 RULE: assert ONLY the parent-side effect
            % (len_) -- NEVER inspect the deleted proxy afterward (its element_ is []
            % post-delete, so proxy property access errors, matching Python's
            % AttributeError; the detached element itself survives, as in Python).
            d = mat2doc.Document();
            ls = d.styles.latent_styles();
            before = ls.len_();
            s = ls.add_latent_style("Table Grid Probe");
            testCase.verifyEqual(ls.len_(), before + 1, 'precondition: add -> +1');
            s.delete_latent_style();   % detach the w:lsdException; do NOT touch s after this line
            testCase.verifyEqual(ls.len_(), before, ...
                'delete_latent_style -> parent lsdException count back to before (parent-side effect)');
        end

        function test_gc_safety_latent_proxy_layer(testCase)
            % Regression (proxy-layer GC-safety property): minting many transient
            % LatentStyle_ proxies (getitem_ / to_array) in a loop and letting them
            % go out of scope leaves the styles part bit-for-bit UNCHANGED (349458 B,
            % 02d71a68... == M1). Because no method named `delete` is overridden on
            % LatentStyle_ OR the element (only delete_latent_style /
            % delete_lsd_exception exist), proxy GC has no tree effect (H17 dissolved).
            % A regression that re-introduced a `delete` override (detaching the live
            % w:lsdException on ordinary iteration) would go RED here. NOTE: only READ proxies are
            % minted (getitem_/to_array) -- add_latent_style mutates and is excluded.
            d = mat2doc.Document();
            styles0 = testCase.saveStylesBytes(d);
            testCase.verifyEqual(numel(styles0), testCase.SIZE_STYLES_M1, ...
                'precondition: fresh styles.xml size (M1)');
            testCase.verifyEqual(sha256hex(styles0), testCase.SHA_STYLES_M1, ...
                'precondition: fresh styles.xml SHA (M1)');
            ls = d.styles.latent_styles();
            for iter = 1:25
                a   = ls.getitem_("Normal");    %#ok<NASGU>
                bx  = ls.getitem_("Heading 1"); %#ok<NASGU>
                arr = ls.to_array();            %#ok<NASGU>  137 transient proxies
                clear a bx arr
            end
            styles1 = testCase.saveStylesBytes(d);
            testCase.verifyEqual(numel(styles1), testCase.SIZE_STYLES_M1, ...
                'styles.xml size UNCHANGED after 25x transient latent proxies (GC-safe)');
            testCase.verifyEqual(sha256hex(styles1), testCase.SHA_STYLES_M1, ...
                'styles.xml SHA-256 UNCHANGED -- LatentStyle_ GC does not touch the tree (H17)');
        end

        % =============================================================== %
        % 4. Latent WRITE-path byte pin (s0034)                            %
        % =============================================================== %

        function test_latent_mutation_styles_byte_pin(testCase)
            % Regression (the latent WRITE-path byte proof, s0034): the full latent
            % mutation sequence (every default_* + load_count setter + add_latent_style
            % + every LatentStyle_ setter) saved -> word/styles.xml is byte-identical
            % to the frozen s0034 reference (349578 B, SHA 3981d463..., whole-bytes),
            % and word/document.xml is byte-UNCHANGED from M1. M2 itself never
            % exercises this path (it resolves styles by name), so this is the
            % dedicated proof that every latent setter serializes byte-for-byte to
            % python-docx.
            [blobs, names] = buildLatentMutation();
            styles = entryBlob(blobs, names, "word/styles.xml");
            testCase.verifyEqual(numel(styles), testCase.SIZE_STYLES_LATENT, ...
                sprintf('latent-mutation styles.xml must be exactly %d B', testCase.SIZE_STYLES_LATENT));
            testCase.verifyEqual(sha256hex(styles), testCase.SHA_STYLES_LATENT, ...
                'latent-mutation styles.xml SHA-256 == frozen s0034 oracle (byte-identical L1)');
            want = loadFixture('s0034_styles.xml');
            verifyByteIdentical(testCase, styles, want, ...
                'latent-mutation word/styles.xml == frozen s0034 reference (whole-bytes)');

            % changed-part discipline: word/document.xml UNCHANGED from M1 (no body).
            doc = entryBlob(blobs, names, "word/document.xml");
            testCase.verifyEqual(sha256hex(doc), ...
                "0e4dd503bc095219cd24e8312b372c588ca56bb834aacac53eafb7bf47836327", ...
                'latent mutation leaves word/document.xml byte-UNCHANGED from M1');
        end

        % =============================================================== %
        % 5. Styles.latent_styles resolution (un-stub)                     %
        % =============================================================== %

        function test_styles_latent_styles_resolves(testCase)
            % Regression (un-stub, styles.py 100-105): Styles.latent_styles now
            % RESOLVES to a LatentStyles over get_or_add_latentStyles (no longer a
            % mat2doc:notYetPorted stub -- see the re-pin in
            % Test_p4_7a_styles_api/test_unstub_resolution). The wrapped element is
            % the w:latentStyles child of the styles part.
            d = mat2doc.Document();
            ls = d.styles.latent_styles();
            testCase.verifyClass(ls, 'mat2doc.styles.LatentStyles', ...
                'Styles.latent_styles RESOLVES to a LatentStyles (un-stubbed at P4-7b)');
            testCase.verifyEqual(string(ls.element().local_part), "latentStyles", ...
                'LatentStyles wraps the w:latentStyles element');
        end

        % =============================================================== %
        % 6. add_heading level -> style mapping + ValueError               %
        % =============================================================== %

        function test_add_heading_levels_and_valueerror(testCase)
            % Nominal + Edge (document.py 90-101, H1/H14): add_heading maps level 0
            % -> "Title", 1 -> "Heading 1", 2 -> "Heading 2", 9 -> "Heading 9", the
            % default (omitted) -> "Heading 1"; each resolves to the right pStyle id.
            % level -1 and 10 -> mat2doc:ValueError with the VERBATIM message. H1:
            % `level` is DATA (no ±1 shift).
            d = mat2doc.Document();
            testCase.verifyEqual(d.add_heading("T0", 0).style.style_id, "Title",    'level 0 -> Title');
            testCase.verifyEqual(d.add_heading("T1", 1).style.style_id, "Heading1", 'level 1 -> Heading1');
            testCase.verifyEqual(d.add_heading("T2", 2).style.style_id, "Heading2", 'level 2 -> Heading2');
            testCase.verifyEqual(d.add_heading("T9", 9).style.style_id, "Heading9", 'level 9 -> Heading9');
            testCase.verifyEqual(d.add_heading("Td").style.style_id, "Heading1", ...
                'default level (omitted) -> Heading1');

            % error path: -1 and 10 (identifier + byte-verbatim message)
            MEn = captureError(@() d.add_heading("t", -1));
            testCase.verifyEqual(string(MEn.identifier), "mat2doc:ValueError", 'level -1 -> ValueError id');
            testCase.verifyEqual(string(MEn.message), "level must be in range 0-9, got -1", ...
                'level -1 -> verbatim message');
            MEt = captureError(@() d.add_heading("t", 10));
            testCase.verifyEqual(string(MEt.identifier), "mat2doc:ValueError", 'level 10 -> ValueError id');
            testCase.verifyEqual(string(MEt.message), "level must be in range 0-9, got 10", ...
                'level 10 -> verbatim message');
        end

        % =============================================================== %
        % 7. add_paragraph text/style/empty + paragraphs list             %
        % =============================================================== %

        function test_add_paragraph_text_style_and_paragraphs(testCase)
            % Nominal + Edge (document.py 109-119, blkcntnr.py 45-59; H4/H3):
            % add_paragraph("hello") -> 1 run text "hello"; add_paragraph("") -> 0
            % runs (H4 truthiness); add_paragraph() (no arg) -> 0 runs;
            % add_paragraph("styled","Heading 1") -> style_id Heading1 + 1 run;
            % add_paragraph("plain") with no style -> Normal. paragraphs() lists them
            % in document order (H1), each a Paragraph.
            d = mat2doc.Document();
            p1 = d.add_paragraph("hello");
            testCase.verifyEqual(numel(p1.runs), 1, 'add_paragraph("hello") -> 1 run');
            testCase.verifyEqual(p1.runs(1).text, "hello", 'run text == "hello"');
            p2 = d.add_paragraph("");
            testCase.verifyEqual(numel(p2.runs), 0, 'add_paragraph("") -> 0 runs (H4)');
            p3 = d.add_paragraph();
            testCase.verifyEqual(numel(p3.runs), 0, 'add_paragraph() no-arg -> 0 runs (H13 default "")');
            p4 = d.add_paragraph("styled", "Heading 1");
            testCase.verifyEqual(p4.style.style_id, "Heading1", 'styled add_paragraph -> Heading1');
            testCase.verifyEqual(numel(p4.runs), 1, 'styled add_paragraph -> 1 run');
            p5 = d.add_paragraph("plain");
            testCase.verifyEqual(p5.style.style_id, "Normal", 'unstyled add_paragraph -> Normal');

            % paragraphs list (H1 document order; each a Paragraph)
            ps = d.body_().paragraphs();
            testCase.verifyEqual(numel(ps), 5, 'body has 5 paragraphs (document order, H1)');
            testCase.verifyClass(ps, 'mat2doc.text.Paragraph', 'paragraphs -> Paragraph array');
            testCase.verifyEqual(ps(1).runs(1).text, "hello", 'paragraphs(1) is the first added (order preserved)');
        end

        % =============================================================== %
        % 8. Adversarial byte pins -- escaping + UTF-8 (independent oracle) %
        % =============================================================== %

        function test_add_paragraph_xml_escaping_bytes(testCase)
            % Edge / Regression (H7 escaping; Gate-3 para variant, 17/17 byte-identical):
            % add_paragraph with markup-significant characters serializes the run text
            % with BYTE-EXACT OOXML escaping -- & -> &amp;, < -> &lt;, > -> &gt;, while
            % ' and " are NOT escaped in text content. Asserted as a byte substring of
            % the shipped word/document.xml. Independent oracle (the OOXML escaping
            % rules, not a MATLAB-derived hash): RED if the serializer mis-escapes.
            d = mat2doc.Document();
            d.add_paragraph('A & B < C > "quoted" it''s <notatag/>');
            doc = testCase.saveDocumentXml(d);
            expected = uint8('A &amp; B &lt; C &gt; "quoted" it''s &lt;notatag/&gt;');
            testCase.verifyTrue(containsBytes(doc, expected), ...
                'document.xml must carry the byte-exact OOXML-escaped run text (& < > escaped; '' " literal)');
        end

        function test_add_paragraph_utf8_bytes(testCase)
            % Edge / Regression (H2 UTF-8; Gate-3 para variant, 17/17 byte-identical):
            % non-ASCII run text (é + CJK + an astral-plane emoji requiring a UTF-16
            % surrogate pair) serializes to its EXACT UTF-8 byte encoding in the
            % shipped word/document.xml -- no mojibake, correct surrogate handling.
            % The input and the expected bytes are BOTH built from the same pure-ASCII
            % UTF-8 byte literal (native2unicode), so the source file is ASCII-only and
            % the check is encoding-independent of the .m file's own charset.
            utf8 = uint8([99 97 102 195 169 32 ...   % "caf" + e-acute (U+00E9)
                          230 177 137 32 ...          % CJK U+6C49
                          229 173 151 32 ...          % CJK U+5B57
                          240 159 142 137]);          % emoji U+1F389 (surrogate pair in UTF-16)
            txt = native2unicode(utf8, 'UTF-8');
            d = mat2doc.Document();
            d.add_paragraph(txt);
            doc = testCase.saveDocumentXml(d);
            testCase.verifyTrue(containsBytes(doc, utf8), ...
                'document.xml must carry the exact UTF-8 encoding of "cafe/CJK/emoji" (no mojibake, surrogate-correct)');
        end

        function test_add_paragraph_tab_newline_run_mapping(testCase)
            % Edge / Regression (blkcntnr add_run text mapping; Gate-3 para variant):
            % a run whose text carries a tab and a newline maps them to <w:tab/> and
            % <w:br/> in the serialized run (the CT_R text-splitting proven at P4-5b,
            % reached here through the un-stubbed add_paragraph). One paragraph, one
            % run; document.xml carries both control elements.
            d = mat2doc.Document();
            p = d.add_paragraph(sprintf("tab\there and\nnewline"));
            testCase.verifyEqual(numel(p.runs), 1, 'tab/newline text -> a single run');
            doc = char(testCase.saveDocumentXml(d));
            testCase.verifyTrue(contains(doc, '<w:tab/>'), 'run tab -> <w:tab/>');
            testCase.verifyTrue(contains(doc, '<w:br/>'),  'run newline -> <w:br/>');
        end

        % =============================================================== %
        % 9. EQUIVALENCE -- full s0035 probe vs the frozen oracle          %
        % =============================================================== %

        function test_equivalence_full_probe_vs_frozen_oracle(testCase)
            % Equivalence: replay the ENTIRE s0035 probe battery (runProbes -- the .m
            % twin's body VERBATIM: LatentStyles read + setters, add_latent_style +
            % LatentStyle_ setters + delete_latent_style, Styles.latent_styles resolution,
            % add_heading levels + ValueError, add_paragraph + paragraphs, H5
            % identity) and compare EVERY one of the 53 keys to the frozen python-docx
            % 1.2.0 oracle copied into data\s0035_probe_oracle.json. Gate-3 found ZERO
            % divergences (probe_diff MATCH exit 0), so every value must be
            % byte/value-identical.
            port   = runProbes();
            oracle = loadProbeOracle();
            pKeys = sort(fieldnames(port));
            oKeys = sort(fieldnames(oracle));
            testCase.verifyEqual(pKeys, oKeys, ...
                'the replayed probe and the frozen oracle must have identical keys');
            % Non-trivial count guard (guards a silent-empty replay): 53 keys.
            testCase.verifyEqual(numel(oKeys), 53, ...
                'the frozen s0035 oracle must expose all 53 probe keys');
            for i = 1:numel(oKeys)
                k = oKeys{i};
                testCase.verifyEqual(string(port.(k)), string(oracle.(k)), ...
                    sprintf('probe key %s must be value-identical to the frozen oracle', k));
            end
        end

    end

    % ===================== instance helpers ============================ %
    methods (Access = private)

        function [blobs, names] = m2Candidate(testCase)
            % The cached M2 milestone candidate (built once in TestClassSetup).
            assert(testCase.m2Built_, 'M2 candidate not built in TestClassSetup');
            blobs = testCase.m2Blobs_;
            names = testCase.m2Names_;
        end

        function bytes = saveStylesBytes(~, d)
            % d.save to a temp .docx, extract word/styles.xml, return raw bytes.
            bytes = saveAndExtract(d, fullfile('word', 'styles.xml'));
        end

        function bytes = saveDocumentXml(~, d)
            % d.save to a temp .docx, extract word/document.xml, return raw bytes.
            bytes = saveAndExtract(d, fullfile('word', 'document.xml'));
        end
    end
end

% ===================== file-local helpers ============================== %

function [blobs, names] = buildM2Candidate()
    % Replay the s0033 M2 milestone body VERBATIM (validation\mat2doc\scenarios\
    % s0033_m2_hello.m): the exact strings + call order, save to a temp .docx, and
    % return the zip entries in stream (write) order. The temp file is removed.
    d = mat2doc.Document();
    d.add_heading("Document Title", 0);   % -> "Title"     pStyle val "Title"
    d.add_heading("First Section", 1);    % -> "Heading 1" pStyle val "Heading1"
    d.add_heading("A Subsection", 2);     % -> "Heading 2" pStyle val "Heading2"
    d.add_paragraph("Body paragraph text.");               % plain body paragraph
    [blobs, names] = saveAndZipList(d);
end

function [blobs, names] = buildLatentMutation()
    % Replay the s0034 latent WRITE-path body VERBATIM (validation\mat2doc\
    % scenarios\s0034_latent_mutation.m): mutate every latent default + add and
    % mutate a w:lsdException, save, return the zip entries in stream order.
    d = mat2doc.Document();
    ls = d.styles.latent_styles();
    ls.default_to_hidden = true;             % @w:defSemiHidden
    ls.default_to_locked = false;            % @w:defLockedState
    ls.default_to_quick_style = true;        % @w:defQFormat
    ls.default_to_unhide_when_used = true;   % @w:defUnhideWhenUsed
    ls.default_priority = 42;                % @w:defUIPriority
    ls.load_count = 300;                     % @w:count
    s = ls.add_latent_style("Table Grid");   % new <w:lsdException w:name="Table Grid"/>
    s.priority = 59;                         % @w:uiPriority
    s.hidden = false;                        % @w:semiHidden
    s.locked = false;                        % @w:locked
    s.quick_style = true;                    % @w:qFormat
    s.unhide_when_used = true;               % @w:unhideWhenUsed
    [blobs, names] = saveAndZipList(d);
end

function [blobs, names] = saveAndZipList(d)
    % Save a Document to a BINARY-mode temp .docx (the writer opens "wb"), enumerate
    % the zip entries in stream (write) order, delete the temp file.
    tmp = [tempname '.docx'];
    cleanup = onCleanup(@() deleteIfExists(tmp)); %#ok<NASGU>
    d.save(tmp);
    zipBytes = readBytes(tmp);
    [blobs, names] = zipEntryList(zipBytes);
end

function bytes = saveAndExtract(d, relpath)
    % Save a Document to a temp .docx, unzip to a temp dir, return the named part's
    % raw bytes. Both temp artifacts are cleaned up on exit.
    tmp = [tempname '.docx'];
    cleanTmp = onCleanup(@() deleteIfExists(tmp)); %#ok<NASGU>
    d.save(tmp);
    exdir = tempname;
    cleanEx = onCleanup(@() rmdirIfExists(exdir)); %#ok<NASGU>
    unzip(tmp, exdir);
    bytes = readBytes(fullfile(exdir, relpath));
end

function [blobs, names] = zipEntryList(zipBytes)
    % Enumerate an in-memory zip in stream (write) order -> ordered {names, blobs}.
    % java.util.zip.ZipInputStream reads local file headers in physical order, so
    % `names` is the true zip-entry write sequence. Kept file-local so the order pin
    % is independent of the reader under test. (Copied from Test_p1_8_skeleton_m1.m.)
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

function blob = entryBlob(blobs, names, membername)
    i = find(names == string(membername), 1);
    assert(~isempty(i), 'zip entry not found: %s', membername);
    blob = blobs{i};
end

function want = loadFixture(name)
    % A co-located byte fixture under tests\styles\data (binary-pinned .gitattributes).
    here = fileparts(mfilename('fullpath'));
    want = readBytes(fullfile(here, 'data', name));
end

function b = readBytes(p)
    f = fopen(p, 'r', 'n');            % binary read (no CRLF translation)
    assert(f >= 0, 'could not open for read: %s', p);
    b = fread(f, Inf, '*uint8')';
    fclose(f);
end

function deleteIfExists(p)
    if isfile(p)
        delete(p);
    end
end

function rmdirIfExists(p)
    if isfolder(p)
        rmdir(p, 's');
    end
end

function tf = containsBytes(hay, needle)
    % True if the uint8 row `needle` occurs contiguously inside the uint8 row `hay`.
    hay = uint8(hay(:)'); needle = uint8(needle(:)');
    tf = ~isempty(strfind(hay, needle)); %#ok<STREMP> % byte-subsequence search
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

function ME = captureError(fn)
    % Run fn and RETURN the caught MException (asserts it actually raised).
    raised = true;
    try
        fn();
        raised = false;
    catch ME
        return
    end
    if ~raised
        error('mat2doc:test:noRaise', 'expected an error but none was raised');
    end
end

function o = loadProbeOracle()
    % Read the co-located frozen s0035 oracle in BINARY mode (no CRLF translation)
    % and decode UTF-8 -> struct. jsondecode is whitespace/line-ending agnostic, so
    % no `* binary` pin is needed (s0031 precedent).
    here = fileparts(mfilename('fullpath'));
    p = fullfile(here, 'data', 's0035_probe_oracle.json');
    fid = fopen(p, 'r', 'n');
    if fid < 0
        error('mat2doc:test:openFail', 'cannot open frozen oracle %s', p);
    end
    raw = fread(fid, Inf, '*uint8')';
    fclose(fid);
    o = jsondecode(native2unicode(raw, 'UTF-8'));
end

% ---- s0035 probe replay (for the Equivalence leg) --------------------- %

function o = runProbes()
    % Replay the s0035 probe sequence (the .m twin's body, VERBATIM tags/inputs/
    % order) and return a flat struct of tagged STRING-encoded values (None->"None",
    % bools->"true"/"false", ints->decimal, strings raw; class token normalized;
    % exception id -> type token). Embedded so the Equivalence leg is self-contained
    % (the validation-folder scenario is NOT on the toolbox path). Mirrors
    % validation\mat2doc\scenarios\s0035_p4_7b_probe.m lines 19-134.
    o = struct();

    % ---- Section A: LatentStyles read (fresh Document) -------------------
    d = mat2doc.Document();
    ls = d.styles.latent_styles();
    o.ls_len = enc(ls.len_());
    o.ls_default_priority = enc(ls.default_priority);
    o.ls_load_count = enc(ls.load_count);
    o.ls_default_to_hidden = enc(ls.default_to_hidden);
    o.ls_default_to_locked = enc(ls.default_to_locked);
    o.ls_default_to_quick_style = enc(ls.default_to_quick_style);
    o.ls_default_to_unhide_when_used = enc(ls.default_to_unhide_when_used);
    arr = ls.to_array();
    o.ls_first_name = enc(arr(1).name);
    o.ls_get_normal_name = enc(ls.getitem_("Normal").name);
    h1 = ls.getitem_("Heading 1");
    o.ls_get_h1_name = enc(h1.name);
    o.ls_get_h1_priority = enc(h1.priority);
    o.ls_get_h1_hidden = enc(h1.hidden);
    o.ls_get_h1_locked = enc(h1.locked);
    o.ls_get_h1_quick_style = enc(h1.quick_style);
    o.ls_get_h1_unhide_when_used = enc(h1.unhide_when_used);
    try
        ls.getitem_("No Such Style Xyz");
        o.ls_keyerror_msg = "NO-ERROR-RAISED";
        o.ls_keyerror_id = "NO-ERROR-RAISED";
    catch ME
        o.ls_keyerror_msg = enc(string(ME.message));
        o.ls_keyerror_id = errtok(ME.identifier);
    end

    % ---- Section B: add_latent_style + LatentStyle_ setters + delete_latent_style ----
    d = mat2doc.Document();
    ls = d.styles.latent_styles();
    o.add_len_before = enc(ls.len_());
    fresh = ls.add_latent_style("Table Grid Probe");
    o.add_fresh_name = enc(fresh.name);
    o.add_fresh_hidden = enc(fresh.hidden);
    o.add_fresh_priority = enc(fresh.priority);
    o.add_fresh_locked = enc(fresh.locked);
    o.add_len_after = enc(ls.len_());
    fresh.priority = 7;
    o.set_priority_readback = enc(fresh.priority);
    fresh.hidden = true;
    o.set_hidden_readback = enc(fresh.hidden);
    fresh.locked = false;
    o.set_locked_readback = enc(fresh.locked);
    fresh.quick_style = true;
    o.set_quick_style_readback = enc(fresh.quick_style);
    fresh.unhide_when_used = true;
    o.set_unhide_readback = enc(fresh.unhide_when_used);
    fresh.delete_latent_style();
    o.after_delete_len = enc(ls.len_());

    % ---- Section C: LatentStyles setters read-back (fresh Document) ------
    d = mat2doc.Document();
    ls = d.styles.latent_styles();
    ls.default_priority = 42;
    o.set_default_priority_readback = enc(ls.default_priority);
    ls.default_to_hidden = true;
    o.set_default_to_hidden_readback = enc(ls.default_to_hidden);
    ls.default_to_locked = false;
    o.set_default_to_locked_readback = enc(ls.default_to_locked);
    ls.load_count = 300;
    o.set_load_count_readback = enc(ls.load_count);

    % ---- Section D: Styles.latent_styles resolution ---------------------
    d = mat2doc.Document();
    ls = d.styles.latent_styles();
    o.latent_styles_class = clstok(class(ls));
    o.latent_styles_localname = enc(string(ls.element().local_part));

    % ---- Section E: add_heading -----------------------------------------
    d = mat2doc.Document();
    p = d.add_heading("T0", 0); o.heading0_style_id = enc(p.style.style_id);
    p = d.add_heading("T1", 1); o.heading1_style_id = enc(p.style.style_id);
    p = d.add_heading("T2", 2); o.heading2_style_id = enc(p.style.style_id);
    p = d.add_heading("T9", 9); o.heading9_style_id = enc(p.style.style_id);
    p = d.add_heading("Td");    o.heading_default_style_id = enc(p.style.style_id);
    try
        d.add_heading("t", -1);
        o.heading_neg1_msg = "NO-ERROR-RAISED";
        o.heading_neg1_id = "NO-ERROR-RAISED";
    catch ME
        o.heading_neg1_msg = enc(string(ME.message));
        o.heading_neg1_id = errtok(ME.identifier);
    end
    try
        d.add_heading("t", 10);
        o.heading_10_msg = "NO-ERROR-RAISED";
        o.heading_10_id = "NO-ERROR-RAISED";
    catch ME
        o.heading_10_msg = enc(string(ME.message));
        o.heading_10_id = errtok(ME.identifier);
    end

    % ---- Section F: add_paragraph ---------------------------------------
    d = mat2doc.Document();
    p1 = d.add_paragraph("hello");
    o.para_text_runs = enc(numel(p1.runs));
    o.para_text_value = enc(p1.runs(1).text);
    p2 = d.add_paragraph("");
    o.para_empty_runs = enc(numel(p2.runs));
    p3 = d.add_paragraph();
    o.para_noarg_runs = enc(numel(p3.runs));
    p4 = d.add_paragraph("styled", "Heading 1");
    o.para_styled_style_id = enc(p4.style.style_id);
    o.para_styled_runs = enc(numel(p4.runs));
    o.paragraphs_len = enc(numel(d.body_().paragraphs()));

    % ---- Section G: H5 identity -----------------------------------------
    d = mat2doc.Document();
    lsA = d.styles.latent_styles();
    lsB = d.styles.latent_styles();
    o.h5_eq_same_element = enc(lsA == lsB);
    o.h5_item_eq = enc(lsA.getitem_("Normal") == lsB.getitem_("Normal"));
end

function s = enc(v)
    % None->"None", logicals->"true"/"false", ints->decimal string, strings raw.
    if isequal(v, [])
        s = "None";
    elseif islogical(v)
        if v, s = "true"; else, s = "false"; end
    elseif isstring(v) || ischar(v)
        s = string(v);
    elseif isnumeric(v)
        s = string(sprintf('%d', v));
    else
        s = string(v);
    end
end

function tok = clstok(name)
    % normalized class token: last dotted component, strip a trailing '_'.
    parts = split(string(name), ".");
    tok = parts(end);
    if endsWith(tok, "_"), tok = extractBefore(tok, strlength(tok)); end
end

function tok = errtok(idstr)
    % exception type token: part after the last ':' (mat2doc:KeyError -> KeyError).
    parts = split(string(idstr), ":");
    tok = parts(end);
end

classdef Test_p2_2_storypart_parts < matlab.unittest.TestCase
% TEST_P2_2_STORYPART_PARTS  Gate-4 permanent unit tests for Mat2Doc P2-2
%   (un-stub the OBJECT GRAPH of the M1 walking skeleton -- NOT the FEATURES).
%
%   Freezes the P2-2 guarantees ported in +mat2doc\+parts\ and +mat2doc\+opc\
%   from python-docx v1.2.0: the StoryPart tier inserted above DocumentPart
%   (`DocumentPart < StoryPart < XmlPart`, parts/story.py + parts/document.py);
%   `DocumentPart._styles_part`/`_settings_part` un-stubbed to real thin parts;
%   the thin `StylesPart`/`SettingsPart`/`NumberingPart` XmlPart shells
%   (parts/{styles,settings,numbering}.py); the 3 byte-neutral PartFactory row
%   flips (WML_STYLES/SETTINGS/NUMBERING -> the shells, docx/__init__.py 49-51);
%   `Part.drop_rel` made live + `Relationships.delitem` (opc/{part,rel}.py); and
%   `StoryPart.next_id` (H1). The #60 xpath hoist that these depend on is frozen
%   separately in tests\oxml\Test_p2_2_xpath_hoist.m.
%
%   Provenance (Gate-1..3, all 2026-07-26):
%     * Audit    : validation\mat2doc\audit_P2-2_storypart.md
%                  (Porter Gate-1 + Opus Gate-2 [R] APPROVE -- StoryPart naming
%                  RATIFIED; escalation NOT triggered; 15/15 Python + 42/42 MATLAB
%                  probes; 218/218 targeted regression; zero new D-numbers).
%     * Validate : validation\mat2doc\validate_P2-2_storypart.md (Gate-3 PASS --
%                  17/17 byte-neutrality escalation gate HELD; bars 3-6 MATCH the
%                  python-docx oracle; drop_rel threshold `< 2`; regression
%                  383/383).
%     * Scenario : validation\mat2doc\scenarios\s0014_p2_2_storypart.{py,m}
%                  (identical probe sequences: next_id, drop_rel/delitem truth
%                  table, part graph, NumberingPart.new).
%     * Frozen refs (python-docx 1.2.0 oracle, frozen ONCE):
%         references\s0014\probe.json -- the truth-table oracle (next_id vectors,
%           drop_rel/delitem outcomes, part-graph classes, NumberingPart.new).
%           Its scalar values are embedded below (NEXTID_* constants + inline
%           expected values) so this suite is self-contained.
%         references\s0001\ -- the frozen 17-part M1 acceptance set. Its
%           name|size|sha256 manifest (frozen zip-entry order) is embedded below
%           as M1_MANIFEST (identical to Test_p1_8_skeleton_m1.M1_MANIFEST, same
%           frozen oracle); SHA-256 equality IS byte-identity (L1).
%     * Template : +mat2doc\templates\default.docx -- ships in the toolbox, so
%           the sweep + graph probes are self-contained relative to the worktree.
%
%   Coverage taxonomy
%   -----------------
%   * Regression (17/17 byte-neutrality, THE escalation gate) -- the reparent +
%     3 flips + #60 hoist moved ZERO bytes: mat2doc.Document().save() is still
%     byte-identical (size + SHA-256) to the frozen s0001 oracle, all 17 parts.
%     REINFORCEMENT of the headline sweep that Test_p1_8_skeleton_m1 (public
%     entry) and Test_p1_6b_package_part (OpcPackage layer) own end-to-end; here
%     it is the P2-2-specific escalation pin (RED if the reparent/flips ever
%     drift a byte). (validate_P2-2 Bar 1.)
%   * Regression (3 flips land on the shells) -- PartFactory dispatches
%     WML_STYLES->StylesPart, WML_SETTINGS->SettingsPart,
%     WML_NUMBERING->NumberingPart (exact-class), and the opened default.docx
%     loads styles/settings/numbering.xml as those classes; all three shells
%     is-a XmlPart (so the reserialize bucket -- hence the bytes -- is unchanged).
%     (validate_P2-2 Bar 6; the moved exact-class pins in Test_p1_6b are ratified
%     by that class, this cross-checks the loaded-class side.)
%   * Equivalence/Regression (drop_rel/delitem truth table, threshold `< 2`) --
%     paired to the python-docx oracle: refcount 2 -> KEPT (the `< 2` threshold
%     pin), 1 -> dropped, 0-present -> dropped, absent -> mat2doc:KeyError;
%     delitem prunes `contains` but LEAVES related_parts STALE (faithful
%     dict.__delitem__), sibling intact, missing -> mat2doc:KeyError; a base Part
%     (refcount always 0) drops. (validate_P2-2 Bar 3.)
%   * Equivalence/Regression (next_id H1, 7 vectors) -- through
%     DocumentPart.next_id: {7,12,notnum}->13, {}->1, {007}->8, {""}->1, {-3}->1,
%     {3,3}->4, {1,5}->6 (data arithmetic on id VALUES, gaps not filled). Matches
%     the frozen oracle exactly. (validate_P2-2 Bar 4.)
%   * Equivalence (DocumentPart graph + handle identity) -- main_document_part
%     is-a StoryPart is-a XmlPart; styles_part_ -> StylesPart, settings_part_ ->
%     SettingsPart; SAME handle on repeat (H5/H9). (validate_P2-2 Bar 5.)
%   * Edge/Regression (thin-scope stub-safety) -- the 12 still-stubbed feature
%     accessors each raise mat2doc:notYetPorted; NumberingPart.new raises
%     mat2doc:NotImplementedError (faithful upstream, NOT a port stub); a clean
%     Document().save() fires ZERO stubs. (validate_P2-2 Bar 6.)
%
%   Deviations exercised (adopt-only, ZERO new -- Gate-3): D-001 / D-serializer-
%   nsdecl / D-zip-time via the unchanged XmlPart blob path (re-proven by 17/17);
%   the next_id ASCII-isdigit grammar note is D-002-family, not output-visible.
%   No new D-number.
%
%   Determinism: no network, no absolute paths -- the template resolves relative
%   to this file via fileparts(mfilename('fullpath')). Written artifacts are
%   tempname .docx, deleted via onCleanup; every read/write is binary ('r'/'w'
%   ,'n'; the writer opens "wb") -- no CRLF translation, no 'wt'.

    properties (Constant)
        % Frozen s0001 M1 manifest (python-docx Document().save(), frozen
        % 2026-07-25), embedded so the 17/17 sweep is self-contained. Col 1 =
        % part name in the frozen zip-entry order; col 2 = exact byte size; col 3
        % = lowercase SHA-256 (SHA-256 equality == byte-identity, the L1 pin).
        % Identical to Test_p1_8_skeleton_m1.M1_MANIFEST (same frozen oracle).
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

        % Frozen s0014 probe.json next_id oracle: {input id vector} -> next_id.
        % Col 1 = space-joined id values (empty string = no ids); col 2 =
        % expected next_id.
        NEXTID_ORACLE = { ...
            ["7" "12" "notnum"], 13; ...   % v_7_12_notnum: max{7,12}+1 (notnum excluded)
            string.empty(1, 0),   1; ...   % v_empty:       no ids -> 1
            "007",                8; ...   % v_007:         int("007")+1
            "",                   1; ...   % v_blank_id:    "".isdigit() False -> 1
            "-3",                 1; ...   % v_neg3:        "-3".isdigit() False -> 1
            ["3" "3"],            4; ...   % v_dup_3_3:     dup -> max{3}+1
            ["1" "5"],            6};      % v_gap_1_5:     gap not filled -> max{1,5}+1
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied from tests\api\Test_p1_8_skeleton_m1.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\parts
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. 17/17 byte-neutrality (THE escalation gate, P2-2 reinforce)   %
        % =============================================================== %

        function test_m1_byte_neutrality_17_of_17(testCase)
            % Regression (validate_P2-2 Bar 1, escalation gate): after the
            % StoryPart reparent + 3 PartFactory flips + #60 hoist,
            % mat2doc.Document().save() STILL emits all 17 parts byte-identical
            % (size + SHA-256) to the frozen s0001 python-docx oracle. SHA-256
            % equality is a byte-level (L1) assertion. RED on any single-byte
            % drift from the object-graph change.
            [blobs, names] = testCase.publicSaveEntries();
            testCase.verifyEqual(numel(names), 17, ...
                'mat2doc.Document().save must still emit exactly 17 parts');
            for k = 1:size(testCase.M1_MANIFEST, 1)
                nm       = testCase.M1_MANIFEST(k, 1);
                wantSize = str2double(testCase.M1_MANIFEST(k, 2));
                wantSha  = testCase.M1_MANIFEST(k, 3);
                got = entryBlob(blobs, names, nm);
                testCase.verifyEqual(numel(got), wantSize, ...
                    sprintf('part %s must be exactly %d B (byte-neutral)', nm, wantSize));
                testCase.verifyEqual(sha256hex(got), wantSha, ...
                    sprintf('part %s SHA-256 must equal frozen s0001 (byte-neutral L1/bin)', nm));
            end
        end

        function test_three_flips_dispatch_to_shells(testCase)
            % Regression (validate_P2-2 Bar 6; docx/__init__.py 49-51): the 3
            % flipped content types dispatch to the ported shells (exact-class),
            % AND the opened default.docx loads those parts as those classes.
            % Cross-checks the loaded-class side of the "row flip moves the pin"
            % that Test_p1_6b ratifies at the factory side.
            CT = mat2doc.opc.CONTENT_TYPE;
            % factory side (exact-class dispatch)
            testCase.verifyEqual(mat2doc.opc.PartFactory.part_cls_for_(CT.WML_STYLES), ...
                "mat2doc.parts.StylesPart", 'WML_STYLES must flip to StylesPart');
            testCase.verifyEqual(mat2doc.opc.PartFactory.part_cls_for_(CT.WML_SETTINGS), ...
                "mat2doc.parts.SettingsPart", 'WML_SETTINGS must flip to SettingsPart');
            testCase.verifyEqual(mat2doc.opc.PartFactory.part_cls_for_(CT.WML_NUMBERING), ...
                "mat2doc.parts.NumberingPart", 'WML_NUMBERING must flip to NumberingPart');
            % loaded-class side (opened default.docx)
            pkg = testCase.openPkg();
            testCase.verifyEqual(class(testCase.partByCT(pkg, CT.WML_STYLES)), ...
                'mat2doc.parts.StylesPart', 'styles.xml must load as StylesPart');
            testCase.verifyEqual(class(testCase.partByCT(pkg, CT.WML_SETTINGS)), ...
                'mat2doc.parts.SettingsPart', 'settings.xml must load as SettingsPart');
            testCase.verifyEqual(class(testCase.partByCT(pkg, CT.WML_NUMBERING)), ...
                'mat2doc.parts.NumberingPart', 'numbering.xml must load as NumberingPart');
        end

        function test_three_shells_isa_xmlpart(testCase)
            % Regression (byte-neutral rationale): each flipped shell is-a XmlPart
            % (inherits the parse+serialize_part_xml blob), so the reserialize
            % bucket -- and hence the emitted bytes -- is unchanged. This is WHY
            % the flips are byte-neutral (test above proves they are).
            pkg = testCase.openPkg();
            CT = mat2doc.opc.CONTENT_TYPE;
            for ct = [CT.WML_STYLES, CT.WML_SETTINGS, CT.WML_NUMBERING]
                p = testCase.partByCT(pkg, ct);
                testCase.verifyTrue(isa(p, 'mat2doc.opc.XmlPart'), ...
                    sprintf('flipped shell for %s must be-a XmlPart (byte-neutral bucket)', ct));
            end
        end

        % =============================================================== %
        % 2. drop_rel / delitem refcount truth table (threshold `< 2`)     %
        % =============================================================== %

        function test_drop_rel_refcount_truth_table(testCase)
            % Equivalence/Regression (validate_P2-2 Bar 3): the exact threshold is
            % `< 2` (part.py:81). An XmlPart carrying r:id rId9x2, rId3x1, with
            % rels {rId9, rId3, rId5(0 refs in XML)}:
            %   drop_rel(rId9) refcount 2 -> NOT < 2 -> KEPT   (the `< 2` pin)
            %   drop_rel(rId3) refcount 1 -> 1 < 2  -> dropped
            %   drop_rel(rId5) refcount 0 -> 0 < 2  -> dropped (implicit rel)
            %   drop_rel(rId404) absent   -> del rels[..] -> mat2doc:KeyError
            [xp, ~] = testCase.buildRefcountPart();
            xp.drop_rel("rId9");
            testCase.verifyTrue(xp.rels().contains("rId9"), ...
                'refcount 2 must be KEPT (not < 2) -- the exact `< 2` threshold pin');
            xp.drop_rel("rId3");
            testCase.verifyFalse(xp.rels().contains("rId3"), ...
                'refcount 1 (1 < 2) must be dropped');
            xp.drop_rel("rId5");
            testCase.verifyFalse(xp.rels().contains("rId5"), ...
                'refcount 0 present (0 < 2) must be dropped (implicit rel)');
            caught = testCase.catchCall(@() xp.drop_rel("rId404"));
            testCase.assertNotEmpty(caught, 'drop_rel on an absent rId must raise');
            testCase.verifyEqual(caught.identifier, 'mat2doc:KeyError', ...
                'drop_rel on an absent rId must raise mat2doc:KeyError');
        end

        function test_base_part_drop_rel_refcount0_dropped(testCase)
            % Regression: a base Part has rel_ref_count_ == 0 always (no XML
            % references), so drop_rel on any present rel drops it (0 < 2).
            pkg = testCase.openPkg();
            RT = mat2doc.opc.RELATIONSHIP_TYPE;
            tgt = mat2doc.opc.Part(mat2doc.opc.PackURI("/word/styles.xml"), ...
                "application/xml", [], pkg);
            bp = mat2doc.opc.Part(mat2doc.opc.PackURI("/word/document.xml"), ...
                mat2doc.opc.CONTENT_TYPE.WML_DOCUMENT_MAIN, [], pkg);
            bp.load_rel(RT.STYLES, tgt, "rIdX");
            bp.drop_rel("rIdX");
            testCase.verifyFalse(bp.rels().contains("rIdX"), ...
                'base Part (refcount 0) drop_rel must drop the rel');
        end

        function test_delitem_prunes_contains_leaves_related_stale(testCase)
            % Equivalence/Regression (validate_P2-2 Bar 3 quirk row): delitem is
            % the inherited dict.__delitem__ -- it prunes ONLY the dict entry
            % (rIds_/rels_) and does NOT touch the parallel target_parts_by_rId_
            % map, so a stale related_parts entry survives a drop EXACTLY as
            % upstream. Sibling intact; missing key -> mat2doc:KeyError.
            pkg = testCase.openPkg();
            RT = mat2doc.opc.RELATIONSHIP_TYPE;
            tgt = mat2doc.opc.Part(mat2doc.opc.PackURI("/word/styles.xml"), ...
                "application/xml", [], pkg);
            p = mat2doc.opc.Part(mat2doc.opc.PackURI("/word/document.xml"), ...
                mat2doc.opc.CONTENT_TYPE.WML_DOCUMENT_MAIN, [], pkg);
            p.load_rel(RT.STYLES,   tgt, "rIdA");   % internal -> recorded in related_parts
            p.load_rel(RT.SETTINGS, tgt, "rIdB");
            rels = p.rels();
            testCase.verifyTrue(rels.contains("rIdA"), 'precondition: rIdA present');
            testCase.verifyTrue(rels.related_parts.isKey("rIdA"), ...
                'precondition: rIdA in related_parts');
            rels.delitem("rIdA");
            testCase.verifyFalse(rels.contains("rIdA"), ...
                'delitem must prune the dict entry (contains -> false)');
            testCase.verifyTrue(rels.related_parts.isKey("rIdA"), ...
                'delitem must LEAVE related_parts STALE (faithful dict.__delitem__)');
            testCase.verifyTrue(rels.contains("rIdB"), 'sibling rIdB must be intact');
            caught = testCase.catchCall(@() rels.delitem("rIdZ"));
            testCase.assertNotEmpty(caught, 'delitem on a missing key must raise');
            testCase.verifyEqual(caught.identifier, 'mat2doc:KeyError', ...
                'delitem on a missing key must raise mat2doc:KeyError');
        end

        % =============================================================== %
        % 3. next_id H1 (7 vectors, frozen oracle)                        %
        % =============================================================== %

        function test_next_id_seven_vectors(testCase)
            % Equivalence/Regression (validate_P2-2 Bar 4): DocumentPart.next_id
            % over 7 synthetic bodies matches the frozen python-docx oracle
            % exactly -- data arithmetic on id VALUES (max+1 / 1), gaps NOT
            % filled, non-numeric ids excluded (ASCII isdigit).
            pkg = testCase.openPkg();
            dp = pkg.main_document_part();
            pn = dp.partname();
            ct = dp.content_type();
            for k = 1:size(testCase.NEXTID_ORACLE, 1)
                ids  = testCase.NEXTID_ORACLE{k, 1};
                want = testCase.NEXTID_ORACLE{k, 2};
                got = testCase.nextIdFor(pn, ct, pkg, ids);
                testCase.verifyEqual(got, want, ...
                    sprintf('next_id for [%s] must be %d (frozen s0014 oracle)', ...
                    strjoin(cellstr(ids), ","), want));
            end
        end

        % =============================================================== %
        % 4. DocumentPart graph + handle identity                         %
        % =============================================================== %

        function test_documentpart_graph_and_handle_identity(testCase)
            % Equivalence (validate_P2-2 Bar 5): main_document_part is-a StoryPart
            % is-a XmlPart; styles_part_ -> StylesPart, settings_part_ ->
            % SettingsPart; and each returns the SAME handle on repeat (H5/H9
            % identity via the live rels). The exact-class pin also guards the
            % inherited-static trap (a silent base-XmlPart build goes RED).
            pkg = testCase.openPkg();
            dp = pkg.main_document_part();
            testCase.verifyEqual(class(dp), 'mat2doc.parts.DocumentPart', ...
                'main_document_part must be a DocumentPart');
            testCase.verifyTrue(isa(dp, 'mat2doc.parts.StoryPart'), ...
                'DocumentPart must be-a StoryPart (P2-2 reparent)');
            testCase.verifyTrue(isa(dp, 'mat2doc.opc.XmlPart'), ...
                'DocumentPart must be-a XmlPart');
            sp1 = dp.styles_part_();
            sp2 = dp.styles_part_();
            testCase.verifyEqual(class(sp1), 'mat2doc.parts.StylesPart', ...
                'styles_part_ must return a StylesPart');
            testCase.verifyTrue(sp1 == sp2, ...
                'styles_part_ must return the SAME handle on repeat (H5/H9)');
            st1 = dp.settings_part_();
            st2 = dp.settings_part_();
            testCase.verifyEqual(class(st1), 'mat2doc.parts.SettingsPart', ...
                'settings_part_ must return a SettingsPart');
            testCase.verifyTrue(st1 == st2, ...
                'settings_part_ must return the SAME handle on repeat (H5/H9)');
        end

        % =============================================================== %
        % 5. Thin-scope stub-safety                                       %
        % =============================================================== %

        function test_twelve_feature_stubs_raise_notyetported(testCase)
            % Edge/Regression (validate_P2-2 Bar 6): the still-stubbed feature
            % accessors each raise mat2doc:notYetPorted (thin-scope: the graph is
            % live, the FEATURE surface is not). Instances are read off the live
            % graph so the delegation chain is real.
            %
            % REGISTRY-FLIP RE-PIN (P4-7a Gate-4): P4-7a un-stubbed the styles
            % delegation, so FOUR of the original twelve now RESOLVE and moved to the
            % positive block below -- DocumentPart.styles / DocumentPart.get_style /
            % DocumentPart.get_style_id / StylesPart.styles.
            % REGISTRY-FLIP RE-PIN (P5-1 Gate-4): P5-1 un-stubbed the settings
            % delegation, so TWO MORE now RESOLVE and moved to the positive block
            % below -- DocumentPart.settings / SettingsPart.settings (8 -> 6).
            % REGISTRY-FLIP RE-PIN (P5-3b Gate-4): P5-3b un-stubbed the hdr/ftr part
            % factories, so add_header_part / add_footer_part now RESOLVE (return
            % [part, rId]) and moved to the positive block below (6 -> 4).
            % REGISTRY-FLIP RE-PIN (P7-4, the picture milestone WP; validate_P7-4 s6
            % re-pin 5): StoryPart.get_or_add_image is now LIVE -> it reaches the image
            % loader instead of the notYetPorted stub and moved to the positive block
            % below (4 -> 3).
            % REGISTRY-FLIP RE-PIN (P8-1 Gate-4): P8-1 un-stubbed BOTH
            % DocumentPart.numbering_part (default.docx ships a numbering part, so it
            % now RESOLVES to a NumberingPart via part_related_by(NUMBERING)) AND
            % NumberingPart.numbering_definitions (returns a NumberingDefinitions_).
            % Both moved to the positive block below (3 -> 1). The ONE remaining
            % genuine stub is DocumentPart.inline_shapes (the inline-shapes tier,
            % P8-2+). (registry-flip stale-pins lesson; the full numbering surface is
            % pinned in tests\oxml\Test_p8_1_numbering.m.)
            pkg = testCase.openPkg();
            dp = pkg.main_document_part();
            sp = dp.styles_part_();      % real StylesPart
            st = dp.settings_part_();    % real SettingsPart
            np = testCase.partByCT(pkg, mat2doc.opc.CONTENT_TYPE.WML_NUMBERING); % NumberingPart
            calls = { ...
                @() dp.inline_shapes(),          'DocumentPart.inline_shapes'};
            testCase.verifyEqual(size(calls, 1), 1, 'exactly 1 genuine feature stub remains after the P4-7a styles + P5-1 settings + P5-3b hdr/ftr + P7-4 image + P8-1 numbering un-stubs');
            for k = 1:size(calls, 1)
                caught = testCase.catchCall(calls{k, 1});
                testCase.assertNotEmpty(caught, ...
                    sprintf('stub %s must raise', calls{k, 2}));
                testCase.verifyEqual(caught.identifier, 'mat2doc:notYetPorted', ...
                    sprintf('stub %s must raise mat2doc:notYetPorted', calls{k, 2}));
            end

            % --- the four styles paths P4-7a un-stubbed now RESOLVE (no stub) ---
            PARA = mat2doc.enum.style.WD_STYLE_TYPE.PARAGRAPH;
            testCase.verifyClass(dp.styles(), 'mat2doc.styles.Styles', ...
                'DocumentPart.styles RESOLVES (un-stubbed at P4-7a)');
            testCase.verifyClass(sp.styles(), 'mat2doc.styles.Styles', ...
                'StylesPart.styles RESOLVES (un-stubbed at P4-7a)');
            testCase.verifyClass(dp.get_style("Heading1", PARA), 'mat2doc.styles.ParagraphStyle', ...
                'DocumentPart.get_style RESOLVES (un-stubbed at P4-7a)');
            testCase.verifyEqual(dp.get_style_id("Heading 1", PARA), "Heading1", ...
                'DocumentPart.get_style_id RESOLVES (un-stubbed at P4-7a)');

            % --- the two settings paths P5-1 un-stubbed now RESOLVE (no stub) ---
            testCase.verifyClass(dp.settings(), 'mat2doc.settings.Settings', ...
                'DocumentPart.settings RESOLVES to a Settings proxy (un-stubbed at P5-1)');
            testCase.verifyClass(st.settings(), 'mat2doc.settings.Settings', ...
                'SettingsPart.settings RESOLVES to a Settings proxy (un-stubbed at P5-1)');
            % H5 eq-chain: DocumentPart.settings delegates to SettingsPart.settings
            % over the SAME CT_Settings element, so they compare == True.
            testCase.verifyTrue(dp.settings() == st.settings(), ...
                'H5: DocumentPart.settings == SettingsPart.settings (same element)');

            % --- the two hdr/ftr part factories P5-3b un-stubbed now RESOLVE ---
            % add_header_part / add_footer_part return [part, rId]: a
            % HeaderPart/FooterPart (HeaderPart/FooterPart.new via next_partname)
            % plus the string rId from relate_to (RT.HEADER / RT.FOOTER). The full
            % add/get/drop surface + the header/footer byte packages are pinned in
            % tests\section\Test_p5_3b_hdrftr_api.m; here it is only the un-stub.
            [hdrPart, hRid] = dp.add_header_part();
            testCase.verifyClass(hdrPart, 'mat2doc.parts.HeaderPart', ...
                'DocumentPart.add_header_part RESOLVES to a HeaderPart (un-stubbed at P5-3b)');
            testCase.verifyClass(hRid, 'string', 'add_header_part returns a string rId');
            [ftrPart, fRid] = dp.add_footer_part();
            testCase.verifyClass(ftrPart, 'mat2doc.parts.FooterPart', ...
                'DocumentPart.add_footer_part RESOLVES to a FooterPart (un-stubbed at P5-3b)');
            testCase.verifyClass(fRid, 'string', 'add_footer_part returns a string rId');

            % --- the StoryPart.get_or_add_image path P7-4 un-stubbed now RESOLVES ---
            % It is LIVE at the picture milestone: it reaches the image loader
            % (package.get_or_add_image_part -> Image.from_file) instead of the
            % notYetPorted stub. On a BOGUS path it fails with the faithful
            % mat2doc:FileNotFoundError (image.py 35-50), NOT mat2doc:notYetPorted --
            % the identifier change IS the un-stub proof. The byte-exact [rId, Image]
            % return over a real image is pinned in tests\parts\Test_p7_4_add_picture.m.
            caughtImg = testCase.catchCall(@() dp.get_or_add_image("no_such_image.png"));
            testCase.assertNotEmpty(caughtImg, ...
                'StoryPart.get_or_add_image on a bogus path must still error (image load)');
            testCase.verifyEqual(caughtImg.identifier, 'mat2doc:FileNotFoundError', ...
                'StoryPart.get_or_add_image is LIVE (P7-4): bogus path -> FileNotFoundError, NOT notYetPorted');

            % --- the two numbering paths P8-1 un-stubbed now RESOLVE (no stub) ---
            % DocumentPart.numbering_part RESOLVES to the package's NumberingPart
            % (default.docx ships numbering.xml, related via RT.NUMBERING, so the
            % faithful NotImplementedError new() branch is never reached).
            % NumberingPart.numbering_definitions RESOLVES to a NumberingDefinitions_
            % (len == 9 <w:num> on default.docx). The full numbering surface +
            % lazyproperty identity is pinned in tests\oxml\Test_p8_1_numbering.m.
            testCase.verifyClass(dp.numbering_part(), 'mat2doc.parts.NumberingPart', ...
                'DocumentPart.numbering_part RESOLVES to a NumberingPart (un-stubbed at P8-1)');
            nd = np.numbering_definitions();
            testCase.verifyClass(nd, 'mat2doc.parts.NumberingDefinitions_', ...
                'NumberingPart.numbering_definitions RESOLVES to a NumberingDefinitions_ (un-stubbed at P8-1)');
            testCase.verifyEqual(nd.len_(), 9, ...
                'default.docx numbering_definitions len == 9 <w:num>');
        end

        function test_numberingpart_new_raises_notimplemented(testCase)
            % Regression (validate_P2-2 Bar 6): NumberingPart.new raises
            % mat2doc:NotImplementedError -- this is FAITHFUL upstream behaviour
            % (python-docx numbering.py 11-14 itself `raise NotImplementedError`),
            % NOT a port stub. The identifier distinguishes it from notYetPorted.
            caught = testCase.catchCall(@() mat2doc.parts.NumberingPart.new());
            testCase.assertNotEmpty(caught, 'NumberingPart.new must raise');
            testCase.verifyEqual(caught.identifier, 'mat2doc:NotImplementedError', ...
                'NumberingPart.new must raise mat2doc:NotImplementedError (faithful upstream)');
        end

        function test_clean_save_fires_zero_stubs(testCase)
            % Regression (validate_P2-2 Bar 6 / thin-scope): a clean
            % mat2doc.Document().save() completes with ZERO stubs on the open/save
            % path -- none of the P2-2 graph un-stubbing pulled a feature stub
            % onto the byte path. A fired stub would throw; the written file IS
            % the mechanical no-stub proof.
            d = mat2doc.Document();
            tmp = [tempname '.docx'];
            cleanup = onCleanup(@() deleteIfExists(tmp)); %#ok<NASGU>
            caught = testCase.catchCall(@() d.save(tmp));
            testCase.verifyEmpty(caught, ...
                'clean Document().save must NOT fire any stub');
            testCase.verifyTrue(isfile(tmp), ...
                'clean Document().save must write the output file');
        end

    end

    % ===================== instance helpers ============================ %
    methods (Access = private)

        function pkg = openPkg(testCase)
            % Open the shipped default.docx into a mat2doc.package.Package.
            pkg = mat2doc.package.Package.open(char(testCase.templatePath()));
        end

        function [blobs, names] = publicSaveEntries(testCase)
            % mat2doc.Document().save() -> the zip entries in stream (write) order.
            d = mat2doc.Document();                    % no arg -> default template
            tmp = [tempname '.docx'];
            cleanup = onCleanup(@() deleteIfExists(tmp)); %#ok<NASGU>
            d.save(tmp);
            zipBytes = readBytes(tmp);
            [blobs, names] = zipEntryList(zipBytes);
        end

        function [xp, tgt] = buildRefcountPart(testCase)
            % An XmlPart whose element carries r:id rId9x2, rId3x1, related to a
            % base Part via rels {rId9, rId3, rId5} -- the drop_rel truth-table
            % fixture (mirrors scenario s0014). rId5 has 0 references in the XML.
            pkg = testCase.openPkg();
            RT = mat2doc.opc.RELATIONSHIP_TYPE;
            W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            rc = "<w:document xmlns:w=""" + W + """ xmlns:r=""" + R + """><w:body>" + ...
                "<w:x r:id=""rId9""/><w:y r:id=""rId9""/><w:z r:id=""rId3""/>" + ...
                "</w:body></w:document>";
            xp = mat2doc.opc.XmlPart.load(mat2doc.opc.PackURI("/word/document.xml"), ...
                mat2doc.opc.CONTENT_TYPE.WML_DOCUMENT_MAIN, ...
                uint8(unicode2native(char(rc), "UTF-8")), pkg);
            tgt = mat2doc.opc.Part(mat2doc.opc.PackURI("/word/styles.xml"), ...
                "application/xml", [], pkg);
            xp.load_rel(RT.STYLES,    tgt, "rId9");
            xp.load_rel(RT.SETTINGS,  tgt, "rId3");
            xp.load_rel(RT.NUMBERING, tgt, "rId5");
        end

        function n = nextIdFor(~, pn, ct, pkg, ids)
            % DocumentPart.next_id over a synthetic w:document body carrying the
            % given id values (mirrors scenario s0014 next_id_for).
            W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            body = "";
            for i = 1:numel(ids)
                body = body + "<w:p id=""" + ids(i) + """/>";
            end
            xml = "<w:document xmlns:w=""" + W + """><w:body>" + body + "</w:body></w:document>";
            syn = mat2doc.parts.DocumentPart.load(pn, ct, ...
                uint8(unicode2native(char(xml), "UTF-8")), pkg);
            n = syn.next_id();
        end

        function p = partByCT(~, pkg, ct)
            % First reachable part whose content_type == ct, off the live graph.
            parts = pkg.iter_parts();
            p = [];
            for k = 1:numel(parts)
                if string(parts(k).content_type()) == string(ct)
                    p = parts(k);
                    return
                end
            end
            assert(~isempty(p), 'part with content-type %s not found in graph', ct);
        end

        function caught = catchCall(~, fn)
            % Run fn; return the caught MException, or [] if it did not raise.
            caught = [];
            try
                fn();
            catch ME
                caught = ME;
            end
        end

        function p = templatePath(~)
            here = fileparts(mfilename('fullpath'));   % tests\parts
            root = fileparts(fileparts(here));         % worktree root
            p = fullfile(root, '+mat2doc', 'templates', 'default.docx');
        end
    end
end

% ===================== file-local helpers ============================== %

function [blobs, names] = zipEntryList(zipBytes)
    % Enumerate an in-memory zip in stream (write) order -> ordered {names, blobs}.
    % (Copied from Test_p1_8_skeleton_m1.m; kept file-local so the enumeration is
    % independent of the reader under test.)
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

function h = sha256hex(bytes)
    % Lowercase hex SHA-256 of a uint8 vector (base MATLAB / Java MessageDigest).
    md  = java.security.MessageDigest.getInstance('SHA-256');
    dig = md.digest(typecast(uint8(bytes(:)'), 'int8'));
    b   = typecast(dig, 'uint8');
    h   = lower(string(sprintf('%02x', b)));
end

classdef Test_p4_1a_oxml_font < matlab.unittest.TestCase
% TEST_P4_1A_OXML_FONT  Gate-4 permanent unit tests for Mat2Doc P4-1a
%   (src/docx/oxml/text/font.py -> +mat2doc\+oxml\+text\CT_RPr + the rPr child
%   classes CT_Color/CT_Fonts/CT_Highlight/CT_HpsMeasure/CT_Underline/
%   CT_VerticalAlignRun, plus the shared CT_OnOff/CT_String registered for the
%   28 font-block tags).
%
%   P4-1a is the FIRST P4 WP and the FIRST real w:-content element surface; it
%   sits on the M2 byte-critical path. The two escalation risks this class
%   permanently freezes:
%
%     (H11 ORDERING, the crux) CT_RPr's 27 ZeroOrOne child descriptors carry
%     NON-CONTIGUOUS successor slices of a single 39-entry _tag_seq. Adding
%     children in ANY scrambled order MUST re-sort them into canonical OOXML
%     schema order (insert_element_before / first_child_found_in), so
%     document.xml / styles.xml stay byte-identical. A future edit to a
%     successor slice goes RED here.
%
%     (M1 BYTE-NEUTRALITY) Registering the 28 font-block tags flips the PARSE
%     CLASS of every w:rPr / w:b / w:color / ... inside a real styles.xml from
%     generic XmlElement to the new CT_* classes. All CT_* exit through the
%     identical +oxml\serialize_part_xml walk, so the flip is byte-neutral.
%     This class pins word/styles.xml (349458 B) specifically -- it is the part
%     P4-1a could break via the CT_RPr parse path (the full 17/17 M1 sweep is
%     owned by Test_p1_8_skeleton_m1; this class adds the P4-1a-specific
%     styles.xml pin + a parse->serialize L1 self-check that exercises the
%     CT_RPr path directly).
%
%   Provenance (Gate-1..3, all 2026-07-26):
%     * Audit    : validation\mat2doc\audit_P4-1a_oxml_font.md (Porter Gate-1 +
%                  Opus Gate-2 adversarial APPROVE on both escalation gates;
%                  F1/F2 comment fixes in place).
%     * Validate : validation\mat2doc\validate_P4-1a_oxml_font.md (Gate-3 PASS --
%                  L1/exact throughout; M1 17/17 byte-neutrality [styles.xml
%                  349458 B & document.xml 1548 B L1]; H11 four scrambled
%                  constructions byte-exact; real-styles.xml round-trip L1;
%                  probe_diff MATCH; 0 equivalence FAIL; 0 new D-numbers;
%                  regression 535/535).
%     * Scenario : validation\mat2doc\scenarios\s0020_p4_1a_oxml_font.{py,m}
%                  (the probe sequence replayed VERBATIM by runProbes() below).
%     * Frozen refs (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE):
%         references\s0020\probe.json -- copied verbatim (self-contained) into
%           tests\oxml\data\s0020_probe_oracle.json (the Equivalence replay set);
%         references\s0001\parts\word\styles.xml (349458 B, sha
%           02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384) --
%           the M1 styles.xml byte reference; NOT copied (the byte pin compares
%           against what mat2doc.Document().save() itself emits, so no 349 KB
%           fixture / `* binary` .gitattributes is needed).
%
%   Coverage taxonomy
%   -----------------
%   * Nominal / Regression -- H11 re-sort of scrambled adds to schema order,
%     CT_RPr @property setters, child CT_* round-trips, CT_OnOff tri-state,
%     _new_color seeding, getter reads: each pinned to a hard-coded expected
%     serialized XML string or value taken verbatim from the frozen oracle.
%   * Edge -- empty rPr (setters from nothing), None/[] removal branches, the
%     non-contiguous parsed_nondesc H11 case with real non-descriptor siblings,
%     the rFonts_hAnsi asymmetry, u_val remove-ALL-then-add against a duplicated
%     <w:u>, the CT_OnOff True/None attr-removal, single-token from_xml set.
%   * Equivalence -- test_equivalence_full_battery_vs_frozen_oracle replays the
%     ENTIRE s0020 battery live (runProbes, the .m twin's body verbatim) and
%     flatten-compares every leaf to the frozen python-docx 1.2.0 oracle copied
%     into data\s0020_probe_oracle.json (0 divergences: byte/value-identical).
%   * Regression (bytes) -- styles.xml SHA-256 + size; serialized XML strings
%     compared byte-for-byte (all ASCII, so string-equality == byte-equality L1).
%
%   Byte-level (L1) note: every serialized-XML assertion compares the FULL
%   serialize_part_xml output (XML declaration + body) as a decoded string; the
%   content is pure ASCII, so string-equality is a byte-identical (L1) assertion.
%   No D-number granted any L2 relaxation in this WP (Gate-3 §7: zero new, none
%   exercised at L2), so every pin here is L1.
%
%   Determinism: no network, no absolute paths. The worktree root and the
%   co-located oracle resolve relative to this file via
%   fileparts(mfilename('fullpath')); saves go to tempname .docx deleted via
%   onCleanup; every file read is binary ('r','n'); no 'wt'. The +mat2doc
%   package resolves via the MANDATORY PathFixture(worktree-root) in
%   TestClassSetup (WP9-F4 lesson).

    properties (Constant)
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        % --- frozen s0001 word/styles.xml byte reference (M1, the P4-1a risk) ---
        STYLES_SIZE = 349458
        STYLES_SHA  = "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"

        % --- canonical OOXML schema order of CT_RPr's 27 ZeroOrOne descriptors
        %     (frozen s0020 h11.loose_reversed27 / loose_shuffle27 result_localnames):
        %     ANY scrambled add-order must converge to exactly this sequence. ---
        SCHEMA_ORDER27 = ["rStyle" "rFonts" "b" "bCs" "i" "iCs" "caps" "smallCaps" ...
            "strike" "dstrike" "outline" "shadow" "emboss" "imprint" "noProof" ...
            "snapToGrid" "vanish" "webHidden" "color" "sz" "highlight" "u" ...
            "vertAlign" "rtl" "cs" "specVanish" "oMath"]
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied verbatim from tests\oxml\Test_p2_3_document_shell.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\oxml
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. H11 successor ordering (THE M2-ordering crux)                %
        % =============================================================== %

        function test_h11_loose_scrambles_reorder_to_schema(testCase)
            % Nominal / Regression (H11, s0020 h11.loose_*): rebuild a loose CT_RPr
            % and add ZeroOrOne children in three scrambled orders -- an 8-child
            % scramble, the full 27 descriptors REVERSED, and a fixed 27-descriptor
            % SHUFFLE. Each must re-sort to canonical schema order (child localnames
            % AND serialized bytes). The 8-child case is pinned inline; the two
            % 27-child cases pin localnames inline (SCHEMA_ORDER27) and bytes vs the
            % frozen oracle. A successor-slice regression moves the order -> RED.
            oracle = loadOracle();

            % -- loose_scramble8: u,b,color,sz,rStyle,vertAlign,highlight,i --
            e = buildRPr("", ["u" "b" "color" "sz" "rStyle" "vertAlign" "highlight" "i"]);
            testCase.verifyEqual(childLocalnames(e), ...
                ["rStyle" "b" "i" "color" "sz" "highlight" "u" "vertAlign"], ...
                'scramble8 must re-sort to schema order');
            testCase.verifyEqual(ser(e), decl() + newline + ...
                "<w:rPr xmlns:w=""" + testCase.W + """>" + ...
                "<w:rStyle/><w:b/><w:i/><w:color w:val=""000000""/><w:sz/>" + ...
                "<w:highlight/><w:u/><w:vertAlign/></w:rPr>", ...
                'scramble8 serialized bytes (L1) must match the frozen oracle');

            % -- loose_reversed27: all 27 descriptors added in REVERSE schema order --
            eRev = buildRPr("", flip(testCase.SCHEMA_ORDER27));
            testCase.verifyEqual(childLocalnames(eRev), testCase.SCHEMA_ORDER27, ...
                'reversed-27 must fully re-sort to canonical schema order');
            testCase.verifyEqual(ser(eRev), string(oracle.h11.loose_reversed27.serialized), ...
                'reversed-27 serialized bytes (L1) must match the frozen oracle');

            % -- loose_shuffle27: a fixed 27-descriptor shuffle --
            shuffle27 = ["vertAlign" "b" "oMath" "color" "rStyle" "highlight" "i" "sz" ...
                "u" "specVanish" "caps" "webHidden" "rFonts" "strike" "bCs" "rtl" ...
                "vanish" "iCs" "cs" "smallCaps" "dstrike" "noProof" "outline" "shadow" ...
                "emboss" "imprint" "snapToGrid"];
            eSh = buildRPr("", shuffle27);
            testCase.verifyEqual(childLocalnames(eSh), testCase.SCHEMA_ORDER27, ...
                'shuffle-27 must fully re-sort to canonical schema order');
            testCase.verifyEqual(ser(eSh), string(oracle.h11.loose_shuffle27.serialized), ...
                'shuffle-27 serialized bytes (L1) must match the frozen oracle');
        end

        function test_h11_parsed_nondesc_noncontiguous_slices(testCase)
            % Edge / Regression (H11, DECISIVE -- s0020 h11.parsed_nondesc): add 8
            % descriptors into a PARSED rPr already holding NON-DESCRIPTOR siblings
            % w:spacing / w:szCs / w:lang. This exercises the NON-CONTIGUOUS successor
            % slices (color=[19:], sz=[24:], highlight=[26:], u=[27:], vertAlign=[32:],
            % specVanish=[38:]) against REAL intervening siblings: the new descriptors
            % must interleave correctly AROUND the pre-existing tags. Result order:
            % rStyle,b,color,spacing,sz,szCs,highlight,u,vertAlign,lang,specVanish.
            e = buildRPr("<w:spacing/><w:szCs/><w:lang/>", ...
                ["specVanish" "vertAlign" "u" "highlight" "sz" "color" "b" "rStyle"]);
            testCase.verifyEqual(childLocalnames(e), ...
                ["rStyle" "b" "color" "spacing" "sz" "szCs" "highlight" "u" ...
                 "vertAlign" "lang" "specVanish"], ...
                'parsed_nondesc: new descriptors must interleave around w:spacing/w:szCs/w:lang');
            testCase.verifyEqual(ser(e), decl() + newline + ...
                "<w:rPr xmlns:w=""" + testCase.W + """>" + ...
                "<w:rStyle/><w:b/><w:color w:val=""000000""/><w:spacing/><w:sz/>" + ...
                "<w:szCs/><w:highlight/><w:u/><w:vertAlign/><w:lang/><w:specVanish/></w:rPr>", ...
                'parsed_nondesc serialized bytes (L1) must match the frozen oracle');
        end

        % =============================================================== %
        % 2. M1 styles.xml byte-neutrality (THE parse-class-flip risk)     %
        % =============================================================== %

        function test_styles_xml_byte_identical_after_registration(testCase)
            % Regression (byte-neutrality, L1): mat2doc.Document().save() emits
            % word/styles.xml at EXACTLY 349458 B with the frozen s0001 SHA-256 --
            % byte-identical DESPITE the CT_RPr / CT_OnOff / CT_Color / ... parse-class
            % flip on its 464 w:rPr blocks. SHA-256 equality is a byte-level (L1)
            % assertion; RED on any single-byte drift through the reparsed-through-CT
            % part. (validate_P4-1a §1; Test_p1_8_skeleton_m1 owns the full 17/17 M1 sweep.)
            bytes = testCase.emitPart('styles.xml');
            testCase.verifyEqual(numel(bytes), testCase.STYLES_SIZE, ...
                sprintf('word/styles.xml must be exactly %d B after font-block registration', ...
                    testCase.STYLES_SIZE));
            testCase.verifyEqual(sha256hex(bytes), testCase.STYLES_SHA, ...
                'word/styles.xml SHA-256 must equal the frozen s0001 oracle (byte-identical L1)');
        end

        function test_styles_xml_parse_serialize_roundtrip_L1(testCase)
            % Regression (byte-neutrality, L1): parse the emitted styles.xml back
            % through mat2doc.oxml.parse_xml (instantiating hundreds of CT_RPr and its
            % CT_* children) and re-serialize -- the output must be byte-identical to
            % the input. This directly exercises the CT_RPr parse->serialize path at
            % part scale (the exact P4-1a byte risk), independently of save(). Also
            % asserts many real w:rPr blocks are present so the pin cannot pass on an
            % empty/short part. (validate_P4-1a Real-rPr round-trip.)
            inBytes = testCase.emitPart('styles.xml');
            root    = mat2doc.oxml.parse_xml(inBytes);
            outBytes = mat2doc.oxml.serialize_part_xml(root);
            testCase.verifyEqual(uint8(outBytes(:)'), uint8(inBytes(:)'), ...
                'styles.xml parse->serialize must be byte-identical (CT_RPr path is byte-neutral)');
            nRpr = numel(root.xpath('.//w:rPr'));
            % Looser-than-byte assertion (justified): the EXACT block count is an
            % implementation detail of the shipped default template; pinning "many
            % CT_RPr blocks were actually parsed" is what matters for this leg.
            testCase.verifyGreaterThan(nRpr, 100, ...
                'styles.xml must parse many real w:rPr blocks (CT_RPr path exercised)');
        end

        % =============================================================== %
        % 3. _new_color override (font.py 149-151)                        %
        % =============================================================== %

        function test_new_color_seeds_black_and_idempotent(testCase)
            % Nominal / Regression (s0020 new_color): CT_RPr._new_color WINS over the
            % generic engine creator -- get_or_add_color() on a fresh rPr seeds
            % <w:color w:val="000000"/> (localname color, hex 000000, RGB black), the
            % serialized rPr is byte-exact, and get_or_add_color() is idempotent
            % (returns the SAME live child on re-call).
            e = buildRPr("", strings(1, 0));
            color = e.get_or_add_color();
            testCase.verifyEqual(string(color.local_part), "color");
            testCase.verifyEqual(color.val.str_(), "000000", ...
                '_new_color must seed hex 000000 (RGB black)');
            testCase.verifyEqual( ...
                [double(color.val.r) double(color.val.g) double(color.val.b)], [0 0 0]);
            testCase.verifyEqual(ser(e), decl() + newline + ...
                "<w:rPr xmlns:w=""" + testCase.W + """><w:color w:val=""000000""/></w:rPr>", ...
                '_new_color seeded rPr serialized bytes (L1) must match the frozen oracle');
            testCase.verifyTrue(e.get_or_add_color() == color, ...
                'get_or_add_color must be idempotent (same live child handle)');
        end

        % =============================================================== %
        % 4. Child CT_* round-trips                                       %
        % =============================================================== %

        function test_children_color_and_ct_string(testCase)
            % Regression (s0020 children): CT_Color val(RGB 2F2F80)+themeColor(accent1
            % -> name ACCENT_1); absent themeColor -> []. CT_String.new builds a
            % w:rStyle with @w:val (exact class CT_String). NOTE: CT_String is
            % FULLY-QUALIFIED inline -- a specific (non-".*") import resolves at
            % suite-CREATION parse time, before PathFixture, and would error
            % MATLAB:undefinedVarOrClass on a cold run (project memory lesson).
            c = el("w:color");
            c.val = mat2doc.shared.RGBColor(hex2dec("2F"), hex2dec("2F"), hex2dec("80"));
            c.themeColor = mat2doc.enum.dml.MSO_THEME_COLOR.ACCENT_1;
            testCase.verifyEqual(ser(c), decl() + newline + ...
                "<w:color xmlns:w=""" + testCase.W + """ w:val=""2F2F80"" w:themeColor=""accent1""/>", ...
                'CT_Color val+themeColor serialized bytes (L1)');
            testCase.verifyEqual(c.val.str_(), "2F2F80");
            testCase.verifyEqual(string(c.themeColor), "ACCENT_1");
            testCase.verifyTrue(isequal(el("w:color").themeColor, []), ...
                'absent w:themeColor must read [] (-> "none")');

            rs = mat2doc.oxml.shared.CT_String.new("w:rStyle", "Emphasis");
            testCase.verifyEqual(class(rs), 'mat2doc.oxml.shared.CT_String', ...
                'CT_String.new must return a CT_String');
            testCase.verifyEqual(rs.val, "Emphasis");
            testCase.verifyEqual(ser(rs), decl() + newline + ...
                "<w:rStyle xmlns:w=""" + testCase.W + """ w:val=""Emphasis""/>", ...
                'CT_String.new serialized bytes (L1)');
        end

        function test_children_fonts_highlight_hps_underline_vertalign(testCase)
            % Regression (s0020 children): CT_Fonts ascii/hAnsi (+absent->[]);
            % CT_Highlight val (yellow, enum name YELLOW); CT_HpsMeasure val (Pt 12 ->
            % "24", 152400 EMU); CT_Underline val (single, enum SINGLE, +absent->[]);
            % CT_VerticalAlignRun val (superscript). Each serialized byte-exact.
            f = el("w:rFonts"); f.ascii = "Arial"; f.hAnsi = "Calibri";
            testCase.verifyEqual(ser(f), decl() + newline + ...
                "<w:rFonts xmlns:w=""" + testCase.W + """ w:ascii=""Arial"" w:hAnsi=""Calibri""/>");
            testCase.verifyEqual(f.ascii, "Arial");
            testCase.verifyEqual(f.hAnsi, "Calibri");
            testCase.verifyTrue(isequal(el("w:rFonts").ascii, []), 'absent w:ascii -> []');

            h = el("w:highlight"); h.val = mat2doc.enum.text.WD_COLOR_INDEX.YELLOW;
            testCase.verifyEqual(ser(h), decl() + newline + ...
                "<w:highlight xmlns:w=""" + testCase.W + """ w:val=""yellow""/>");
            testCase.verifyEqual(string(h.val), "YELLOW");

            s = el("w:sz"); s.val = mat2doc.shared.Pt(12);
            testCase.verifyEqual(ser(s), decl() + newline + ...
                "<w:sz xmlns:w=""" + testCase.W + """ w:val=""24""/>", ...
                'CT_HpsMeasure Pt(12) -> half-points "24"');
            testCase.verifyEqual(round(double(s.val)), 152400, 'Pt(12) == 152400 EMU');

            u = el("w:u"); u.val = mat2doc.enum.text.WD_UNDERLINE.SINGLE;
            testCase.verifyEqual(ser(u), decl() + newline + ...
                "<w:u xmlns:w=""" + testCase.W + """ w:val=""single""/>");
            testCase.verifyEqual(string(u.val), "SINGLE");
            testCase.verifyTrue(isequal(el("w:u").val, []), 'absent w:u/@val -> []');

            v = el("w:vertAlign"); v.val = "superscript";
            testCase.verifyEqual(ser(v), decl() + newline + ...
                "<w:vertAlign xmlns:w=""" + testCase.W + """ w:val=""superscript""/>");
            testCase.verifyEqual(v.val, "superscript");
        end

        % =============================================================== %
        % 5. CT_RPr @property setters                                     %
        % =============================================================== %

        function test_ct_rpr_sz_val_setter(testCase)
            % Regression (s0020 ct_rpr_setters sz_*): sz_val add (nothing -> Pt 12 ->
            % w:sz w:val="24"); change (24 -> Pt 18 -> "36"); remove ([] removes w:sz).
            e = buildRPr("", strings(1,0));   e.sz_val = mat2doc.shared.Pt(12);
            testCase.verifyEqual(ser(e), rprWrap(testCase, "<w:sz w:val=""24""/>"), 'sz_nothing');
            e = buildRPr("<w:sz w:val=""24""/>", strings(1,0)); e.sz_val = mat2doc.shared.Pt(18);
            testCase.verifyEqual(ser(e), rprWrap(testCase, "<w:sz w:val=""36""/>"), 'sz_change');
            e = buildRPr("<w:sz w:val=""36""/>", strings(1,0)); e.sz_val = [];
            testCase.verifyEqual(ser(e), rprWrap(testCase, ""), 'sz_remove');
        end

        function test_ct_rpr_subscript_superscript_truth_table(testCase)
            % Regression (s0020 ct_rpr_setters subscript_*/superscript_*): the
            % vertAlign truth table -- True adds sub/super; False/None on empty is a
            % no-op; subscript=False REMOVES only a matching subscript (keeps a
            % superscript) and vice-versa. Pinned byte-exact.
            e = buildRPr("", strings(1,0));  e.subscript = true;
            testCase.verifyEqual(ser(e), rprWrap(testCase, "<w:vertAlign w:val=""subscript""/>"), 'subscript_true');
            e = buildRPr("", strings(1,0));  e.subscript = false;
            testCase.verifyEqual(ser(e), rprWrap(testCase, ""), 'subscript_false_on_empty (no-op)');
            e = buildRPr("", strings(1,0));  e.subscript = [];
            testCase.verifyEqual(ser(e), rprWrap(testCase, ""), 'subscript_none_on_empty (no-op)');
            e = buildRPr("<w:vertAlign w:val=""subscript""/>", strings(1,0));  e.subscript = false;
            testCase.verifyEqual(ser(e), rprWrap(testCase, ""), 'subscript_false removes matching subscript');
            e = buildRPr("<w:vertAlign w:val=""superscript""/>", strings(1,0)); e.subscript = false;
            testCase.verifyEqual(ser(e), rprWrap(testCase, "<w:vertAlign w:val=""superscript""/>"), ...
                'subscript_false must KEEP a superscript');
            e = buildRPr("", strings(1,0));  e.superscript = true;
            testCase.verifyEqual(ser(e), rprWrap(testCase, "<w:vertAlign w:val=""superscript""/>"), 'superscript_true');
            e = buildRPr("<w:vertAlign w:val=""subscript""/>", strings(1,0)); e.superscript = false;
            testCase.verifyEqual(ser(e), rprWrap(testCase, "<w:vertAlign w:val=""subscript""/>"), ...
                'superscript_false must KEEP a subscript');
        end

        function test_ct_rpr_style_and_rfonts_asymmetry(testCase)
            % Regression (s0020 ct_rpr_setters style_*/rFonts_*): style add(_add_rStyle
            % on val), change, remove. rFonts_ascii add/change/None-removes-w:rFonts,
            % vs the ASYMMETRIC rFonts_hAnsi=None which KEEPS w:rFonts (removes only
            % @w:hAnsi) and is a no-op when w:rFonts is absent. Pinned byte-exact.
            e = buildRPr("", strings(1,0));  e.style = "Heading1";
            testCase.verifyEqual(ser(e), rprWrap(testCase, "<w:rStyle w:val=""Heading1""/>"), 'style_add');
            e = buildRPr("<w:rStyle w:val=""A""/>", strings(1,0));  e.style = "B";
            testCase.verifyEqual(ser(e), rprWrap(testCase, "<w:rStyle w:val=""B""/>"), 'style_change');
            e = buildRPr("<w:rStyle w:val=""A""/>", strings(1,0));  e.style = [];
            testCase.verifyEqual(ser(e), rprWrap(testCase, ""), 'style_remove');

            e = buildRPr("", strings(1,0));  e.rFonts_ascii = "Foo";
            testCase.verifyEqual(ser(e), rprWrap(testCase, "<w:rFonts w:ascii=""Foo""/>"), 'rFonts_ascii_add');
            e = buildRPr("<w:rFonts w:hAnsi=""Foo""/>", strings(1,0));  e.rFonts_ascii = "Bar";
            testCase.verifyEqual(ser(e), rprWrap(testCase, "<w:rFonts w:hAnsi=""Foo"" w:ascii=""Bar""/>"), 'rFonts_ascii_change');
            e = buildRPr("<w:rFonts w:ascii=""X""/>", strings(1,0));  e.rFonts_ascii = [];
            testCase.verifyEqual(ser(e), rprWrap(testCase, ""), 'rFonts_ascii=None REMOVES w:rFonts');
            e = buildRPr("<w:rFonts w:ascii=""X"" w:hAnsi=""Y""/>", strings(1,0));  e.rFonts_hAnsi = [];
            testCase.verifyEqual(ser(e), rprWrap(testCase, "<w:rFonts w:ascii=""X""/>"), ...
                'rFonts_hAnsi=None KEEPS w:rFonts (ASYMMETRY -- removes only @w:hAnsi)');
            e = buildRPr("", strings(1,0));  e.rFonts_hAnsi = [];
            testCase.verifyEqual(ser(e), rprWrap(testCase, ""), 'rFonts_hAnsi=None on empty is a no-op');
        end

        function test_ct_rpr_u_val_remove_all_then_add_and_highlight(testCase)
            % Regression (s0020 ct_rpr_setters u_val_*/highlight_*): u_val add; remove
            % ([]); the remove-ALL-then-add branch collapses a DUPLICATED <w:u> input
            % into a single <w:u w:val="wave">. highlight add/remove. Pinned byte-exact.
            e = buildRPr("", strings(1,0));  e.u_val = mat2doc.enum.text.WD_UNDERLINE.DOUBLE;
            testCase.verifyEqual(ser(e), rprWrap(testCase, "<w:u w:val=""double""/>"), 'u_val_add');
            e = buildRPr("<w:u w:val=""single""/>", strings(1,0));  e.u_val = [];
            testCase.verifyEqual(ser(e), rprWrap(testCase, ""), 'u_val_remove');
            e = buildRPr("<w:u w:val=""single""/><w:u w:val=""double""/>", strings(1,0));
            e.u_val = mat2doc.enum.text.WD_UNDERLINE.WAVY;
            testCase.verifyEqual(ser(e), rprWrap(testCase, "<w:u w:val=""wave""/>"), ...
                'u_val setter removes ALL w:u then adds one (WAVY -> "wave")');
            e = buildRPr("", strings(1,0));  e.highlight_val = mat2doc.enum.text.WD_COLOR_INDEX.YELLOW;
            testCase.verifyEqual(ser(e), rprWrap(testCase, "<w:highlight w:val=""yellow""/>"), 'highlight_add');
            e = buildRPr("<w:highlight w:val=""yellow""/>", strings(1,0));  e.highlight_val = [];
            testCase.verifyEqual(ser(e), rprWrap(testCase, ""), 'highlight_remove');
        end

        % =============================================================== %
        % 6. CT_OnOff tri-state (D-delta-1)                               %
        % =============================================================== %

        function test_ct_onoff_tristate_and_from_xml_tokens(testCase)
            % Regression (s0020 ct_onoff, D-delta-1): CT_OnOff.val tri-state --
            %   absent @w:val get -> True (non-None default);
            %   set False        -> w:val="0";
            %   set True(==default) -> @w:val REMOVED (bare <w:b/>);
            %   set None/[]       -> @w:val REMOVED;
            % and the six from_xml tokens: 1/true/on -> True, 0/false/off -> False.
            b = el("w:b");
            testCase.verifyTrue(b.val, 'absent @w:val must read True (D-delta-1 non-None default)');

            b2 = el("w:b"); b2.val = false;
            testCase.verifyEqual(ser(b2), decl() + newline + ...
                "<w:b xmlns:w=""" + testCase.W + """ w:val=""0""/>", 'set False -> w:val="0"');
            b3 = el("w:b"); b3.val = true;
            testCase.verifyEqual(ser(b3), decl() + newline + ...
                "<w:b xmlns:w=""" + testCase.W + """/>", 'set True(==default) REMOVES @w:val');
            b4 = rparse("<w:b xmlns:w=""" + testCase.W + """ w:val=""1""/>"); b4.val = [];
            testCase.verifyEqual(ser(b4), decl() + newline + ...
                "<w:b xmlns:w=""" + testCase.W + """/>", 'set None REMOVES @w:val');

            trueToks  = ["1" "true" "on"];
            falseToks = ["0" "false" "off"];
            for t = trueToks
                bt = rparse("<w:b xmlns:w=""" + testCase.W + """ w:val=""" + t + """/>");
                testCase.verifyTrue(bt.val, sprintf('from_xml token "%s" -> True', t));
            end
            for t = falseToks
                bt = rparse("<w:b xmlns:w=""" + testCase.W + """ w:val=""" + t + """/>");
                testCase.verifyFalse(bt.val, sprintf('from_xml token "%s" -> False', t));
            end
        end

        % =============================================================== %
        % 7. _set_bool_val battery + getter reads (vs frozen oracle)       %
        % =============================================================== %

        function test_bool_setters_full_battery(testCase)
            % Regression (s0020 bool_setters): the 19-case set_bool_val_ battery
            % (nothing/default/True/False starts x True/False/None targets across
            % many descriptor tags) -- each serialized rPr compared byte-for-byte to
            % the frozen oracle. Freezes both the CT_OnOff removal rules AND the H11
            % placement of the (re)added bool child.
            oracle = loadOracle();
            BS = { ...
                "",                        "caps",      true,  "nothing_to_True"; ...
                "",                        "b",         false, "nothing_to_False"; ...
                "",                        "i",         [],    "nothing_to_None"; ...
                "<w:cs/>",                 "cs",        true,  "default_to_True"; ...
                "<w:bCs/>",                "bCs",       false, "default_to_False"; ...
                "<w:iCs/>",                "iCs",       [],    "default_to_None"; ...
                "<w:dstrike w:val=""1""/>", "dstrike",  true,  "True_to_True"; ...
                "<w:emboss w:val=""on""/>", "emboss",   false, "True_to_False"; ...
                "<w:vanish w:val=""1""/>",  "vanish",   [],    "True_to_None"; ...
                "<w:i w:val=""false""/>",   "i",        true,  "False_to_True"; ...
                "<w:imprint w:val=""0""/>", "imprint",  false, "False_to_False"; ...
                "<w:oMath w:val=""off""/>", "oMath",    [],    "False_to_None"; ...
                "<w:noProof w:val=""1""/>", "noProof",  false, "mix_noProof"; ...
                "",                        "outline",   true,  "mix_outline"; ...
                "<w:rtl w:val=""true""/>",  "rtl",      false, "mix_rtl"; ...
                "<w:smallCaps/>",          "smallCaps", false, "mix_smallCaps"; ...
                "<w:snapToGrid/>",         "snapToGrid",true,  "mix_snapToGrid"; ...
                "<w:strike w:val=""foo""/>","strike",   true,  "mix_strike"; ...
                "<w:webHidden/>",          "webHidden", false, "mix_webHidden"};
            for k = 1:size(BS,1)
                e = buildRPr(BS{k,1}, strings(1,0));
                e.set_bool_val_(BS{k,2}, BS{k,3});
                testCase.verifyEqual(ser(e), string(oracle.bool_setters.(BS{k,4})), ...
                    sprintf('bool_setters.%s serialized bytes (L1)', BS{k,4}));
            end
        end

        function test_reads_getters_full(testCase)
            % Regression (s0020 reads): getter reads -- get_bool_val_ tri-state across
            % absent/bare/on/off/1/0/true/false; subscript/superscript per vertAlign;
            % sz_val (present EMU / -1 sentinel absent); rFonts_ascii; style. Compared
            % to the frozen oracle (tri-state -> "true"/"false"/"none", EMU int).
            oracle = loadOracle();

            BR = { ...
                "",                          "caps",     "absent"; ...
                "<w:caps/>",                 "caps",     "present_bare"; ...
                "<w:caps w:val=""on""/>",    "caps",     "on"; ...
                "<w:caps w:val=""off""/>",   "caps",     "off"; ...
                "<w:b w:val=""1""/>",        "b",        "one"; ...
                "<w:i w:val=""0""/>",        "i",        "zero"; ...
                "<w:cs w:val=""true""/>",    "cs",       "true"; ...
                "<w:bCs w:val=""false""/>",  "bCs",      "false"; ...
                "<w:iCs w:val=""on""/>",     "iCs",      "on"; ...
                "<w:dstrike w:val=""off""/>","dstrike",  "off"; ...
                "<w:emboss w:val=""1""/>",   "emboss",   "one"; ...
                "<w:vanish w:val=""0""/>",   "vanish",   "zero"; ...
                "<w:oMath w:val=""on""/>",   "oMath",    "on"; ...
                "<w:strike w:val=""1""/>",   "strike",   "one"; ...
                "<w:webHidden w:val=""0""/>","webHidden","zero"};
            for k = 1:size(BR,1)
                got = tri(buildRPr(BR{k,1}, strings(1,0)).get_bool_val_(BR{k,2}));
                testCase.verifyEqual(got, string(oracle.reads.bool_states.(BR{k,3})), ...
                    sprintf('reads.bool_states.%s', BR{k,3}));
            end

            % subscript / superscript per vertAlign
            testCase.verifyEqual(tri(buildRPr("", strings(1,0)).subscript), ...
                string(oracle.reads.subscript.absent), 'subscript absent');
            testCase.verifyEqual(tri(buildRPr("<w:vertAlign w:val=""baseline""/>", strings(1,0)).subscript), ...
                string(oracle.reads.subscript.baseline), 'subscript baseline');
            testCase.verifyEqual(tri(buildRPr("<w:vertAlign w:val=""subscript""/>", strings(1,0)).subscript), ...
                string(oracle.reads.subscript.subscript), 'subscript subscript');
            testCase.verifyEqual(tri(buildRPr("<w:vertAlign w:val=""superscript""/>", strings(1,0)).subscript), ...
                string(oracle.reads.subscript.superscript), 'subscript superscript');
            testCase.verifyEqual(tri(buildRPr("<w:vertAlign w:val=""superscript""/>", strings(1,0)).superscript), ...
                string(oracle.reads.superscript.superscript), 'superscript superscript');

            % sz_val EMU / -1 sentinel
            testCase.verifyEqual(emu(buildRPr("", strings(1,0)).sz_val), ...
                double(oracle.reads.sz_val.absent), 'sz_val absent -> -1 sentinel');
            testCase.verifyEqual(emu(buildRPr("<w:sz w:val=""28""/>", strings(1,0)).sz_val), ...
                double(oracle.reads.sz_val.present), 'sz_val present EMU');

            % rFonts_ascii / style
            testCase.verifyEqual(sn(buildRPr("<w:rFonts w:ascii=""Arial""/>", strings(1,0)).rFonts_ascii), ...
                string(oracle.reads.rFonts_ascii.present), 'rFonts_ascii present');
            testCase.verifyEqual(sn(buildRPr("", strings(1,0)).rFonts_ascii), ...
                string(oracle.reads.rFonts_ascii.absent_no_rFonts), 'rFonts_ascii absent (no rFonts)');
            testCase.verifyEqual(sn(buildRPr("<w:rStyle w:val=""Emphasis""/>", strings(1,0)).style), ...
                string(oracle.reads.style.present), 'style present');
            testCase.verifyEqual(sn(buildRPr("", strings(1,0)).style), ...
                string(oracle.reads.style.absent), 'style absent');
        end

        % =============================================================== %
        % 8. EQUIVALENCE -- full s0020 battery vs the frozen oracle        %
        % =============================================================== %

        function test_equivalence_full_battery_vs_frozen_oracle(testCase)
            % Equivalence: replay the ENTIRE s0020 battery (runProbes -- the .m twin's
            % body VERBATIM: h11 / new_color / children / ct_rpr_setters / bool_setters
            % / reads / ct_onoff) and flatten-compare EVERY leaf to the frozen
            % python-docx 1.2.0 oracle copied into data\s0020_probe_oracle.json. Gate-3
            % found ZERO divergences, so every leaf must be byte/value-identical. This
            % single method is the class's Equivalence leg, tying the suite to the
            % Gate-3 reference output.
            port   = runProbes();
            oracle = loadOracle();
            pMap = containers.Map('KeyType','char','ValueType','char');
            oMap = containers.Map('KeyType','char','ValueType','char');
            flattenLeaves(port,   '', pMap);
            flattenLeaves(oracle, '', oMap);

            pKeys = sort(pMap.keys());
            oKeys = sort(oMap.keys());
            testCase.verifyEqual(pKeys, oKeys, ...
                'the replayed battery and the frozen oracle must have identical leaf keys');
            % Non-trivial size guard (the battery is large -- guards a silent-empty replay).
            testCase.verifyGreaterThan(numel(oKeys), 90, ...
                'the flattened oracle must expose the full battery of leaves');
            for i = 1:numel(oKeys)
                k = oKeys{i};
                testCase.verifyTrue(isKey(pMap, k), sprintf('port is missing leaf %s', k));
                testCase.verifyEqual(pMap(k), oMap(k), ...
                    sprintf('leaf %s must be byte/value-identical to the frozen oracle', k));
            end
        end

    end

    % ===================== instance helpers ============================ %
    methods (Access = private)
        function bytes = emitPart(~, partLeaf)
            % mat2doc.Document().save() to a temp .docx, extract word/<partLeaf>,
            % return its raw bytes. Base-MATLAB unzip (no toolbox) into a temp dir,
            % both cleaned up on exit. (Idiom copied from Test_p2_3_document_shell.m;
            % tempname paths are absolute so no cwd handling is needed.)
            d = mat2doc.Document();
            tmp = [tempname '.docx'];
            cleanTmp = onCleanup(@() deleteIfExists(tmp)); %#ok<NASGU>
            d.save(tmp);
            exdir = tempname;
            cleanEx = onCleanup(@() rmdirIfExists(exdir)); %#ok<NASGU>
            unzip(tmp, exdir);
            bytes = readBytes(fullfile(exdir, 'word', partLeaf));
        end
    end
end

% ===================== file-local helpers ============================== %

function s = decl()
    s = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>";
end

function s = W_()
    s = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
end

function e = rparse(xml)
    e = mat2doc.oxml.parse_xml(uint8(unicode2native(char(xml), "UTF-8")));
end

function e = rpr(inner)
    % A loose CT_RPr with optional inner XML (mirrors the s0020 py/m rpr() helper).
    xml = "<w:rPr xmlns:w=""" + W_() + """>" + string(inner) + "</w:rPr>";
    e = rparse(xml);
end

function e = el(nsptag)
    % A loose child element carrying its own xmlns:w (mirrors s0020 el()/py elm()).
    xml = "<" + string(nsptag) + " xmlns:w=""" + W_() + """/>";
    e = rparse(xml);
end

function e = buildRPr(inner, order)
    % Build a loose CT_RPr from `inner` and apply get_or_add_<tag> for each name in
    % `order` (string row), driving the H11 re-sort. Used by the dedicated pins.
    e = rpr(inner);
    for k = 1:numel(order)
        feval("get_or_add_" + order(k), e);
    end
end

function s = ser(e)
    % Decode serialize_part_xml bytes to a string (ASCII content -> string-equality
    % is a byte-identical L1 assertion).
    s = string(native2unicode(mat2doc.oxml.serialize_part_xml(e), "UTF-8"));
end

function s = rprWrap(testCase, body)
    % decl + newline + <w:rPr xmlns:w="...">BODY</w:rPr>. An EMPTY body serializes
    % self-closing (<w:rPr .../>), matching lxml/serialize_part_xml and the frozen
    % oracle (e.g. sz_remove) -- NOT an empty-element pair.
    if strlength(string(body)) == 0
        s = decl() + newline + "<w:rPr xmlns:w=""" + testCase.W + """/>";
    else
        s = decl() + newline + "<w:rPr xmlns:w=""" + testCase.W + """>" + string(body) + "</w:rPr>";
    end
end

function names = childLocalnames(e)
    kids = e.xpath("./*");
    names = strings(1, numel(kids));
    for k = 1:numel(kids)
        names(k) = string(kids(k).local_part);
    end
end

function s = tri(v)
    % tri-state (True/False/[]) -> "true"/"false"/"none" (matches s0020 tri()).
    if isequal(v, [])
        s = "none";
    elseif v
        s = "true";
    else
        s = "false";
    end
end

function n = emu(v)
    if isequal(v, [])
        n = -1;
    else
        n = round(double(v));
    end
end

function out = sn(v)
    % [] -> "none", else string(v) (matches s0020 sn()).
    if isequal(v, [])
        out = "none";
    else
        out = string(v);
    end
end

function o = loadOracle()
    % Read the co-located frozen oracle in BINARY mode (no CRLF translation) and
    % decode UTF-8 -> struct. jsondecode is whitespace/line-ending agnostic, so no
    % `* binary` .gitattributes pin is needed (value-based fixture, s0017 precedent).
    here = fileparts(mfilename('fullpath'));
    p = fullfile(here, 'data', 's0020_probe_oracle.json');
    fid = fopen(p, 'r', 'n');
    if fid < 0
        error('mat2doc:test:openFail', 'cannot open frozen oracle %s', p);
    end
    raw = fread(fid, Inf, '*uint8')';
    fclose(fid);
    o = jsondecode(native2unicode(raw, 'UTF-8'));
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

function h = sha256hex(bytes)
    % Lowercase hex SHA-256 of a uint8 vector (base MATLAB / Java MessageDigest).
    md  = java.security.MessageDigest.getInstance('SHA-256');
    dig = md.digest(typecast(uint8(bytes(:)'), 'int8'));
    b   = typecast(dig, 'uint8');
    h   = lower(string(sprintf('%02x', b)));
end

function flattenLeaves(s, prefix, map)
    % Recursively flatten a (possibly nested) struct / cell / array into
    % map(dotted.path) -> canonical string, so a MATLAB-built probe struct and the
    % jsondecode'd oracle compare leaf-by-leaf regardless of container types.
    if isstruct(s)
        fn = fieldnames(s);
        for i = 1:numel(fn)
            if isempty(prefix)
                child = fn{i};
            else
                child = [prefix '.' fn{i}];
            end
            flattenLeaves(s.(fn{i}), child, map);
        end
    elseif iscell(s)
        parts = strings(1, numel(s));
        for i = 1:numel(s)
            parts(i) = canonScalar(s{i});
        end
        map(prefix) = char(join(parts, "|")); %#ok<NASGU>
    elseif (isstring(s) && ~isscalar(s))
        map(prefix) = char(join(s(:).', "|")); %#ok<NASGU>
    elseif isnumeric(s) && ~isscalar(s)
        map(prefix) = char(join(string(s(:).'), ",")); %#ok<NASGU>
    else
        map(prefix) = char(canonScalar(s)); %#ok<NASGU>
    end
end

function s = canonScalar(v)
    if islogical(v)
        if v, s = "true"; else, s = "false"; end
    elseif isnumeric(v)
        s = string(num2str(v));
    else
        s = string(v);
    end
end

function P = runProbes()
    % Replay the s0020 probe sequence (the .m twin's P-struct body, VERBATIM
    % tags/inputs/order) and return the nested struct of tagged canonical values.
    % Embedded here so the Equivalence leg is self-contained (the validation-folder
    % scenario is NOT on the toolbox path and must not be a dependency). Mirrors
    % validation\mat2doc\scenarios\s0020_p4_1a_oxml_font.m lines 22-207.
    P = struct();

    % ===================== H11 scrambled add-order =======================
    ALL27 = ["rStyle","rFonts","b","bCs","i","iCs","caps","smallCaps","strike", ...
        "dstrike","outline","shadow","emboss","imprint","noProof","snapToGrid", ...
        "vanish","webHidden","color","sz","highlight","u","vertAlign","rtl", ...
        "cs","specVanish","oMath"];
    SHUFFLE27 = ["vertAlign","b","oMath","color","rStyle","highlight","i","sz","u", ...
        "specVanish","caps","webHidden","rFonts","strike","bCs","rtl","vanish", ...
        "iCs","cs","smallCaps","dstrike","noProof","outline","shadow","emboss", ...
        "imprint","snapToGrid"];
    SCRAMBLE8 = ["u","b","color","sz","rStyle","vertAlign","highlight","i"];

    P.h11 = struct( ...
        "loose_scramble8",  h11_case(SCRAMBLE8, ""), ...
        "loose_reversed27", h11_case(flip(ALL27), ""), ...
        "loose_shuffle27",  h11_case(SHUFFLE27, ""), ...
        "parsed_nondesc",   h11_case( ...
            ["specVanish","vertAlign","u","highlight","sz","color","b","rStyle"], ...
            "<w:spacing/><w:szCs/><w:lang/>"));

    % ===================== _new_color override ===========================
    e = rpr("");
    color = e.get_or_add_color();
    val = color.val;
    P.new_color = struct( ...
        "localname", string(color.local_part), ...
        "hex", val.str_(), ...
        "rgb", [double(val.r) double(val.g) double(val.b)], ...
        "serialized_rPr", ser(e), ...
        "get_or_add_is_idempotent", e.get_or_add_color() == color);

    % ===================== children CT_* round-trips =====================
    ch = struct();
    c = el("w:color");
    c.val = mat2doc.shared.RGBColor(hex2dec("2F"), hex2dec("2F"), hex2dec("80"));
    c.themeColor = mat2doc.enum.dml.MSO_THEME_COLOR.ACCENT_1;
    ch.color_set = struct("serialized", ser(c), "val_hex", c.val.str_(), ...
        "themeColor_name", string(c.themeColor));
    ch.color_absent_themeColor = sn(el("w:color").themeColor);

    f = el("w:rFonts"); f.ascii = "Arial"; f.hAnsi = "Calibri";
    ch.rFonts_set = struct("serialized", ser(f), "ascii", f.ascii, "hAnsi", f.hAnsi);
    ch.rFonts_ascii_absent = sn(el("w:rFonts").ascii);

    h = el("w:highlight"); h.val = mat2doc.enum.text.WD_COLOR_INDEX.YELLOW;
    ch.highlight_set = struct("serialized", ser(h), "val_name", string(h.val));

    s = el("w:sz"); s.val = mat2doc.shared.Pt(12);
    ch.sz_set = struct("serialized", ser(s), "val_emu", emu(s.val));

    u = el("w:u"); u.val = mat2doc.enum.text.WD_UNDERLINE.SINGLE;
    ch.u_set = struct("serialized", ser(u), "val_name", string(u.val));
    ch.u_absent = sn(el("w:u").val);

    v = el("w:vertAlign"); v.val = "superscript";
    ch.vertAlign_set = struct("serialized", ser(v), "val", v.val);

    rs = mat2doc.oxml.shared.CT_String.new("w:rStyle", "Emphasis");
    ch.rStyle_new = struct("serialized", ser(rs), "val", rs.val, "class", barecls(rs));
    P.children = ch;

    % ===================== CT_RPr @property setters ======================
    st = struct();
    e = rpr("");                     e.sz_val = mat2doc.shared.Pt(12); st.sz_nothing = ser(e);
    e = rpr("<w:sz w:val=""24""/>"); e.sz_val = mat2doc.shared.Pt(18); st.sz_change = ser(e);
    e = rpr("<w:sz w:val=""36""/>"); e.sz_val = [];                    st.sz_remove = ser(e);

    e = rpr("");  e.subscript = true;   st.subscript_true = ser(e);
    e = rpr("");  e.subscript = false;  st.subscript_false_on_empty = ser(e);
    e = rpr("");  e.subscript = [];     st.subscript_none_on_empty = ser(e);
    e = rpr("<w:vertAlign w:val=""subscript""/>");   e.subscript = false; st.subscript_false_removes = ser(e);
    e = rpr("<w:vertAlign w:val=""superscript""/>"); e.subscript = false; st.subscript_false_keeps_super = ser(e);
    e = rpr("");  e.superscript = true; st.superscript_true = ser(e);
    e = rpr("<w:vertAlign w:val=""subscript""/>");   e.superscript = false; st.superscript_false_keeps_sub = ser(e);

    e = rpr("");  e.style = "Heading1"; st.style_add = ser(e);
    e = rpr("<w:rStyle w:val=""A""/>"); e.style = "B"; st.style_change = ser(e);
    e = rpr("<w:rStyle w:val=""A""/>"); e.style = [];  st.style_remove = ser(e);

    e = rpr("");  e.rFonts_ascii = "Foo"; st.rFonts_ascii_add = ser(e);
    e = rpr("<w:rFonts w:hAnsi=""Foo""/>"); e.rFonts_ascii = "Bar"; st.rFonts_ascii_change = ser(e);
    e = rpr("<w:rFonts w:ascii=""X""/>");   e.rFonts_ascii = [];     st.rFonts_ascii_none_removes_rFonts = ser(e);
    e = rpr("<w:rFonts w:ascii=""X"" w:hAnsi=""Y""/>"); e.rFonts_hAnsi = []; st.rFonts_hAnsi_none_keeps_rFonts = ser(e);
    e = rpr("");  e.rFonts_hAnsi = [];   st.rFonts_hAnsi_none_on_empty_noop = ser(e);

    e = rpr("");  e.u_val = mat2doc.enum.text.WD_UNDERLINE.DOUBLE; st.u_val_add = ser(e);
    e = rpr("<w:u w:val=""single""/>"); e.u_val = []; st.u_val_remove = ser(e);
    e = rpr("<w:u w:val=""single""/><w:u w:val=""double""/>");
    e.u_val = mat2doc.enum.text.WD_UNDERLINE.WAVY; st.u_val_remove_all_then_add = ser(e);

    e = rpr("");  e.highlight_val = mat2doc.enum.text.WD_COLOR_INDEX.YELLOW; st.highlight_add = ser(e);
    e = rpr("<w:highlight w:val=""yellow""/>"); e.highlight_val = []; st.highlight_remove = ser(e);
    P.ct_rpr_setters = st;

    % ===================== _set_bool_val battery =========================
    BS = { ...
        "",                        "caps", true,  "nothing_to_True"; ...
        "",                        "b",    false, "nothing_to_False"; ...
        "",                        "i",    [],    "nothing_to_None"; ...
        "<w:cs/>",                 "cs",   true,  "default_to_True"; ...
        "<w:bCs/>",                "bCs",  false, "default_to_False"; ...
        "<w:iCs/>",                "iCs",  [],    "default_to_None"; ...
        "<w:dstrike w:val=""1""/>","dstrike",true,"True_to_True"; ...
        "<w:emboss w:val=""on""/>","emboss",false,"True_to_False"; ...
        "<w:vanish w:val=""1""/>", "vanish",[],   "True_to_None"; ...
        "<w:i w:val=""false""/>",  "i",    true,  "False_to_True"; ...
        "<w:imprint w:val=""0""/>","imprint",false,"False_to_False"; ...
        "<w:oMath w:val=""off""/>","oMath",[],    "False_to_None"; ...
        "<w:noProof w:val=""1""/>","noProof",false,"mix_noProof"; ...
        "",                        "outline",true,"mix_outline"; ...
        "<w:rtl w:val=""true""/>", "rtl",  false, "mix_rtl"; ...
        "<w:smallCaps/>",          "smallCaps",false,"mix_smallCaps"; ...
        "<w:snapToGrid/>",         "snapToGrid",true,"mix_snapToGrid"; ...
        "<w:strike w:val=""foo""/>","strike",true,"mix_strike"; ...
        "<w:webHidden/>",          "webHidden",false,"mix_webHidden"};
    bs = struct();
    for k = 1:size(BS,1)
        e = rpr(BS{k,1});
        e.set_bool_val_(BS{k,2}, BS{k,3});
        bs.(BS{k,4}) = ser(e);
    end
    P.bool_setters = bs;

    % ===================== reads (getters) ===============================
    BR = { ...
        "",                        "caps", "absent"; ...
        "<w:caps/>",               "caps", "present_bare"; ...
        "<w:caps w:val=""on""/>",  "caps", "on"; ...
        "<w:caps w:val=""off""/>", "caps", "off"; ...
        "<w:b w:val=""1""/>",      "b",    "one"; ...
        "<w:i w:val=""0""/>",      "i",    "zero"; ...
        "<w:cs w:val=""true""/>",  "cs",   "true"; ...
        "<w:bCs w:val=""false""/>","bCs",  "false"; ...
        "<w:iCs w:val=""on""/>",   "iCs",  "on"; ...
        "<w:dstrike w:val=""off""/>","dstrike","off"; ...
        "<w:emboss w:val=""1""/>", "emboss","one"; ...
        "<w:vanish w:val=""0""/>", "vanish","zero"; ...
        "<w:oMath w:val=""on""/>", "oMath","on"; ...
        "<w:strike w:val=""1""/>", "strike","one"; ...
        "<w:webHidden w:val=""0""/>","webHidden","zero"};
    br = struct();
    for k = 1:size(BR,1)
        br.(BR{k,3}) = tri(rpr(BR{k,1}).get_bool_val_(BR{k,2}));
    end
    rd = struct();
    rd.bool_states = br;
    rd.subscript = struct( ...
        "absent",      tri(rpr("").subscript), ...
        "baseline",    tri(rpr("<w:vertAlign w:val=""baseline""/>").subscript), ...
        "subscript",   tri(rpr("<w:vertAlign w:val=""subscript""/>").subscript), ...
        "superscript", tri(rpr("<w:vertAlign w:val=""superscript""/>").subscript));
    rd.superscript = struct( ...
        "baseline",    tri(rpr("<w:vertAlign w:val=""baseline""/>").superscript), ...
        "subscript",   tri(rpr("<w:vertAlign w:val=""subscript""/>").superscript), ...
        "superscript", tri(rpr("<w:vertAlign w:val=""superscript""/>").superscript));
    rd.sz_val = struct( ...
        "absent",  emu(rpr("").sz_val), ...
        "present", emu(rpr("<w:sz w:val=""28""/>").sz_val));
    rd.rFonts_ascii = struct( ...
        "absent_no_rFonts", sn(rpr("").rFonts_ascii), ...
        "absent_attr",      sn(rpr("<w:rFonts/>").rFonts_ascii), ...
        "present",          sn(rpr("<w:rFonts w:ascii=""Arial""/>").rFonts_ascii));
    rd.style = struct( ...
        "absent",  sn(rpr("").style), ...
        "present", sn(rpr("<w:rStyle w:val=""Emphasis""/>").style));
    P.reads = rd;

    % ===================== CT_OnOff D-delta-1 tri-state ==================
    b = el("w:b");
    oo = struct("absent_get", tri(b.val));
    b2 = el("w:b"); b2.val = false; oo.set_false = ser(b2);
    b3 = el("w:b"); b3.val = true;  oo.set_true_removes_attr = ser(b3);
    b4 = rparse("<w:b xmlns:w=""" + W_() + """ w:val=""1""/>"); b4.val = [];
    oo.set_none_removes_attr = ser(b4);
    toks = ["1","true","on","0","false","off"];
    fx = struct();
    for k = 1:numel(toks)
        bt = rparse("<w:b xmlns:w=""" + W_() + """ w:val=""" + toks(k) + """/>");
        fx.("tok_" + toks(k)) = tri(bt.val);
    end
    oo.from_xml = fx;
    P.ct_onoff = oo;
end

function out = h11_case(order, inner)
    e = rpr(inner);
    for k = 1:numel(order)
        feval("get_or_add_" + order(k), e);
    end
    kids = e.xpath("./*");
    names = strings(1, numel(kids));
    for k = 1:numel(kids)
        names(k) = string(kids(k).local_part);
    end
    out = struct("add_order", {cellstr(order(:).')}, ...
                 "result_localnames", {cellstr(names)}, ...
                 "serialized", ser(e));
end

function name = barecls(obj)
    parts = split(string(class(obj)), ".");
    name = parts(end);
end

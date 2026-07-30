classdef Test_p4_4a_text_font < matlab.unittest.TestCase
% TEST_P4_4A_TEXT_FONT  Gate-4 permanent unit tests for Mat2Doc P4-4a
%   (src/docx/text/font.py -> +mat2doc\+text\Font and src/docx/dml/color.py ->
%   +mat2doc\+dml\ColorFormat -- the FIRST P4 API/proxy-tier WP).
%
%   Font and ColorFormat are ElementProxy proxies that read/write the already
%   byte-validated CT_RPr / CT_Color tree (P4-1a). They add NO register_element_cls
%   row, NO oxml class and NO serialization-path change -> API/proxy tier: no bytes,
%   no registry, no M1 risk (validate_P4-4a §2 confirms M1 17/17 unchanged, so this
%   class does NOT re-pin styles.xml/document.xml -- Test_p1_8_skeleton_m1 owns the
%   full 17/17 M1 sweep). This class permanently FREEZES the P4-4a BEHAVIORAL surface
%   -- the API-tier equivalence pin -- byte/value-identical to python-docx 1.2.0.
%
%   The guarantees frozen here (each verified byte/value-identical at Gate-3,
%   probe_diff MATCH 228/228, zero divergences, zero new D-numbers):
%
%     (FONT bool tri-state -- the API-tier equivalence pin) the 20 boolean
%     properties (all_caps..web_hidden) over the full tri-state: absent -> None ([]);
%     True -> bare <w:tag/> + get True; False -> <w:tag w:val="0"/> + get False;
%     None -> tag removed (<w:rPr/>) + get None. Serialized r.xml pinned as UPPERCASE
%     serhex vs the frozen s0024 oracle (a byte pin -- masks nothing, includes the
%     <?xml ... standalone='yes'?> prolog) PLUS hard-coded XML-string regression pins.
%
%     (SIZE) Pt(12) -> "24" half-points + get 152400 EMU; Pt(10.5) -> "21";
%     Pt(11.7) -> "23" (23.4 half-points, ST_HpsMeasure `fix` truncation, H6) + get
%     146050 EMU; Emu(152400) -> "24"; None -> removed.
%
%     (NAME) ascii+hAnsi asymmetry ("Calibri" -> both @w:ascii/@w:hAnsi); the
%     non-ASCII "M-fullwidth Serif e-acute" UTF-8 round-trip (byte pin via serhex,
%     built with char(...) so the source stays ASCII); None -> removed.
%
%     (UNDERLINE) True -> @w:val="single"/get True; False -> "none"/get False;
%     None -> removed; the WD_UNDERLINE domain DOUBLE/WAVY/WORDS/DOTTED/DASH (each
%     serhex + get member NAME); parsed edges <w:u/> (no @val) -> get None and parsed
%     @w:val="single" -> get True.
%
%     (HIGHLIGHT) YELLOW -> "yellow"; TURQUOISE -> "cyan"; None -> removed.
%
%     (SUB/SUPER truth table) subscript=True -> <w:vertAlign w:val="subscript"/>;
%     superscript=True SWITCHES to "superscript"; subscript=False while super is set
%     does NOT remove (byte-unchanged); superscript=False REMOVES vertAlign; a fresh
%     superscript=True then None removes.
%
%     (COLORFORMAT -- the precedence pin) rgb/theme_color/type detection over
%     RGB/THEME/AUTO/None with the THEME-before-AUTO-before-RGB order; setting
%     theme_color after rgb RETAINS the RGB @w:val (Word's good guess) yet type reports
%     THEME; theme_color=None removes color; parsed @w:val="auto" -> AUTO/rgb None;
%     rgb=None / theme_color=None no-ops on a fresh run create NO rPr; Font.color
%     returns a FRESH ColorFormat that mutates the run.
%
%     (COLORFORMAT InvalidXmlError edge -- Gate-3 bonus) on a parsed
%     <w:color w:themeColor="accent1"/> (RequiredAttribute @w:val ABSENT), ColorFormat.rgb
%     raises mat2doc:InvalidXmlError with the verbatim "required 'w:val' attribute not
%     present..." message, while .type/.theme_color on the SAME tree read THEME/ACCENT_1
%     safely (type checks themeColor before val).
%
%     (H5 identity) Font/ColorFormat == by element handle (inherited from ElementProxy):
%     Font.color==Font.color True (fresh proxy, same run); Font==Font.color True;
%     Font(r)==Font(r) True; Font(r1)==Font(r2) False; ColorFormat(r)==ColorFormat(r) True.
%
%   Provenance (Gate-1..3, all 2026-07-30):
%     * Audit    : validation\mat2doc\audit_P4-4a_text_font.md (Porter Gate-1 +
%                  Fable Gate-2 adversarial APPROVE, zero defects; 20/20 tag map,
%                  size/name/underline/sub-super/ColorFormat all verified).
%     * Validate : validation\mat2doc\validate_P4-4a_text_font.md (Gate-3 PASS --
%                  probe_diff MATCH 228/228 leaf assertions; M1 17/17 byte-identical;
%                  H5 identity; regression 585/585; ZERO divergences, ZERO new D-numbers).
%     * Scenario : validation\mat2doc\scenarios\s0024_p4_4a_text_font.{py,m}
%                  (the probe sequence replayed VERBATIM by runProbes() below).
%     * Frozen ref (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE):
%         references\s0024\probe.json -- copied verbatim (self-contained) into
%           tests\text\data\s0024_probe_oracle.json (the Equivalence replay set;
%           value/serhex JSON, jsondecode is line-ending agnostic -> no `* binary`
%           .gitattributes needed, per the s0020/s0023 precedent).
%
%   Coverage taxonomy
%   -----------------
%   * Nominal      -- documented happy path: bold/size/name/underline/color set,
%                     ColorFormat rgb, Font.color wiring.
%   * Edge         -- empty/absent-get -> None; None/[]-removal branches; the non-ASCII
%                     name UTF-8 round-trip; the parsed <w:u/> no-@val -> None edge; the
%                     no-op-creates-no-rPr paths; and the InvalidXmlError error path
%                     (identifier mat2doc:InvalidXmlError + verbatim message, not just
%                     that it throws).
%   * Equivalence  -- test_equivalence_full_battery_vs_frozen_oracle replays the ENTIRE
%                     s0024 battery live (runProbes, the .m twin's body verbatim) and
%                     flatten-compares every leaf to the frozen python-docx 1.2.0 oracle
%                     (Gate-3 found ZERO divergences: every leaf byte/value-identical).
%   * Regression   -- hard-coded expected serialized-XML strings + UPPERCASE serhex of
%                     the raw UTF-8 shipping bytes vs the frozen oracle; hard-coded EMU
%                     property values.
%   * Upstream     -- the Font tri-state property surface (bool None/True/False, the
%                     WD_UNDERLINE domain, the sub/super truth table) and the ColorFormat
%                     rgb/theme_color/type precedence ARE the python-docx font.py/color.py
%                     API surface; the frozen oracle IS lxml's expected output for this
%                     API sequence.
%
%   Byte-level (L1) note: every serialized-XML assertion is either the FULL
%   serialize_part_xml output as an ASCII-decoded string (string-equality ==
%   byte-equality L1) or its UPPERCASE hex (serhex) vs the frozen oracle -- the serhex
%   pins include the non-ASCII name where the shipping bytes are NOT ASCII. No D-number
%   granted any L2 relaxation in this WP (Gate-3: zero new, none engaged), so every pin
%   here is L1.
%
%   Determinism: no network, no absolute paths. The co-located oracle resolves relative
%   to this file via fileparts(mfilename('fullpath')); every file read is binary
%   ('r','n'); no 'wt'; non-ASCII built via char(codepoints). The +mat2doc package
%   resolves via the MANDATORY PathFixture(worktree-root) in TestClassSetup (WP9-F4
%   lesson).

    properties (Constant)
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        % The 20 boolean tri-state Font properties (order == s0024 boolProps).
        BOOL_PROPS = ["all_caps" "bold" "complex_script" "cs_bold" "cs_italic" ...
            "double_strike" "emboss" "hidden" "italic" "imprint" "math" "no_proof" ...
            "outline" "rtl" "shadow" "small_caps" "snap_to_grid" "spec_vanish" ...
            "strike" "web_hidden"]

        % Font-property -> w: local tag (for the hard-coded XML-string regression pins).
        BOOL_TAG = struct( ...
            "all_caps","caps", "bold","b", "complex_script","cs", "cs_bold","bCs", ...
            "cs_italic","iCs", "double_strike","dstrike", "emboss","emboss", ...
            "hidden","vanish", "italic","i", "imprint","imprint", "math","oMath", ...
            "no_proof","noProof", "outline","outline", "rtl","rtl", "shadow","shadow", ...
            "small_caps","smallCaps", "snap_to_grid","snapToGrid", ...
            "spec_vanish","specVanish", "strike","strike", "web_hidden","webHidden")
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run cannot
            % resolve the +mat2doc package (MATLAB:undefinedVarOrClass). here is
            % tests\text; the worktree root is two levels up.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\text
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. Font 20 boolean tri-state properties (the API-tier pin)       %
        % =============================================================== %

        function test_bool_props_tristate_vs_frozen_oracle(testCase)
            % Nominal + Edge + Regression (font.py 26-416, s0024 bool_props): every
            % one of the 20 bool props driven absent -> True -> False -> None on a
            % FRESH run. serhex (byte pin) and the get repr are compared to the frozen
            % oracle for all three set states; the absent-get is [] (None). This is the
            % 20-property API-tier equivalence surface (140 oracle leaves).
            oracle = loadOracle();
            for i = 1:numel(testCase.BOOL_PROPS)
                name = testCase.BOOL_PROPS(i);
                od = oracle.bool_props.(char(name));

                [r, f] = newFont();
                testCase.verifyTrue(isequal(f.(name), []), ...
                    sprintf('%s absent get -> [] (None)', name));
                testCase.verifyEqual(string(od.absent), "None", ...
                    sprintf('%s oracle absent sanity', name));

                f.(name) = true;
                testCase.verifyEqual(hx_e(r), string(od.true_serhex), ...
                    sprintf('%s set True serialized bytes (L1) vs oracle', name));
                testCase.verifyEqual(rv(f.(name)), string(od.true_get), ...
                    sprintf('%s True get vs oracle', name));

                f.(name) = false;
                testCase.verifyEqual(hx_e(r), string(od.false_serhex), ...
                    sprintf('%s set False serialized bytes (L1) vs oracle', name));
                testCase.verifyEqual(rv(f.(name)), string(od.false_get), ...
                    sprintf('%s False get vs oracle', name));

                f.(name) = [];
                testCase.verifyEqual(hx_e(r), string(od.none_serhex), ...
                    sprintf('%s set None serialized bytes (L1) vs oracle', name));
                testCase.verifyEqual(rv(f.(name)), string(od.none_get), ...
                    sprintf('%s None get vs oracle', name));
            end
        end

        function test_bool_props_hardcoded_xml_regression(testCase)
            % Regression (hard-coded L1): for every bool prop, the three tri-state
            % serializations must equal the hard-coded expected r.xml strings (pure
            % ASCII -> string-equality is byte-identical), INDEPENDENT of the oracle.
            % True -> bare <w:tag/>; False -> <w:tag w:val="0"/>; None -> <w:rPr/>.
            for i = 1:numel(testCase.BOOL_PROPS)
                name = testCase.BOOL_PROPS(i);
                tag  = testCase.BOOL_TAG.(char(name));

                [r, f] = newFont();
                f.(name) = true;
                testCase.verifyEqual(ser(r), rWrap(testCase, ...
                    "<w:rPr><w:" + tag + "/></w:rPr>"), ...
                    sprintf('%s True -> bare <w:%s/> (hard-coded L1)', name, tag));
                f.(name) = false;
                testCase.verifyEqual(ser(r), rWrap(testCase, ...
                    "<w:rPr><w:" + tag + " w:val=""0""/></w:rPr>"), ...
                    sprintf('%s False -> <w:%s w:val="0"/> (hard-coded L1)', name, tag));
                f.(name) = [];
                testCase.verifyEqual(ser(r), rWrap(testCase, "<w:rPr/>"), ...
                    sprintf('%s None -> <w:rPr/> (hard-coded L1)', name));
            end
        end

        % =============================================================== %
        % 2. size -- half-points, EMU read, H6 fix-truncation              %
        % =============================================================== %

        function test_size_halfpoints_emu_and_h6(testCase)
            % Nominal + Edge + Regression (font.py 254-278, s0024 size): absent -> None;
            % Pt(12) -> "24" + get 152400 EMU; Pt(10.5) -> "21"; Pt(11.7) -> "23"
            % (23.4 half-points, ST_HpsMeasure fix truncation, H6) + get 146050 EMU;
            % Emu(152400) -> "24"; None -> removed. serhex + EMU vs oracle + hard-coded.
            o = loadOracle().size;

            [r, f] = newFont();
            testCase.verifyTrue(isequal(f.size, []), 'size absent get -> [] (None)');

            f.size = mat2doc.shared.Pt(12);
            testCase.verifyEqual(hx_e(r), string(o.pt12_serhex), 'Pt(12) serhex vs oracle');
            testCase.verifyEqual(ser(r), rWrap(testCase, "<w:rPr><w:sz w:val=""24""/></w:rPr>"), ...
                'Pt(12) -> <w:sz w:val="24"/> half-points (hard-coded L1)');
            testCase.verifyEqual(rv(f.size), string(o.pt12_get_emu), 'Pt(12) get EMU vs oracle');
            testCase.verifyEqual(round(double(f.size)), 152400, 'Pt(12) == 152400 EMU (hard-coded)');

            [r, f] = newFont();  f.size = mat2doc.shared.Pt(10.5);
            testCase.verifyEqual(hx_e(r), string(o.pt105_serhex), 'Pt(10.5) -> "21" serhex vs oracle');

            [r, f] = newFont();  f.size = mat2doc.shared.Pt(11.7);   % 23.4 hp -> fix -> "23"
            testCase.verifyEqual(hx_e(r), string(o.pt117_serhex), ...
                'Pt(11.7) -> "23" (H6 fix-truncation) serhex vs oracle');
            testCase.verifyEqual(ser(r), rWrap(testCase, "<w:rPr><w:sz w:val=""23""/></w:rPr>"), ...
                'Pt(11.7) -> <w:sz w:val="23"/> H6 truncation (hard-coded L1)');
            testCase.verifyEqual(rv(f.size), string(o.pt117_get_emu), 'Pt(11.7) get EMU vs oracle');
            testCase.verifyEqual(round(double(f.size)), 146050, 'Pt(11.7) read back == 146050 EMU');

            [r, f] = newFont();  f.size = mat2doc.shared.Emu(152400);
            testCase.verifyEqual(hx_e(r), string(o.emu152400_serhex), 'Emu(152400) -> "24" serhex vs oracle');
            f.size = [];
            testCase.verifyEqual(hx_e(r), string(o.none_serhex), 'size None -> removed serhex vs oracle');
            testCase.verifyEqual(ser(r), rWrap(testCase, "<w:rPr/>"), 'size None -> <w:rPr/> (hard-coded L1)');
            testCase.verifyTrue(isequal(f.size, []), 'size None get -> [] (None)');
        end

        % =============================================================== %
        % 3. name -- ascii/hAnsi + non-ASCII UTF-8 round-trip              %
        % =============================================================== %

        function test_name_ascii_hansi_and_nonascii_roundtrip(testCase)
            % Nominal + Edge + Regression (font.py 184-200, s0024 name): absent -> None;
            % "Calibri" -> BOTH @w:ascii/@w:hAnsi + get "Calibri"; the NON-ASCII name
            % (fullwidth-M U+FF2D + " Serif " + e-acute U+00E9), built via char(...) so
            % the source file stays ASCII, round-trips byte-identical (UTF-8 shipping
            % bytes pinned in serhex -- NOT ASCII, so this is a true byte pin) + get
            % identical; None -> removed.
            o = loadOracle().name;
            nonascii = string(char([65325 32 83 101 114 105 102 32 233]));  % "M(FF2D) Serif e(00E9)"

            [r, f] = newFont();
            testCase.verifyTrue(isequal(f.name, []), 'name absent get -> [] (None)');

            f.name = "Calibri";
            testCase.verifyEqual(hx_e(r), string(o.calibri_serhex), 'Calibri serhex vs oracle');
            testCase.verifyEqual(ser(r), rWrap(testCase, ...
                "<w:rPr><w:rFonts w:ascii=""Calibri"" w:hAnsi=""Calibri""/></w:rPr>"), ...
                'name="Calibri" -> ascii+hAnsi (hard-coded L1)');
            testCase.verifyEqual(rv(f.name), string(o.calibri_get), 'Calibri get vs oracle');

            [r, f] = newFont();
            f.name = nonascii;
            testCase.verifyEqual(hx_e(r), string(o.nonascii_serhex), ...
                'non-ASCII name UTF-8 shipping bytes (serhex byte pin) vs oracle');
            testCase.verifyEqual(rv(f.name), string(o.nonascii_get), 'non-ASCII name get round-trip vs oracle');
            testCase.verifyEqual(rv(f.name), nonascii, 'non-ASCII name round-trips exactly (hard-coded)');
            f.name = [];
            testCase.verifyEqual(hx_e(r), string(o.none_serhex), 'name None -> removed serhex vs oracle');
            testCase.verifyTrue(isequal(f.name, []), 'name None get -> [] (None)');
        end

        % =============================================================== %
        % 4. underline -- tri-state + WD_UNDERLINE domain + parsed edges   %
        % =============================================================== %

        function test_underline_tristate_domain_and_parsed(testCase)
            % Nominal + Edge + Regression (font.py 369-403, s0024 underline): absent
            % -> None; True -> @w:val="single"/get True; False -> "none"/get False;
            % None -> removed; the WD_UNDERLINE domain DOUBLE/WAVY/WORDS/DOTTED/DASH
            % (serhex + get member NAME); parsed <w:u/> (no @val) -> get None and
            % parsed @w:val="single" -> get True.
            o = loadOracle().underline;

            [r, f] = newFont();
            testCase.verifyTrue(isequal(f.underline, []), 'underline absent get -> [] (None)');
            f.underline = true;
            testCase.verifyEqual(hx_e(r), string(o.true_serhex), 'underline True serhex vs oracle');
            testCase.verifyEqual(ser(r), rWrap(testCase, "<w:rPr><w:u w:val=""single""/></w:rPr>"), ...
                'underline True -> @w:val="single" (hard-coded L1)');
            testCase.verifyEqual(rv(f.underline), string(o.true_get), 'underline True get vs oracle');
            f.underline = false;
            testCase.verifyEqual(hx_e(r), string(o.false_serhex), 'underline False serhex vs oracle');
            testCase.verifyEqual(ser(r), rWrap(testCase, "<w:rPr><w:u w:val=""none""/></w:rPr>"), ...
                'underline False -> @w:val="none" (hard-coded L1)');
            testCase.verifyEqual(rv(f.underline), string(o.false_get), 'underline False get vs oracle');
            f.underline = [];
            testCase.verifyEqual(hx_e(r), string(o.none_serhex), 'underline None -> removed serhex vs oracle');
            testCase.verifyTrue(isequal(f.underline, []), 'underline None get -> [] (None)');

            % WD_UNDERLINE domain (member -> xml + get member NAME)
            dom = { ...
                "double", mat2doc.enum.text.WD_UNDERLINE.DOUBLE; ...
                "wavy",   mat2doc.enum.text.WD_UNDERLINE.WAVY; ...
                "words",  mat2doc.enum.text.WD_UNDERLINE.WORDS; ...
                "dotted", mat2doc.enum.text.WD_UNDERLINE.DOTTED; ...
                "dash",   mat2doc.enum.text.WD_UNDERLINE.DASH};
            for i = 1:size(dom, 1)
                tag = dom{i, 1};
                [r, f] = newFont();
                f.underline = dom{i, 2};
                testCase.verifyEqual(hx_e(r), string(o.(tag + "_serhex")), ...
                    sprintf('underline %s serhex vs oracle', tag));
                testCase.verifyEqual(rv(f.underline), string(o.(tag + "_get")), ...
                    sprintf('underline %s get member NAME vs oracle', tag));
            end
            % DOUBLE hard-coded byte pin
            [r, f] = newFont();  f.underline = mat2doc.enum.text.WD_UNDERLINE.DOUBLE;
            testCase.verifyEqual(ser(r), rWrap(testCase, "<w:rPr><w:u w:val=""double""/></w:rPr>"), ...
                'underline DOUBLE -> @w:val="double" (hard-coded L1)');

            % parsed edges: <w:u/> (no @val) -> get None; @w:val="single" -> get True
            f = mat2doc.text.Font(parseR("<w:rPr><w:u/></w:rPr>"));
            testCase.verifyEqual(rv(f.underline), string(o.parsed_noval_get), ...
                'parsed <w:u/> (no @val) -> get None vs oracle');
            testCase.verifyTrue(isequal(f.underline, []), 'parsed <w:u/> -> [] (None)');
            f = mat2doc.text.Font(parseR("<w:rPr><w:u w:val=""single""/></w:rPr>"));
            testCase.verifyEqual(rv(f.underline), string(o.parsed_single_get), ...
                'parsed @w:val="single" -> get True vs oracle');
        end

        % =============================================================== %
        % 5. highlight_color                                              %
        % =============================================================== %

        function test_highlight_color(testCase)
            % Nominal + Edge + Regression (font.py 133-144, s0024 highlight): absent
            % -> None; YELLOW -> "yellow"/get YELLOW; TURQUOISE -> "cyan"/get TURQUOISE;
            % None -> removed. serhex + get vs oracle + hard-coded.
            o = loadOracle().highlight;

            [r, f] = newFont();
            testCase.verifyTrue(isequal(f.highlight_color, []), 'highlight absent get -> [] (None)');
            f.highlight_color = mat2doc.enum.text.WD_COLOR_INDEX.YELLOW;
            testCase.verifyEqual(hx_e(r), string(o.yellow_serhex), 'highlight YELLOW serhex vs oracle');
            testCase.verifyEqual(ser(r), rWrap(testCase, "<w:rPr><w:highlight w:val=""yellow""/></w:rPr>"), ...
                'highlight YELLOW -> @w:val="yellow" (hard-coded L1)');
            testCase.verifyEqual(rv(f.highlight_color), string(o.yellow_get), 'highlight YELLOW get vs oracle');

            [r, f] = newFont();
            f.highlight_color = mat2doc.enum.text.WD_COLOR_INDEX.TURQUOISE;
            testCase.verifyEqual(hx_e(r), string(o.turquoise_serhex), ...
                'highlight TURQUOISE -> "cyan" serhex vs oracle');
            testCase.verifyEqual(rv(f.highlight_color), string(o.turquoise_get), 'highlight TURQUOISE get vs oracle');
            f.highlight_color = [];
            testCase.verifyEqual(hx_e(r), string(o.none_serhex), 'highlight None -> removed serhex vs oracle');
            testCase.verifyTrue(isequal(f.highlight_color, []), 'highlight None get -> [] (None)');
        end

        % =============================================================== %
        % 6. subscript / superscript truth table                          %
        % =============================================================== %

        function test_subscript_superscript_truth_table(testCase)
            % Nominal + Edge + Regression (font.py 334-367, s0024 sub_super): the
            % vertAlign truth table on ONE run -- subscript=True -> "subscript"
            % (sub T/sup F); superscript=True SWITCHES to "superscript" (sub F/sup T);
            % subscript=False while super set does NOT remove (byte-unchanged, sub F/sup
            % T); superscript=False REMOVES vertAlign (sub None/sup None); then a fresh
            % superscript=True -> "superscript", None removes. serhex + reads vs oracle.
            o = loadOracle().sub_super;

            [r, f] = newFont();
            f.subscript = true;
            testCase.verifyEqual(hx_e(r), string(o.sub_true_serhex), 'sub=True serhex vs oracle');
            testCase.verifyEqual(ser(r), rWrap(testCase, "<w:rPr><w:vertAlign w:val=""subscript""/></w:rPr>"), ...
                'subscript=True -> <w:vertAlign w:val="subscript"/> (hard-coded L1)');
            testCase.verifyEqual(rv(f.subscript), string(o.sub_true_sub), 'sub=True -> sub True vs oracle');
            testCase.verifyEqual(rv(f.superscript), string(o.sub_true_sup), 'sub=True -> sup False vs oracle');

            f.superscript = true;                 % SWITCH
            testCase.verifyEqual(hx_e(r), string(o.sup_switch_serhex), 'sup switch serhex vs oracle');
            testCase.verifyEqual(ser(r), rWrap(testCase, "<w:rPr><w:vertAlign w:val=""superscript""/></w:rPr>"), ...
                'superscript=True SWITCHES to "superscript" (hard-coded L1)');
            testCase.verifyEqual(rv(f.subscript), string(o.sup_switch_sub), 'switch -> sub False vs oracle');
            testCase.verifyEqual(rv(f.superscript), string(o.sup_switch_sup), 'switch -> sup True vs oracle');

            f.subscript = false;                  % does NOT remove (sup set)
            testCase.verifyEqual(hx_e(r), string(o.sub_false_serhex), ...
                'subscript=False while super set does NOT remove (byte-unchanged) vs oracle');
            testCase.verifyEqual(rv(f.subscript), string(o.sub_false_sub), 'sub=False -> sub False vs oracle');
            testCase.verifyEqual(rv(f.superscript), string(o.sub_false_sup), 'sub=False keeps sup True vs oracle');

            f.superscript = false;                % removes vertAlign
            testCase.verifyEqual(hx_e(r), string(o.sup_false_serhex), 'sup=False removes vertAlign serhex vs oracle');
            testCase.verifyEqual(ser(r), rWrap(testCase, "<w:rPr/>"), ...
                'superscript=False REMOVES vertAlign -> <w:rPr/> (hard-coded L1)');
            testCase.verifyEqual(rv(f.subscript), string(o.sup_false_sub), 'sup=False -> sub None vs oracle');
            testCase.verifyEqual(rv(f.superscript), string(o.sup_false_sup), 'sup=False -> sup None vs oracle');

            [r, f] = newFont();
            f.superscript = true;
            testCase.verifyEqual(hx_e(r), string(o.fresh_sup_serhex), 'fresh sup=True serhex vs oracle');
            f.superscript = [];
            testCase.verifyEqual(hx_e(r), string(o.sup_none_serhex), 'sup=None removes vs oracle');
            testCase.verifyEqual(rv(f.superscript), string(o.sup_none_get), 'sup=None get -> None vs oracle');
        end

        % =============================================================== %
        % 7. ColorFormat -- rgb/theme_color/type precedence (the pin)      %
        % =============================================================== %

        function test_colorformat_precedence_vs_frozen_oracle(testCase)
            % Nominal + Edge + Regression (color.py 29-101, s0024 colorformat): the
            % ColorFormat precedence pin. absent -> type/rgb/theme None; rgb=FF0000
            % -> <w:color w:val="FF0000"/>, type RGB, rgb "FF0000"; theme_color=ACCENT_1
            % after rgb RETAINS the RGB @w:val (Word's good guess) yet type reports THEME
            % (checked before val), rgb still "FF0000", theme ACCENT_1; theme_color=None
            % removes color, type None; parsed @w:val="auto" -> AUTO, rgb None;
            % rgb=None no-op on fresh -> NO rPr; theme_color=None no-op on fresh -> NO
            % rPr; rgb=None after parsed auto -> color removed, type None. serhex + reads
            % vs oracle + hard-coded.
            o = loadOracle().colorformat;

            [r, c] = newCf();
            testCase.verifyEqual(rv(c.type), string(o.absent_type), 'absent type None vs oracle');
            testCase.verifyEqual(rv(c.rgb), string(o.absent_rgb), 'absent rgb None vs oracle');
            testCase.verifyEqual(rv(c.theme_color), string(o.absent_theme), 'absent theme None vs oracle');

            c.rgb = mat2doc.shared.RGBColor(255, 0, 0);
            testCase.verifyEqual(hx_e(r), string(o.rgb_serhex), 'rgb=FF0000 serhex vs oracle');
            testCase.verifyEqual(ser(r), rWrap(testCase, "<w:rPr><w:color w:val=""FF0000""/></w:rPr>"), ...
                'rgb=FF0000 -> <w:color w:val="FF0000"/> (hard-coded L1)');
            testCase.verifyEqual(rv(c.type), string(o.rgb_type), 'rgb -> type RGB vs oracle');
            testCase.verifyEqual(c.rgb.str_(), string(o.rgb_val), 'rgb value "FF0000" vs oracle');

            c.theme_color = mat2doc.enum.dml.MSO_THEME_COLOR.ACCENT_1;   % RGB val RETAINED
            testCase.verifyEqual(hx_e(r), string(o.theme_serhex), 'theme after rgb serhex (val retained) vs oracle');
            testCase.verifyEqual(ser(r), rWrap(testCase, ...
                "<w:rPr><w:color w:val=""FF0000"" w:themeColor=""accent1""/></w:rPr>"), ...
                'theme_color after rgb RETAINS @w:val (Word good-guess) (hard-coded L1)');
            testCase.verifyEqual(rv(c.type), string(o.theme_type), ...
                'type THEME (themeColor checked BEFORE val) vs oracle');
            testCase.verifyEqual(c.rgb.str_(), string(o.theme_rgb_retained), 'rgb still "FF0000" vs oracle');
            testCase.verifyEqual(rv(c.theme_color), string(o.theme_name), 'theme_color ACCENT_1 vs oracle');

            c.theme_color = [];                   % removes color
            testCase.verifyEqual(hx_e(r), string(o.theme_none_serhex), 'theme_color=None removes color serhex vs oracle');
            testCase.verifyEqual(rv(c.type), string(o.theme_none_type), 'theme_color=None -> type None vs oracle');

            % parsed @w:val="auto" -> AUTO / rgb None
            c = mat2doc.dml.ColorFormat(parseR("<w:rPr><w:color w:val=""auto""/></w:rPr>"));
            testCase.verifyEqual(rv(c.type), string(o.auto_type), 'parsed auto -> type AUTO vs oracle');
            testCase.verifyEqual(rv(c.rgb), string(o.auto_rgb), 'parsed auto -> rgb None vs oracle');

            % no-op paths create NO rPr on a fresh run
            [r, c] = newCf();  c.rgb = [];
            testCase.verifyEqual(hx_e(r), string(o.rgb_none_noop_serhex), ...
                'rgb=None no-op on fresh -> NO rPr created serhex vs oracle');
            testCase.verifyEqual(ser(r), decl() + newline + "<w:r xmlns:w=""" + testCase.W + """/>", ...
                'rgb=None no-op leaves a bare <w:r/> (hard-coded L1)');
            [r, c] = newCf();  c.theme_color = [];
            testCase.verifyEqual(hx_e(r), string(o.theme_none_noop_serhex), ...
                'theme_color=None no-op on fresh -> NO rPr created serhex vs oracle');

            % rgb=None after parsed auto -> color removed, type None
            r = parseR("<w:rPr><w:color w:val=""auto""/></w:rPr>");
            c = mat2doc.dml.ColorFormat(r);
            c.rgb = [];
            testCase.verifyEqual(hx_e(r), string(o.rgb_none_after_auto_serhex), ...
                'rgb=None after parsed auto -> color removed serhex vs oracle');
            testCase.verifyEqual(rv(c.type), string(o.rgb_none_after_auto_type), ...
                'rgb=None after auto -> type None vs oracle');
        end

        % =============================================================== %
        % 8. ColorFormat.rgb InvalidXmlError edge (Gate-3 bonus)           %
        % =============================================================== %

        function test_colorformat_themeonly_rgb_raises_invalidxmlerror(testCase)
            % Edge / error-path (Gate-3 bonus, s0024 colorformat.themeonly_*): on a
            % PARSED <w:color w:themeColor="accent1"/> (RequiredAttribute @w:val ABSENT),
            % ColorFormat.rgb reads the required @w:val and RAISES -- verify the
            % IDENTIFIER mat2doc:InvalidXmlError AND the verbatim python-docx message
            % (not merely that it throws). .type/.theme_color on the SAME tree read
            % THEME/ACCENT_1 safely (type checks themeColor before val).
            o = loadOracle().colorformat;
            c = mat2doc.dml.ColorFormat(parseR("<w:rPr><w:color w:themeColor=""accent1""/></w:rPr>"));

            % safe reads first (type checks themeColor before the absent @w:val)
            testCase.verifyEqual(rv(c.type), string(o.themeonly_type), 'themeonly type THEME (safe) vs oracle');
            testCase.verifyEqual(rv(c.theme_color), string(o.themeonly_theme), 'themeonly theme ACCENT_1 (safe) vs oracle');

            % .rgb reads the RequiredAttribute @w:val (absent) -> raises
            ME = captureError(@() c.rgb);
            testCase.verifyEqual(string(ME.identifier), "mat2doc:InvalidXmlError", ...
                'ColorFormat.rgb on val-absent themeColor -> mat2doc:InvalidXmlError');
            testCase.verifyEqual(string(ME.message), string(o.themeonly_rgb_msg), ...
                'InvalidXmlError message verbatim vs oracle');
            % oracle sanity: the raise flag is recorded True
            testCase.verifyEqual(string(o.themeonly_rgb_raises), "True", 'oracle themeonly_rgb_raises sanity');
        end

        % =============================================================== %
        % 9. Font.color wiring (fresh ColorFormat mutates the run)         %
        % =============================================================== %

        function test_color_via_font(testCase)
            % Nominal + Regression (font.py 50-54, s0024 color_via_font): Font.color
            % returns a FRESH ColorFormat wrapping the run each access; setting .rgb
            % through it mutates the run, and a subsequent fresh .color access reads it
            % back type RGB / "0066CC". serhex + reads vs oracle + hard-coded.
            o = loadOracle().color_via_font;

            [r, f] = newFont();
            testCase.verifyEqual(rv(f.color.type), string(o.fresh_type), 'fresh Font.color.type None vs oracle');
            f.color.rgb = mat2doc.shared.RGBColor(0, 102, 204);   % 0066CC
            testCase.verifyEqual(hx_e(r), string(o.set_serhex), 'color via Font serhex vs oracle');
            testCase.verifyEqual(ser(r), rWrap(testCase, "<w:rPr><w:color w:val=""0066CC""/></w:rPr>"), ...
                'f.color.rgb mutates the run -> <w:color w:val="0066CC"/> (hard-coded L1)');
            testCase.verifyEqual(rv(f.color.type), string(o.type), 'Font.color.type RGB (fresh proxy) vs oracle');
            testCase.verifyEqual(f.color.rgb.str_(), string(o.rgb), 'Font.color.rgb "0066CC" vs oracle');
        end

        % =============================================================== %
        % 10. H5 identity (ElementProxy element-handle ==)                 %
        % =============================================================== %

        function test_h5_identity(testCase)
            % Regression (shared.py __eq__/__ne__ inherited, s0024 h5_identity):
            % Font.color==Font.color True (fresh proxy each access, same run);
            % Font==Font.color True (both wrap the same run); Font(r)==Font(r) True;
            % Font(r1)==Font(r2) False (different elements); ColorFormat(r)==ColorFormat(r)
            % True. Compared to the frozen oracle + asserted directly.
            o = loadOracle().h5_identity;

            [r, f] = newFont();
            testCase.verifyTrue(f.color == f.color, 'Font.color == Font.color (same run) True');
            testCase.verifyEqual(rv(f.color == f.color), string(o.color_eq_color), 'color_eq_color vs oracle');
            testCase.verifyTrue(f == f.color, 'Font == Font.color (same run) True');
            testCase.verifyEqual(rv(f == f.color), string(o.font_eq_color), 'font_eq_color vs oracle');
            testCase.verifyTrue(mat2doc.text.Font(r) == mat2doc.text.Font(r), 'Font(r) == Font(r) True');
            testCase.verifyEqual(rv(mat2doc.text.Font(r) == mat2doc.text.Font(r)), ...
                string(o.font_eq_font_same), 'font_eq_font_same vs oracle');

            [r2, ~] = newFont();
            testCase.verifyFalse(mat2doc.text.Font(r) == mat2doc.text.Font(r2), 'Font(r1) == Font(r2) False');
            testCase.verifyEqual(rv(mat2doc.text.Font(r) == mat2doc.text.Font(r2)), ...
                string(o.font_ne_font_diff), 'font_ne_font_diff vs oracle');
            testCase.verifyTrue(mat2doc.dml.ColorFormat(r) == mat2doc.dml.ColorFormat(r), ...
                'ColorFormat(r) == ColorFormat(r) True');
            testCase.verifyEqual(rv(mat2doc.dml.ColorFormat(r) == mat2doc.dml.ColorFormat(r)), ...
                string(o.cf_eq_cf_same), 'cf_eq_cf_same vs oracle');
        end

        % =============================================================== %
        % 11. M1 sanity (light) -- ElementProxy lineage, API/proxy tier    %
        % =============================================================== %

        function test_font_colorformat_are_elementproxy(testCase)
            % M1 sanity (light): P4-4a is API/proxy tier -- byte-neutral, no styles.xml/
            % document.xml pin needed here (the full 17/17 M1 sweep stays owned by
            % Test_p1_8_skeleton_m1; validate_P4-4a §2 confirms M1 17/17 unchanged).
            % Spot-check that both new classes ARE ElementProxy subclasses (so H5
            % identity + reference semantics come from the shared base) and expose their
            % surface.
            r = mat2doc.oxml.OxmlElement("w:r");
            f = mat2doc.text.Font(r);
            c = mat2doc.dml.ColorFormat(r);
            testCase.verifyClass(f, 'mat2doc.text.Font');
            testCase.verifyClass(c, 'mat2doc.dml.ColorFormat');
            testCase.verifyTrue(isa(f, 'mat2doc.shared.ElementProxy'), ...
                'Font must be a mat2doc.shared.ElementProxy');
            testCase.verifyTrue(isa(c, 'mat2doc.shared.ElementProxy'), ...
                'ColorFormat must be a mat2doc.shared.ElementProxy');
            % Font.color returns a ColorFormat (also an ElementProxy); spot behavior.
            testCase.verifyClass(f.color, 'mat2doc.dml.ColorFormat');
            f.bold = true;
            testCase.verifyTrue(f.bold, 'Font.bold spot round-trip');
        end

        % =============================================================== %
        % 12. EQUIVALENCE -- full s0024 battery vs the frozen oracle       %
        % =============================================================== %

        function test_equivalence_full_battery_vs_frozen_oracle(testCase)
            % Equivalence: replay the ENTIRE s0024 battery (runProbes -- the .m twin's
            % body VERBATIM: bool_props / size / name / underline / highlight / sub_super
            % / color_via_font / colorformat [incl. the InvalidXmlError message leaf] /
            % h5_identity) and flatten-compare EVERY leaf to the frozen python-docx 1.2.0
            % oracle copied into data\s0024_probe_oracle.json. Gate-3 found ZERO
            % divergences (probe_diff MATCH 228/228), so every leaf must be
            % byte/value-identical. This single method is the class's Equivalence leg,
            % tying the suite to the Gate-3 reference output.
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
            testCase.verifyGreaterThan(numel(oKeys), 180, ...
                'the flattened oracle must expose the full battery of leaves');
            for i = 1:numel(oKeys)
                k = oKeys{i};
                testCase.verifyTrue(isKey(pMap, k), sprintf('port is missing leaf %s', k));
                testCase.verifyEqual(pMap(k), oMap(k), ...
                    sprintf('leaf %s must be byte/value-identical to the frozen oracle', k));
            end
        end

    end
end

% ===================== file-local helpers ============================== %

function s = decl()
    s = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>";
end

function W = W_()
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
end

function [r, f] = newFont()
    r = mat2doc.oxml.OxmlElement("w:r");
    f = mat2doc.text.Font(r);
end

function [r, c] = newCf()
    r = mat2doc.oxml.OxmlElement("w:r");
    c = mat2doc.dml.ColorFormat(r);
end

function e = parseR(inner)
    % Parse a <w:r> carrying its own xmlns:w with the given inner XML (mirrors the
    % s0024 twin's parseR()).
    xml = "<w:r xmlns:w=""" + W_() + """>" + string(inner) + "</w:r>";
    e = mat2doc.oxml.parse_xml(uint8(unicode2native(char(xml), "UTF-8")));
end

function s = ser(e)
    % Decode serialize_part_xml bytes to a string (ASCII content -> string-equality
    % is a byte-identical L1 assertion).
    s = string(native2unicode(mat2doc.oxml.serialize_part_xml(e), "UTF-8"));
end

function h = hx_e(e)
    % UPPERCASE hex (serhex) of the raw UTF-8 serialize_part_xml shipping bytes of the
    % whole w:r -- a byte pin (matches the s0024 oracle / Python bytes.hex().upper()).
    h = string(sprintf('%02X', uint8(mat2doc.oxml.serialize_part_xml(e))));
end

function s = rWrap(testCase, body)
    % decl + newline + <w:r xmlns:w="...">BODY</w:r>. Used by the hard-coded XML-string
    % regression pins (the run is created via OxmlElement("w:r") -> declares xmlns:w).
    s = decl() + newline + "<w:r xmlns:w=""" + testCase.W + """>" + string(body) + "</w:r>";
end

function s = rv(x)
    % Uniform accessor repr mirroring the s0024 rv(): None->"None",
    % bool->"True"/"False", Length/int->decimal, enum member->NAME, string->itself.
    if isequal(x, [])
        s = "None";
    elseif islogical(x)
        if x, s = "True"; else, s = "False"; end
    elseif isnumeric(x)
        s = string(sprintf('%.0f', double(x)));
    else
        s = string(x);
    end
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

function o = loadOracle()
    % Read the co-located frozen oracle in BINARY mode (no CRLF translation) and decode
    % UTF-8 -> struct. jsondecode is whitespace/line-ending agnostic, so no `* binary`
    % .gitattributes pin is needed (value/serhex fixture, s0020/s0023 precedent).
    here = fileparts(mfilename('fullpath'));
    p = fullfile(here, 'data', 's0024_probe_oracle.json');
    fid = fopen(p, 'r', 'n');
    if fid < 0
        error('mat2doc:test:openFail', 'cannot open frozen oracle %s', p);
    end
    raw = fread(fid, Inf, '*uint8')';
    fclose(fid);
    o = jsondecode(native2unicode(raw, 'UTF-8'));
end

function P = runProbes()
    % Replay the s0024 probe sequence (the .m twin's P-struct body, VERBATIM
    % tags/inputs/order) and return the nested struct of tagged canonical values.
    % Embedded here so the Equivalence leg is self-contained (the validation-folder
    % scenario is NOT on the toolbox path and must not be a dependency). Mirrors
    % validation\mat2doc\scenarios\s0024_p4_4a_text_font.m.
    P = struct();

    boolProps = ["all_caps" "bold" "complex_script" "cs_bold" "cs_italic" ...
        "double_strike" "emboss" "hidden" "italic" "imprint" "math" "no_proof" ...
        "outline" "rtl" "shadow" "small_caps" "snap_to_grid" "spec_vanish" ...
        "strike" "web_hidden"];

    % ===================== bool_props =====================================
    bp = struct();
    for i = 1:numel(boolProps)
        name = boolProps(i);
        [r, f] = newFont();
        absent = rv(f.(name));
        f.(name) = true;
        true_serhex = hx_e(r); true_get = rv(f.(name));
        f.(name) = false;
        false_serhex = hx_e(r); false_get = rv(f.(name));
        f.(name) = [];
        none_serhex = hx_e(r); none_get = rv(f.(name));
        bp.(name) = struct( ...
            "absent", absent, ...
            "true_serhex", true_serhex, "true_get", true_get, ...
            "false_serhex", false_serhex, "false_get", false_get, ...
            "none_serhex", none_serhex, "none_get", none_get);
    end
    P.bool_props = bp;

    % ===================== size ============================================
    sz = struct();
    [r, f] = newFont();
    sz.absent = rv(f.size);
    f.size = mat2doc.shared.Pt(12);
    sz.pt12_serhex = hx_e(r); sz.pt12_get_emu = rv(f.size);
    [r, f] = newFont();
    f.size = mat2doc.shared.Pt(10.5);
    sz.pt105_serhex = hx_e(r);
    [r, f] = newFont();
    f.size = mat2doc.shared.Pt(11.7);          % 23.4 half-points -> fix -> "23"
    sz.pt117_serhex = hx_e(r); sz.pt117_get_emu = rv(f.size);
    [r, f] = newFont();
    f.size = mat2doc.shared.Emu(152400);
    sz.emu152400_serhex = hx_e(r);
    f.size = [];
    sz.none_serhex = hx_e(r); sz.none_get = rv(f.size);
    P.size = sz;

    % ===================== name ============================================
    nm = struct();
    [r, f] = newFont();
    nm.absent = rv(f.name);
    f.name = "Calibri";
    nm.calibri_serhex = hx_e(r); nm.calibri_get = rv(f.name);
    [r, f] = newFont();
    f.name = string(char([65325 32 83 101 114 105 102 32 233]));   % non-ASCII, built ASCII-safe
    nm.nonascii_serhex = hx_e(r); nm.nonascii_get = rv(f.name);
    f.name = [];
    nm.none_serhex = hx_e(r); nm.none_get = rv(f.name);
    P.name = nm;

    % ===================== underline ======================================
    ul = struct();
    [r, f] = newFont();
    ul.absent = rv(f.underline);
    f.underline = true;
    ul.true_serhex = hx_e(r); ul.true_get = rv(f.underline);
    f.underline = false;
    ul.false_serhex = hx_e(r); ul.false_get = rv(f.underline);
    f.underline = [];
    ul.none_serhex = hx_e(r); ul.none_get = rv(f.underline);
    umembers = { ...
        "double", mat2doc.enum.text.WD_UNDERLINE.DOUBLE; ...
        "wavy",   mat2doc.enum.text.WD_UNDERLINE.WAVY; ...
        "words",  mat2doc.enum.text.WD_UNDERLINE.WORDS; ...
        "dotted", mat2doc.enum.text.WD_UNDERLINE.DOTTED; ...
        "dash",   mat2doc.enum.text.WD_UNDERLINE.DASH};
    for i = 1:size(umembers, 1)
        tag = umembers{i, 1};
        [r, f] = newFont();
        f.underline = umembers{i, 2};
        ul.(tag + "_serhex") = hx_e(r);
        ul.(tag + "_get") = rv(f.underline);
    end
    f = mat2doc.text.Font(parseR("<w:rPr><w:u/></w:rPr>"));
    ul.parsed_noval_get = rv(f.underline);
    f = mat2doc.text.Font(parseR("<w:rPr><w:u w:val=""single""/></w:rPr>"));
    ul.parsed_single_get = rv(f.underline);
    P.underline = ul;

    % ===================== highlight ======================================
    hi = struct();
    [r, f] = newFont();
    hi.absent = rv(f.highlight_color);
    f.highlight_color = mat2doc.enum.text.WD_COLOR_INDEX.YELLOW;
    hi.yellow_serhex = hx_e(r); hi.yellow_get = rv(f.highlight_color);
    [r, f] = newFont();
    f.highlight_color = mat2doc.enum.text.WD_COLOR_INDEX.TURQUOISE;
    hi.turquoise_serhex = hx_e(r); hi.turquoise_get = rv(f.highlight_color);
    f.highlight_color = [];
    hi.none_serhex = hx_e(r); hi.none_get = rv(f.highlight_color);
    P.highlight = hi;

    % ===================== sub_super truth table ==========================
    ss = struct();
    [r, f] = newFont();
    f.subscript = true;
    ss.sub_true_serhex = hx_e(r);
    ss.sub_true_sub = rv(f.subscript); ss.sub_true_sup = rv(f.superscript);
    f.superscript = true;                      % SWITCH -> superscript
    ss.sup_switch_serhex = hx_e(r);
    ss.sup_switch_sub = rv(f.subscript); ss.sup_switch_sup = rv(f.superscript);
    f.subscript = false;                       % does NOT remove (sup set)
    ss.sub_false_serhex = hx_e(r);
    ss.sub_false_sub = rv(f.subscript); ss.sub_false_sup = rv(f.superscript);
    f.superscript = false;                     % removes vertAlign
    ss.sup_false_serhex = hx_e(r);
    ss.sup_false_sub = rv(f.subscript); ss.sup_false_sup = rv(f.superscript);
    [r, f] = newFont();
    f.superscript = true;
    ss.fresh_sup_serhex = hx_e(r);
    f.superscript = [];
    ss.sup_none_serhex = hx_e(r); ss.sup_none_get = rv(f.superscript);
    P.sub_super = ss;

    % ===================== color_via_font =================================
    cvf = struct();
    [r, f] = newFont();
    cvf.fresh_type = rv(f.color.type);
    f.color.rgb = mat2doc.shared.RGBColor(0, 102, 204);   % 0066CC
    cvf.set_serhex = hx_e(r);
    cvf.type = rv(f.color.type); cvf.rgb = f.color.rgb.str_();
    P.color_via_font = cvf;

    % ===================== colorformat direct =============================
    cf = struct();
    [r, c] = newCf();
    cf.absent_type = rv(c.type); cf.absent_rgb = rv(c.rgb); cf.absent_theme = rv(c.theme_color);
    c.rgb = mat2doc.shared.RGBColor(255, 0, 0);
    cf.rgb_serhex = hx_e(r); cf.rgb_type = rv(c.type); cf.rgb_val = c.rgb.str_();
    c.theme_color = mat2doc.enum.dml.MSO_THEME_COLOR.ACCENT_1;   % val RETAINED
    cf.theme_serhex = hx_e(r); cf.theme_type = rv(c.type);
    cf.theme_rgb_retained = c.rgb.str_(); cf.theme_name = rv(c.theme_color);
    c.theme_color = [];                        % removes color
    cf.theme_none_serhex = hx_e(r); cf.theme_none_type = rv(c.type);
    c = mat2doc.dml.ColorFormat(parseR("<w:rPr><w:color w:val=""auto""/></w:rPr>"));
    cf.auto_type = rv(c.type); cf.auto_rgb = rv(c.rgb);
    c = mat2doc.dml.ColorFormat(parseR("<w:rPr><w:color w:themeColor=""accent1""/></w:rPr>"));
    cf.themeonly_type = rv(c.type); cf.themeonly_theme = rv(c.theme_color);
    % .rgb reads the REQUIRED @w:val -> InvalidXmlError (guarded; message compared)
    try
        v = c.rgb; %#ok<NASGU>
        cf.themeonly_rgb_raises = "False"; cf.themeonly_rgb_msg = "";
    catch ME
        cf.themeonly_rgb_raises = "True"; cf.themeonly_rgb_msg = string(ME.message);
    end
    [r, c] = newCf();
    c.rgb = [];
    cf.rgb_none_noop_serhex = hx_e(r);
    [r, c] = newCf();
    c.theme_color = [];
    cf.theme_none_noop_serhex = hx_e(r);
    r = parseR("<w:rPr><w:color w:val=""auto""/></w:rPr>");
    c = mat2doc.dml.ColorFormat(r);
    c.rgb = [];
    cf.rgb_none_after_auto_serhex = hx_e(r); cf.rgb_none_after_auto_type = rv(c.type);
    P.colorformat = cf;

    % ===================== h5_identity ====================================
    h5 = struct();
    [r, f] = newFont();
    h5.color_eq_color = rv(f.color == f.color);
    h5.font_eq_color = rv(f == f.color);
    h5.font_eq_font_same = rv(mat2doc.text.Font(r) == mat2doc.text.Font(r));
    [r2, ~] = newFont();
    h5.font_ne_font_diff = rv(mat2doc.text.Font(r) == mat2doc.text.Font(r2));
    h5.cf_eq_cf_same = rv(mat2doc.dml.ColorFormat(r) == mat2doc.dml.ColorFormat(r));
    P.h5_identity = h5;
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
        if isempty(s)
            map(prefix) = ''; %#ok<NASGU>
        else
            parts = strings(1, numel(s));
            for i = 1:numel(s)
                parts(i) = canonScalar(s{i});
            end
            map(prefix) = char(join(parts, "|")); %#ok<NASGU>
        end
    elseif (isstring(s) && ~isscalar(s))
        if isempty(s), map(prefix) = ''; else, map(prefix) = char(join(s(:).', "|")); end %#ok<NASGU>
    elseif isnumeric(s) && ~isscalar(s)
        if isempty(s), map(prefix) = ''; else, map(prefix) = char(join(string(s(:).'), ",")); end %#ok<NASGU>
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

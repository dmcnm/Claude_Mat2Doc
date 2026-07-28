classdef Test_p4_2_oxml_parfmt < matlab.unittest.TestCase
% TEST_P4_2_OXML_PARFMT  Gate-4 permanent unit tests for Mat2Doc P4-2
%   (src/docx/oxml/text/parfmt.py -> +mat2doc\+oxml\+text\CT_PPr + the para-
%   property children CT_Ind/CT_Jc/CT_Spacing/CT_TabStop/CT_TabStops, and
%   src/docx/oxml/text/paragraph.py -> CT_P), plus the 12 parfmt/paragraph tags
%   registered in +mat2doc\+oxml\registry.m).
%
%   P4-2 is the M2 byte-critical paragraph surface: add_paragraph / add_heading
%   write w:p / w:pPr / w:pStyle into document.xml. This class permanently freezes
%   the guarantees the prior gates established (Fable Gate-2 APPROVE zero-findings;
%   Gate-3 equivalence PASS -- probe_diff s0022 MATCH, M1 17/17, the H11 crux):
%
%     (H11 CT_PPr ORDERING, the crux) CT_PPr's 12 ZeroOrOne child descriptors
%     carry NON-CONTIGUOUS successor slices of a single 36-entry _tag_seq. Adding
%     children in ANY scrambled order MUST re-sort them into canonical OOXML
%     schema order (insert_element_before / first_child_found_in), so
%     document.xml / styles.xml stay byte-identical. In particular pStyle
%     (successors=_tag_seq[1:]) sorts FIRST even when added LAST -- the M2
%     add_heading path. A future edit to a successor slice goes RED here.
%
%     (M1 BYTE-NEUTRALITY) Registering the 12 parfmt/paragraph tags flips the
%     PARSE CLASS of every real w:pPr in styles.xml (204 of them) from generic
%     XmlElement to CT_PPr, and w:p in document.xml to CT_P. All CT_* exit through
%     the identical +oxml\serialize_part_xml walk, so the flip is byte-neutral.
%     This class pins word/styles.xml (349458 B) specifically -- it is the part
%     P4-2's CT_PPr PARSE path could break (the full 17/17 M1 sweep stays owned by
%     Test_p1_8_skeleton_m1; document.xml 1548 B is pinned here too).
%
%     (CT_TabStop CARRY-FORWARD) Registering w:tab->CT_TabStop closes the P4-1b
%     gap: CT_R.text and CT_P.text over a run containing w:tab now reproduce the
%     tab char ("a\tb"); str_() == "\t". Pinned so a revert goes RED.
%
%   Provenance (Gate-1..3, all 2026-07-27):
%     * Audit    : validation\mat2doc\audit_P4-2_oxml_parfmt.md (Porter Gate-1 +
%                  Fable Gate-2 adversarial APPROVE, 38/38 probes, zero findings;
%                  the 12-slice H11 table + 13 @property accessors verified).
%     * Validate : validation\mat2doc\validate_P4-2_oxml_parfmt.md (Gate-3
%                  equivalence PASS -- every leg byte/value-exact: M1 17/17
%                  [styles.xml 349458 B & document.xml 1548 B L1], the H11
%                  scrambled-add crux 4/4 byte-identical, CT_TabStop str_->"\t"
%                  carry-forward, CT_Ind/CT_Jc/CT_Spacing + all 13 CT_PPr
%                  accessors, CT_P, the 204-pPr styles.xml round-trip; probe_diff
%                  MATCH; 0 new D-numbers. Gate-3's FAIL was ONLY the 8 stale
%                  class-specificity pins [Test_p1_3x_xpath + Test_p2_3], fixed by
%                  Gate-4 test-pin update -- NOT a P4-2 defect).
%     * Scenario : validation\mat2doc\scenarios\s0022_p4_2_oxml_parfmt.{py,m}
%                  (the probe sequence replayed VERBATIM by runProbes() below).
%     * Frozen ref (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE):
%         references\s0022\probe.json -- copied verbatim (self-contained) into
%           tests\oxml\data\s0022_probe_oracle.json (the Equivalence replay set;
%           value/serhex JSON, jsondecode is line-ending agnostic -> no
%           `* binary` .gitattributes needed, per the s0020/s0021 precedent).
%         references\s0001\parts\word\{document,styles}.xml -- the M1 byte
%           references; NOT copied (the byte pins compare SHA-256 of what
%           mat2doc.Document().save() itself emits, so no fixture is needed).
%
%   Coverage taxonomy
%   -----------------
%   * Nominal      -- H11 canonical re-sort (o1), the M2-valued pStyle-first
%                     add_heading case (o2), CT_Ind/CT_Jc/CT_Spacing attr
%                     round-trips, CT_P pPr-first, text accessor.
%   * Edge         -- pStyle added LAST sorts FIRST (o3), the full 12-child
%                     scramble incl. NON-CONTIGUOUS slices (o4), signed w:right
%                     vs unsigned w:firstLine, hanging absent -> None, w:tab over
%                     a run -> "a\tb", leader=SPACES default omits @w:leader,
%                     spacing lineRule-absent-line-present -> MULTIPLE, the
%                     first_line_indent clear-both order, empty pPr accessors,
%                     clear_content keeps only pPr, jc_val/style/onoff removal.
%   * Equivalence  -- test_equivalence_full_battery_vs_frozen_oracle replays the
%                     ENTIRE s0022 battery live (runProbes, the .m twin's body
%                     verbatim -- incl. the document.xml + styles.xml
%                     parse->serialize round-trip legs) and flatten-compares every
%                     leaf to the frozen python-docx 1.2.0 oracle (0 divergences).
%   * Regression   -- hard-coded expected serialized-XML strings (ASCII where
%                     compared as strings, so string-equality is byte-identical L1)
%                     + UPPERCASE serhex of the raw UTF-8 shipping bytes vs the
%                     frozen oracle + SHA-256 of the two M1 parts.
%   * Upstream     -- the CT_PPr successor ordering, the CT_Ind signed/unsigned
%                     twips split, and the CT_TabStop leader default are the
%                     python-docx parfmt.py surface; the frozen oracle IS lxml's
%                     expected output for this API sequence.
%
%   Byte-level (L1) note: every serialized-XML comparison is either the FULL
%   serialize_part_xml output as an ASCII decoded string (string-equality ==
%   byte-equality L1) or its UPPERCASE hex (serhex) vs the frozen oracle. No
%   D-number granted any L2 relaxation in this WP (Gate-3 §4: zero new, none
%   exercised at L2), so every pin here is L1. The nPPr==204 count, the parse
%   count guards, and the equivalence key-count guard are the only looser-than-
%   byte checks and are commented at their site.
%
%   Determinism: no network, no absolute paths. The worktree root and the
%   co-located oracle resolve relative to this file via
%   fileparts(mfilename('fullpath')); saves go to tempname .docx deleted via
%   onCleanup; every file read is binary ('r','n'); no 'wt'. The tab char is
%   built via char(9) (source-encoding independent). The +mat2doc package
%   resolves via the MANDATORY PathFixture(worktree-root) in TestClassSetup
%   (WP9-F4 lesson).

    properties (Constant)
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        % --- frozen s0001 M1 byte references (the P4-2 CT_PPr/CT_P parse risk) ---
        DOC_SIZE    = 1548
        DOC_SHA     = "0e4dd503bc095219cd24e8312b372c588ca56bb834aacac53eafb7bf47836327"
        STYLES_SIZE = 349458
        STYLES_SHA  = "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"

        % --- canonical child order of CT_PPr's full-12-descriptor scramble (o4),
        %     frozen s0022 h11.o4_full_scramble.localnames: ANY scrambled add-order
        %     of these 12 must converge to exactly this sequence (pStyle first,
        %     sectPr last), driven by the NON-CONTIGUOUS _tag_seq successor slices. ---
        SCHEMA_ORDER12 = ["pStyle" "keepNext" "keepLines" "pageBreakBefore" ...
            "widowControl" "numPr" "tabs" "spacing" "ind" "jc" "outlineLvl" "sectPr"]
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied verbatim from tests\oxml\Test_p4_1b_oxml_run.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\oxml
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. CT_PPr H11 successor ordering (THE M2-ordering crux)          %
        % =============================================================== %

        function test_h11_ctppr_scrambled_reorder_to_schema(testCase)
            % Nominal + Edge + Regression (H11, s0022 h11.o1..o4): build a CT_PPr and
            % add ZeroOrOne children via get_or_add_* in SCRAMBLED order -- each must
            % re-sort to canonical schema order (child localnames AND serialized
            % bytes). Covers: o1 empty 4-child scramble; o2 the M2 add_heading path
            % (pStyle VALUED Heading1 added first, jc/ind/spacing valued); o3 pStyle
            % added LAST but sorting FIRST; o4 the full 12-child scramble exercising
            % the NON-CONTIGUOUS successor slices (widowControl@7 skips framePr@5,
            % etc.). Bytes pinned hard-coded (o1/o2/o4) AND vs the frozen oracle.
            oracle = loadOracle();

            % -- o1: jc -> pStyle -> ind -> spacing (empty) -> pStyle,spacing,ind,jc --
            p = newPPr(); p.get_or_add_jc(); p.get_or_add_pStyle(); p.get_or_add_ind(); p.get_or_add_spacing();
            testCase.verifyEqual(childLocalnames(p), ["pStyle" "spacing" "ind" "jc"], ...
                'o1 scramble must re-sort to schema order');
            testCase.verifyEqual(ser(p), ppr(testCase, "<w:pStyle/><w:spacing/><w:ind/><w:jc/>"), ...
                'o1 serialized bytes (L1) hard-coded');
            testCase.verifyEqual(hx_e(p), string(oracle.h11.o1_jc_pStyle_ind_spacing.serhex));

            % -- o2: the M2 add_heading path -- pStyle VALUED, added FIRST, sorts first --
            p = newPPr();
            ps = p.get_or_add_pStyle(); ps.val = "Heading1";
            j  = p.get_or_add_jc();      j.val   = mat2doc.enum.text.WD_ALIGN_PARAGRAPH.CENTER;
            ind = p.get_or_add_ind();    ind.left = mat2doc.shared.Twips(720);
            sp  = p.get_or_add_spacing();sp.after  = mat2doc.shared.Twips(120);
            testCase.verifyEqual(childLocalnames(p), ["pStyle" "spacing" "ind" "jc"], ...
                'o2 (M2 valued) child order');
            testCase.verifyEqual(ser(p), ppr(testCase, ...
                "<w:pStyle w:val=""Heading1""/><w:spacing w:after=""120""/>" + ...
                "<w:ind w:left=""720""/><w:jc w:val=""center""/>"), ...
                'o2 M2 add_heading serialized bytes (L1) hard-coded (pStyle-first)');
            testCase.verifyEqual(hx_e(p), string(oracle.h11.o2_m2_valued.serhex));

            % -- o3: pStyle added LAST but must sort FIRST --
            p = newPPr(); p.get_or_add_spacing(); p.get_or_add_jc(); p.get_or_add_ind(); p.get_or_add_pStyle();
            testCase.verifyEqual(childLocalnames(p), ["pStyle" "spacing" "ind" "jc"], ...
                'o3 pStyle added last must sort FIRST');
            testCase.verifyEqual(hx_e(p), string(oracle.h11.o3_pStyle_last.serhex), ...
                'o3 serialized bytes (L1) vs frozen oracle');

            % -- o4: full 12-child scramble across the tag_seq (NON-CONTIGUOUS slices) --
            p = newPPr();
            p.get_or_add_sectPr(); p.get_or_add_jc(); p.get_or_add_ind(); p.get_or_add_spacing();
            p.get_or_add_tabs(); p.get_or_add_numPr(); p.get_or_add_widowControl();
            p.get_or_add_pageBreakBefore(); p.get_or_add_keepLines(); p.get_or_add_keepNext();
            p.get_or_add_pStyle(); p.get_or_add_outlineLvl();
            testCase.verifyEqual(childLocalnames(p), testCase.SCHEMA_ORDER12, ...
                'o4 full-12 scramble must fully re-sort to canonical schema order');
            testCase.verifyEqual(ser(p), ppr(testCase, ...
                "<w:pStyle/><w:keepNext/><w:keepLines/><w:pageBreakBefore/><w:widowControl/>" + ...
                "<w:numPr/><w:tabs/><w:spacing/><w:ind/><w:jc/><w:outlineLvl/><w:sectPr/>"), ...
                'o4 serialized bytes (L1) hard-coded');
            testCase.verifyEqual(hx_e(p), string(oracle.h11.o4_full_scramble.serhex));
        end

        % =============================================================== %
        % 2. M1 byte-neutrality (parse-class flip: registration)          %
        % =============================================================== %

        function test_m1_styles_xml_byte_identical(testCase)
            % Regression (byte-neutrality, L1): mat2doc.Document().save() emits
            % word/styles.xml at EXACTLY 349458 B with the frozen s0001 SHA-256 --
            % byte-identical DESPITE its 204 real w:pPr now parsing to CT_PPr (P4-2's
            % specific parse-path risk). SHA-256 equality is a byte-level (L1)
            % assertion. (validate_P4-2 §3; Test_p1_8 owns the full 17/17 M1 sweep.)
            bytes = testCase.emitPart('styles.xml');
            testCase.verifyEqual(numel(bytes), testCase.STYLES_SIZE, ...
                sprintf('word/styles.xml must be exactly %d B after parfmt registration', ...
                    testCase.STYLES_SIZE));
            testCase.verifyEqual(sha256hex(bytes), testCase.STYLES_SHA, ...
                'word/styles.xml SHA-256 must equal the frozen s0001 oracle (byte-identical L1)');
        end

        function test_m1_document_xml_byte_identical(testCase)
            % Regression (byte-neutrality, L1): word/document.xml stays 1548 B with
            % the frozen s0001 SHA-256 -- the w:p->CT_P / w:pPr->CT_PPr registration
            % re-routes its parse but must not move a byte.
            bytes = testCase.emitPart('document.xml');
            testCase.verifyEqual(numel(bytes), testCase.DOC_SIZE, ...
                sprintf('word/document.xml must be exactly %d B', testCase.DOC_SIZE));
            testCase.verifyEqual(sha256hex(bytes), testCase.DOC_SHA, ...
                'word/document.xml SHA-256 must equal the frozen s0001 oracle (byte-identical L1)');
        end

        function test_styles_xml_parse_serialize_roundtrip_L1(testCase)
            % Regression (byte-neutrality, L1): parse the emitted styles.xml back
            % through mat2doc.oxml.parse_xml (instantiating 204 CT_PPr and their
            % children) and re-serialize -- the output must be byte-identical to the
            % input. This directly exercises the CT_PPr parse->serialize path at part
            % scale (the exact P4-2 byte risk), independently of save().
            inBytes  = testCase.emitPart('styles.xml');
            root     = mat2doc.oxml.parse_xml(inBytes);
            outBytes = mat2doc.oxml.serialize_part_xml(root);
            testCase.verifyEqual(uint8(outBytes(:)'), uint8(inBytes(:)'), ...
                'styles.xml parse->serialize must be byte-identical (CT_PPr path is byte-neutral)');
            % Count guard (looser-than-byte, justified): the frozen default template
            % ships EXACTLY 204 w:pPr; pinning it proves the CT_PPr path was actually
            % exercised at scale, not that a short part slipped through. (Gate-3 §3.)
            nPPr = numel(root.xpath('.//w:pPr'));
            testCase.verifyEqual(nPPr, 204, ...
                'styles.xml must parse exactly 204 real w:pPr (CT_PPr path exercised at scale)');
        end

        % =============================================================== %
        % 3. CT_TabStop + str_->"\t" carry-forward (P4-1b gap CLOSED)      %
        % =============================================================== %

        function test_ct_tabstop_and_carry_forward(testCase)
            % Nominal + Edge + Regression (s0022 tabstop): CT_TabStop attribute
            % serialization (val/pos + leader default D-delta-1 omit / DOTS write),
            % str_()=="\t", the CARRY-FORWARD closure (CT_R.text AND CT_P.text over a
            % run containing w:tab -> "a\tb"), and CT_TabStops.insert_tab_in_order
            % sorting by pos (720,240,480 inserted -> 240,480,720 serialized). Bytes
            % pinned hard-coded AND vs the frozen oracle.
            oracle = loadOracle();
            TAB = string(char(9));

            % -- tab attrs: leader=SPACES default OMITS @w:leader (D-delta-1) --
            tab = mat2doc.oxml.OxmlElement("w:tab");
            tab.val = mat2doc.enum.text.WD_TAB_ALIGNMENT.LEFT;
            tab.pos = mat2doc.shared.Twips(720);
            tab.leader = mat2doc.enum.text.WD_TAB_LEADER.SPACES;
            testCase.verifyEqual(ser(tab), decl() + newline + ...
                "<w:tab xmlns:w=""" + testCase.W + """ w:val=""left"" w:pos=""720""/>", ...
                'leader=SPACES (default) must OMIT @w:leader (D-delta-1)');
            testCase.verifyEqual(hx_e(tab), string(oracle.tabstop.tab_leader_default.serhex));
            testCase.verifyEqual(string(tab.val), "LEFT");
            testCase.verifyEqual(string(tab.leader), "SPACES", 'leader read-when-absent -> SPACES');

            % -- leader=DOTS WRITES @w:leader="dot" --
            tab = mat2doc.oxml.OxmlElement("w:tab");
            tab.val = mat2doc.enum.text.WD_TAB_ALIGNMENT.RIGHT;
            tab.pos = mat2doc.shared.Twips(1440);
            tab.leader = mat2doc.enum.text.WD_TAB_LEADER.DOTS;
            testCase.verifyEqual(ser(tab), decl() + newline + ...
                "<w:tab xmlns:w=""" + testCase.W + """ w:val=""right"" w:pos=""1440"" w:leader=""dot""/>", ...
                'leader=DOTS must write @w:leader="dot"');
            testCase.verifyEqual(string(tab.leader), "DOTS");

            % -- str_() == "\t" (char 9, not the literal two-char "\t") --
            testCase.verifyEqual(mat2doc.oxml.OxmlElement("w:tab").str_(), TAB, ...
                'CT_TabStop.str_() must be a real TAB char(9)');

            % -- CARRY-FORWARD: CT_R.text AND CT_P.text over w:tab -> "a\tb" --
            r = rparse("<w:r xmlns:w=""" + testCase.W + """><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r>");
            testCase.verifyEqual(r.text, "a" + TAB + "b", ...
                'CT_R.text over w:tab -> "a\tb" (P4-1b gap closed by w:tab->CT_TabStop)');
            testCase.verifyEqual(lnsCell(r), {'t','tab','t'}, 'run child sequence [t,tab,t]');
            pWithTab = rparse("<w:p xmlns:w=""" + testCase.W + ...
                """><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p>");
            testCase.verifyEqual(pWithTab.text, "a" + TAB + "b", ...
                'CT_P.text over a run containing w:tab -> "a\tb"');

            % -- insert_tab_in_order: inserted 720,240,480 -> serialized 240,480,720 --
            tabs = mat2doc.oxml.OxmlElement("w:tabs");
            tabs.insert_tab_in_order(mat2doc.shared.Twips(720), ...
                mat2doc.enum.text.WD_TAB_ALIGNMENT.LEFT, mat2doc.enum.text.WD_TAB_LEADER.SPACES);
            tabs.insert_tab_in_order(mat2doc.shared.Twips(240), ...
                mat2doc.enum.text.WD_TAB_ALIGNMENT.LEFT, mat2doc.enum.text.WD_TAB_LEADER.SPACES);
            tabs.insert_tab_in_order(mat2doc.shared.Twips(480), ...
                mat2doc.enum.text.WD_TAB_ALIGNMENT.CENTER, mat2doc.enum.text.WD_TAB_LEADER.DOTS);
            testCase.verifyEqual(ser(tabs), decl() + newline + ...
                "<w:tabs xmlns:w=""" + testCase.W + """>" + ...
                "<w:tab w:pos=""240"" w:val=""left""/>" + ...
                "<w:tab w:pos=""480"" w:val=""center"" w:leader=""dot""/>" + ...
                "<w:tab w:pos=""720"" w:val=""left""/></w:tabs>", ...
                'insert_tab_in_order must sort by pos (240,480,720)');
            testCase.verifyEqual(hx_e(tabs), string(oracle.tabstop.insert_in_order.serhex));
        end

        % =============================================================== %
        % 4. CT_Ind / CT_Jc / CT_Spacing child attribute serialization    %
        % =============================================================== %

        function test_ct_ind_jc_spacing_children(testCase)
            % Regression (s0022 children): CT_Ind left/right (signed
            % ST_SignedTwipsMeasure -> w:right="-60") vs firstLine/hanging (unsigned
            % ST_TwipsMeasure); hanging absent -> None. CT_Jc val=CENTER -> "center".
            % CT_Spacing after/before/line + lineRule=EXACTLY -> "exact". Bytes pinned
            % hard-coded AND vs the frozen oracle; EMU reads spot-checked.
            oracle = loadOracle();

            ind = mat2doc.oxml.OxmlElement("w:ind");
            ind.left = mat2doc.shared.Twips(720);
            ind.right = mat2doc.shared.Twips(-60);        % SIGNED
            ind.firstLine = mat2doc.shared.Twips(240);    % UNSIGNED
            testCase.verifyEqual(ser(ind), decl() + newline + ...
                "<w:ind xmlns:w=""" + testCase.W + """ w:left=""720"" w:right=""-60"" w:firstLine=""240""/>", ...
                'CT_Ind signed w:right vs unsigned w:firstLine serialized (L1)');
            testCase.verifyEqual(hx_e(ind), string(oracle.children.ind_all.serhex));
            testCase.verifyEqual(round(double(ind.right)), -38100, 'w:right=-60tw -> -38100 EMU (signed)');
            testCase.verifyTrue(isequal(ind.hanging, []), 'absent w:hanging -> [] (None)');

            jc = mat2doc.oxml.OxmlElement("w:jc");
            jc.val = mat2doc.enum.text.WD_ALIGN_PARAGRAPH.CENTER;
            testCase.verifyEqual(ser(jc), decl() + newline + ...
                "<w:jc xmlns:w=""" + testCase.W + """ w:val=""center""/>");
            testCase.verifyEqual(hx_e(jc), string(oracle.children.jc_center.serhex));
            testCase.verifyEqual(string(jc.val), "CENTER");

            sp = mat2doc.oxml.OxmlElement("w:spacing");
            sp.after = mat2doc.shared.Twips(120); sp.before = mat2doc.shared.Twips(60);
            sp.line = mat2doc.shared.Twips(360);
            sp.lineRule = mat2doc.enum.text.WD_LINE_SPACING.EXACTLY;
            testCase.verifyEqual(ser(sp), decl() + newline + ...
                "<w:spacing xmlns:w=""" + testCase.W + """ w:after=""120"" w:before=""60"" " + ...
                "w:line=""360"" w:lineRule=""exact""/>", ...
                'CT_Spacing all-attrs + lineRule=EXACTLY -> "exact" serialized (L1)');
            testCase.verifyEqual(hx_e(sp), string(oracle.children.spacing_all.serhex));
            testCase.verifyEqual(string(sp.lineRule), "EXACTLY");
        end

        % =============================================================== %
        % 5. CT_PPr 13 @property accessors                                %
        % =============================================================== %

        function test_ct_ppr_accessors(testCase)
            % Regression (s0022 accessors): the 13 CT_PPr @property accessors --
            % first_line_indent (neg -> w:hanging, pos -> w:firstLine, [] clears BOTH
            % keeping an empty w:ind), ind_left/ind_right, jc_val (+[] removes w:jc,
            % empty pPr), the CT_OnOff quartet keepLines/keepNext/pageBreakBefore=True
            % + widowControl=False (w:val="0") and keepNext=[] removes the child,
            % spacing_after/before/line/lineRule + the lineRule-ABSENT-line-PRESENT ->
            % MULTIPLE rule, style (+[] removes w:pStyle). Bytes pinned hard-coded.
            oracle = loadOracle();

            % first_line_indent: neg -> w:hanging="360"
            p = newPPr(); p.first_line_indent = mat2doc.shared.Twips(-360);
            testCase.verifyEqual(round(double(p.first_line_indent)), -228600, 'fli neg read EMU');
            testCase.verifyEqual(ser(p), ppr(testCase, "<w:ind w:hanging=""360""/>"), 'fli_neg -> w:hanging');
            % pos -> w:firstLine="360"
            p = newPPr(); p.first_line_indent = mat2doc.shared.Twips(360);
            testCase.verifyEqual(ser(p), ppr(testCase, "<w:ind w:firstLine=""360""/>"), 'fli_pos -> w:firstLine');
            % [] clears BOTH firstLine & hanging, keeping an empty w:ind (the clear order)
            p.first_line_indent = [];
            testCase.verifyTrue(isequal(p.first_line_indent, []), 'fli cleared -> None');
            testCase.verifyEqual(ser(p), ppr(testCase, "<w:ind/>"), 'fli_clear keeps empty w:ind');

            % ind_left / ind_right (signed)
            p = newPPr(); p.ind_left = mat2doc.shared.Twips(720); p.ind_right = mat2doc.shared.Twips(-45);
            testCase.verifyEqual(round(double(p.ind_left)), 457200);
            testCase.verifyEqual(round(double(p.ind_right)), -28575, 'ind_right signed');
            testCase.verifyEqual(hx_e(p), string(oracle.accessors.ind_lr_serhex), ...
                'ind_left/right serialized bytes (L1) vs frozen oracle');

            % jc_val + clear
            p = newPPr(); p.jc_val = mat2doc.enum.text.WD_ALIGN_PARAGRAPH.RIGHT;
            testCase.verifyEqual(string(p.jc_val), "RIGHT");
            p.jc_val = [];
            testCase.verifyTrue(isequal(p.jc_val, []), 'jc_val cleared -> None');
            testCase.verifyEqual(childLocalnames(p), strings(1,0), 'jc_val=[] removes w:jc (empty pPr)');

            % CT_OnOff quartet
            p = newPPr();
            p.keepLines_val = true; p.keepNext_val = true;
            p.pageBreakBefore_val = true; p.widowControl_val = false;
            testCase.verifyTrue(p.keepLines_val); testCase.verifyTrue(p.keepNext_val);
            testCase.verifyTrue(p.pageBreakBefore_val); testCase.verifyFalse(p.widowControl_val);
            testCase.verifyEqual(ser(p), ppr(testCase, ...
                "<w:keepNext/><w:keepLines/><w:pageBreakBefore/><w:widowControl w:val=""0""/>"), ...
                'onoff quartet serialized (L1): True bare, False -> w:val="0"');
            testCase.verifyEqual(hx_e(p), string(oracle.accessors.onoff.serhex));
            p.keepNext_val = [];
            testCase.verifyTrue(isequal(p.keepNext_val, []), 'keepNext=[] removes child -> None');

            % spacing accessors + lineRule-absent -> MULTIPLE
            p = newPPr();
            p.spacing_after = mat2doc.shared.Twips(200); p.spacing_before = mat2doc.shared.Twips(100);
            p.spacing_line = mat2doc.shared.Twips(276);
            p.spacing_lineRule = mat2doc.enum.text.WD_LINE_SPACING.AT_LEAST;
            testCase.verifyEqual(string(p.spacing_lineRule), "AT_LEAST");
            testCase.verifyEqual(ser(p), ppr(testCase, ...
                "<w:spacing w:after=""200"" w:before=""100"" w:line=""276"" w:lineRule=""atLeast""/>"), ...
                'spacing accessors serialized (L1)');
            testCase.verifyEqual(hx_e(p), string(oracle.accessors.spacing.serhex));
            p2 = newPPr(); p2.spacing_line = mat2doc.shared.Twips(360);   % line present, lineRule absent
            testCase.verifyEqual(string(p2.spacing_lineRule), "MULTIPLE", ...
                'line present + lineRule absent -> MULTIPLE (parfmt.py 296-298)');

            % style + clear
            p = newPPr(); p.style = "Heading1";
            testCase.verifyEqual(p.style, "Heading1");
            p.style = [];
            testCase.verifyTrue(isequal(p.style, []), 'style cleared -> None');
            testCase.verifyEqual(childLocalnames(p), strings(1,0), 'style=[] removes w:pStyle (empty pPr)');
        end

        % =============================================================== %
        % 6. CT_P (paragraph.py surface)                                  %
        % =============================================================== %

        function test_ct_p_surface(testCase)
            % Nominal + Edge + Regression (s0022 ct_p): CT_P pPr-first (_insert_pPr
            % index 0 -- append w:r THEN set style adds pPr at FRONT), text accessor
            % over runs ("Hello world") and over a w:tab run ("a\tb"), alignment
            % get/set via jc_val, clear_content (keeps only pPr), add_p_before (body
            % count 2), set_sectPr (grandchild order pStyle,sectPr under pPr). Bytes
            % pinned hard-coded AND vs the frozen oracle.
            oracle = loadOracle();
            TAB = string(char(9));

            % pPr-first: add_r THEN style -> child order [pPr, r]
            p = mat2doc.oxml.OxmlElement("w:p"); p.add_r(); p.style = "Heading1";
            testCase.verifyEqual(childLocalnames(p), ["pPr" "r"], ...
                'CT_P _insert_pPr forces pPr to index 0 (before the pre-existing w:r)');
            testCase.verifyEqual(ser(p), decl() + newline + ...
                "<w:p xmlns:w=""" + testCase.W + """><w:pPr><w:pStyle w:val=""Heading1""/></w:pPr><w:r/></w:p>", ...
                'CT_P pPr-first serialized bytes (L1) hard-coded');
            testCase.verifyEqual(hx_e(p), string(oracle.ct_p.pPr_first.serhex));

            % text accessor
            p = rparse("<w:p xmlns:w=""" + testCase.W + """>" + ...
                "<w:r><w:t>Hello</w:t></w:r><w:r><w:t> world</w:t></w:r></w:p>");
            testCase.verifyEqual(p.text, "Hello world", 'CT_P.text joins run text');
            p = rparse("<w:p xmlns:w=""" + testCase.W + """><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p>");
            testCase.verifyEqual(p.text, "a" + TAB + "b", 'CT_P.text over w:tab -> "a\tb"');

            % alignment get/set (via jc_val)
            p = mat2doc.oxml.OxmlElement("w:p"); p.alignment = mat2doc.enum.text.WD_ALIGN_PARAGRAPH.CENTER;
            testCase.verifyEqual(string(p.alignment), "CENTER");
            testCase.verifyEqual(ser(p), decl() + newline + ...
                "<w:p xmlns:w=""" + testCase.W + """><w:pPr><w:jc w:val=""center""/></w:pPr></w:p>", ...
                'CT_P.alignment set serialized bytes (L1)');

            % clear_content keeps only pPr
            p = rparse("<w:p xmlns:w=""" + testCase.W + """>" + ...
                "<w:pPr><w:pStyle w:val=""Heading1""/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>");
            p.clear_content();
            testCase.verifyEqual(childLocalnames(p), "pPr", 'clear_content keeps only w:pPr');
            testCase.verifyEqual(hx_e(p), string(oracle.ct_p.clear_content.serhex));

            % add_p_before inside a body -> 2 paragraphs
            body = rparse("<w:body xmlns:w=""" + testCase.W + """><w:p><w:r><w:t>orig</w:t></w:r></w:p></w:body>");
            origs = body.xpath("./w:p"); origs(1).add_p_before();
            testCase.verifyEqual(numel(body.xpath("./w:p")), 2, 'add_p_before -> 2 body paragraphs');

            % set_sectPr -> pPr grandchildren [pStyle, sectPr]
            p = mat2doc.oxml.OxmlElement("w:p"); p.style = "Heading1";
            p.set_sectPr(mat2doc.oxml.OxmlElement("w:sectPr"));
            testCase.verifyEqual(childLocalnames(p.pPr), ["pStyle" "sectPr"], ...
                'set_sectPr appends w:sectPr LAST under pPr (after pStyle)');
            testCase.verifyEqual(ser(p), decl() + newline + ...
                "<w:p xmlns:w=""" + testCase.W + """><w:pPr><w:pStyle w:val=""Heading1""/>" + ...
                "<w:sectPr/></w:pPr></w:p>", 'set_sectPr serialized bytes (L1) hard-coded');
            testCase.verifyEqual(hx_e(p), string(oracle.ct_p.set_sectPr.serhex));
        end

        % =============================================================== %
        % 7. EQUIVALENCE -- full s0022 battery vs the frozen oracle        %
        % =============================================================== %

        function test_equivalence_full_battery_vs_frozen_oracle(testCase)
            % Equivalence: replay the ENTIRE s0022 battery (runProbes -- the .m twin's
            % body VERBATIM: h11 / children / tabstop / accessors / ct_p + the
            % document.xml & styles.xml parse->serialize round-trip legs, the latter
            % over what Document().save() itself emits) and flatten-compare EVERY leaf
            % to the frozen python-docx 1.2.0 oracle copied into
            % data\s0022_probe_oracle.json. Gate-3 found ZERO divergences, so every
            % leaf must be byte/value-identical. Ties the suite to the Gate-3 output.
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
            % Non-trivial size guard (guards a silent-empty replay).
            testCase.verifyGreaterThan(numel(oKeys), 40, ...
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
            % both cleaned up on exit. (Idiom from Test_p4_1b_oxml_run.m; tempname
            % paths are absolute so no cwd handling is needed.)
            bytes = emitDocPart(partLeaf);
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

function p = newPPr()
    p = mat2doc.oxml.OxmlElement("w:pPr");
end

function e = rparse(xml)
    e = mat2doc.oxml.parse_xml(uint8(unicode2native(char(xml), "UTF-8")));
end

function s = ser(e)
    % Decode serialize_part_xml bytes to a string (ASCII content -> string-equality
    % is a byte-identical L1 assertion).
    s = string(native2unicode(mat2doc.oxml.serialize_part_xml(e), "UTF-8"));
end

function h = hx(raw)
    % UPPERCASE hex of raw UTF-8 bytes (matches Python bytes.hex().upper()).
    h = string(sprintf('%02X', uint8(raw)));
end

function h = hx_e(e)
    h = hx(mat2doc.oxml.serialize_part_xml(e));
end

function s = ppr(testCase, body)
    % decl + newline + <w:pPr xmlns:w="...">BODY</w:pPr>.
    s = decl() + newline + "<w:pPr xmlns:w=""" + testCase.W + """>" + string(body) + "</w:pPr>";
end

function names = childLocalnames(e)
    kids = e.xpath("./*");
    names = strings(1, numel(kids));
    for k = 1:numel(kids)
        names(k) = string(kids(k).local_part);
    end
end

function C = lnsCell(e)
    % Ordered child localnames as a 1xN cell row of char ({} when no children).
    kids = e.xpath("./*");
    if isempty(kids)
        C = cell(1, 0);
        return
    end
    C = cell(1, numel(kids));
    for k = 1:numel(kids)
        C{k} = char(kids(k).local_part);
    end
end

function s = rv(x)
    % Uniform accessor repr mirroring the s0022 rv(): None->"None", enum->member
    % name, bool->"True"/"False", Length/int->decimal EMU string.
    if isequal(x, [])
        s = "None";
    elseif isa(x, 'mat2doc.enum.base.BaseXmlEnum')
        s = string(x);
    elseif islogical(x)
        if x, s = "True"; else, s = "False"; end
    elseif isnumeric(x)
        s = string(sprintf('%.0f', double(x)));
    else
        s = string(x);
    end
end

function s = tf(b)
    if b, s = "True"; else, s = "False"; end
end

function o = loadOracle()
    % Read the co-located frozen oracle in BINARY mode (no CRLF translation) and
    % decode UTF-8 -> struct. jsondecode is whitespace/line-ending agnostic, so no
    % `* binary` .gitattributes pin is needed (value-based fixture, s0021 precedent).
    here = fileparts(mfilename('fullpath'));
    p = fullfile(here, 'data', 's0022_probe_oracle.json');
    fid = fopen(p, 'r', 'n');
    if fid < 0
        error('mat2doc:test:openFail', 'cannot open frozen oracle %s', p);
    end
    raw = fread(fid, Inf, '*uint8')';
    fclose(fid);
    o = jsondecode(native2unicode(raw, 'UTF-8'));
end

function bytes = emitDocPart(partLeaf)
    % Document().save() to a temp .docx, unzip, return word/<partLeaf> raw bytes.
    d = mat2doc.Document();
    tmp = [tempname '.docx'];
    cleanTmp = onCleanup(@() deleteIfExists(tmp)); %#ok<NASGU>
    d.save(tmp);
    exdir = tempname;
    cleanEx = onCleanup(@() rmdirIfExists(exdir)); %#ok<NASGU>
    unzip(tmp, exdir);
    bytes = readBytes(fullfile(exdir, 'word', partLeaf));
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

function P = runProbes()
    % Replay the s0022 probe sequence (the .m twin's P-struct body, VERBATIM
    % tags/inputs/order) and return the nested struct of tagged canonical values.
    % Embedded here so the Equivalence leg is self-contained (the validation-folder
    % scenario is NOT on the toolbox path and must not be a dependency). Mirrors
    % validation\mat2doc\scenarios\s0022_p4_2_oxml_parfmt.m lines 27-191. The two
    % round-trip legs read the parts that Document().save() itself emits (rather
    % than external file args), so no fixture path is needed.
    T  = @(tw) mat2doc.shared.Twips(tw);
    AP = @(name) mat2doc.enum.text.WD_ALIGN_PARAGRAPH.(name);
    LS = @(name) mat2doc.enum.text.WD_LINE_SPACING.(name);
    TA = @(name) mat2doc.enum.text.WD_TAB_ALIGNMENT.(name);
    TL = @(name) mat2doc.enum.text.WD_TAB_LEADER.(name);

    P = struct();

    % ===================== h11: CT_PPr scrambled order (CRUX) ===============
    h11 = struct();
    p = newPPr(); p.get_or_add_jc(); p.get_or_add_pStyle(); p.get_or_add_ind(); p.get_or_add_spacing();
    h11.o1_jc_pStyle_ind_spacing = struct("localnames", {lnsCell(p)}, "serhex", hx_e(p));
    p = newPPr();
    ps = p.get_or_add_pStyle(); ps.val = "Heading1";
    j = p.get_or_add_jc(); j.val = AP("CENTER");
    ind = p.get_or_add_ind(); ind.left = T(720);
    sp = p.get_or_add_spacing(); sp.after = T(120);
    h11.o2_m2_valued = struct("localnames", {lnsCell(p)}, "serhex", hx_e(p));
    p = newPPr(); p.get_or_add_spacing(); p.get_or_add_jc(); p.get_or_add_ind(); p.get_or_add_pStyle();
    h11.o3_pStyle_last = struct("localnames", {lnsCell(p)}, "serhex", hx_e(p));
    p = newPPr();
    p.get_or_add_sectPr(); p.get_or_add_jc(); p.get_or_add_ind(); p.get_or_add_spacing();
    p.get_or_add_tabs(); p.get_or_add_numPr(); p.get_or_add_widowControl();
    p.get_or_add_pageBreakBefore(); p.get_or_add_keepLines(); p.get_or_add_keepNext();
    p.get_or_add_pStyle(); p.get_or_add_outlineLvl();
    h11.o4_full_scramble = struct("localnames", {lnsCell(p)}, "serhex", hx_e(p));
    P.h11 = h11;

    % ===================== children: CT_Ind / CT_Jc / CT_Spacing ============
    ch = struct();
    ind = mat2doc.oxml.OxmlElement("w:ind");
    ind.left = T(720); ind.right = T(-60); ind.firstLine = T(240);
    ch.ind_all = struct("serhex", hx_e(ind), "left", rv(ind.left), ...
        "right", rv(ind.right), "firstLine", rv(ind.firstLine), "hanging", rv(ind.hanging));
    jc = mat2doc.oxml.OxmlElement("w:jc"); jc.val = AP("CENTER");
    ch.jc_center = struct("serhex", hx_e(jc), "val", rv(jc.val));
    sp = mat2doc.oxml.OxmlElement("w:spacing");
    sp.after = T(120); sp.before = T(60); sp.line = T(360); sp.lineRule = LS("EXACTLY");
    ch.spacing_all = struct("serhex", hx_e(sp), "after", rv(sp.after), ...
        "before", rv(sp.before), "line", rv(sp.line), "lineRule", rv(sp.lineRule));
    P.children = ch;

    % ===================== tabstop: CT_TabStop + carry-forward + tabs =======
    ts = struct();
    tab = mat2doc.oxml.OxmlElement("w:tab");
    tab.val = TA("LEFT"); tab.pos = T(720); tab.leader = TL("SPACES");
    ts.tab_leader_default = struct("serhex", hx_e(tab), "val", rv(tab.val), ...
        "pos", rv(tab.pos), "leader_readback", rv(tab.leader));
    tab = mat2doc.oxml.OxmlElement("w:tab");
    tab.val = TA("RIGHT"); tab.pos = T(1440); tab.leader = TL("DOTS");
    ts.tab_leader_dots = struct("serhex", hx_e(tab), "leader_readback", rv(tab.leader));
    ts.tab_str = mat2doc.oxml.OxmlElement("w:tab").str_();
    r = rparse("<w:r xmlns:w=""" + W_() + """><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r>");
    ts.ct_r_text_over_tab = struct("text", r.text, "localnames", {lnsCell(r)});
    tabs = mat2doc.oxml.OxmlElement("w:tabs");
    tabs.insert_tab_in_order(T(720), TA("LEFT"), TL("SPACES"));
    tabs.insert_tab_in_order(T(240), TA("LEFT"), TL("SPACES"));
    tabs.insert_tab_in_order(T(480), TA("CENTER"), TL("DOTS"));
    tl = tabs.tab_lst; pos = cell(1, numel(tl));
    for k = 1:numel(tl); pos{k} = rv(tl(k).pos); end
    ts.insert_in_order = struct("positions", {pos}, "serhex", hx_e(tabs));
    P.tabstop = ts;

    % ===================== accessors: all 13 CT_PPr @property ================
    ac = struct();
    p = newPPr(); p.first_line_indent = T(-360);
    ac.fli_neg = struct("read", rv(p.first_line_indent), "serhex", hx_e(p));
    p = newPPr(); p.first_line_indent = T(360);
    ac.fli_pos = struct("read", rv(p.first_line_indent), "serhex", hx_e(p));
    p.first_line_indent = [];
    ac.fli_clear = struct("read", rv(p.first_line_indent), "serhex", hx_e(p));
    p = newPPr(); p.ind_left = T(720); p.ind_right = T(-45);
    ac.ind_left = rv(p.ind_left);
    ac.ind_right = rv(p.ind_right);
    ac.ind_lr_serhex = hx_e(p);
    p = newPPr(); p.jc_val = AP("RIGHT");
    ac.jc_val = rv(p.jc_val);
    p.jc_val = [];
    ac.jc_val_cleared = struct("read", rv(p.jc_val), "localnames", {lnsCell(p)});
    p = newPPr();
    p.keepLines_val = true; p.keepNext_val = true; p.pageBreakBefore_val = true; p.widowControl_val = false;
    ac.onoff = struct("keepLines", rv(p.keepLines_val), "keepNext", rv(p.keepNext_val), ...
        "pageBreakBefore", rv(p.pageBreakBefore_val), "widowControl", rv(p.widowControl_val), ...
        "serhex", hx_e(p));
    p.keepNext_val = [];
    ac.keepNext_cleared = rv(p.keepNext_val);
    p = newPPr();
    p.spacing_after = T(200); p.spacing_before = T(100); p.spacing_line = T(276); p.spacing_lineRule = LS("AT_LEAST");
    ac.spacing = struct("after", rv(p.spacing_after), "before", rv(p.spacing_before), ...
        "line", rv(p.spacing_line), "lineRule", rv(p.spacing_lineRule), "serhex", hx_e(p));
    p = newPPr(); p.spacing_line = T(360);
    ac.lineRule_absent_multiple = rv(p.spacing_lineRule);
    p = newPPr(); p.style = "Heading1";
    ac.style = rv(p.style);
    p.style = [];
    ac.style_cleared = struct("read", rv(p.style), "localnames", {lnsCell(p)});
    P.accessors = ac;

    % ===================== ct_p: CT_P element ===============================
    cp = struct();
    p = mat2doc.oxml.OxmlElement("w:p"); p.add_r(); p.style = "Heading1";
    cp.pPr_first = struct("localnames", {lnsCell(p)}, "serhex", hx_e(p));
    p = rparse("<w:p xmlns:w=""" + W_() + """><w:r><w:t>Hello</w:t></w:r><w:r><w:t> world</w:t></w:r></w:p>");
    cp.text = p.text;
    p = rparse("<w:p xmlns:w=""" + W_() + """><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p>");
    cp.text_with_tab = p.text;
    p = mat2doc.oxml.OxmlElement("w:p"); p.alignment = AP("CENTER");
    cp.alignment = struct("read", rv(p.alignment), "serhex", hx_e(p));
    p = rparse("<w:p xmlns:w=""" + W_() + """><w:pPr><w:pStyle w:val=""Heading1""/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>");
    p.clear_content();
    cp.clear_content = struct("localnames", {lnsCell(p)}, "serhex", hx_e(p));
    body = rparse("<w:body xmlns:w=""" + W_() + """><w:p><w:r><w:t>orig</w:t></w:r></w:p></w:body>");
    origs = body.xpath("./w:p"); origs(1).add_p_before();
    cp.add_p_before = struct("body_p_count", numel(body.xpath("./w:p")));
    p = mat2doc.oxml.OxmlElement("w:p"); p.style = "Heading1";
    p.set_sectPr(mat2doc.oxml.OxmlElement("w:sectPr"));
    cp.set_sectPr = struct("pPr_localnames", {lnsCell(p.pPr)}, "serhex", hx_e(p));
    P.ct_p = cp;

    % ===================== document.xml parse->serialize L1 =================
    dr = struct();
    db = emitDocPart('document.xml');
    elm = mat2doc.oxml.parse_xml(db);
    out = mat2doc.oxml.serialize_part_xml(elm);
    dr.roundtrip_byte_identical = tf(isequal(uint8(out(:)'), uint8(db(:)')));
    dr.n_wp  = numel(elm.xpath(".//w:p"));
    dr.n_pPr = numel(elm.xpath(".//w:pPr"));
    P.document_roundtrip = dr;

    % ===================== styles.xml parse->serialize L1 (204 pPr) =========
    sr = struct();
    sb = emitDocPart('styles.xml');
    elm = mat2doc.oxml.parse_xml(sb);
    out = mat2doc.oxml.serialize_part_xml(elm);
    sr.roundtrip_byte_identical = tf(isequal(uint8(out(:)'), uint8(sb(:)')));
    sr.n_pPr = numel(elm.xpath(".//w:pPr"));
    P.styles_roundtrip = sr;
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

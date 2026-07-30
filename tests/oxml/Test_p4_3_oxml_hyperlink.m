classdef Test_p4_3_oxml_hyperlink < matlab.unittest.TestCase
% TEST_P4_3_OXML_HYPERLINK  Gate-4 permanent unit tests for Mat2Doc P4-3
%   (src/docx/oxml/text/hyperlink.py -> +mat2doc\+oxml\+text\CT_Hyperlink and
%   src/docx/oxml/text/pagebreak.py -> +mat2doc\+oxml\+text\CT_LastRenderedPageBreak,
%   plus the two registry rows w:hyperlink -> CT_Hyperlink and
%   w:lastRenderedPageBreak -> CT_LastRenderedPageBreak).
%
%   P4-3 is the LAST oxml WP of P4. This class permanently freezes the guarantees
%   the prior gates established (Porter Gate-1 + Opus Gate-2 adversarial APPROVE
%   zero-findings; Gate-3 equivalence PASS-DEVIATION -- one contrived, unreachable
%   serializer namespace-declaration residual ACCEPTED, see the s7 pin below):
%
%     (CT_LastRenderedPageBreak DETECTION -- the meat) precedes_all_content /
%     follows_all_content over 8 hand-built trees exercising every branch of
%     pagebreak.py:52-99 -- the ./w:r[1] first-run predicate, the (./w:r)[last()]
%     last-run predicate, the hyperlink short-circuit (atomic), the empty-run-1 /
%     break-start-of-run-2 subtlety (t5), the two-breaks n_breaks=2 count (t6), and
%     the contrived sole-content T/T edge (t7). Assert vs the frozen s0023 oracle.
%
%     (FRAGMENT SPLIT -- byte level) preceding_fragment_p / following_fragment_p
%     produce loose CT_P fragments whose .text AND serialized bytes (serhex) are
%     pinned on the 6 split-relevant trees, including the atomic-hyperlink split
%     ("prexy"/"post" -- whole w:hyperlink into the preceding fragment, break
%     removed from inside) and the w:pPr-retained-in-both-fragments case (t8).
%
%     (SECOND-BREAK GUARD) preceding_/following_fragment_p on a NON-first break both
%     raise ValueError -- identifier mat2doc:ValueError + the verbatim python-docx
%     message.
%
%     (CT_Hyperlink H3 TRI-STATE) the 8-step rId/anchor/history mutation sequence,
%     serhex-pinned at every step, byte-identical to lxml EXCEPT step s7 -- see the
%     KNOWN-DEVIATION pin in test_ct_hyperlink_h3_tristate. history has a NON-None
%     default=True (the docx-delta tri-state).
%
%     (CT_Hyperlink.text / lastRenderedPageBreaks) text over direct child runs
%     (w:br -> "\n"); lastRenderedPageBreaks count via ./w:r/w:lastRenderedPageBreak.
%
%     (CT_P.text-over-hyperlink closure -- P4-2 VERIFY) CT_P.text on a paragraph
%     bearing a w:hyperlink returns the concatenated hyperlink run text; registering
%     w:hyperlink alone (no CT_P edit) closes this. Pinned so a revert goes RED.
%
%     (M1 BYTE-NEUTRALITY -- light) the two new registry rows are a pure lookup
%     addition: the default template carries ZERO w:hyperlink and ZERO
%     w:lastRenderedPageBreak, so neither new class is instantiated during
%     Document().save(), and every part exits the identical serialize_part_xml walk.
%     document.xml (1548 B) + styles.xml (349458 B) are pinned byte-identical here;
%     the full 17/17 M1 sweep stays owned by Test_p1_8_skeleton_m1.
%
%   ------------------------------------------------------------------------------
%   THE s7 KNOWN-DEVIATION (ACCEPTED -- see the decision doc)
%   ------------------------------------------------------------------------------
%   On a LOOSE created w:hyperlink serialized STANDALONE: setting rId mints an
%   xmlns:ns0 for the relationships URI (lxml auto-prefix); after rId is CLEARED
%   (s7) lxml's sticky nsmap RETAINS the now-orphaned, unused xmlns:ns0, whereas the
%   Mat2Doc serializer recomputes the actually-used namespaces and DROPS it. The two
%   forms are the SAME empty w:hyperlink (no attributes, no rId/anchor/history),
%   exclusive-C14N-equal; the difference is namespace-DECLARATION-emission only. It
%   is a generation-side (created-element, minted-then-orphaned) manifestation of the
%   signed D-serializer-nsdecl class -- NO new D-number -- and is UNREACHABLE on
%   every real path (real hyperlinks keep their rId; document-tree hyperlinks live
%   under a root already declaring xmlns:r so never mint ns0; template has 0
%   hyperlinks). ACCEPTED (deferred fix) per
%   validation\summary\decision_2026-07-28_nsdecl_created_element_orphan.md.
%   Gate-4 therefore pins the PORT'S ACTUAL s7 output (the orphan dropped -> byte-
%   identical to a fresh empty w:hyperlink) with a KNOWN-DEVIATION guard
%   (verifyNotEqual vs lxml's frozen s7 bytes) so a future regression/engine change
%   that alters it goes RED and re-surfaces the decision. It does NOT assert lxml's
%   s7 bytes (that would fail by design).
%
%   Provenance (Gate-1..3, all 2026-07-28):
%     * Audit    : validation\mat2doc\audit_P4-3_oxml_hyperlink.md (Porter Gate-1 +
%                  Opus Gate-2 adversarial APPROVE, zero findings; CT_Hyperlink H3 +
%                  the pagebreak detection/fragment surface verified).
%     * Validate : validation\mat2doc\validate_P4-3_oxml_hyperlink.md (Gate-3
%                  equivalence PASS-DEVIATION -- M1 17/17 L1 [document.xml 1548 B &
%                  styles.xml 349458 B], the 8-tree detection, the 6-tree fragment
%                  serhex byte pins, the second-break guard, the CT_Hyperlink H3
%                  tri-state 7/8 steps byte-identical [s7 the accepted residual], the
%                  CT_Hyperlink.text/lrpb legs, and the CT_P.text-over-hyperlink
%                  closure -- all byte/value-identical; regression 575/575 GREEN,
%                  ZERO flips; exactly ONE probe leaf mismatch = the s7 orphaned-ns0).
%     * Scenario : validation\mat2doc\scenarios\s0023_p4_3_oxml_hyperlink.{py,m}
%                  (the probe sequence replayed VERBATIM by runProbes() below).
%     * Frozen ref (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE):
%         references\s0023\probe.json -- copied verbatim (self-contained) into
%           tests\oxml\data\s0023_probe_oracle.json (the Equivalence replay set;
%           value/serhex JSON, jsondecode is line-ending agnostic -> no
%           `* binary` .gitattributes needed, per the s0020/s0021/s0022 precedent).
%           NOTE: the frozen oracle keeps lxml's s7 xmlns:ns0 bytes VERBATIM (no
%           hand-edit) -- the s7 pin below asserts the port differs from it.
%
%   Coverage taxonomy
%   -----------------
%   * Nominal      -- detection break-start (t1 T/F) / break-end (t2 F/T), the
%                     leading/trailing fragment splits, CT_Hyperlink.text over runs,
%                     the CT_P.text-over-hyperlink closure, M1 byte-neutrality.
%   * Edge         -- sole-content T/T (t7), empty-run-1/break-start-of-run-2 (t5),
%                     mid-run F/F (t3), in-hyperlink atomic F/F (t4), two-breaks
%                     n=2 (t6), break-with-pPr (t8), the atomic-hyperlink fragment
%                     split (whole hyperlink -> preceding frag), w:pPr retained in
%                     both fragments, w:br -> "\n" in .text, the second-break error
%                     path (identifier + message), the history=True==default removal
%                     and the None/[] clears.
%   * Equivalence  -- test_equivalence_full_battery_vs_frozen_oracle replays the
%                     ENTIRE s0023 battery live (runProbes, the .m twin's body
%                     verbatim incl. the document.xml + styles.xml parse->serialize
%                     round-trip legs) and flatten-compares every leaf to the frozen
%                     python-docx 1.2.0 oracle, ASSERTING that the ONLY divergence is
%                     the accepted s7 leaf (mirroring Gate-3's exactly-one-mismatch).
%   * Regression   -- hard-coded expected serialized-XML strings + UPPERCASE serhex
%                     of the raw UTF-8 shipping bytes vs the frozen oracle + SHA-256
%                     of the two M1 parts.
%   * Upstream     -- the CT_LastRenderedPageBreak precedes/follows predicates, the
%                     fragment-split rules, and the CT_Hyperlink history tri-state
%                     ARE the python-docx pagebreak.py / hyperlink.py surface; the
%                     frozen oracle IS lxml's expected output for this API sequence.
%
%   Byte-level (L1) note: every serialized-XML comparison is either the FULL
%   serialize_part_xml output as an ASCII-decoded string (string-equality ==
%   byte-equality L1) or its UPPERCASE hex (serhex) vs the frozen oracle. The ONLY
%   deliberately-not-byte-identical-vs-lxml pin is the s7 known deviation, which is
%   asserted byte-EXACT vs the port's own documented output (still L1 against the
%   accepted reference) and byte-DIFFERENT vs lxml -- both commented at their site.
%
%   Determinism: no network, no absolute paths. The worktree root and the co-located
%   oracle resolve relative to this file via fileparts(mfilename('fullpath')); saves
%   go to tempname .docx deleted via onCleanup; every file read is binary ('r','n');
%   no 'wt'. The +mat2doc package resolves via the MANDATORY PathFixture(worktree-
%   root) in TestClassSetup (WP9-F4 lesson).

    properties (Constant)
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

        % --- frozen s0001 M1 byte references (P4-3 registration is byte-neutral) ---
        DOC_SIZE    = 1548
        DOC_SHA     = "0e4dd503bc095219cd24e8312b372c588ca56bb834aacac53eafb7bf47836327"
        STYLES_SIZE = 349458
        STYLES_SHA  = "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"

        % --- expected detection outcomes over the 8 trees (frozen s0023 oracle);
        %     [precedes, follows, n_breaks] as strings mirroring rv() ------------
        DET_EXPECT = struct( ...
            't1_break_start',                 {{"True",  "False", 1}}, ...
            't2_break_end_sole_run',          {{"False", "True",  1}}, ...
            't3_break_mid',                   {{"False", "False", 1}}, ...
            't4_break_in_hyperlink',          {{"False", "False", 1}}, ...
            't5_break_start_run2_empty_run1', {{"False", "False", 1}}, ...
            't6_two_breaks',                  {{"False", "False", 2}}, ...
            't7_sole_content_break',          {{"True",  "True",  1}}, ...
            't8_break_with_pPr',              {{"True",  "False", 1}})
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied verbatim from tests\oxml\Test_p4_2_oxml_parfmt.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\oxml
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. CT_LastRenderedPageBreak detection (THE MEAT -- all 8 trees)  %
        % =============================================================== %

        function test_pagebreak_detection_8_trees(testCase)
            % Nominal + Edge + Regression (pagebreak.py:52-99, s0023 detection):
            % precedes_all_content / follows_all_content / n_breaks on the FIRST
            % break of each of the 8 hand-built trees. Every outcome is asserted
            % BOTH against a hard-coded [precedes,follows,n] table (DET_EXPECT) AND
            % against the frozen s0023 oracle -- so the whole branch matrix is pinned:
            %   t1 break-start   -> T/F   (./w:r[1] first-run predicate matches)
            %   t2 break-end      -> F/T   ((./w:r)[last()] last-run predicate)
            %   t3 mid-run        -> F/F   (content on both sides)
            %   t4 in-hyperlink   -> F/F   (is_in_hyperlink short-circuit, atomic)
            %   t5 empty-run-1 /
            %      break-start-run-2 -> F/F (the w:r[1]-only subtlety: run-1 empty ->
            %                                first-run predicate finds no break)
            %   t6 two-breaks     -> F/F, n_breaks=2
            %   t7 sole-content   -> T/T   (contrived edge, not precluded by spec)
            %   t8 break-with-pPr -> T/F   (w:pPr present, break still precedes all)
            oracle = loadOracle();
            [ids, inners] = treesTable();
            for i = 1:numel(ids)
                tid = ids(i);
                p = rparse(pXml(testCase, inners(i)));
                lrpbs = p.xpath(".//w:lastRenderedPageBreak");
                lrpb = lrpbs(1);
                precedes = rv(lrpb.precedes_all_content());
                follows  = rv(lrpb.follows_all_content());
                nBreaks  = numel(lrpbs);

                exp = testCase.DET_EXPECT.(char(tid));
                % -- hard-coded regression table --
                testCase.verifyEqual(precedes, exp{1}, ...
                    sprintf('%s precedes_all_content (hard-coded)', tid));
                testCase.verifyEqual(follows, exp{2}, ...
                    sprintf('%s follows_all_content (hard-coded)', tid));
                testCase.verifyEqual(nBreaks, exp{3}, ...
                    sprintf('%s n_breaks (hard-coded)', tid));
                % -- vs the frozen s0023 oracle --
                od = oracle.detection.(char(tid));
                testCase.verifyEqual(precedes, string(od.precedes_all_content), ...
                    sprintf('%s precedes_all_content vs frozen oracle', tid));
                testCase.verifyEqual(follows, string(od.follows_all_content), ...
                    sprintf('%s follows_all_content vs frozen oracle', tid));
                testCase.verifyEqual(nBreaks, double(od.n_breaks), ...
                    sprintf('%s n_breaks vs frozen oracle', tid));
            end
        end

        % =============================================================== %
        % 2. Fragment split -- byte level (preceding/following loose CT_P) %
        % =============================================================== %

        function test_fragment_splits_byte_level(testCase)
            % Nominal + Edge + Regression (pagebreak.py:101-118 + split helpers,
            % s0023 fragments): preceding_fragment_p / following_fragment_p on the
            % FIRST break of the 6 split-relevant trees, compared BOTH as fragment
            % .text AND as serhex (UPPERCASE hex of the raw UTF-8 serialize_part_xml
            % bytes -- a byte pin) AND as child localnames, all vs the frozen oracle.
            % Highlights: t4 atomic-hyperlink split ("prexy"/"post", whole hyperlink
            % into the preceding fragment with the break removed from inside;
            % preceding_localnames [r,hyperlink]); t8 w:pPr retained in BOTH fragments
            % (localnames [pPr,r]).
            oracle = loadOracle();
            fragIds = ["t1_break_start" "t2_break_end_sole_run" "t3_break_mid" ...
                       "t4_break_in_hyperlink" "t7_sole_content_break" "t8_break_with_pPr"];
            [ids, inners] = treesTable();
            for i = 1:numel(fragIds)
                tid = fragIds(i);
                inner = inners(ids == tid);
                p = rparse(pXml(testCase, inner));
                lrpbs = p.xpath(".//w:lastRenderedPageBreak");
                lrpb = lrpbs(1);
                pre = lrpb.preceding_fragment_p();
                fol = lrpb.following_fragment_p();

                of = oracle.fragments.(char(tid));
                testCase.verifyEqual(pre.text, string(of.preceding_text), ...
                    sprintf('%s preceding_fragment .text', tid));
                testCase.verifyEqual(fol.text, string(of.following_text), ...
                    sprintf('%s following_fragment .text', tid));
                testCase.verifyEqual(hx_e(pre), string(of.preceding_serhex), ...
                    sprintf('%s preceding_fragment serialized bytes (L1)', tid));
                testCase.verifyEqual(hx_e(fol), string(of.following_serhex), ...
                    sprintf('%s following_fragment serialized bytes (L1)', tid));
                testCase.verifyEqual(lnsCell(pre), oracleCell(of.preceding_localnames), ...
                    sprintf('%s preceding_localnames', tid));
                testCase.verifyEqual(lnsCell(fol), oracleCell(of.following_localnames), ...
                    sprintf('%s following_localnames', tid));
            end

            % -- spotlight the atomic-hyperlink split (t4) explicitly --
            p = rparse(pXml(testCase, inners(ids == "t4_break_in_hyperlink")));
            lrpb = p.xpath(".//w:lastRenderedPageBreak"); lrpb = lrpb(1);
            pre = lrpb.preceding_fragment_p();
            testCase.verifyEqual(pre.text, "prexy", ...
                'atomic hyperlink: whole w:hyperlink kept in preceding frag, break removed -> "prexy"');
            testCase.verifyEqual(lnsCell(pre), {'r','hyperlink'}, ...
                'atomic hyperlink preceding_localnames [r, hyperlink]');
            testCase.verifyEqual(lrpb.following_fragment_p().text, "post", ...
                'atomic hyperlink: following frag drops the hyperlink -> "post"');

            % -- spotlight w:pPr retention in both fragments (t8) --
            p = rparse(pXml(testCase, inners(ids == "t8_break_with_pPr")));
            lrpb = p.xpath(".//w:lastRenderedPageBreak"); lrpb = lrpb(1);
            testCase.verifyEqual(lnsCell(lrpb.preceding_fragment_p()), {'pPr','r'}, ...
                't8 preceding frag retains w:pPr [pPr, r]');
            testCase.verifyEqual(lnsCell(lrpb.following_fragment_p()), {'pPr','r'}, ...
                't8 following frag retains w:pPr [pPr, r]');
        end

        % =============================================================== %
        % 3. Second-break ValueError guard                                %
        % =============================================================== %

        function test_second_break_valueerror_guard(testCase)
            % Edge / error-path (pagebreak.py first-break guard, s0023 guard):
            % preceding_/following_fragment_p on a NON-first break both raise
            % ValueError. Verify the IDENTIFIER (mat2doc:ValueError) AND the verbatim
            % python-docx message -- not merely that it throws.
            [ids, inners] = treesTable();
            p = rparse(pXml(testCase, inners(ids == "t6_two_breaks")));
            lrpbs = p.xpath(".//w:lastRenderedPageBreak");
            second = lrpbs(2);
            expMsg = "only defined on first rendered page-break in paragraph";

            mePre = captureError(@() second.preceding_fragment_p());
            testCase.verifyEqual(string(mePre.identifier), "mat2doc:ValueError", ...
                'preceding_fragment_p on 2nd break must raise mat2doc:ValueError');
            testCase.verifyEqual(string(mePre.message), expMsg, ...
                'preceding_fragment_p 2nd-break message verbatim');

            meFol = captureError(@() second.following_fragment_p());
            testCase.verifyEqual(string(meFol.identifier), "mat2doc:ValueError", ...
                'following_fragment_p on 2nd break must raise mat2doc:ValueError');
            testCase.verifyEqual(string(meFol.message), expMsg, ...
                'following_fragment_p 2nd-break message verbatim');
        end

        % =============================================================== %
        % 4. CT_Hyperlink H3 tri-state (+ the s7 KNOWN DEVIATION)          %
        % =============================================================== %

        function test_ct_hyperlink_h3_tristate(testCase)
            % Nominal + Edge + Regression + KNOWN-DEVIATION (hyperlink.py rId/anchor/
            % history, s0023 hyperlink.mutations): fresh read (rId/anchor absent -> []
            % [None], history absent -> True [the NON-None default]) then the 8-step
            % mutation sequence, serhex-pinned at every step vs the frozen oracle --
            % byte-identical EXCEPT step s7 (see below).
            oracle = loadOracle();

            % -- fresh: <w:hyperlink xmlns:w="..."/> ; rId/anchor None, history True --
            h = mat2doc.oxml.OxmlElement("w:hyperlink");
            testCase.verifyTrue(isequal(h.rId, []),    'fresh rId -> [] (None)');
            testCase.verifyTrue(isequal(h.anchor, []), 'fresh anchor -> [] (None)');
            testCase.verifyTrue(h.history,             'fresh history -> True (non-None default)');
            testCase.verifyEqual(ser(h), decl() + newline + ...
                "<w:hyperlink xmlns:w=""" + testCase.W + """/>", ...
                'fresh empty w:hyperlink serialized bytes (L1) hard-coded');
            testCase.verifyEqual(hx_e(h), string(oracle.hyperlink.fresh.serhex));

            % -- the 8-step mutation sequence on ONE loose element --
            h = mat2doc.oxml.OxmlElement("w:hyperlink");
            h.rId = "rId7";
            testCase.verifyEqual(string(h.rId), "rId7");
            testCase.verifyEqual(hx_e(h), string(oracle.hyperlink.mutations.s1_set_rId.serhex), ...
                's1 set rId -- lxml mints xmlns:ns0 + ns0:id (byte-identical)');

            h.anchor = "section1";
            testCase.verifyEqual(string(h.anchor), "section1");
            testCase.verifyEqual(hx_e(h), string(oracle.hyperlink.mutations.s2_set_anchor.serhex), 's2 set anchor');

            h.history = false;
            testCase.verifyFalse(h.history);
            testCase.verifyEqual(hx_e(h), string(oracle.hyperlink.mutations.s3_history_false.serhex), ...
                's3 history=False -> @w:history="0"');

            h.history = true;                 % == default -> @w:history REMOVED
            testCase.verifyTrue(h.history);
            testCase.verifyEqual(hx_e(h), string(oracle.hyperlink.mutations.s4_history_true_removed.serhex), ...
                's4 history=True (==default) -> @w:history REMOVED');

            h.history = [];                   % None -> stays removed
            testCase.verifyTrue(h.history,    's5 history=None read-back -> True (default)');
            testCase.verifyEqual(hx_e(h), string(oracle.hyperlink.mutations.s5_history_none_removed.serhex), 's5');

            h.anchor = [];                    % remove @w:anchor
            testCase.verifyTrue(isequal(h.anchor, []), 's6 anchor cleared -> [] (None)');
            testCase.verifyEqual(hx_e(h), string(oracle.hyperlink.mutations.s6_anchor_none_removed.serhex), 's6');

            % ---- s7: rId cleared -- THE ACCEPTED KNOWN DEVIATION ----
            % lxml retains the now-orphaned xmlns:ns0 (sticky nsmap); the Mat2Doc
            % serializer recomputes used-namespaces and DROPS it, yielding a fresh
            % empty <w:hyperlink xmlns:w="..."/>. ACCEPTED as a generation-side
            % (created-element, minted-then-orphaned) residual of D-serializer-nsdecl,
            % no new D-number, UNREACHABLE on real paths -- per
            % validation\summary\decision_2026-07-28_nsdecl_created_element_orphan.md.
            % We pin the PORT'S ACTUAL output (== the fresh empty hyperlink) and
            % assert it DIFFERS from lxml's frozen s7 bytes, so any future change is
            % caught. Asserting lxml's s7 bytes here would fail BY DESIGN.
            h.rId = [];
            testCase.verifyTrue(isequal(h.rId, []), 's7 rId cleared -> [] (None)');
            portS7 = hx_e(h);
            testCase.verifyEqual(ser(h), decl() + newline + ...
                "<w:hyperlink xmlns:w=""" + testCase.W + """/>", ...
                's7 KNOWN DEVIATION: port drops the orphaned xmlns:ns0 -> fresh empty w:hyperlink (L1 hard-coded)');
            testCase.verifyEqual(portS7, string(oracle.hyperlink.fresh.serhex), ...
                's7 port output is byte-identical to a fresh empty w:hyperlink');
            testCase.verifyNotEqual(portS7, string(oracle.hyperlink.mutations.s7_rId_none_removed.serhex), ...
                's7 KNOWN DEVIATION: port intentionally DIFFERS from lxml (orphaned ns0 retained by lxml, dropped by port) -- see decision_2026-07-28_nsdecl_created_element_orphan.md');

            % -- s8: history re-read after full clear -> True (default) --
            testCase.verifyTrue(h.history, 's8 history absent -> True (default)');
            testCase.verifyEqual(rv(h.history), string(oracle.hyperlink.mutations.s8_history_reread_default.history));
        end

        % =============================================================== %
        % 5. CT_Hyperlink parse read-back (ST_OnOff history)               %
        % =============================================================== %

        function test_ct_hyperlink_parse_readback(testCase)
            % Regression (hyperlink.py, s0023 parse_*): parse a w:hyperlink and read
            % history / rId / anchor. ST_OnOff: w:history="0" -> False,
            % w:history="on" -> True, absent -> True (default). rId/anchor read exact.
            oracle = loadOracle();

            h = rparse(hlinkXml(testCase, "r:id=""rId5"" w:anchor=""sec"" w:history=""0""", "<w:r><w:t>x</w:t></w:r>"));
            testCase.verifyFalse(h.history, 'w:history="0" -> False');
            testCase.verifyEqual(string(h.rId), "rId5");
            testCase.verifyEqual(string(h.anchor), "sec");
            testCase.verifyEqual(rv(h.history), string(oracle.hyperlink.parse_hist0.history));
            testCase.verifyEqual(rv(h.rId),     string(oracle.hyperlink.parse_hist0.rId));
            testCase.verifyEqual(rv(h.anchor),  string(oracle.hyperlink.parse_hist0.anchor));

            h = rparse(hlinkXml(testCase, "w:history=""on""", "<w:r><w:t>x</w:t></w:r>"));
            testCase.verifyTrue(h.history, 'w:history="on" -> True');
            testCase.verifyEqual(rv(h.history), string(oracle.hyperlink.parse_hist_on.history));

            h = rparse(hlinkXml(testCase, "", "<w:r><w:t>x</w:t></w:r>"));
            testCase.verifyTrue(h.history, 'absent w:history -> True (default)');
            testCase.verifyTrue(isequal(h.rId, []), 'absent r:id -> [] (None)');
            testCase.verifyEqual(rv(h.rId), string(oracle.hyperlink.parse_hist_absent.rId));
        end

        % =============================================================== %
        % 6. CT_Hyperlink.text + lastRenderedPageBreaks                    %
        % =============================================================== %

        function test_ct_hyperlink_text_and_lrpb(testCase)
            % Nominal + Edge + Regression (hyperlink.py text/lastRenderedPageBreaks,
            % s0023 hyperlink): .text joins the CT_R.text of each DIRECT child w:r;
            % w:br -> "\n"; lastRenderedPageBreaks counts ./w:r/w:lastRenderedPageBreak.
            oracle = loadOracle();

            h = rparse(hlinkXml(testCase, "", "<w:r><w:t>Go </w:t></w:r><w:r><w:t>here</w:t></w:r>"));
            testCase.verifyEqual(h.text, "Go here", 'CT_Hyperlink.text joins direct-child run text');
            testCase.verifyEqual(h.text, string(oracle.hyperlink.text_runs));

            h = rparse(hlinkXml(testCase, "", "<w:r><w:t>Go </w:t><w:br/><w:t>there</w:t></w:r>"));
            testCase.verifyEqual(h.text, "Go " + newline + "there", 'w:br -> "\n" in CT_Hyperlink.text');
            testCase.verifyEqual(h.text, string(oracle.hyperlink.text_with_br));

            h = rparse(hlinkXml(testCase, "", ...
                "<w:r><w:lastRenderedPageBreak/><w:t>a</w:t></w:r><w:r><w:lastRenderedPageBreak/><w:t>b</w:t></w:r>"));
            testCase.verifyEqual(numel(h.lastRenderedPageBreaks()), 2, 'two runs each with an lrpb -> count 2');
            testCase.verifyEqual(numel(h.lastRenderedPageBreaks()), double(oracle.hyperlink.lrpb_count));
        end

        % =============================================================== %
        % 7. CT_P.text-over-hyperlink closure (P4-2 VERIFY)                %
        % =============================================================== %

        function test_ct_p_text_over_hyperlink_closure(testCase)
            % Nominal + Regression (P4-2 VERIFY, s0023 ct_p_over_hyperlink):
            % registering w:hyperlink alone (no CT_P edit) flips CT_P.text over a
            % hyperlink-bearing paragraph from generic char data to the concatenated
            % hyperlink run text.
            oracle = loadOracle();
            p = rparse(pXml(testCase, ...
                "<w:r><w:t>Before </w:t></w:r><w:hyperlink r:id=""rId1"">" + ...
                "<w:r><w:t>link</w:t></w:r></w:hyperlink><w:r><w:t> after</w:t></w:r>"));
            testCase.verifyEqual(p.text, "Before link after", ...
                'CT_P.text over a w:hyperlink returns the run text (closure)');
            testCase.verifyEqual(p.text, string(oracle.ct_p_over_hyperlink.text));
        end

        % =============================================================== %
        % 8. M1 byte-neutrality (registration is a pure lookup addition)  %
        % =============================================================== %

        function test_m1_document_xml_byte_identical(testCase)
            % Regression (byte-neutrality, L1): mat2doc.Document().save() emits
            % word/document.xml at EXACTLY 1548 B with the frozen s0001 SHA-256 --
            % the two new registry rows add no drift (template has 0 hyperlinks /
            % 0 lastRenderedPageBreaks; neither new class is instantiated).
            bytes = emitDocPart('document.xml');
            testCase.verifyEqual(numel(bytes), testCase.DOC_SIZE, ...
                sprintf('word/document.xml must be exactly %d B after P4-3 registration', testCase.DOC_SIZE));
            testCase.verifyEqual(sha256hex(bytes), testCase.DOC_SHA, ...
                'word/document.xml SHA-256 must equal the frozen s0001 oracle (byte-identical L1)');
        end

        function test_m1_styles_xml_byte_identical(testCase)
            % Regression (byte-neutrality, L1): word/styles.xml stays 349458 B with
            % the frozen s0001 SHA-256. (Full 17/17 M1 sweep owned by Test_p1_8.)
            bytes = emitDocPart('styles.xml');
            testCase.verifyEqual(numel(bytes), testCase.STYLES_SIZE, ...
                sprintf('word/styles.xml must be exactly %d B', testCase.STYLES_SIZE));
            testCase.verifyEqual(sha256hex(bytes), testCase.STYLES_SHA, ...
                'word/styles.xml SHA-256 must equal the frozen s0001 oracle (byte-identical L1)');
        end

        % =============================================================== %
        % 9. EQUIVALENCE -- full s0023 battery vs the frozen oracle        %
        % =============================================================== %

        function test_equivalence_full_battery_vs_frozen_oracle(testCase)
            % Equivalence: replay the ENTIRE s0023 battery (runProbes -- the .m twin's
            % body VERBATIM: detection / fragments / guard / hyperlink / ct_p closure +
            % the document.xml & styles.xml parse->serialize round-trip legs over what
            % Document().save() itself emits) and flatten-compare EVERY leaf to the
            % frozen python-docx 1.2.0 oracle. Gate-3 found EXACTLY ONE mismatch --
            % the accepted s7 orphaned-ns0 leaf -- so this test asserts every leaf is
            % byte/value-identical EXCEPT that single s7 serhex leaf, which must (a)
            % differ from lxml and (b) equal the fresh empty-hyperlink bytes. This
            % ties the suite to the Gate-3 output including the shape of the deviation.
            S7KEY = 'hyperlink.mutations.s7_rId_none_removed.serhex';
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

            freshHex = char(oMap('hyperlink.fresh.serhex'));
            nDiff = 0;
            for i = 1:numel(oKeys)
                k = oKeys{i};
                testCase.verifyTrue(isKey(pMap, k), sprintf('port is missing leaf %s', k));
                if strcmp(k, S7KEY)
                    % THE accepted known deviation -- must differ from lxml and equal
                    % the fresh empty-hyperlink bytes (see test_ct_hyperlink_h3_tristate).
                    nDiff = nDiff + 1;
                    testCase.verifyNotEqual(pMap(k), oMap(k), ...
                        's7 serhex must DIFFER from lxml (accepted D-serializer-nsdecl residual)');
                    testCase.verifyEqual(pMap(k), freshHex, ...
                        's7 port serhex must equal the fresh empty-hyperlink bytes');
                else
                    testCase.verifyEqual(pMap(k), oMap(k), ...
                        sprintf('leaf %s must be byte/value-identical to the frozen oracle', k));
                end
            end
            % Gate-3 crux: EXACTLY ONE leaf mismatch across the whole probe.
            testCase.verifyEqual(nDiff, 1, ...
                'exactly ONE leaf (the s7 orphaned-ns0) may diverge from lxml -- Gate-3 finding');
        end

    end
end

% ===================== file-local helpers ============================== %

function s = decl()
    s = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>";
end

function s = nsWR(testCase)
    s = "xmlns:w=""" + testCase.W + """ xmlns:r=""" + testCase.R + """";
end

function x = pXml(testCase, inner)
    x = "<w:p " + nsWR(testCase) + ">" + string(inner) + "</w:p>";
end

function x = hlinkXml(testCase, attrs, inner)
    a = string(attrs);
    if strlength(a) > 0
        a = " " + a;
    end
    x = "<w:hyperlink " + nsWR(testCase) + a + ">" + string(inner) + "</w:hyperlink>";
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
    % UPPERCASE hex of raw bytes (matches Python bytes.hex().upper()).
    h = string(sprintf('%02X', uint8(raw)));
end

function h = hx_e(e)
    h = hx(mat2doc.oxml.serialize_part_xml(e));
end

function [ids, inners] = treesTable()
    % The 8 detection trees (inner content of a <w:p>); IDENTICAL to the s0023 twin.
    trees = { ...
        "t1_break_start",                "<w:r><w:lastRenderedPageBreak/><w:t>text</w:t></w:r>"; ...
        "t2_break_end_sole_run",         "<w:r><w:t>text</w:t><w:lastRenderedPageBreak/></w:r>"; ...
        "t3_break_mid",                  "<w:r><w:t>a</w:t><w:lastRenderedPageBreak/><w:t>b</w:t></w:r>"; ...
        "t4_break_in_hyperlink",         "<w:r><w:t>pre</w:t></w:r><w:hyperlink r:id=""rId1""><w:r><w:t>x</w:t><w:lastRenderedPageBreak/><w:t>y</w:t></w:r></w:hyperlink><w:r><w:t>post</w:t></w:r>"; ...
        "t5_break_start_run2_empty_run1","<w:r/><w:r><w:lastRenderedPageBreak/><w:t>x</w:t></w:r>"; ...
        "t6_two_breaks",                 "<w:r><w:t>a</w:t><w:lastRenderedPageBreak/><w:t>b</w:t><w:lastRenderedPageBreak/><w:t>c</w:t></w:r>"; ...
        "t7_sole_content_break",         "<w:r><w:lastRenderedPageBreak/></w:r>"; ...
        "t8_break_with_pPr",             "<w:pPr><w:pStyle w:val=""Heading1""/></w:pPr><w:r><w:lastRenderedPageBreak/><w:t>text</w:t></w:r>" ...
    };
    ids    = string(trees(:, 1))';
    inners = string(trees(:, 2))';
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

function C = oracleCell(v)
    % Normalize a jsondecode'd localnames leaf (cell of char, or a bare char for a
    % single-element JSON array) into a 1xN cell row of char, matching lnsCell.
    if iscell(v)
        C = cell(1, numel(v));
        for k = 1:numel(v)
            C{k} = char(v{k});
        end
    elseif isempty(v)
        C = cell(1, 0);
    else
        C = {char(v)};
    end
end

function s = rv(x)
    % Uniform accessor repr mirroring the s0023 rv(): None->"None",
    % bool->"True"/"False", int->decimal string.
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

function s = tf(b)
    if b, s = "True"; else, s = "False"; end
end

function ME = captureError(fn)
    % Run fn and RETURN the caught MException (asserts it actually raised).
    try
        fn();
        error('mat2doc:test:noRaise', 'expected an error but none was raised');
    catch ME %#ok<NASGU>
    end
    if strcmp(ME.identifier, 'mat2doc:test:noRaise')
        error('mat2doc:test:noRaise', 'expected a mat2doc:ValueError but none was raised');
    end
end

function o = loadOracle()
    % Read the co-located frozen oracle in BINARY mode (no CRLF translation) and
    % decode UTF-8 -> struct. jsondecode is whitespace/line-ending agnostic, so no
    % `* binary` .gitattributes pin is needed (value-based fixture, s0022 precedent).
    here = fileparts(mfilename('fullpath'));
    p = fullfile(here, 'data', 's0023_probe_oracle.json');
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
    % Base-MATLAB unzip (no toolbox) into a temp dir, both cleaned on exit.
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
    % Replay the s0023 probe sequence (the .m twin's P-struct body, VERBATIM
    % tags/inputs/order) and return the nested struct of tagged canonical values.
    % Embedded here so the Equivalence leg is self-contained (the validation-folder
    % scenario is NOT on the toolbox path and must not be a dependency). Mirrors
    % validation\mat2doc\scenarios\s0023_p4_3_oxml_hyperlink.m. The two round-trip
    % legs read the parts Document().save() itself emits (rather than external file
    % args), so no fixture path is needed.
    P = struct();
    [ids, inners] = treesTable();
    fragIds = ["t1_break_start" "t2_break_end_sole_run" "t3_break_mid" ...
               "t4_break_in_hyperlink" "t7_sole_content_break" "t8_break_with_pPr"];

    % ===================== detection =======================================
    det = struct();
    for i = 1:numel(ids)
        tid = ids(i);
        p = rparse(pXml_local(inners(i)));
        lrpbs = p.xpath(".//w:lastRenderedPageBreak");
        lrpb = lrpbs(1);
        det.(char(tid)) = struct( ...
            "precedes_all_content", rv(lrpb.precedes_all_content()), ...
            "follows_all_content",  rv(lrpb.follows_all_content()), ...
            "n_breaks", numel(lrpbs));
    end
    P.detection = det;

    % ===================== fragments =======================================
    frag = struct();
    for i = 1:numel(fragIds)
        tid = fragIds(i);
        inner = inners(ids == tid);
        p = rparse(pXml_local(inner));
        lrpbs = p.xpath(".//w:lastRenderedPageBreak");
        lrpb = lrpbs(1);
        pre = lrpb.preceding_fragment_p();
        fol = lrpb.following_fragment_p();
        frag.(char(tid)) = struct( ...
            "preceding_text", pre.text, ...
            "following_text", fol.text, ...
            "preceding_serhex", hx_e(pre), ...
            "following_serhex", hx_e(fol), ...
            "preceding_localnames", {lnsCell(pre)}, ...
            "following_localnames", {lnsCell(fol)});
    end
    P.fragments = frag;

    % ===================== guard: second break raises ValueError ===========
    g = struct();
    p = rparse(pXml_local(inners(ids == "t6_two_breaks")));
    lrpbs = p.xpath(".//w:lastRenderedPageBreak");
    second = lrpbs(2);
    g.preceding_fragment_p = guardCall(@() second.preceding_fragment_p());
    g.following_fragment_p = guardCall(@() second.following_fragment_p());
    P.guard = g;

    % ===================== hyperlink: H3 tri-state + text + count ===========
    hl = struct();
    h = mat2doc.oxml.OxmlElement("w:hyperlink");
    hl.fresh = struct("serhex", hx_e(h), "rId", rv(h.rId), ...
        "anchor", rv(h.anchor), "history", rv(h.history));

    seq = struct();
    h = mat2doc.oxml.OxmlElement("w:hyperlink");
    h.rId = "rId7";
    seq.s1_set_rId = struct("serhex", hx_e(h), "rId", rv(h.rId));
    h.anchor = "section1";
    seq.s2_set_anchor = struct("serhex", hx_e(h), "anchor", rv(h.anchor));
    h.history = false;
    seq.s3_history_false = struct("serhex", hx_e(h), "history", rv(h.history));
    h.history = true;                       % == default -> @w:history REMOVED
    seq.s4_history_true_removed = struct("serhex", hx_e(h), "history", rv(h.history));
    h.history = [];                         % None -> removed (stays removed)
    seq.s5_history_none_removed = struct("serhex", hx_e(h), "history", rv(h.history));
    h.anchor = [];                          % remove @w:anchor
    seq.s6_anchor_none_removed = struct("serhex", hx_e(h), "anchor", rv(h.anchor));
    h.rId = [];                             % remove @r:id (s7 -- known deviation)
    seq.s7_rId_none_removed = struct("serhex", hx_e(h), "rId", rv(h.rId));
    seq.s8_history_reread_default = struct("history", rv(h.history));  % absent -> True
    hl.mutations = seq;

    h = rparse(hlinkXml_local("r:id=""rId5"" w:anchor=""sec"" w:history=""0""", "<w:r><w:t>x</w:t></w:r>"));
    hl.parse_hist0 = struct("history", rv(h.history), "rId", rv(h.rId), "anchor", rv(h.anchor));
    h = rparse(hlinkXml_local("w:history=""on""", "<w:r><w:t>x</w:t></w:r>"));
    hl.parse_hist_on = struct("history", rv(h.history));
    h = rparse(hlinkXml_local("", "<w:r><w:t>x</w:t></w:r>"));
    hl.parse_hist_absent = struct("history", rv(h.history), "rId", rv(h.rId));

    h = rparse(hlinkXml_local("", "<w:r><w:t>Go </w:t></w:r><w:r><w:t>here</w:t></w:r>"));
    hl.text_runs = h.text;                  % "Go here"
    h = rparse(hlinkXml_local("", "<w:r><w:t>Go </w:t><w:br/><w:t>there</w:t></w:r>"));
    hl.text_with_br = h.text;               % "Go \nthere"
    h = rparse(hlinkXml_local("", ...
        "<w:r><w:lastRenderedPageBreak/><w:t>a</w:t></w:r><w:r><w:lastRenderedPageBreak/><w:t>b</w:t></w:r>"));
    hl.lrpb_count = numel(h.lastRenderedPageBreaks());   % 2
    P.hyperlink = hl;

    % ===================== ct_p_over_hyperlink: P4-2 VERIFY closure =========
    p = rparse(pXml_local( ...
        "<w:r><w:t>Before </w:t></w:r><w:hyperlink r:id=""rId1""><w:r><w:t>link</w:t></w:r></w:hyperlink><w:r><w:t> after</w:t></w:r>"));
    P.ct_p_over_hyperlink = struct("text", p.text);   % "Before link after"

    % ===================== document.xml parse->serialize L1 =================
    dr = struct();
    db = emitDocPart('document.xml');
    elm = mat2doc.oxml.parse_xml(db);
    out = mat2doc.oxml.serialize_part_xml(elm);
    dr.roundtrip_byte_identical = tf(isequal(uint8(out(:)'), uint8(db(:)')));
    dr.n_hyperlink = numel(elm.xpath(".//w:hyperlink"));
    dr.n_lrpb      = numel(elm.xpath(".//w:lastRenderedPageBreak"));
    P.document_roundtrip = dr;

    % ===================== styles.xml parse->serialize L1 ===================
    sr = struct();
    sb = emitDocPart('styles.xml');
    elm = mat2doc.oxml.parse_xml(sb);
    out = mat2doc.oxml.serialize_part_xml(elm);
    sr.roundtrip_byte_identical = tf(isequal(uint8(out(:)'), uint8(sb(:)')));
    sr.n_lrpb = numel(elm.xpath(".//w:lastRenderedPageBreak"));
    P.styles_roundtrip = sr;
end

function s = nsWR_local()
    s = "xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""" + ...
        " xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""";
end

function x = pXml_local(inner)
    x = "<w:p " + nsWR_local() + ">" + string(inner) + "</w:p>";
end

function x = hlinkXml_local(attrs, inner)
    a = string(attrs);
    if strlength(a) > 0
        a = " " + a;
    end
    x = "<w:hyperlink " + nsWR_local() + a + ">" + string(inner) + "</w:hyperlink>";
end

function out = guardCall(fn)
    % Capture ValueError raise + message (identifier is Python-absent; compare msg).
    try
        fn();
        out = struct("raised", "False", "msg", "");
    catch ME
        out = struct("raised", "True", "msg", string(ME.message));
    end
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

classdef Test_p4_5a_parfmt_api < matlab.unittest.TestCase
% TEST_P4_5A_PARFMT_API  Gate-4 permanent unit tests for Mat2Doc P4-5a
%   (src/docx/text/parfmt.py -> +mat2doc\+text\ParagraphFormat and
%   src/docx/text/tabstops.py -> +mat2doc\+text\TabStops / TabStop -- a P4
%   API/proxy-tier WP).
%
%   ParagraphFormat / TabStops / TabStop are pure API/proxy classes over the
%   already byte-validated CT_PPr / CT_TabStops / CT_TabStop (P4-2). None adds a
%   register_element_cls row, an oxml class, or a serialization-path change ->
%   API/proxy tier: no bytes, no registry, no M1 risk (validate_P4-5a section 3
%   confirms M1 17/17 unchanged, so this class does NOT re-pin styles.xml /
%   document.xml on the default template -- Test_p1_8_skeleton_m1 owns the full
%   17/17 M1 sweep). This class permanently FREEZES the P4-5a BEHAVIORAL surface
%   -- the API-tier equivalence pin -- byte/value-identical to python-docx 1.2.0,
%   PLUS the G-scenario document byte pin (four formatted paragraphs written into
%   a real word/document.xml through the live save path).
%
%   The guarantees frozen here (each verified byte/value-identical at Gate-3:
%   probe_diff MATCH 120/120 leaves + G-scenario 17/17 parts byte-identical +
%   M1 17/17; zero divergences, zero new D-numbers):
%
%     (line_spacing -- the meaty logic, the equivalence pin)
%       float MULTIPLE matrix 1.0/1.5/1.75/2.0 -> Twips 240/360/420/480,
%         w:lineRule="auto"; get 1/1.5/1.75/2 (%.12g, D-STYPE-1); rule maps
%         SINGLE/ONE_POINT_FIVE/DOUBLE and bare MULTIPLE for 1.75.
%       EXACTLY via Pt(18) -> <w:spacing w:line="360" w:lineRule="exact"/>,
%         get the absolute Length EMU 228600, rule EXACTLY.
%       AT_LEAST-preserved: rule set FIRST then a Length keeps lineRule="atLeast"
%         (NOT overwritten to exact) and serializes **w:lineRule BEFORE w:line**
%         (the attribute insertion order is byte-load-bearing).
%       []->reset keeps an EMPTY <w:spacing/> (attrs removed, element retained);
%         get None / rule None.
%       line_spacing_rule SET SINGLE/ONE_POINT_FIVE/DOUBLE/AT_LEAST/EXACTLY/None.
%
%     (alignment + the four bool tri-states) alignment CENTER -> <w:jc
%       w:val="center"/>, []-> removed (empty <w:pPr/>); keep_together /
%       keep_with_next / page_break_before / widow_control each True -> <w:x/>
%       (val absent), False -> <w:x w:val="0"/>, [] -> removed.
%
%     (indents, SIGN both ways) first_line_indent Pt(24)->@firstLine="480"
%       (get 304800) / Pt(-24)->@hanging="480" (get -304800) / Pt(0)->
%       @firstLine="0"; left_indent Pt(36)/Pt(-36) -> @left="720"/"-720";
%       right_indent Pt(18)/Pt(-18) -> @right="360"/"-360". space_before Pt(6)
%       -> @before="120"; space_after Pt(12) -> @after="240".
%
%     (TabStops sequence + H1) empty len_()==0, getitem_(0) -> mat2doc:IndexError
%       "TabStops object is empty"; add_tab_stop pos-SORTS (out-of-order 2/0.5/1in
%       -> pos 720/1440/2880); getitem_(0)/(-1)/(1) index incl. negative wrap;
%       getitem_(3)/(-4) -> mat2doc:IndexError "list index out of range" (verbatim);
%       delitem_ middle / delitem_(9) -> "tab index out of range" / down-to-empty
%       removes <w:tabs>; clear_all removes <w:tabs>; len_.
%
%     (TabStop props) default add_tab_stop(Inches(1)) -> position 914400 /
%       alignment LEFT (w:val="left" present) / leader SPACES with **@w:leader
%       ABSENT** (OptionalAttribute default-removal); explicit
%       add_tab_stop(Inches(3),CENTER,DOTS) -> <w:tab w:pos="4320" w:val="center"
%       w:leader="dot"/>; the position-setter quirk: t==t2 stays TRUE after
%       t.position=Inches(1.5) (both proxies track the OLD detached element --
%       python-docx tabstops.py 118-123, faithfully NOT "corrected").
%
%     (G-scenario document byte pin -- the M2 de-risker) a fresh mat2doc.Document()
%       + four formatted paragraphs (P1 1.75-multiple; P2 AT_LEAST+Pt(18); P3
%       hanging-indent + 3 pos-sorted tab stops; P4 ONE_POINT_FIVE/center/widow/
%       keepLines) appended at the oxml level (body.add_p(), Paragraph proxy is
%       P4-6) and .save()d -> the extracted word/document.xml is byte-identical
%       (SHA-256 + size + full-bytes) to the frozen s0027 python-docx reference.
%       This is the closest pin to M2's add_paragraph formatting path available
%       before the Paragraph proxy lands (P4-6). Exercises the CT_PPr child
%       ordering hazard (keepLines/widowControl/spacing/jc schema sequence;
%       w:tabs BEFORE w:ind) and pos-sorted w:tabs on the actual serializer.
%
%   Provenance (Gate-1..3, all 2026-07-30):
%     * Audit    : validation\mat2doc\audit_P4-5a_parfmt_api.md (Porter Gate-1 +
%                  Fable Gate-2 adversarial APPROVE, zero defects; 64-key porter
%                  self-sanity + 142-key auditor battery byte-identical; the three
%                  porter flags -- VERIFY-COLLECTION interim surface, D-STYPE-1
%                  adopt-only, the position-setter quirk genuine -- all RULED).
%     * Validate : validation\mat2doc\validate_P4-5a_parfmt_api.md (Gate-3 PASS --
%                  probe_diff MATCH 120/120; G-scenario 17/17 parts byte-identical
%                  incl. word/document.xml; M1 17/17; regression 611/611; ZERO
%                  divergences, ZERO new D-numbers).
%     * Scenarios: validation\mat2doc\scenarios\s0026_p4_5a_parfmt_probe.{py,m}
%                  (the property/behavioral probe replayed VERBATIM by runProbes()
%                  below) and s0027_p4_5a_parfmt_gscenario.{py,m} (the four-
%                  paragraph document byte pin replayed by buildGScenario() below).
%     * Frozen refs (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE):
%         references\s0026\probe.json -- copied verbatim (self-contained) into
%           tests\text\data\s0026_probe_oracle.json (value/serhex JSON; jsondecode
%           is line-ending agnostic -> no `* binary` pin, s0020/s0023/s0024
%           precedent).
%         references\s0027\parts\word\document.xml (SHA-256
%           0f7a0519f6aefcde29b1baaf11800c21b7e26c3e61129b3bee28dbf3c1a17a01,
%           1998 B) -- copied byte-for-byte into tests\text\data\s0027_document.xml
%           (co-located `* binary` .gitattributes) as the G-scenario byte fixture,
%           AND its SHA/size embedded below as SHA_S0027_DOCUMENT / SIZE_.
%
%   Coverage taxonomy
%   -----------------
%   * Nominal      -- line_spacing multiples, EXACTLY, the rule SET matrix,
%                     alignment/bools True paths, positive indents, spacing,
%                     add_tab_stop, TabStop default/explicit props, the G-scenario.
%   * Edge         -- []->empty <w:spacing/> reset, [] removals (alignment/bools),
%                     AT_LEAST-preserved attr-order, NEGATIVE + ZERO indents
%                     (Pt(-24)/Pt(0)), non-ASCII is N/A (numeric surface), empty
%                     TabStops, delitem_ down-to-empty, the SPACES-default
%                     @w:leader absence, the position-setter eq-tracks-original
%                     quirk, and the error paths below.
%   * Error path   -- getitem_ empty / out-of-range (+ negative OOR) and delitem_
%                     out-of-range each verify the IDENTIFIER mat2doc:IndexError
%                     (a mat2ppt:<PyExceptionName>-class id, not merely that it
%                     throws) AND the verbatim Python list message.
%   * Equivalence  -- test_equivalence_full_battery_vs_frozen_oracle replays the
%                     ENTIRE s0026 battery live (runProbes, the .m twin's body
%                     verbatim) and flatten-compares every leaf to the frozen
%                     python-docx 1.2.0 oracle (Gate-3 found ZERO divergences).
%   * Regression   -- hard-coded expected serialized-XML strings + UPPERCASE serhex
%                     of the raw UTF-8 shipping bytes vs the frozen oracle; the
%                     G-scenario document.xml SHA-256/size/full-bytes pin.
%   * Upstream     -- the line_spacing float<->Twips matrix, the firstLine/hanging
%                     sign convention, the pos-sorted w:tabs, the leader-absent
%                     default and the verbatim IndexError messages ARE the
%                     python-docx parfmt.py / tabstops.py contract; the frozen
%                     oracle IS lxml's expected output for this API sequence.
%
%   Byte-level (L1) note: every serialized-XML assertion is either the FULL
%   serialize_part_xml output as a UTF-8-decoded string (string-equality == byte
%   equality L1) or its UPPERCASE hex (serhex) vs the frozen oracle, and the
%   G-scenario is a SHA-256 + whole-bytes pin. No D-number granted any L2
%   relaxation in this WP (Gate-3: zero new; D-STYPE-1 value-exact/output-invisible,
%   D-serializer-nsdecl NON-engaged since every wrapped element is w:-only), so
%   every pin is L1.
%
%   Determinism: no network, no absolute paths. The co-located oracle + byte
%   fixture resolve relative to this file via fileparts(mfilename('fullpath'));
%   every file read is binary ('r','n'); no 'wt'. The +mat2doc package resolves
%   via the MANDATORY PathFixture(worktree-root) in TestClassSetup (WP9-F4 lesson).

    properties (Constant)
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        % The frozen s0027 word/document.xml pin (python-docx 1.2.0, frozen
        % 2026-07-30). SHA-256 equality == byte-identity (L1). Embedded so the
        % G-scenario pin is self-contained even without the byte fixture.
        SHA_S0027_DOCUMENT = "0f7a0519f6aefcde29b1baaf11800c21b7e26c3e61129b3bee28dbf3c1a17a01"
        SIZE_S0027_DOCUMENT = 1998
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied from tests\text\Test_p4_4b_text_run.m. here is
            % tests\text; the worktree root is two levels up.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\text
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. line_spacing -- the meaty logic (the equivalence pin)         %
        % =============================================================== %

        function test_line_spacing_multiple_matrix(testCase)
            % Nominal + Regression (parfmt.py 102-160, s0026 ls_multiple): the
            % float MULTIPLE matrix. Each multiple -> w:line Twips + lineRule="auto"
            % (serhex byte pin vs oracle + hard-coded L1); get reads back the float
            % (%.12g); rule maps the special members and bare MULTIPLE for 1.75.
            o = loadOracle().ls_multiple;
            cases = { ...
                "v10", 1.0,  "240", "1",    "SINGLE"; ...
                "v15", 1.5,  "360", "1.5",  "ONE_POINT_FIVE"; ...
                "v175",1.75, "420", "1.75", "MULTIPLE"; ...
                "v20", 2.0,  "480", "2",    "DOUBLE"};
            for i = 1:size(cases, 1)
                key = cases{i, 1};
                [p, pf] = newPf();
                pf.line_spacing = cases{i, 2};
                od = o.(key);
                testCase.verifyEqual(hx_e(p), string(od.serhex), ...
                    sprintf('line_spacing=%g serialized bytes (L1) vs oracle', cases{i,2}));
                testCase.verifyEqual(ser(p), pWrap(testCase, ...
                    "<w:pPr><w:spacing w:line=""" + cases{i,3} + """ w:lineRule=""auto""/></w:pPr>"), ...
                    sprintf('line_spacing=%g -> exact <w:p> (hard-coded L1)', cases{i,2}));
                testCase.verifyEqual(rv(pf.line_spacing), string(od.ls_get), 'ls get vs oracle');
                testCase.verifyEqual(rv(pf.line_spacing), cases{i,4}, 'ls get (hard-coded)');
                testCase.verifyEqual(rv(pf.line_spacing_rule), string(od.rule_get), 'rule get vs oracle');
                testCase.verifyEqual(rv(pf.line_spacing_rule), cases{i,5}, 'rule get (hard-coded)');
            end
        end

        function test_line_spacing_exactly_atleast_none_reset(testCase)
            % Nominal + Edge + Regression (parfmt.py 102-160, s0026 ls_exactly_pt18
            % / ls_atleast / ls_none_reset): EXACTLY via Pt(18); AT_LEAST-preserved
            % (rule FIRST then a Length keeps lineRule="atLeast" AND serializes
            % lineRule BEFORE line -- the attribute-order pin); []-reset retains an
            % EMPTY <w:spacing/>.
            LS = @(m) mat2doc.enum.text.WD_LINE_SPACING.(m);

            [p, pf] = newPf();
            pf.line_spacing = mat2doc.shared.Pt(18);
            o = loadOracle().ls_exactly_pt18;
            testCase.verifyEqual(hx_e(p), string(o.serhex), 'EXACTLY serhex vs oracle');
            testCase.verifyEqual(ser(p), pWrap(testCase, ...
                "<w:pPr><w:spacing w:line=""360"" w:lineRule=""exact""/></w:pPr>"), ...
                'line_spacing=Pt(18) -> lineRule="exact" (hard-coded L1)');
            testCase.verifyEqual(rv(pf.line_spacing), "228600", 'EXACTLY get -> absolute Length EMU 228600');
            testCase.verifyEqual(rv(pf.line_spacing_rule), "EXACTLY", 'EXACTLY rule');

            % AT_LEAST-preserved + attr order lineRule-before-line
            [p, pf] = newPf();
            pf.line_spacing_rule = LS("AT_LEAST");     % rule FIRST (no line)
            pf.line_spacing = mat2doc.shared.Pt(18);   % Length: AT_LEAST NOT overwritten
            o = loadOracle().ls_atleast;
            testCase.verifyEqual(hx_e(p), string(o.serhex), 'AT_LEAST-preserved serhex vs oracle');
            testCase.verifyEqual(ser(p), pWrap(testCase, ...
                "<w:pPr><w:spacing w:lineRule=""atLeast"" w:line=""360""/></w:pPr>"), ...
                'AT_LEAST-preserved: w:lineRule BEFORE w:line (attr-order pin, hard-coded L1)');
            testCase.verifyEqual(rv(pf.line_spacing), "228600", 'AT_LEAST get -> Length EMU 228600');
            testCase.verifyEqual(rv(pf.line_spacing_rule), "AT_LEAST", 'AT_LEAST rule preserved');

            % []-reset -> empty <w:spacing/>
            [p, pf] = newPf();
            pf.line_spacing = 1.5;
            pf.line_spacing = [];                      % attrs removed, element retained
            o = loadOracle().ls_none_reset;
            testCase.verifyEqual(hx_e(p), string(o.serhex), 'None-reset serhex vs oracle');
            testCase.verifyEqual(ser(p), pWrap(testCase, "<w:pPr><w:spacing/></w:pPr>"), ...
                'line_spacing=[] retains an EMPTY <w:spacing/> (hard-coded L1)');
            testCase.verifyTrue(isequal(pf.line_spacing, []), 'reset ls get -> [] (None)');
            testCase.verifyTrue(isequal(pf.line_spacing_rule, []), 'reset rule get -> [] (None)');
            testCase.verifyEqual(rv(pf.line_spacing), string(o.ls_get), 'reset ls get vs oracle');
            testCase.verifyEqual(rv(pf.line_spacing_rule), string(o.rule_get), 'reset rule get vs oracle');
        end

        function test_line_spacing_rule_set(testCase)
            % Nominal + Edge (parfmt.py 133-160, s0026 rule_set): SINGLE/1.5/DOUBLE
            % write the matching Twips line + MULTIPLE; AT_LEAST/EXACTLY write
            % lineRule only (get None); None removes lineRule leaving residue.
            LS = @(m) mat2doc.enum.text.WD_LINE_SPACING.(m);
            o = loadOracle().rule_set;
            rows = { ...
                "single",   "SINGLE"; ...
                "one_five", "ONE_POINT_FIVE"; ...
                "double",   "DOUBLE"; ...
                "at_least", "AT_LEAST"; ...
                "exactly",  "EXACTLY"};
            for i = 1:size(rows, 1)
                [p, pf] = newPf();
                pf.line_spacing_rule = LS(rows{i, 2});
                od = o.(rows{i, 1});
                testCase.verifyEqual(hx_e(p), string(od.serhex), ...
                    sprintf('rule=%s serhex vs oracle', rows{i,2}));
                testCase.verifyEqual(rv(pf.line_spacing), string(od.ls_get), ...
                    sprintf('rule=%s ls get vs oracle', rows{i,2}));
                testCase.verifyEqual(rv(pf.line_spacing_rule), string(od.rule_get), ...
                    sprintf('rule=%s rule get vs oracle', rows{i,2}));
            end
            % None removes lineRule (residual line -> get 2 / rule DOUBLE)
            [p, pf] = newPf();
            pf.line_spacing_rule = LS("DOUBLE");
            pf.line_spacing_rule = [];                 % assigns lineRule=None (removes)
            testCase.verifyEqual(hx_e(p), string(o.none.serhex), 'rule=None serhex vs oracle');
            testCase.verifyEqual(ser(p), pWrap(testCase, "<w:pPr><w:spacing w:line=""480""/></w:pPr>"), ...
                'rule=[] removes only lineRule, line residue kept (hard-coded L1)');
            testCase.verifyEqual(rv(pf.line_spacing), string(o.none.ls_get), 'rule=None ls get vs oracle');
            testCase.verifyEqual(rv(pf.line_spacing_rule), string(o.none.rule_get), 'rule=None rule get vs oracle');
        end

        % =============================================================== %
        % 2. alignment + the four bool tri-states                          %
        % =============================================================== %

        function test_alignment_and_bool_tristates(testCase)
            % Nominal + Edge (parfmt.py 12-28, 49-81, 162-176, 240-254, s0026
            % alignment/bools): alignment CENTER round-trip + []-removal; each of
            % keep_together/keep_with_next/page_break_before/widow_control True ->
            % <w:x/> (val absent) / False -> <w:x w:val="0"/> / [] -> removed.
            AL = mat2doc.enum.text.WD_ALIGN_PARAGRAPH.CENTER;

            [p, pf] = newPf();
            pf.alignment = AL;
            oa = loadOracle().alignment;
            testCase.verifyEqual(hx_e(p), string(oa.center_serhex), 'alignment=CENTER serhex vs oracle');
            testCase.verifyEqual(ser(p), pWrap(testCase, "<w:pPr><w:jc w:val=""center""/></w:pPr>"), ...
                'alignment=CENTER -> <w:jc w:val="center"/> (hard-coded L1)');
            testCase.verifyEqual(rv(pf.alignment), "CENTER", 'alignment get CENTER');
            pf.alignment = [];
            testCase.verifyEqual(hx_e(p), string(oa.none_serhex), 'alignment=[] serhex vs oracle');
            testCase.verifyEqual(ser(p), pWrap(testCase, "<w:pPr/>"), ...
                'alignment=[] removes w:jc -> empty <w:pPr/> (hard-coded L1)');
            testCase.verifyTrue(isequal(pf.alignment, []), 'alignment=[] get -> [] (None)');

            ob = loadOracle().bools;
            % localname of the child element each bool writes
            names = { ...
                "keep_together",     "keepLines"; ...
                "keep_with_next",    "keepNext"; ...
                "page_break_before", "pageBreakBefore"; ...
                "widow_control",     "widowControl"};
            for i = 1:size(names, 1)
                prop = names{i, 1}; ln = names{i, 2};
                od = ob.(prop);
                [p, pf] = newPf();
                pf.(prop) = true;
                testCase.verifyEqual(hx_e(p), string(od.true_serhex), ...
                    sprintf('%s=true serhex vs oracle', prop));
                testCase.verifyEqual(ser(p), pWrap(testCase, "<w:pPr><w:" + ln + "/></w:pPr>"), ...
                    sprintf('%s=true -> <w:%s/> (val absent, hard-coded L1)', prop, ln));
                testCase.verifyTrue(pf.(prop), sprintf('%s=true get True', prop));
                pf.(prop) = false;
                testCase.verifyEqual(hx_e(p), string(od.false_serhex), ...
                    sprintf('%s=false serhex vs oracle', prop));
                testCase.verifyEqual(ser(p), pWrap(testCase, "<w:pPr><w:" + ln + " w:val=""0""/></w:pPr>"), ...
                    sprintf('%s=false -> <w:%s w:val="0"/> (hard-coded L1)', prop, ln));
                testCase.verifyFalse(pf.(prop), sprintf('%s=false get False', prop));
                pf.(prop) = [];
                testCase.verifyEqual(hx_e(p), string(od.none_serhex), ...
                    sprintf('%s=[] serhex vs oracle', prop));
                testCase.verifyEqual(ser(p), pWrap(testCase, "<w:pPr/>"), ...
                    sprintf('%s=[] removes element -> empty <w:pPr/> (hard-coded L1)', prop));
                testCase.verifyTrue(isequal(pf.(prop), []), sprintf('%s=[] get None', prop));
            end
        end

        % =============================================================== %
        % 3. indents (sign both ways) + spacing                            %
        % =============================================================== %

        function test_indents_sign_both_ways(testCase)
            % Nominal + Edge (parfmt.py 30-47, 83-100, 178-195, s0026 indents):
            % firstLine/hanging SIGN both ways incl. Pt(0); signed left/right.
            Pt = @(n) mat2doc.shared.Pt(n);
            o = loadOracle().indents;
            % {oracle-key, property, Pt-arg, exact-<w:ind>-attr, get-EMU}
            rows = { ...
                "fli_pos",  "first_line_indent", Pt(24),  "w:firstLine=""480""", "304800"; ...
                "fli_neg",  "first_line_indent", Pt(-24), "w:hanging=""480""",   "-304800"; ...
                "fli_zero", "first_line_indent", Pt(0),   "w:firstLine=""0""",   "0"; ...
                "left_pos", "left_indent",       Pt(36),  "w:left=""720""",      "457200"; ...
                "left_neg", "left_indent",       Pt(-36), "w:left=""-720""",     "-457200"; ...
                "right_pos","right_indent",      Pt(18),  "w:right=""360""",     "228600"; ...
                "right_neg","right_indent",      Pt(-18), "w:right=""-360""",    "-228600"};
            for i = 1:size(rows, 1)
                [p, pf] = newPf();
                pf.(rows{i, 2}) = rows{i, 3};
                od = o.(rows{i, 1});
                testCase.verifyEqual(hx_e(p), string(od.serhex), ...
                    sprintf('%s serhex vs oracle', rows{i,1}));
                testCase.verifyEqual(ser(p), pWrap(testCase, "<w:pPr><w:ind " + rows{i,4} + "/></w:pPr>"), ...
                    sprintf('%s -> exact <w:ind> (hard-coded L1)', rows{i,1}));
                testCase.verifyEqual(rv(pf.(rows{i, 2})), string(od.get), ...
                    sprintf('%s get (signed EMU) vs oracle', rows{i,1}));
                testCase.verifyEqual(rv(pf.(rows{i, 2})), rows{i,5}, ...
                    sprintf('%s get (signed EMU, hard-coded)', rows{i,1}));
            end
        end

        function test_spacing_before_after(testCase)
            % Nominal (parfmt.py 197-231, s0026 spacing): space_before Pt(6) ->
            % @before="120" (76200 EMU); space_after Pt(12) -> @after="240" (152400).
            Pt = @(n) mat2doc.shared.Pt(n);
            o = loadOracle().spacing;

            [p, pf] = newPf();
            pf.space_before = Pt(6);
            testCase.verifyEqual(hx_e(p), string(o.before_serhex), 'space_before serhex vs oracle');
            testCase.verifyEqual(ser(p), pWrap(testCase, "<w:pPr><w:spacing w:before=""120""/></w:pPr>"), ...
                'space_before=Pt(6) -> @before="120" (hard-coded L1)');
            testCase.verifyEqual(rv(pf.space_before), "76200", 'space_before get -> 76200 EMU');

            [p, pf] = newPf();
            pf.space_after = Pt(12);
            testCase.verifyEqual(hx_e(p), string(o.after_serhex), 'space_after serhex vs oracle');
            testCase.verifyEqual(ser(p), pWrap(testCase, "<w:pPr><w:spacing w:after=""240""/></w:pPr>"), ...
                'space_after=Pt(12) -> @after="240" (hard-coded L1)');
            testCase.verifyEqual(rv(pf.space_after), "152400", 'space_after get -> 152400 EMU');
        end

        % =============================================================== %
        % 4. TabStops sequence -- add / pos-order / len_ / to_array        %
        % =============================================================== %

        function test_tabstops_empty_and_add_pos_order(testCase)
            % Nominal + Edge (tabstops.py 39-65, s0026 tabs_empty/tabs_add_order):
            % empty len_()==0 + to_array empty; add_tab_stop pos-SORTS three
            % out-of-order stops (2in/0.5in/1in -> pos 720/1440/2880), serhex byte
            % pin + hard-coded <w:tabs>; len_ / to_array positions.
            In = @(n) mat2doc.shared.Inches(n);

            [~, pf] = newPf();
            ts = pf.tab_stops;
            oe = loadOracle().tabs_empty;
            testCase.verifyEqual(ts.len_(), 0, 'empty TabStops len_()==0');
            testCase.verifyEqual(rv(ts.len_()), string(oe.len), 'empty len vs oracle');
            testCase.verifyEqual(numel(ts.to_array()), 0, 'empty to_array numel 0');

            [p, pf] = newPf();
            ts = pf.tab_stops;
            ts.add_tab_stop(In(2));                    % deliberately out of order
            ts.add_tab_stop(In(0.5));
            ts.add_tab_stop(In(1));
            oo = loadOracle().tabs_add_order;
            testCase.verifyEqual(hx_e(p), string(oo.serhex), 'add pos-order serhex vs oracle');
            testCase.verifyEqual(ser(p), pWrap(testCase, ...
                "<w:pPr><w:tabs>" + ...
                "<w:tab w:pos=""720"" w:val=""left""/>" + ...
                "<w:tab w:pos=""1440"" w:val=""left""/>" + ...
                "<w:tab w:pos=""2880"" w:val=""left""/>" + ...
                "</w:tabs></w:pPr>"), ...
                'add_tab_stop POS-SORTS the three stops (hard-coded L1)');
            testCase.verifyEqual(ts.len_(), 3, 'len_ after 3 adds');
            testCase.verifyEqual(positions(ts), {"457200", "914400", "1828800"}, ...
                'to_array positions in pos order (hard-coded)');
            % index reads (0-based getitem_ H1 surface)
            testCase.verifyEqual(rv(ts.getitem_(0).position),  string(oo.getitem0),     'getitem_(0) vs oracle');
            testCase.verifyEqual(rv(ts.getitem_(-1).position), string(oo.getitem_neg1), 'getitem_(-1) neg-wrap vs oracle');
            testCase.verifyEqual(rv(ts.getitem_(1).position),  string(oo.getitem1),     'getitem_(1) vs oracle');
        end

        function test_tabstops_getitem_h1_indexerror(testCase)
            % Edge / error-path (tabstops.py 20-37, H1, s0026 tabs_empty/
            % tabs_add_order): THE TabStops-H1 pin. getitem_ on empty raises
            % mat2doc:IndexError "TabStops object is empty"; on a 3-stop list
            % getitem_(3)/(-4) raise mat2doc:IndexError "list index out of range".
            % Verify the IDENTIFIER (a mat2ppt:<PyExceptionName>-class id) AND the
            % verbatim Python list message -- not merely that it throws.
            In = @(n) mat2doc.shared.Inches(n);
            oe = loadOracle().tabs_empty;
            oo = loadOracle().tabs_add_order;

            [~, pf] = newPf();
            ts = pf.tab_stops;
            ME = captureError(@() ts.getitem_(0));
            testCase.verifyEqual(string(ME.identifier), "mat2doc:IndexError", ...
                'getitem_ on empty -> mat2doc:IndexError identifier');
            testCase.verifyEqual(string(ME.message), "TabStops object is empty", ...
                'getitem_ on empty -> verbatim "TabStops object is empty"');
            testCase.verifyEqual(string(ME.message), string(oe.getitem0_err), ...
                'empty getitem message vs oracle');

            [~, pf] = newPf();
            ts = pf.tab_stops;
            ts.add_tab_stop(In(2)); ts.add_tab_stop(In(0.5)); ts.add_tab_stop(In(1));
            for idx = [3, -4]
                ME = captureError(@() ts.getitem_(idx));
                testCase.verifyEqual(string(ME.identifier), "mat2doc:IndexError", ...
                    sprintf('getitem_(%d) OOR -> mat2doc:IndexError identifier', idx));
                testCase.verifyEqual(string(ME.message), "list index out of range", ...
                    sprintf('getitem_(%d) OOR -> verbatim "list index out of range"', idx));
            end
            testCase.verifyEqual(string(captureError(@() ts.getitem_(3)).message), ...
                string(oo.getitem3_err), 'getitem_(3) OOR message vs oracle');
            testCase.verifyEqual(string(captureError(@() ts.getitem_(-4)).message), ...
                string(oo.getitem_neg4_err), 'getitem_(-4) OOR message vs oracle');
        end

        function test_tabstops_delitem_and_clear_all(testCase)
            % Nominal + Edge + error-path (tabstops.py 20-29, 67-69, s0026
            % tabs_delitem/tabs_clear_all): delitem_ middle; delitem_(9) ->
            % mat2doc:IndexError "tab index out of range"; delitem_ down-to-empty
            % REMOVES <w:tabs> (empty <w:pPr/>); clear_all removes <w:tabs>.
            In = @(n) mat2doc.shared.Inches(n);
            od = loadOracle().tabs_delitem;

            [p, pf] = newPf();
            ts = pf.tab_stops;
            ts.add_tab_stop(In(0.5)); ts.add_tab_stop(In(1)); ts.add_tab_stop(In(2));
            ts.delitem_(1);                            % remove the middle (1in) stop
            testCase.verifyEqual(hx_e(p), string(od.del_mid_serhex), 'delitem_ middle serhex vs oracle');
            testCase.verifyEqual(positions(ts), {"457200", "1828800"}, ...
                'delitem_ middle leaves [0.5in, 2in] (hard-coded)');

            ME = captureError(@() ts.delitem_(9));
            testCase.verifyEqual(string(ME.identifier), "mat2doc:IndexError", ...
                'delitem_(9) OOR -> mat2doc:IndexError identifier');
            testCase.verifyEqual(string(ME.message), "tab index out of range", ...
                'delitem_(9) OOR -> verbatim "tab index out of range"');
            testCase.verifyEqual(string(ME.message), string(od.del_oor_err), 'delitem OOR message vs oracle');

            ts.delitem_(0); ts.delitem_(0);            % down to empty
            testCase.verifyEqual(hx_e(p), string(od.del_to_empty_serhex), 'delitem down-to-empty serhex vs oracle');
            testCase.verifyEqual(ser(p), pWrap(testCase, "<w:pPr/>"), ...
                'delitem_ down to empty REMOVES <w:tabs> -> empty <w:pPr/> (hard-coded L1)');
            testCase.verifyEqual(ts.len_(), 0, 'len_ 0 after down-to-empty');

            % clear_all
            [p, pf] = newPf();
            ts = pf.tab_stops;
            ts.add_tab_stop(In(1)); ts.add_tab_stop(In(2));
            ts.clear_all();
            oc = loadOracle().tabs_clear_all;
            testCase.verifyEqual(hx_e(p), string(oc.serhex), 'clear_all serhex vs oracle');
            testCase.verifyEqual(ser(p), pWrap(testCase, "<w:pPr/>"), ...
                'clear_all removes <w:tabs> -> empty <w:pPr/> (hard-coded L1)');
            testCase.verifyEqual(ts.len_(), 0, 'len_ 0 after clear_all');
        end

        % =============================================================== %
        % 5. TabStop properties + the position-setter quirk                %
        % =============================================================== %

        function test_tabstop_properties_default_and_explicit(testCase)
            % Nominal + Edge (tabstops.py 78-123, s0026 tab_props): default
            % add_tab_stop(Inches(1)) -> position 914400 / alignment LEFT / leader
            % SPACES with **@w:leader ABSENT** (OptionalAttribute default-removal,
            % w:val="left" present); explicit CENTER/DOTS -> the full attr set.
            In = @(n) mat2doc.shared.Inches(n);
            TA = mat2doc.enum.text.WD_TAB_ALIGNMENT.CENTER;
            TL = mat2doc.enum.text.WD_TAB_LEADER.DOTS;
            o = loadOracle().tab_props;

            [p, pf] = newPf();
            ts = pf.tab_stops;
            ts.add_tab_stop(In(1));                    % LEFT / SPACES defaults
            t0 = ts.getitem_(0);
            testCase.verifyEqual(rv(t0.position), "914400", 'default position 914400');
            testCase.verifyEqual(rv(t0.alignment), "LEFT",  'default alignment LEFT');
            testCase.verifyEqual(rv(t0.leader),    "SPACES",'default leader SPACES');
            testCase.verifyEqual(rv(t0.position), string(o.default_pos),   'default pos vs oracle');
            testCase.verifyEqual(rv(t0.alignment), string(o.default_align), 'default align vs oracle');
            testCase.verifyEqual(rv(t0.leader),    string(o.default_leader),'default leader vs oracle');
            % @w:leader ABSENT (default-removal) while w:val="left" present
            testCase.verifyFalse(contains(ser(p), "w:leader"), ...
                'SPACES-default leader is @w:leader-ABSENT in the XML');
            testCase.verifyEqual(ternary(contains(ser(p), "w:leader")), string(o.default_has_leader), ...
                'default_has_leader vs oracle');
            testCase.verifyTrue(contains(ser(p), "w:val=""left"""), 'w:val="left" present (RequiredAttribute)');

            [p, pf] = newPf();
            ts = pf.tab_stops;
            ts.add_tab_stop(In(3), TA, TL);
            tE = ts.getitem_(0);
            testCase.verifyEqual(hx_e(p), string(o.explicit_serhex), 'explicit CENTER/DOTS serhex vs oracle');
            testCase.verifyEqual(ser(p), pWrap(testCase, ...
                "<w:pPr><w:tabs><w:tab w:pos=""4320"" w:val=""center"" w:leader=""dot""/></w:tabs></w:pPr>"), ...
                'explicit -> <w:tab w:pos="4320" w:val="center" w:leader="dot"/> (hard-coded L1)');
            testCase.verifyEqual(rv(tE.position),  string(o.explicit_pos),   'explicit pos vs oracle');
            testCase.verifyEqual(rv(tE.alignment), string(o.explicit_align), 'explicit align vs oracle');
            testCase.verifyEqual(rv(tE.leader),    string(o.explicit_leader),'explicit leader vs oracle');
            testCase.verifyEqual(rv(tE.position), "2743200", 'explicit position 2743200 (hard-coded)');
        end

        function test_tabstop_position_setter_quirk(testCase)
            % Edge (tabstops.py 118-123, s0026 tab_quirk): the position-setter quirk
            % is a GENUINE python-docx behavior (Gate-2 RULED do-NOT-correct). Two
            % fresh proxies over the SAME element compare ==; after t.position=
            % Inches(1.5) (re-inserts at 1.5in, removes the old 0.5in element) both
            % proxies' eq basis stays the OLD detached element so t==t2 is STILL
            % True, while t.position reads the NEW 1371600 via the reassigned tab_,
            % and to_array re-sorts to [914400, 1371600, 1828800].
            In = @(n) mat2doc.shared.Inches(n);
            o = loadOracle().tab_quirk;

            [~, pf] = newPf();
            ts = pf.tab_stops;
            ts.add_tab_stop(In(0.5)); ts.add_tab_stop(In(1)); ts.add_tab_stop(In(2));
            t  = ts.getitem_(0);                       % the 0.5in stop
            t2 = ts.getitem_(0);                       % fresh proxy, SAME element
            testCase.verifyTrue(t == t2, 'two proxies over the same element compare == (before)');
            testCase.verifyEqual(rv(t == t2), string(o.eq_before), 'eq_before vs oracle');

            t.position = In(1.5);                      % re-sort: removes 0.5, inserts 1.5
            testCase.verifyTrue(t == t2, ...
                'after position set both proxies still track the OLD detached element (quirk) -> ==');
            testCase.verifyEqual(rv(t == t2), string(o.eq_after), 'eq_after vs oracle');
            testCase.verifyEqual(rv(t.position), "1371600", 't.position reads the NEW 1371600 via reassigned tab_');
            testCase.verifyEqual(rv(t.position), string(o.pos_after), 'pos_after vs oracle');
            testCase.verifyEqual(positions(ts), {"914400", "1371600", "1828800"}, ...
                'to_array re-sorts after the position set (hard-coded)');
        end

        % =============================================================== %
        % 6. G-scenario document byte pin (the M2 de-risker)               %
        % =============================================================== %

        function test_gscenario_document_byte_pin(testCase)
            % Regression (THE G-scenario byte pin, s0027): a fresh mat2doc.Document()
            % + four formatted paragraphs (buildGScenario, the s0027 .m twin body
            % verbatim) saved -> the extracted word/document.xml is byte-identical
            % (size + SHA-256 + whole-bytes) to the frozen python-docx 1.2.0
            % reference. Closest pin to M2's add_paragraph formatting path (P4-6).
            docxBytes = buildGScenario();              % mat2doc.Document().save bytes
            [blobs, names] = zipEntryList(docxBytes);
            i = find(names == "word/document.xml", 1);
            testCase.assertNotEmpty(i, 'saved package must contain word/document.xml');
            got = blobs{i};

            testCase.verifyEqual(numel(got), testCase.SIZE_S0027_DOCUMENT, ...
                sprintf('word/document.xml must be exactly %d B', testCase.SIZE_S0027_DOCUMENT));
            testCase.verifyEqual(sha256hex(got), testCase.SHA_S0027_DOCUMENT, ...
                'word/document.xml SHA-256 must equal the frozen s0027 oracle (byte-identical L1)');
            want = loadDocumentFixture();
            verifyByteIdentical(testCase, got, want, ...
                'G-scenario word/document.xml == frozen s0027 reference');
        end

        % =============================================================== %
        % 7. M1 sanity (light) + StoryChild/ElementProxy lineage           %
        % =============================================================== %

        function test_paragraphformat_is_elementproxy_m1_light(testCase)
            % M1 sanity (light): P4-5a is API/proxy tier -- byte-neutral, no
            % styles.xml/document.xml default-template pin needed here (the full
            % 17/17 M1 sweep stays owned by Test_p1_8_skeleton_m1; validate_P4-5a
            % section 3 confirms M1 17/17 unchanged). Spot-check that all three
            % proxies are ElementProxy (so H5 eq/identity + the wrapped-element
            % accessor come from the shared base) and a live round-trip works.
            [p, pf] = newPf();
            testCase.verifyClass(pf, 'mat2doc.text.ParagraphFormat');
            testCase.verifyTrue(isa(pf, 'mat2doc.shared.ElementProxy'), ...
                'ParagraphFormat must be a mat2doc.shared.ElementProxy');
            ts = pf.tab_stops;
            testCase.verifyClass(ts, 'mat2doc.text.TabStops');
            testCase.verifyTrue(isa(ts, 'mat2doc.shared.ElementProxy'), ...
                'TabStops must be a mat2doc.shared.ElementProxy');
            ts.add_tab_stop(mat2doc.shared.Inches(1));
            t0 = ts.getitem_(0);
            testCase.verifyClass(t0, 'mat2doc.text.TabStop');
            testCase.verifyTrue(isa(t0, 'mat2doc.shared.ElementProxy'), ...
                'TabStop must be a mat2doc.shared.ElementProxy');
            % live spot round-trip through the wrapped element
            pf.alignment = mat2doc.enum.text.WD_ALIGN_PARAGRAPH.CENTER;
            testCase.verifyEqual(rv(pf.alignment), "CENTER", 'ParagraphFormat.alignment spot round-trip');
            testCase.verifyTrue(pf == pf, 'ElementProxy == (H5 element identity) holds');
        end

        % =============================================================== %
        % 8. EQUIVALENCE -- full s0026 battery vs the frozen oracle         %
        % =============================================================== %

        function test_equivalence_full_battery_vs_frozen_oracle(testCase)
            % Equivalence: replay the ENTIRE s0026 battery (runProbes -- the .m
            % twin's body VERBATIM: ls_multiple / ls_exactly_pt18 / ls_atleast /
            % ls_none_reset / rule_set / alignment / bools / indents / spacing /
            % tabs_empty / tabs_add_order / tab_props / tab_quirk / tabs_delitem /
            % tabs_clear_all) and flatten-compare EVERY leaf to the frozen
            % python-docx 1.2.0 oracle copied into data\s0026_probe_oracle.json.
            % Gate-3 found ZERO divergences (probe_diff MATCH 120/120), so every
            % leaf must be byte/value-identical.
            port   = runProbes();
            oracle = loadOracle();
            pMap = containers.Map('KeyType', 'char', 'ValueType', 'char');
            oMap = containers.Map('KeyType', 'char', 'ValueType', 'char');
            flattenLeaves(port,   '', pMap);
            flattenLeaves(oracle, '', oMap);

            pKeys = sort(pMap.keys());
            oKeys = sort(oMap.keys());
            testCase.verifyEqual(pKeys, oKeys, ...
                'the replayed battery and the frozen oracle must have identical leaf keys');
            % Non-trivial size guard (guards against a silent-empty replay).
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
end

% ===================== file-local helpers ============================== %

function s = decl()
    s = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>";
end

function [p, pf] = newPf()
    p  = mat2doc.oxml.OxmlElement("w:p");
    pf = mat2doc.text.ParagraphFormat(p);
end

function s = ser(e)
    % Decode serialize_part_xml bytes to a string (string-equality == byte-identical L1).
    s = string(native2unicode(mat2doc.oxml.serialize_part_xml(e), "UTF-8"));
end

function h = hx(raw)
    % UPPERCASE hex of raw UTF-8 shipping bytes (matches Python bytes.hex().upper()).
    h = string(sprintf('%02X', uint8(raw)));
end

function h = hx_e(e)
    % serhex of the whole element's serialize_part_xml shipping bytes -- a byte pin.
    h = hx(mat2doc.oxml.serialize_part_xml(e));
end

function s = pWrap(testCase, body)
    % decl + newline + <w:p xmlns:w="...">BODY</w:p>. The paragraph is created via
    % OxmlElement("w:p") -> declares xmlns:w. Used by the hard-coded XML regression pins.
    s = decl() + newline + "<w:p xmlns:w=""" + testCase.W + """>" + string(body) + "</w:p>";
end

function L = positions(ts)
    % 1xN cell of EMU-string positions over ts.to_array() (rv repr).
    arr = ts.to_array();
    L = cell(1, numel(arr));
    for k = 1:numel(arr)
        L{k} = rv(arr(k).position);
    end
end

function s = ternary(tf)
    if tf, s = "True"; else, s = "False"; end
end

function s = rv(x)
    % Uniform accessor repr mirroring the s0026 rv(). Order matters: Length
    % (a double subclass) is tested BEFORE the plain-double branch. None->"None",
    % bool->"True"/"False", Length->EMU int, line-spacing float multiple->"%.12g",
    % enum member->NAME.
    if isequal(x, [])
        s = "None";
    elseif islogical(x)
        if x, s = "True"; else, s = "False"; end
    elseif isa(x, "mat2doc.shared.Length")
        s = string(sprintf('%.0f', double(x)));        % EMU integer
    elseif isnumeric(x)
        s = string(sprintf('%.12g', double(x)));       % line-spacing float multiple
    else
        s = string(x);                                 % enum member NAME
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
    % Read the co-located frozen s0026 oracle in BINARY mode (no CRLF translation)
    % and decode UTF-8 -> struct. jsondecode is whitespace/line-ending agnostic, so
    % no `* binary` pin is needed for this value/serhex fixture (s0024 precedent).
    here = fileparts(mfilename('fullpath'));
    p = fullfile(here, 'data', 's0026_probe_oracle.json');
    fid = fopen(p, 'r', 'n');
    if fid < 0
        error('mat2doc:test:openFail', 'cannot open frozen oracle %s', p);
    end
    raw = fread(fid, Inf, '*uint8')';
    fclose(fid);
    o = jsondecode(native2unicode(raw, 'UTF-8'));
end

function want = loadDocumentFixture()
    % The frozen s0027 word/document.xml byte fixture (co-located `* binary`
    % .gitattributes so it is checked out byte-for-byte). Read in BINARY mode.
    here = fileparts(mfilename('fullpath'));
    p = fullfile(here, 'data', 's0027_document.xml');
    want = readBytes(p);
end

function docxBytes = buildGScenario()
    % Replay the s0027 G-scenario body VERBATIM (validation\mat2doc\scenarios\
    % s0027_p4_5a_parfmt_gscenario.m): a fresh mat2doc.Document() + four formatted
    % paragraphs appended at the oxml level (body.add_p(), inserted before the
    % w:sectPr; Paragraph proxy is P4-6). Returns the whole-package .save() bytes.
    LS = @(m) mat2doc.enum.text.WD_LINE_SPACING.(m);
    AL = mat2doc.enum.text.WD_ALIGN_PARAGRAPH.CENTER;
    TA_CENTER = mat2doc.enum.text.WD_TAB_ALIGNMENT.CENTER;
    TA_RIGHT  = mat2doc.enum.text.WD_TAB_ALIGNMENT.RIGHT;
    TL_DOTS   = mat2doc.enum.text.WD_TAB_LEADER.DOTS;
    Pt = @(n) mat2doc.shared.Pt(n);
    In = @(n) mat2doc.shared.Inches(n);

    d = mat2doc.Document();
    body = d.element().body;

    % ---- P1: 1.75 multiple ----
    pf = mat2doc.text.ParagraphFormat(body.add_p());
    pf.line_spacing = 1.75;

    % ---- P2: AT_LEAST + Pt(18) ----
    pf = mat2doc.text.ParagraphFormat(body.add_p());
    pf.line_spacing_rule = LS("AT_LEAST");
    pf.line_spacing = Pt(18);

    % ---- P3: hanging indent + 3 tab stops ----
    pf = mat2doc.text.ParagraphFormat(body.add_p());
    pf.first_line_indent = Pt(-18);
    ts = pf.tab_stops;
    ts.add_tab_stop(In(1));                            % LEFT / SPACES defaults
    ts.add_tab_stop(In(3), TA_CENTER, TL_DOTS);
    ts.add_tab_stop(In(2), TA_RIGHT);

    % ---- P4: 1.5 / center / widow / keepLines ----
    pf = mat2doc.text.ParagraphFormat(body.add_p());
    pf.line_spacing_rule = LS("ONE_POINT_FIVE");
    pf.alignment = AL;
    pf.widow_control = true;
    pf.keep_together = true;

    tmp = [tempname '.docx'];
    cleanup = onCleanup(@() deleteIfExists(tmp)); %#ok<NASGU>
    d.save(tmp);
    docxBytes = readBytes(tmp);
end

function P = runProbes()
    % Replay the s0026 probe sequence (the .m twin's P-struct body, VERBATIM
    % tags/inputs/order) and return the nested struct of tagged canonical values.
    % Embedded so the Equivalence leg is self-contained (the validation-folder
    % scenario is NOT on the toolbox path and must not be a dependency). Mirrors
    % validation\mat2doc\scenarios\s0026_p4_5a_parfmt_probe.m.
    P = struct();
    AL  = mat2doc.enum.text.WD_ALIGN_PARAGRAPH.CENTER;
    TA  = mat2doc.enum.text.WD_TAB_ALIGNMENT.CENTER;
    TL  = mat2doc.enum.text.WD_TAB_LEADER.DOTS;
    lsm = @(m) mat2doc.enum.text.WD_LINE_SPACING.(m);
    Pt  = @(n) mat2doc.shared.Pt(n);
    In  = @(n) mat2doc.shared.Inches(n);

    % ===================== line_spacing float MULTIPLE matrix ==============
    lm = struct();
    mults = {"v10", 1.0; "v15", 1.5; "v175", 1.75; "v20", 2.0};
    for k = 1:size(mults, 1)
        [p, pf] = newPf();
        pf.line_spacing = mults{k, 2};
        lm.(mults{k, 1}) = struct( ...
            "serhex",   hx_e(p), ...
            "ls_get",   rv(pf.line_spacing), ...
            "rule_get", rv(pf.line_spacing_rule));
    end
    P.ls_multiple = lm;

    % ===================== line_spacing = Pt(18) -> EXACTLY ================
    [p, pf] = newPf();
    pf.line_spacing = Pt(18);
    P.ls_exactly_pt18 = struct("serhex", hx_e(p), ...
        "ls_get", rv(pf.line_spacing), "rule_get", rv(pf.line_spacing_rule));

    % ===================== AT_LEAST preserved + attr order ================
    [p, pf] = newPf();
    pf.line_spacing_rule = lsm("AT_LEAST");
    pf.line_spacing = Pt(18);
    P.ls_atleast = struct("serhex", hx_e(p), ...
        "ls_get", rv(pf.line_spacing), "rule_get", rv(pf.line_spacing_rule));

    % ===================== None reset -> empty <w:spacing/> ===============
    [p, pf] = newPf();
    pf.line_spacing = 1.5;
    pf.line_spacing = [];
    P.ls_none_reset = struct("serhex", hx_e(p), ...
        "ls_get", rv(pf.line_spacing), "rule_get", rv(pf.line_spacing_rule));

    % ===================== line_spacing_rule SET ==========================
    rs = struct();
    rmembers = {"single","SINGLE"; "one_five","ONE_POINT_FIVE"; "double","DOUBLE"; ...
                "at_least","AT_LEAST"; "exactly","EXACTLY"};
    for k = 1:size(rmembers, 1)
        [p, pf] = newPf();
        pf.line_spacing_rule = lsm(rmembers{k, 2});
        rs.(rmembers{k, 1}) = struct("serhex", hx_e(p), ...
            "ls_get", rv(pf.line_spacing), "rule_get", rv(pf.line_spacing_rule));
    end
    [p, pf] = newPf();
    pf.line_spacing_rule = lsm("DOUBLE");
    pf.line_spacing_rule = [];
    rs.none = struct("serhex", hx_e(p), ...
        "ls_get", rv(pf.line_spacing), "rule_get", rv(pf.line_spacing_rule));
    P.rule_set = rs;

    % ===================== alignment ======================================
    al = struct();
    [p, pf] = newPf();
    pf.alignment = AL;
    al.center_serhex = hx_e(p); al.center_get = rv(pf.alignment);
    pf.alignment = [];
    al.none_serhex = hx_e(p); al.none_get = rv(pf.alignment);
    P.alignment = al;

    % ===================== bool tri-states ================================
    bl = struct();
    bnames = ["keep_together","keep_with_next","page_break_before","widow_control"];
    for name = bnames
        [p, pf] = newPf();
        pf.(name) = true;
        d = struct("true_serhex", hx_e(p), "true_get", rv(pf.(name)));
        pf.(name) = false;
        d.false_serhex = hx_e(p); d.false_get = rv(pf.(name));
        pf.(name) = [];
        d.none_serhex = hx_e(p); d.none_get = rv(pf.(name));
        bl.(name) = d;
    end
    P.bools = bl;

    % ===================== indents (sign both ways) =======================
    ind = struct();
    ind.fli_pos   = doInd("first_line_indent", Pt(24));
    ind.fli_neg   = doInd("first_line_indent", Pt(-24));
    ind.fli_zero  = doInd("first_line_indent", Pt(0));
    ind.left_pos  = doInd("left_indent", Pt(36));
    ind.left_neg  = doInd("left_indent", Pt(-36));
    ind.right_pos = doInd("right_indent", Pt(18));
    ind.right_neg = doInd("right_indent", Pt(-18));
    P.indents = ind;

    % ===================== spacing before/after ===========================
    sp = struct();
    [p, pf] = newPf();
    pf.space_before = Pt(6);
    sp.before_serhex = hx_e(p); sp.before_get = rv(pf.space_before);
    [p, pf] = newPf();
    pf.space_after = Pt(12);
    sp.after_serhex = hx_e(p); sp.after_get = rv(pf.space_after);
    P.spacing = sp;

    % ===================== TabStops: empty ================================
    te = struct();
    [~, pf] = newPf();
    ts = pf.tab_stops;
    te.len = rv(ts.len_());
    te.getitem0_err = errMsg(@() ts.getitem_(0));
    te.toarray_len = rv(numel(ts.to_array()));
    P.tabs_empty = te;

    % ===================== TabStops: add + pos-order ======================
    ao = struct();
    [p, pf] = newPf();
    ts = pf.tab_stops;
    ts.add_tab_stop(In(2));
    ts.add_tab_stop(In(0.5));
    ts.add_tab_stop(In(1));
    ao.serhex = hx_e(p);
    ao.len = rv(ts.len_());
    ao.positions = positions(ts);
    t0 = ts.getitem_(0);    ao.getitem0 = rv(t0.position);
    tn1 = ts.getitem_(-1);  ao.getitem_neg1 = rv(tn1.position);
    t1 = ts.getitem_(1);    ao.getitem1 = rv(t1.position);
    ao.getitem3_err = errMsg(@() ts.getitem_(3));
    ao.getitem_neg4_err = errMsg(@() ts.getitem_(-4));
    P.tabs_add_order = ao;

    % ===================== TabStop properties =============================
    tp = struct();
    [p, pf] = newPf();
    ts = pf.tab_stops;
    ts.add_tab_stop(In(1));
    t0 = ts.getitem_(0);
    tp.default_pos = rv(t0.position);
    tp.default_align = rv(t0.alignment);
    tp.default_leader = rv(t0.leader);
    tp.default_has_leader = ternary(contains(ser(p), "w:leader"));
    [p, pf] = newPf();
    ts = pf.tab_stops;
    ts.add_tab_stop(In(3), TA, TL);
    tE = ts.getitem_(0);
    tp.explicit_serhex = hx_e(p);
    tp.explicit_pos = rv(tE.position);
    tp.explicit_align = rv(tE.alignment);
    tp.explicit_leader = rv(tE.leader);
    P.tab_props = tp;

    % ===================== TabStop.position quirk =========================
    q = struct();
    [~, pf] = newPf();
    ts = pf.tab_stops;
    ts.add_tab_stop(In(0.5));
    ts.add_tab_stop(In(1));
    ts.add_tab_stop(In(2));
    t = ts.getitem_(0);
    t2 = ts.getitem_(0);
    q.eq_before = rv(t == t2);
    t.position = In(1.5);
    q.eq_after = rv(t == t2);
    q.pos_after = rv(t.position);
    q.positions_after = positions(ts);
    P.tab_quirk = q;

    % ===================== TabStops.delitem_ ==============================
    dl = struct();
    [p, pf] = newPf();
    ts = pf.tab_stops;
    ts.add_tab_stop(In(0.5));
    ts.add_tab_stop(In(1));
    ts.add_tab_stop(In(2));
    ts.delitem_(1);
    dl.del_mid_serhex = hx_e(p);
    dl.del_mid_positions = positions(ts);
    dl.del_oor_err = errMsg(@() ts.delitem_(9));
    ts.delitem_(0);
    ts.delitem_(0);
    dl.del_to_empty_serhex = hx_e(p);
    dl.del_to_empty_len = rv(ts.len_());
    P.tabs_delitem = dl;

    % ===================== TabStops.clear_all =============================
    ca = struct();
    [p, pf] = newPf();
    ts = pf.tab_stops;
    ts.add_tab_stop(In(1));
    ts.add_tab_stop(In(2));
    ts.clear_all();
    ca.serhex = hx_e(p);
    ca.len = rv(ts.len_());
    P.tabs_clear_all = ca;
end

function d = doInd(attr, value)
    [p, pf] = newPf();
    pf.(attr) = value;
    d = struct("serhex", hx_e(p), "get", rv(pf.(attr)));
end

function m = errMsg(fn)
    try
        fn();
        m = "NO-RAISE";
    catch ME
        m = string(ME.message);
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

% -------- zip / byte helpers (file-local, copied from Test_p1_8_skeleton_m1) ----

function [blobs, names] = zipEntryList(zipBytes)
    % Enumerate an in-memory zip in stream (write) order -> ordered {names, blobs}.
    % java.util.zip.ZipInputStream reads local file headers in physical order.
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

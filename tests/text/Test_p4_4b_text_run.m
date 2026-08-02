classdef Test_p4_4b_text_run < matlab.unittest.TestCase
% TEST_P4_4B_TEXT_RUN  Gate-4 permanent unit tests for Mat2Doc P4-4b
%   (src/docx/text/run.py -> +mat2doc\+text\Run and _Text -> +mat2doc\+text\Text_
%   -- a P4 API/proxy-tier WP).
%
%   Run is a pure API/proxy over the already byte-validated CT_R (P4-1b) and Font
%   (P4-4a); _Text is the inert `<w:t>` wrapper. Neither adds a register_element_cls
%   row, an oxml class, or a serialization-path change -> API/proxy tier: no bytes,
%   no registry, no M1 risk (validate_P4-4b §2 confirms M1 17/17 unchanged, so this
%   class does NOT re-pin styles.xml/document.xml -- Test_p1_8_skeleton_m1 owns the
%   full 17/17 M1 sweep). This class permanently FREEZES the P4-4b BEHAVIORAL
%   surface -- the API-tier equivalence pin -- byte/value-identical to
%   python-docx 1.2.0.
%
%   The guarantees frozen here (each verified byte/value-identical at Gate-3,
%   probe_diff MATCH 66/66, stub-safety 4/4, zero divergences, zero new D-numbers):
%
%     (ADD_BREAK map -- the byte pin, Gate-2-requested) all 6 mapped WD_BREAK
%     members + no-arg default + the TEXT_WRAPPING alias, serialized as UPPERCASE
%     serhex vs the frozen s0025 oracle PLUS hard-coded XML-string regression pins:
%       LINE / default -> <w:br/>            (no @w:type, no @w:clear)
%       PAGE           -> <w:br w:type="page"/>
%       COLUMN         -> <w:br w:type="column"/>
%       LINE_CLEAR_LEFT/RIGHT/ALL -> <w:br w:clear="left|right|all"/>
%                         **only @w:clear, NO @w:type** (the load-bearing pin:
%                         type_="textWrapping" is the CT_Br @w:type DEFAULT, so the
%                         OptionalAttribute setter REMOVES it -- D-delta-1)
%       TEXT_WRAPPING (alias of LINE_CLEAR_ALL) -> <w:br w:clear="all"/> (identical)
%     SECTION_CONTINUOUS/EVEN_PAGE/NEXT_PAGE/ODD_PAGE -> raise mat2doc:KeyError
%     (identifier verified, not merely that it throws) with children=0 on the run
%     (the error precedes add_br -- no partial mutation).
%
%     (TEXT get/set) text = "a\tb\nc\rd" -> <w:t>a</w:t><w:tab/><w:t>b</w:t>
%     <w:br/><w:t>c</w:t><w:br/><w:t>d</w:t> (serhex byte pin) and get ->
%     "a\tb\nc\nd" (the \r reads back as \n); empty set -> self-closing <w:r/> +
%     get ""; re-set replaces content (no append).
%
%     (FONT delegation) bold/italic/underline forward to the Font proxy; combined
%     -> <w:rPr><w:b/><w:i/><w:u w:val="single"/></w:rPr> (serhex byte pin) + gets
%     True/True/True; bold=[] removes w:b (get None); bold=false -> <w:b w:val="0"/>
%     (get False). font returns a FRESH Font each access over the SAME element
%     (== True by element identity).
%
%     (CONTAINS_PAGE_BREAK, H4) empty -> False; a parsed <w:lastRenderedPageBreak/>
%     -> True; a HARD <w:br w:type="page"/> -> False (an author page-break is NOT
%     counted).
%
%     (CONTENT ops) add_text(" hi ") -> <w:t xml:space="preserve"> hi </w:t> and
%     returns a mat2doc.text.Text_ proxy; add_text("plain") -> <w:t>plain</w:t>
%     (no xml:space); add_tab() -> <w:tab/>; clear() after bold+add_text keeps
%     <w:rPr><w:b/></w:rPr> only (content dropped, formatting preserved) and RETURNS
%     SELF (handle ==). mark_comment_range(run2, 3) over two runs in a w:p -> child
%     localnames [commentRangeStart, r, r, commentRangeEnd, r], reference-run style
%     "CommentReference", full-parent serhex byte-identical (w:id="3").
%
%     (_Text surface) inert: isprop(t,"text") False, 0 public members (the faithful
%     port of the bare v1.2.0 object subclass, run.py 252-257 -- no text accessor).
%
%     (H5 / StoryChild) run.font == run.font True (fresh proxy, same w:r);
%     Run(r, partStub).part() delegates via StoryChild to parent.part() -> "STUBPART".
%
%     (Stub-safety) add_picture / iter_inner_content / style get / style set each
%     raise mat2doc:notYetPorted (identifier verified, 4/4).
%
%   Provenance (Gate-1..3, all 2026-07-30):
%     * Audit    : validation\mat2doc\audit_P4-4b_text_run.md (Porter Gate-1 +
%                  Fable Gate-2 adversarial APPROVE, zero defects; 31/31 paired
%                  records byte-identical; add_break/text/font/contains_page_break/
%                  content/stubs all verified; _Text no-text-property RATIFIED).
%     * Validate : validation\mat2doc\validate_P4-4b_text_run.md (Gate-3 PASS --
%                  probe_diff MATCH 66/66; M1 17/17 byte-identical; stub-safety 4/4;
%                  H5/StoryChild .part delegation + fresh-Font; regression 598/598;
%                  ZERO divergences, ZERO new D-numbers).
%     * Scenario : validation\mat2doc\scenarios\s0025_p4_4b_text_run.{py,m}
%                  (the probe sequence replayed VERBATIM by runProbes() below).
%     * Frozen ref (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE):
%         references\s0025\probe.json -- copied verbatim (self-contained) into
%           tests\text\data\s0025_probe_oracle.json (the Equivalence replay set;
%           value/serhex JSON, jsondecode is line-ending agnostic -> no `* binary`
%           .gitattributes needed, per the s0020/s0023/s0024 precedent).
%
%   Coverage taxonomy
%   -----------------
%   * Nominal      -- add_break happy paths, text set/get, bold/italic/underline,
%                     add_text/add_tab/clear, mark_comment_range, font/part wiring.
%   * Edge         -- no-arg default break, the TEXT_WRAPPING alias, the LINE_CLEAR_*
%                     @w:type-removal, empty-text set (self-closing run), non-mapped
%                     SECTION_* error path (identifier mat2doc:KeyError, not merely
%                     that it throws), the hard-break-not-counted H4 case, the inert
%                     _Text surface, and the 4 notYetPorted stub paths.
%   * Equivalence  -- test_equivalence_full_battery_vs_frozen_oracle replays the
%                     ENTIRE s0025 battery live (runProbes, the .m twin's body
%                     verbatim) and flatten-compares every leaf to the frozen
%                     python-docx 1.2.0 oracle (Gate-3 found ZERO divergences).
%   * Regression   -- hard-coded expected serialized-XML strings + UPPERCASE serhex
%                     of the raw UTF-8 shipping bytes vs the frozen oracle.
%   * Upstream     -- the add_break (type_, clear) map, the text \t/\n/\r splitting,
%                     the font tri-state, and the mark_comment_range surface ARE the
%                     python-docx run.py API; the frozen oracle IS lxml's expected
%                     output for this API sequence.
%
%   Byte-level (L1) note: every serialized-XML assertion is either the FULL
%   serialize_part_xml output as a UTF-8-decoded string (string-equality == byte
%   equality L1) or its UPPERCASE hex (serhex) vs the frozen oracle. No D-number
%   granted any L2 relaxation in this WP (Gate-3: zero new, D-delta-1/D10 adopt-only,
%   D-serializer-nsdecl NON-engaged since every run is w:-only), so every pin is L1.
%
%   Determinism: no network, no absolute paths. The co-located oracle resolves
%   relative to this file via fileparts(mfilename('fullpath')); every file read is
%   binary ('r','n'); no 'wt'. The +mat2doc package AND the co-located
%   PartStub_p4_4b helper resolve via the MANDATORY PathFixture(worktree-root) +
%   PathFixture(this-folder) in TestClassSetup (WP9-F4 lesson).

    properties (Constant)
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    end

    methods (TestClassSetup)
        function addPaths(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so a COLD run cannot resolve the +mat2doc package without
            % the worktree root on the path (MATLAB:undefinedVarOrClass). here is
            % tests\text; the worktree root is two levels up. A second fixture puts
            % this folder on the path so the co-located PartStub_p4_4b helper resolves
            % regardless of cwd.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\text
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
            testCase.applyFixture(PathFixture(here));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. add_break map -- the byte pin (6 members + default + alias)   %
        % =============================================================== %

        function test_add_break_map_bytepin_vs_oracle(testCase)
            % Nominal + Edge + Regression (run.py 38-57, s0025 add_break): each mapped
            % member + no-arg default + the TEXT_WRAPPING alias serialized on a FRESH
            % run. serhex (byte pin, includes the <?xml ...?> prolog) + the
            % has_wtype/has_wclear substring flags compared to the frozen oracle. This
            % is the Gate-2-requested add_break byte pin: LINE_CLEAR_* carry ONLY
            % @w:clear (the @w:type="textWrapping" default is removed).
            WB = mat2doc.enum.text.WD_BREAK;
            o = loadOracle().add_break;
            cases = { ...
                "line",         {WB.LINE}; ...
                "page",         {WB.PAGE}; ...
                "column",       {WB.COLUMN}; ...
                "clear_left",   {WB.LINE_CLEAR_LEFT}; ...
                "clear_right",  {WB.LINE_CLEAR_RIGHT}; ...
                "clear_all",    {WB.LINE_CLEAR_ALL}; ...
                "default",      {}; ...                  % no arg -> LINE
                "text_wrapping",{WB.TEXT_WRAPPING}};     % alias of LINE_CLEAR_ALL
            for i = 1:size(cases, 1)
                key = cases{i, 1};
                od  = o.(key);
                [r, run] = newRun();
                run.add_break(cases{i, 2}{:});
                testCase.verifyEqual(hx_e(r), string(od.serhex), ...
                    sprintf('add_break %s serialized bytes (L1) vs oracle', key));
                s = ser(r);
                testCase.verifyEqual(ternary(contains(s, "w:type=")), string(od.has_wtype), ...
                    sprintf('add_break %s has @w:type flag vs oracle', key));
                testCase.verifyEqual(ternary(contains(s, "w:clear=")), string(od.has_wclear), ...
                    sprintf('add_break %s has @w:clear flag vs oracle', key));
            end
        end

        function test_add_break_hardcoded_xml_regression(testCase)
            % Regression (hard-coded L1, INDEPENDENT of the oracle): the exact <w:r>
            % serialization for each break variant. The load-bearing pin: LINE_CLEAR_*
            % (and the TEXT_WRAPPING alias) carry <w:br w:clear="..."/> with NO @w:type.
            WB = mat2doc.enum.text.WD_BREAK;
            expect = { ...
                WB.LINE,             "<w:br/>"; ...
                WB.PAGE,             "<w:br w:type=""page""/>"; ...
                WB.COLUMN,           "<w:br w:type=""column""/>"; ...
                WB.LINE_CLEAR_LEFT,  "<w:br w:clear=""left""/>"; ...
                WB.LINE_CLEAR_RIGHT, "<w:br w:clear=""right""/>"; ...
                WB.LINE_CLEAR_ALL,   "<w:br w:clear=""all""/>"; ...
                WB.TEXT_WRAPPING,    "<w:br w:clear=""all""/>"};
            for i = 1:size(expect, 1)
                [r, run] = newRun();
                run.add_break(expect{i, 1});
                testCase.verifyEqual(ser(r), rWrap(testCase, expect{i, 2}), ...
                    sprintf('add_break case %d -> exact <w:r> (hard-coded L1)', i));
            end
            % no-arg default -> LINE -> <w:br/>
            [r, run] = newRun();
            run.add_break();
            testCase.verifyEqual(ser(r), rWrap(testCase, "<w:br/>"), ...
                'add_break() no-arg default -> <w:br/> (hard-coded L1)');
        end

        function test_add_break_section_raises_keyerror(testCase)
            % Edge / error-path (run.py 45-52 KeyError, s0025 add_break_section): the 4
            % unmapped SECTION_* members reproduce Python {...}[break_type] KeyError as
            % mat2doc:KeyError -- verify the IDENTIFIER (not merely that it throws) AND
            % that children==0 (the error precedes add_br: no partial mutation).
            WB = mat2doc.enum.text.WD_BREAK;
            o  = loadOracle().add_break_section;
            secs = { ...
                "section_continuous", WB.SECTION_CONTINUOUS; ...
                "section_even_page",  WB.SECTION_EVEN_PAGE; ...
                "section_next_page",  WB.SECTION_NEXT_PAGE; ...
                "section_odd_page",   WB.SECTION_ODD_PAGE};
            for i = 1:size(secs, 1)
                key = secs{i, 1};
                [r, run] = newRun();
                ME = captureError(@() run.add_break(secs{i, 2}));
                testCase.verifyEqual(string(ME.identifier), "mat2doc:KeyError", ...
                    sprintf('add_break %s -> mat2doc:KeyError identifier', key));
                testCase.verifyEqual(nchild(r), 0, ...
                    sprintf('add_break %s leaves children=0 (no partial mutation)', key));
                testCase.verifyEqual(string(o.(key).raises), "True", ...
                    sprintf('oracle %s raises sanity', key));
                testCase.verifyEqual(double(o.(key).children), 0, ...
                    sprintf('oracle %s children==0 sanity', key));
            end
        end

        % =============================================================== %
        % 2. text get/set                                                 %
        % =============================================================== %

        function test_text_get_set_and_serhex(testCase)
            % Nominal + Edge + Regression (run.py 205-225, s0025 text): "a\tb\nc\rd"
            % -> the split w:t/w:tab/w:br structure (serhex byte pin) and get ->
            % "a\tb\nc\nd" (\r reads back as \n); empty set -> self-closing <w:r/> +
            % get ""; re-set replaces (no append).
            o = loadOracle().text;
            TAB = string(char(9)); LF = string(char(10)); CR = string(char(13));

            [r, run] = newRun();
            run.text = "a" + TAB + "b" + LF + "c" + CR + "d";
            testCase.verifyEqual(hx_e(r), string(o.set_serhex), 'text set serhex vs oracle');
            testCase.verifyEqual(ser(r), rWrap(testCase, ...
                "<w:t>a</w:t><w:tab/><w:t>b</w:t><w:br/><w:t>c</w:t><w:br/><w:t>d</w:t>"), ...
                'text set -> split w:t/w:tab/w:br (hard-coded L1)');
            testCase.verifyEqual(run.text, "a" + TAB + "b" + LF + "c" + LF + "d", ...
                'text get -> "a\tb\nc\nd" (\r reads back as \n, hard-coded)');
            testCase.verifyEqual(run.text, string(o.get), 'text get vs oracle');

            [r, run] = newRun();
            run.text = "";
            testCase.verifyEqual(hx_e(r), string(o.empty_serhex), 'empty text serhex vs oracle');
            testCase.verifyEqual(ser(r), decl() + newline + "<w:r xmlns:w=""" + testCase.W + """/>", ...
                'empty text -> self-closing <w:r/> (hard-coded L1)');
            testCase.verifyEqual(run.text, string(o.empty_get), 'empty text get -> "" vs oracle');
            testCase.verifyEqual(run.text, "", 'empty text get -> "" (hard-coded)');

            [r, run] = newRun();
            run.text = "first";
            run.text = "second";        % replaces, does not append
            testCase.verifyEqual(hx_e(r), string(o.reset_serhex), 'reset text serhex vs oracle');
            testCase.verifyEqual(run.text, string(o.reset_get), 'reset text get -> "second" vs oracle');
            testCase.verifyEqual(run.text, "second", 'reset replaces content (hard-coded)');
        end

        % =============================================================== %
        % 3. font delegation (bold/italic/underline forward to Font)      %
        % =============================================================== %

        function test_font_delegation(testCase)
            % Nominal + Edge + Regression (run.py 98-151, 227-249, s0025
            % font_delegation): bold+italic+underline on one run ->
            % <w:rPr><w:b/><w:i/><w:u w:val="single"/></w:rPr> (serhex byte pin) + gets
            % True/True/True; bold=[] removes w:b (get None); bold=false ->
            % <w:b w:val="0"/> (get False).
            o = loadOracle().font_delegation;

            [r, run] = newRun();
            run.bold = true;
            run.italic = true;
            run.underline = true;
            testCase.verifyEqual(hx_e(r), string(o.combined_serhex), 'combined font serhex vs oracle');
            testCase.verifyEqual(ser(r), rWrap(testCase, ...
                "<w:rPr><w:b/><w:i/><w:u w:val=""single""/></w:rPr>"), ...
                'bold+italic+underline -> <w:rPr><w:b/><w:i/><w:u w:val="single"/></w:rPr> (hard-coded L1)');
            testCase.verifyEqual(rv(run.bold), string(o.bold), 'bold get vs oracle');
            testCase.verifyEqual(rv(run.italic), string(o.italic), 'italic get vs oracle');
            testCase.verifyEqual(rv(run.underline), string(o.underline), 'underline get vs oracle');

            run.bold = [];
            testCase.verifyEqual(hx_e(r), string(o.bold_none_serhex), 'bold=None removes w:b serhex vs oracle');
            testCase.verifyEqual(rv(run.bold), string(o.bold_none_get), 'bold=None get -> None vs oracle');
            testCase.verifyTrue(isequal(run.bold, []), 'bold=None get -> [] (None)');

            [r, run] = newRun();
            run.bold = false;
            testCase.verifyEqual(hx_e(r), string(o.bold_false_serhex), 'bold=False serhex vs oracle');
            testCase.verifyEqual(rv(run.bold), string(o.bold_false_get), 'bold=False get vs oracle');
        end

        % =============================================================== %
        % 4. contains_page_break (H4 -- hard break NOT counted)           %
        % =============================================================== %

        function test_contains_page_break_h4(testCase)
            % Nominal + Edge (run.py 120-131, s0025 contains_page_break, H4): empty ->
            % False; parsed <w:lastRenderedPageBreak/> -> True; a HARD
            % <w:br w:type="page"/> -> False (an author page-break is NOT counted).
            o = loadOracle().contains_page_break;

            [~, run] = newRun();
            testCase.verifyFalse(run.contains_page_break, 'empty run -> contains_page_break False');
            testCase.verifyEqual(rv(run.contains_page_break), string(o.empty), 'empty vs oracle');

            run = mat2doc.text.Run(parseR("<w:lastRenderedPageBreak/>"), []);
            testCase.verifyTrue(run.contains_page_break, 'w:lastRenderedPageBreak -> True');
            testCase.verifyEqual(rv(run.contains_page_break), string(o.with_lrpb), 'with_lrpb vs oracle');

            run = mat2doc.text.Run(parseR("<w:br w:type=""page""/>"), []);
            testCase.verifyFalse(run.contains_page_break, 'hard <w:br w:type="page"/> NOT counted (H4) -> False');
            testCase.verifyEqual(rv(run.contains_page_break), string(o.hard_break), 'hard_break vs oracle');
        end

        % =============================================================== %
        % 5. content ops -- add_text / add_tab / clear                    %
        % =============================================================== %

        function test_content_add_text_add_tab_clear(testCase)
            % Nominal + Edge + Regression (run.py 83-118, s0025 content): add_text(" hi ")
            % -> <w:t xml:space="preserve"> hi </w:t> returning a Text_ proxy;
            % add_text("plain") -> no xml:space; add_tab() -> <w:tab/>; clear() after
            % bold+add_text keeps <w:rPr><w:b/></w:rPr> only and RETURNS SELF.
            o = loadOracle().content;

            [r, run] = newRun();
            t = run.add_text(" hi ");
            testCase.verifyEqual(hx_e(r), string(o.add_text_hi_serhex), 'add_text " hi " serhex vs oracle');
            testCase.verifyEqual(ser(r), rWrap(testCase, "<w:t xml:space=""preserve""> hi </w:t>"), ...
                'add_text " hi " -> xml:space="preserve" (hard-coded L1)');
            testCase.verifyClass(t, 'mat2doc.text.Text_');
            testCase.verifyEqual(ternary(isa(t, "mat2doc.text.Text_")), string(o.add_text_is_proxy), ...
                'add_text returns a Text_ proxy vs oracle');

            [r, run] = newRun();
            run.add_text("plain");
            testCase.verifyEqual(hx_e(r), string(o.add_text_plain_serhex), 'add_text "plain" serhex vs oracle');
            testCase.verifyEqual(ser(r), rWrap(testCase, "<w:t>plain</w:t>"), ...
                'add_text "plain" -> no xml:space (hard-coded L1)');

            [r, run] = newRun();
            run.add_tab();
            testCase.verifyEqual(hx_e(r), string(o.add_tab_serhex), 'add_tab serhex vs oracle');
            testCase.verifyEqual(ser(r), rWrap(testCase, "<w:tab/>"), ...
                'add_tab -> <w:tab/> (hard-coded L1)');

            [r, run] = newRun();
            run.bold = true;
            run.add_text("x");
            ret = run.clear();
            testCase.verifyEqual(hx_e(r), string(o.clear_serhex), 'clear serhex vs oracle');
            testCase.verifyEqual(ser(r), rWrap(testCase, "<w:rPr><w:b/></w:rPr>"), ...
                'clear keeps <w:rPr><w:b/></w:rPr> only, drops content (hard-coded L1)');
            testCase.verifyTrue(ret == run, 'clear returns self (handle ==)');
            testCase.verifyEqual(ternary(ret == run), string(o.clear_returns_self), ...
                'clear returns self vs oracle');
        end

        % =============================================================== %
        % 6. mark_comment_range                                           %
        % =============================================================== %

        function test_mark_comment_range(testCase)
            % Nominal + Regression (run.py 176-186, s0025 content.mark_*): over two runs
            % in a w:p, mark_comment_range(run2, 3) inserts commentRangeStart before the
            % first run and commentRangeEnd + a reference run after the last -> child
            % localnames [commentRangeStart, r, r, commentRangeEnd, r], reference-run
            % style "CommentReference", full-parent serhex byte-identical (w:id="3").
            o = loadOracle().content;

            p  = mat2doc.oxml.OxmlElement("w:p");
            r1 = mat2doc.oxml.OxmlElement("w:r");
            r2 = mat2doc.oxml.OxmlElement("w:r");
            p.append(r1); p.append(r2);
            run1 = mat2doc.text.Run(r1, []);
            run2 = mat2doc.text.Run(r2, []);
            run1.mark_comment_range(run2, 3);

            names  = lns(p);
            oracle = cellstr(string(o.mark_localnames(:))');
            testCase.verifyEqual(names, oracle, 'mark_comment_range child localnames vs oracle');
            testCase.verifyEqual(names, {'commentRangeStart','r','r','commentRangeEnd','r'}, ...
                'mark_comment_range localnames (hard-coded)');

            kids = p.xpath("./*");
            testCase.verifyEqual(string(kids(end).style), string(o.mark_ref_style), ...
                'reference-run style vs oracle');
            testCase.verifyEqual(string(kids(end).style), "CommentReference", ...
                'reference-run style "CommentReference" (hard-coded)');
            testCase.verifyEqual(hx_e(p), string(o.mark_serhex), 'mark_comment_range full-parent serhex vs oracle');
        end

        % =============================================================== %
        % 7. _Text inert surface                                          %
        % =============================================================== %

        function test_text_surface_inert(testCase)
            % Edge (run.py 252-257, s0025 text_surface): _Text/Text_ is the faithful
            % port of a bare v1.2.0 object subclass -- NO public members, no text
            % accessor. isprop(t,"text") False, 0 public members.
            o = loadOracle().text_surface;

            [~, run] = newRun();
            t = run.add_text("z");
            testCase.verifyFalse(isprop(t, "text"), '_Text has no text property');
            testCase.verifyEqual(ternary(isprop(t, "text")), string(o.has_text_attr), ...
                'has_text_attr vs oracle');
            testCase.verifyEqual(numel(properties(t)), 0, '_Text exposes 0 public members');
            testCase.verifyEqual(numel(properties(t)), double(o.public_count), ...
                'public member count vs oracle');
        end

        % =============================================================== %
        % 8. H5 identity + StoryChild .part delegation                    %
        % =============================================================== %

        function test_h5_identity_and_part_delegation(testCase)
            % Regression (run.py 133-137 fresh Font; shared.py StoryChild.part, s0025
            % h5): run.font == run.font True (fresh proxy each access, SAME w:r -> element
            % identity ==); Run(r, partStub).part() delegates via StoryChild to
            % parent.part() -> the "STUBPART" sentinel.
            o = loadOracle().h5;

            [~, run] = newRun();
            testCase.verifyTrue(run.font == run.font, 'run.font == run.font (fresh proxy, same run) True');
            testCase.verifyEqual(rv(run.font == run.font), string(o.font_eq_font), 'font_eq_font vs oracle');

            run2 = mat2doc.text.Run(mat2doc.oxml.OxmlElement("w:r"), PartStub_p4_4b());
            testCase.verifyEqual(string(run2.part()), "STUBPART", ...
                'Run.part delegates via StoryChild to parent.part() (hard-coded)');
            testCase.verifyEqual(string(run2.part()), string(o.part_delegation), 'part_delegation vs oracle');
        end

        % =============================================================== %
        % 9. Stub-safety (2 stubs) + Run.style RESOLVES (P4-7a un-stub)    %
        % =============================================================== %

        function test_stub_safety(testCase)
            % Edge / stub-safety (run.py 59-81, 153-174): the remaining
            % unported-dependency stubs each raise mat2doc:notYetPorted (identifier
            % verified). iter_inner_content -> P4-5b+P7 (still stubbed).
            %
            % REGISTRY-FLIP RE-PIN (P4-7a Gate-4): Run.style GET and SET were
            % un-stubbed at P4-7a (they delegate to the now-live
            % DocumentPart.get_style / get_style_id) and RESOLVE over a run with a
            % real Document part -- re-pinned to the resolved behavior below (the
            % registry-flip stale-pins lesson). The old notYetPorted assertion is
            % gone.
            %
            [~, run] = newRun();
            testCase.verifyEqual(string(captureError(@() run.iter_inner_content()).identifier), ...
                "mat2doc:notYetPorted", 'iter_inner_content -> mat2doc:notYetPorted');

            % --- Run.style get/set now RESOLVE end-to-end over a live Document ---
            % Capture the CT_P so the CT_R rStyle can be reached via xpath (Run
            % exposes no element() accessor; it is a StoryChild).
            d = mat2doc.Document();
            p = d.element().body.add_p();
            para = mat2doc.text.Paragraph(p, d);
            lrun = para.add_run("hi");

            % REGISTRY-FLIP RE-PIN (P7-4, the picture milestone WP; validate_P7-4 s6
            % re-pin 4): Run.add_picture is now LIVE -> it reaches the image loader
            % (part().new_pic_inline -> get_or_add_image -> Image.from_file) instead of
            % the notYetPorted stub. Exercised over `lrun` (a run with a LIVE Document
            % part, so part() resolves before the loader). On a BOGUS path it fails
            % inside Image.from_file with the faithful mat2doc:FileNotFoundError
            % (image.py 35-50), NOT mat2doc:notYetPorted -- the identifier change IS the
            % un-stub proof. The byte-exact InlineShape return over a real image is
            % pinned by tests\parts\Test_p7_4_add_picture.m (a real image is out of this
            % class's data scope).
            testCase.verifyEqual(string(captureError(@() lrun.add_picture("no_such_image.png")).identifier), ...
                "mat2doc:FileNotFoundError", ...
                'add_picture is LIVE (P7-4): a bogus path -> FileNotFoundError, NOT notYetPorted');
            % GET (default): a CharacterStyle (the document default character style)
            testCase.verifyClass(lrun.style, 'mat2doc.styles.CharacterStyle', ...
                'Run.style GET RESOLVES to a CharacterStyle (P4-7a un-stub)');
            % SET by name -> writes the run's rStyle; GET round-trips
            lrun.style = "Emphasis";
            rr = p.xpath('.//w:r');
            testCase.verifyEqual(string(rr(end).style), "Emphasis", ...
                'Run.style SET writes w:rPr/w:rStyle w:val="Emphasis" on the CT_R (byte-level)');
            testCase.verifyEqual(lrun.style.style_id, "Emphasis", ...
                'Run.style GET round-trips the applied character style');
        end

        % =============================================================== %
        % 10. M1 sanity (light) -- StoryChild lineage, API/proxy tier      %
        % =============================================================== %

        function test_run_is_storychild_spot(testCase)
            % M1 sanity (light): P4-4b is API/proxy tier -- byte-neutral, no styles.xml/
            % document.xml pin needed here (the full 17/17 M1 sweep stays owned by
            % Test_p1_8_skeleton_m1; validate_P4-4b §2 confirms M1 17/17 unchanged).
            % Spot-check that Run IS a StoryChild (so H5 identity + .part delegation come
            % from the shared base) and that a live member round-trips.
            r = mat2doc.oxml.OxmlElement("w:r");
            run = mat2doc.text.Run(r, []);
            testCase.verifyClass(run, 'mat2doc.text.Run');
            testCase.verifyTrue(isa(run, 'mat2doc.shared.StoryChild'), ...
                'Run must be a mat2doc.shared.StoryChild');
            run.bold = true;                       % spot round-trip via font delegation
            testCase.verifyTrue(run.bold, 'Run.bold spot round-trip');
            % _Text is a plain handle proxy (not a StoryChild)
            t = run.add_text("q");
            testCase.verifyClass(t, 'mat2doc.text.Text_');
            testCase.verifyTrue(isa(t, 'handle'), 'Text_ must be a handle');
        end

        % =============================================================== %
        % 11. EQUIVALENCE -- full s0025 battery vs the frozen oracle       %
        % =============================================================== %

        function test_equivalence_full_battery_vs_frozen_oracle(testCase)
            % Equivalence: replay the ENTIRE s0025 battery (runProbes -- the .m twin's
            % diffed-P body VERBATIM: add_break / add_break_section / text /
            % font_delegation / contains_page_break / content / text_surface / h5) and
            % flatten-compare EVERY leaf to the frozen python-docx 1.2.0 oracle copied
            % into data\s0025_probe_oracle.json. Gate-3 found ZERO divergences
            % (probe_diff MATCH 66/66), so every leaf must be byte/value-identical.
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
            testCase.verifyGreaterThan(numel(oKeys), 50, ...
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

function [r, run] = newRun()
    r = mat2doc.oxml.OxmlElement("w:r");
    run = mat2doc.text.Run(r, []);
end

function e = parseR(inner)
    % Parse a <w:r> carrying its own xmlns:w with the given inner XML (mirrors the
    % s0025 twin's parseR()).
    xml = "<w:r xmlns:w=""" + W_() + """>" + string(inner) + "</w:r>";
    e = mat2doc.oxml.parse_xml(uint8(unicode2native(char(xml), "UTF-8")));
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

function s = rWrap(testCase, body)
    % decl + newline + <w:r xmlns:w="...">BODY</w:r>. The run is created via
    % OxmlElement("w:r") -> declares xmlns:w. Used by the hard-coded XML regression pins.
    s = decl() + newline + "<w:r xmlns:w=""" + testCase.W + """>" + string(body) + "</w:r>";
end

function n = nchild(e)
    n = numel(e.xpath("./*"));
end

function C = lns(e)
    % local_part of each child element, as a 1xN cellstr.
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

function s = ternary(tf)
    if tf, s = "True"; else, s = "False"; end
end

function s = rv(x)
    % Uniform accessor repr mirroring the s0025 rv(): None->"None",
    % bool->"True"/"False", int->decimal, string->itself.
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

function setStyle(run)
    % Drive the style SETTER (a separate stub from the getter).
    run.style = "Emphasis";
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
    % .gitattributes pin is needed (value/serhex fixture, s0024 precedent).
    here = fileparts(mfilename('fullpath'));
    p = fullfile(here, 'data', 's0025_probe_oracle.json');
    fid = fopen(p, 'r', 'n');
    if fid < 0
        error('mat2doc:test:openFail', 'cannot open frozen oracle %s', p);
    end
    raw = fread(fid, Inf, '*uint8')';
    fclose(fid);
    o = jsondecode(native2unicode(raw, 'UTF-8'));
end

function P = runProbes()
    % Replay the s0025 diffed-P probe sequence (the .m twin's P-struct body, VERBATIM
    % tags/inputs/order) and return the nested struct of tagged canonical values.
    % Embedded here so the Equivalence leg is self-contained (the validation-folder
    % scenario is NOT on the toolbox path and must not be a dependency). Mirrors
    % validation\mat2doc\scenarios\s0025_p4_4b_text_run.m (stub-safety + M1 are NOT
    % part of the diffed P, matching the twin).
    P = struct();
    WB = mat2doc.enum.text.WD_BREAK;

    % ===================== add_break: 6 mapped + default + alias ===========
    ab = struct();
    ab.line          = doBreak(WB.LINE);
    ab.page          = doBreak(WB.PAGE);
    ab.column        = doBreak(WB.COLUMN);
    ab.clear_left    = doBreak(WB.LINE_CLEAR_LEFT);
    ab.clear_right   = doBreak(WB.LINE_CLEAR_RIGHT);
    ab.clear_all     = doBreak(WB.LINE_CLEAR_ALL);
    ab.default       = doBreak();                  % no arg -> LINE
    ab.text_wrapping = doBreak(WB.TEXT_WRAPPING);   % alias of LINE_CLEAR_ALL
    P.add_break = ab;

    % ===================== add_break SECTION_* -> KeyError =================
    SEC = { ...
        "section_continuous", WB.SECTION_CONTINUOUS; ...
        "section_even_page",  WB.SECTION_EVEN_PAGE; ...
        "section_next_page",  WB.SECTION_NEXT_PAGE; ...
        "section_odd_page",   WB.SECTION_ODD_PAGE};
    sk = struct();
    for k = 1:size(SEC, 1)
        [r, run] = newRun();
        raised = "False";
        try
            run.add_break(SEC{k, 2});
        catch ME
            if strcmp(ME.identifier, "mat2doc:KeyError")
                raised = "True";
            else
                rethrow(ME);
            end
        end
        sk.(SEC{k, 1}) = struct("raises", raised, "children", nchild(r));
    end
    P.add_break_section = sk;

    % ===================== text get/set ====================================
    TAB = string(char(9)); LF = string(char(10)); CR = string(char(13));
    tx = struct();
    [r, run] = newRun();
    run.text = "a" + TAB + "b" + LF + "c" + CR + "d";
    tx.set_serhex = hx_e(r); tx.get = run.text;             % get -> "a\tb\nc\nd"
    [r, run] = newRun();
    run.text = "";
    tx.empty_serhex = hx_e(r); tx.empty_get = run.text;
    [r, run] = newRun();
    run.text = "first";
    run.text = "second";                                    % replaces (no append)
    tx.reset_serhex = hx_e(r); tx.reset_get = run.text;
    P.text = tx;

    % ===================== font delegation =================================
    fn = struct();
    [r, run] = newRun();
    run.bold = true;
    run.italic = true;
    run.underline = true;
    fn.combined_serhex = hx_e(r);
    fn.bold = rv(run.bold); fn.italic = rv(run.italic); fn.underline = rv(run.underline);
    run.bold = [];
    fn.bold_none_serhex = hx_e(r); fn.bold_none_get = rv(run.bold);
    [r, run] = newRun();
    run.bold = false;
    fn.bold_false_serhex = hx_e(r); fn.bold_false_get = rv(run.bold);
    P.font_delegation = fn;

    % ===================== contains_page_break (H4) ========================
    cpb = struct();
    [~, run] = newRun();
    cpb.empty = rv(run.contains_page_break);                     % False
    run = mat2doc.text.Run(parseR("<w:lastRenderedPageBreak/>"), []);
    cpb.with_lrpb = rv(run.contains_page_break);                 % True
    run = mat2doc.text.Run(parseR("<w:br w:type=""page""/>"), []);
    cpb.hard_break = rv(run.contains_page_break);                % False (not counted)
    P.contains_page_break = cpb;

    % ===================== content ops =====================================
    co = struct();
    [r, run] = newRun();
    t = run.add_text(" hi ");
    co.add_text_hi_serhex = hx_e(r);                            % xml:space="preserve"
    co.add_text_is_proxy = ternary(isa(t, "mat2doc.text.Text_"));
    [r, run] = newRun();
    run.add_text("plain");
    co.add_text_plain_serhex = hx_e(r);                         % no xml:space
    [r, run] = newRun();
    run.add_tab();
    co.add_tab_serhex = hx_e(r);                                % <w:tab/>
    [r, run] = newRun();
    run.bold = true;
    run.add_text("x");
    ret = run.clear();
    co.clear_serhex = hx_e(r);                                  % <w:rPr><w:b/></w:rPr> only
    co.clear_returns_self = ternary(ret == run);
    % mark_comment_range over two runs in a w:p
    p = mat2doc.oxml.OxmlElement("w:p");
    r1 = mat2doc.oxml.OxmlElement("w:r");
    r2 = mat2doc.oxml.OxmlElement("w:r");
    p.append(r1); p.append(r2);
    run1 = mat2doc.text.Run(r1, []);
    run2 = mat2doc.text.Run(r2, []);
    run1.mark_comment_range(run2, 3);
    co.mark_localnames = lns(p);
    kids = p.xpath("./*");
    co.mark_ref_style = kids(end).style;                        % "CommentReference"
    co.mark_serhex = hx_e(p);
    P.content = co;

    % ===================== Text_ surface ===================================
    ts = struct();
    [~, run] = newRun();
    t = run.add_text("z");
    ts.has_text_attr = ternary(isprop(t, "text"));              % False
    ts.public_count = numel(properties(t));                     % 0
    P.text_surface = ts;

    % ===================== h5 identity + part delegation ===================
    h5 = struct();
    [~, run] = newRun();
    h5.font_eq_font = rv(run.font == run.font);                 % True (same element)
    run2 = mat2doc.text.Run(mat2doc.oxml.OxmlElement("w:r"), PartStub_p4_4b());
    h5.part_delegation = string(run2.part());                   % "STUBPART" via StoryChild
    P.h5 = h5;
end

function e = doBreak(varargin)
    % One add_break case: emit serhex + the has_wtype/has_wclear substring flags,
    % matching the s0025 twin's doBreak().
    [r, run] = newRun();
    run.add_break(varargin{:});
    raw = mat2doc.oxml.serialize_part_xml(r);
    s = string(native2unicode(raw, "UTF-8"));
    e = struct("serhex", hx(raw), ...
               "has_wtype", ternary(contains(s, "w:type=")), ...
               "has_wclear", ternary(contains(s, "w:clear=")));
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

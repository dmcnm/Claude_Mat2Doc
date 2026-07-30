classdef Test_p4_5b_paragraph_api < matlab.unittest.TestCase
% TEST_P4_5B_PARAGRAPH_API  Gate-4 permanent unit tests for Mat2Doc P4-5b
%   (src/docx/text/paragraph.py -> +mat2doc\+text\Paragraph, hyperlink.py ->
%   Hyperlink, pagebreak.py -> RenderedPageBreak -- the LAST API/proxy-tier WP of
%   Phase 4).
%
%   Paragraph / Hyperlink / RenderedPageBreak are pure API/proxy classes over the
%   already-byte-validated CT_P (P4-2) + CT_Hyperlink / CT_LastRenderedPageBreak
%   (P4-3). None adds a register_element_cls row, an oxml class, or a
%   serialization-path change -> API/proxy tier: no bytes-of-the-registry, no M1
%   risk (validate_P4-5b section 3 confirms M1 17/17 unchanged, so this class does
%   NOT re-pin the default-template document.xml -- Test_p1_8_skeleton_m1 owns the
%   full 17/17 M1 sweep). This class permanently FREEZES the P4-5b BEHAVIORAL +
%   serialized-bytes surface -- the API-tier equivalence pin -- byte/value-identical
%   to python-docx 1.2.0, PLUS the G-scenario document byte pin (three paragraphs
%   built THROUGH the NEW Paragraph API into a real word/document.xml).
%
%   The guarantees frozen here (each verified byte/value-identical at Gate-3:
%   probe_diff MATCH 170/170 leaves + G-scenario 17/17 parts byte-identical incl.
%   word/document.xml + M1 17/17; zero divergences, zero new D-numbers):
%
%     (add_run -- serialized-bytes parity) empty run (text None); "Hello";
%       "a\tb" -> <w:tab/>; "x\ny\rz" -> two <w:br/> (text getter maps to
%       "x\ny\nz"); non-ASCII "café ☕ 日本語" UTF-8 round-trip; two runs (the
%       second add_run gets xml:space="preserve" on "Hello "). char-style path
%       (add_run(text, style)) reaches Run.style, the P4-7 STUB -> raises
%       mat2doc:notYetPorted (identifier pinned).
%
%     (alignment) CENTER get/set round-trip -> <w:jc w:val="center"/>; []-reset
%       removes w:jc (H3 tri-state, get -> [] None).
%
%     (clear) returns the SAME handle (H5) and preserves the w:pPr while dropping
%       all content; text getter -> "".
%
%     (contains_page_break) false for w:r / w:r+w:t; true for a lastRenderedPageBreak
%       inside a hyperlink and for two lastRenderedPageBreaks (H4 bool(node-list)).
%
%     (runs / hyperlinks / rendered_page_breaks) homogeneous 1xN object arrays
%       (Run / Hyperlink / RenderedPageBreak); H1 first/last; EMPTY -> a TYPED 1x0
%       array (Class.empty(1,0), NOT []); the interleaved tree yields runs
%       [R1,R2,R3] and hyperlinks [A1,A2].
%
%     (iter_inner_content) a 1xN CELL of Run|Hyperlink in EXACT document order
%       (the interleaved tree -> {Run R1, Hyperlink A1, Run R2, Hyperlink A2,
%       Run R3}); empty -> a 1x0 cell.
%
%     (text) get concatenates inner-content text (tabs -> \t, breaks -> \n); set
%       replaces all content with a SINGLE run holding the text (pPr preserved,
%       "new\ttext" -> "new"<w:tab/>"text"); "" -> a single empty run.
%
%     (insert_paragraph_before) inserts a new paragraph BEFORE this one (body
%       child order [new, orig]); with text -> a single run; empty -> empty; the
%       rootless-<w:p/> case raises mat2doc:TypeError at add_p_before (identifier +
%       verbatim message pinned -- identical on both sides, Gate-3 section 6).
%
%     (Paragraph.style) get AND set both raise mat2doc:notYetPorted -- faithful
%       P4-7 stub delegation through part().get_style / get_style_id (identifier
%       pinned; NOT a stand-in).
%
%     (Hyperlink) address: external r:id -> target_ref resolved LIVE via a real
%       relationship ("https://google.com/"), internal jump (w:anchor only) -> "";
%       fragment ±; url = address / address#fragment / "" (no address); runs;
%       text; contains_page_break T/F.
%
%     (RenderedPageBreak) precedes/follows_all_content guards -> [] (None);
%       preceding/following fragment split incl. the ATOMIC break-inside-hyperlink
%       case (preceding keeps the WHOLE w:hyperlink, following DROPS it); the
%       second-(non-first)-break preceding_paragraph_fragment raises
%       mat2doc:ValueError with the verbatim "only defined on first rendered
%       page-break in paragraph".
%
%     (G-scenario document byte pin -- the M2 de-risker) a fresh mat2doc.Document()
%       whose three paragraphs are built THROUGH the NEW Paragraph API
%       (add_run/text=/alignment=) and .save()d -> the extracted word/document.xml
%       is byte-identical (SHA-256 + size + full-bytes) to the frozen s0029
%       python-docx reference. Closest guard yet to M2's add_paragraph path.
%
%   Provenance (Gate-1..3, all 2026-07-30):
%     * Audit    : validation\mat2doc\audit_P4-5b_paragraph_api.md (Porter Gate-1
%                  58-check self-probe + Fable/mso-auditor Gate-2 76-check
%                  adversarial APPROVE; s7 ns-decl reopen-check CONCUR unreachable;
%                  surface-shape decisions ACCEPT; zero defects).
%     * Validate : validation\mat2doc\validate_P4-5b_paragraph_api.md (Gate-3 PASS
%                  -- probe_diff MATCH 170/170; G-scenario 17/17 parts byte-identical
%                  incl. word/document.xml built through the NEW Paragraph API; M1
%                  17/17; regression 57/57; ZERO divergences, ZERO new D-numbers).
%     * Scenarios: validation\mat2doc\scenarios\s0028_p4_5b_paragraph_probe.{py,m}
%                  (the full-surface behavioral + serhex probe replayed VERBATIM by
%                  runProbes() below) and s0029_p4_5b_paragraph_gscenario.{py,m}
%                  (the three-paragraph NEW-Paragraph-API document byte pin replayed
%                  by buildGScenario() below).
%     * Frozen refs (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE):
%         references\s0028\probe.json -- copied verbatim (self-contained) into
%           tests\text\data\s0028_probe_oracle.json (value/serhex JSON; jsondecode
%           is line-ending agnostic -> no `* binary` pin, s0024/s0026 precedent).
%         references\s0029\parts\word\document.xml (SHA-256
%           8ea393a4ccc8f2835d6c74c26f58a3af56200efaabf25cf6cc4c49f9b197e25e,
%           1814 B) -- copied byte-for-byte into tests\text\data\s0029_document.xml
%           (co-located `* binary` .gitattributes) as the G-scenario byte fixture,
%           AND its SHA/size embedded below as SHA_S0029_DOCUMENT / SIZE_.
%
%   Coverage taxonomy
%   -----------------
%   * Nominal      -- add_run "Hello"/multi; alignment CENTER; clear;
%                     contains_page_break true; runs/hyperlinks lists; iter order;
%                     text get/set; insert_paragraph_before with text; the full
%                     Hyperlink external surface; the RenderedPageBreak mid split;
%                     the G-scenario.
%   * Edge         -- empty run / empty text / empty collections (typed 1x0, NOT
%                     []); non-ASCII (café ☕ 日本語); tab/newline mapping;
%                     alignment []-reset; single-element lists; internal jump
%                     (address ""); leading/trailing fragment None; the atomic
%                     break-inside-hyperlink split.
%   * Error path   -- add_run char-style + Paragraph.style get/set raise
%                     mat2doc:notYetPorted (P4-7 stub); rootless
%                     insert_paragraph_before raises mat2doc:TypeError; the
%                     second-break fragment raises mat2doc:ValueError. Each verifies
%                     the IDENTIFIER (a mat2ppt:<PyExceptionName>-class id, not
%                     merely that it throws) AND the verbatim message.
%   * Equivalence  -- test_equivalence_full_battery_vs_frozen_oracle replays the
%                     ENTIRE s0028 battery live (runProbes, the .m twin's body
%                     verbatim) and flatten-compares every leaf to the frozen
%                     python-docx 1.2.0 oracle (Gate-3 found ZERO divergences).
%   * Regression   -- hard-coded expected serialized-XML strings + UPPERCASE serhex
%                     of the raw UTF-8 shipping bytes vs the frozen oracle; the
%                     G-scenario document.xml SHA-256/size/full-bytes pin.
%   * Upstream     -- the add_run \t/\n mapping, the interleaved runs/hyperlinks/
%                     iter_inner_content document order, the atomic
%                     break-inside-hyperlink fragment split and the verbatim
%                     ValueError message ARE the python-docx paragraph.py /
%                     hyperlink.py / pagebreak.py contract; the frozen oracle IS
%                     lxml's expected output for this API sequence.
%
%   Byte-level (L1) note: every serialized-XML assertion is either the FULL
%   serialize_part_xml output as a UTF-8-decoded string (string-equality == byte
%   equality L1) or its UPPERCASE hex (serhex) vs the frozen oracle, and the
%   G-scenario is a SHA-256 + whole-bytes pin. No D-number granted any L2
%   relaxation in this WP (Gate-3: zero new; D-serializer-nsdecl non-engaged /
%   unreachable -- the s7 ns0 scan found 0 hits; D10 consumed read-only), so every
%   pin is L1.
%
%   Determinism: no network, no absolute paths. The co-located oracle + byte
%   fixture resolve relative to this file via fileparts(mfilename('fullpath'));
%   every file read is binary ('r','n'). The +mat2doc package resolves via the
%   MANDATORY PathFixture(worktree-root) in TestClassSetup (WP9-F4 lesson).

    properties (Constant)
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        % The frozen s0029 word/document.xml pin (python-docx 1.2.0, frozen
        % 2026-07-30). SHA-256 equality == byte-identity (L1). Embedded so the
        % G-scenario pin is self-contained even without the byte fixture.
        SHA_S0029_DOCUMENT = "8ea393a4ccc8f2835d6c74c26f58a3af56200efaabf25cf6cc4c49f9b197e25e"
        SIZE_S0029_DOCUMENT = 1814
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied from tests\text\Test_p4_5a_parfmt_api.m. here is
            % tests\text; the worktree root is two levels up.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\text
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. add_run -- serialized-bytes parity + text mapping            %
        % =============================================================== %

        function test_add_run_text_tab_newline_mapping(testCase)
            % Nominal + Edge + Regression (paragraph.py 30-44, s0028 add_run):
            % empty run / "Hello" / "a\tb" (-> <w:tab/>) / "x\ny\rz" (-> two
            % <w:br/>, text getter maps to "x\ny\nz") / two runs (xml:space on
            % "Hello "). Each pins the whole w:p serhex vs oracle + a hard-coded
            % readable <w:p> (L1) + the text getter + run count.
            o = loadOracle().add_run;
            d = mat2doc.Document();

            % --- empty run (text None) ---
            [p, para] = newParsed(testCase, "", d);
            para.add_run();
            testCase.verifyEqual(hx_e(p), string(o.empty.serhex), 'add_run() empty serhex vs oracle');
            testCase.verifyEqual(ser(p), pWrap(testCase, "<w:r/>"), ...
                'add_run() -> empty <w:r/> (hard-coded L1)');
            testCase.verifyEqual(string(para.text), "", 'add_run() text ""');
            testCase.verifyEqual(numel(para.runs), 1, 'add_run() -> 1 run');

            % --- "Hello" ---
            [p, para] = newParsed(testCase, "", d);
            para.add_run("Hello");
            testCase.verifyEqual(hx_e(p), string(o.plain.serhex), 'add_run("Hello") serhex vs oracle');
            testCase.verifyEqual(ser(p), pWrap(testCase, "<w:r><w:t>Hello</w:t></w:r>"), ...
                'add_run("Hello") -> <w:r><w:t>Hello</w:t></w:r> (hard-coded L1)');
            testCase.verifyEqual(string(para.text), "Hello", 'add_run("Hello") text');
            testCase.verifyEqual(numel(para.runs), 1, 'add_run("Hello") -> 1 run');

            % --- tab: "a\tb" -> <w:tab/> ---
            [p, para] = newParsed(testCase, "", d);
            para.add_run("a" + sprintf('\t') + "b");
            testCase.verifyEqual(hx_e(p), string(o.tab.serhex), 'add_run tab serhex vs oracle');
            testCase.verifyEqual(ser(p), pWrap(testCase, ...
                "<w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r>"), ...
                'add_run("a\tb") -> <w:tab/> between text (hard-coded L1)');
            testCase.verifyEqual(string(para.text), "a" + sprintf('\t') + "b", 'add_run tab text round-trip');

            % --- newline + CR: "x\ny\rz" -> two <w:br/>; getter maps \r->\n ---
            [p, para] = newParsed(testCase, "", d);
            para.add_run("x" + sprintf('\n') + "y" + sprintf('\r') + "z");
            testCase.verifyEqual(hx_e(p), string(o.newline_cr.serhex), 'add_run newline/cr serhex vs oracle');
            testCase.verifyEqual(ser(p), pWrap(testCase, ...
                "<w:r><w:t>x</w:t><w:br/><w:t>y</w:t><w:br/><w:t>z</w:t></w:r>"), ...
                'add_run("x\ny\rz") -> two <w:br/> (hard-coded L1)');
            testCase.verifyEqual(string(para.text), "x" + sprintf('\n') + "y" + sprintf('\n') + "z", ...
                'text getter maps \n and \r both to \n');

            % --- two runs: second gets xml:space="preserve" on "Hello " ---
            [p, para] = newParsed(testCase, "", d);
            para.add_run("Hello ");
            para.add_run("World");
            testCase.verifyEqual(hx_e(p), string(o.multi.serhex), 'add_run multi serhex vs oracle');
            testCase.verifyEqual(ser(p), pWrap(testCase, ...
                "<w:r><w:t xml:space=""preserve"">Hello </w:t></w:r>" + ...
                "<w:r><w:t>World</w:t></w:r>"), ...
                'two add_run -> xml:space="preserve" on "Hello " (hard-coded L1)');
            testCase.verifyEqual(string(para.text), "Hello World", 'multi text');
            testCase.verifyEqual(numel(para.runs), 2, 'multi -> 2 runs');
        end

        function test_add_run_nonascii_roundtrip(testCase)
            % Edge (non-ASCII, paragraph.py 30-44, s0028 add_run.nonascii): a
            % run holding "café ☕ 日本語" serializes as UTF-8 and round-trips
            % through the text getter byte-identical to the oracle.
            o = loadOracle().add_run.nonascii;
            NONASCII = string(char([99 97 102 233 32 9749 32 26085 26412 35486])); % café ☕ 日本語
            d = mat2doc.Document();
            [p, para] = newParsed(testCase, "", d);
            para.add_run(NONASCII);
            testCase.verifyEqual(hx_e(p), string(o.serhex), 'non-ASCII run serhex vs oracle (UTF-8 L1)');
            testCase.verifyEqual(string(para.text), NONASCII, 'non-ASCII text round-trip');
        end

        function test_add_run_charstyle_notYetPorted(testCase)
            % Error path (paragraph.py 42-43, H4 `if style:`): a truthy char-style
            % reaches Run.style, the P4-7 STUB -> mat2doc:notYetPorted. Pin the
            % IDENTIFIER (a mat2ppt:<PyExceptionName>-class id), not merely a throw.
            d = mat2doc.Document();
            [~, para] = newParsed(testCase, "", d);
            ME = captureError(@() para.add_run("t", "Emphasis"));
            testCase.verifyEqual(string(ME.identifier), "mat2doc:notYetPorted", ...
                'add_run char-style -> mat2doc:notYetPorted identifier (P4-7 stub)');
        end

        % =============================================================== %
        % 2. alignment (tri-state) + clear                                %
        % =============================================================== %

        function test_alignment_get_set_none(testCase)
            % Nominal + Edge (paragraph.py 46-59, H3): CENTER round-trip ->
            % <w:jc w:val="center"/>; []-reset removes w:jc (empty <w:pPr/>).
            AL = mat2doc.enum.text.WD_ALIGN_PARAGRAPH.CENTER;
            d = mat2doc.Document();
            [p, para] = newParsed(testCase, "", d);
            testCase.verifyTrue(isequal(para.alignment, []), 'alignment initially [] (None)');
            para.alignment = AL;
            testCase.verifyEqual(ser(p), pWrap(testCase, "<w:pPr><w:jc w:val=""center""/></w:pPr>"), ...
                'alignment=CENTER -> <w:jc w:val="center"/> (hard-coded L1)');
            testCase.verifyEqual(string(para.alignment), "CENTER", 'alignment get CENTER');
            para.alignment = [];
            testCase.verifyEqual(ser(p), pWrap(testCase, "<w:pPr/>"), ...
                'alignment=[] removes w:jc -> empty <w:pPr/> (hard-coded L1)');
            testCase.verifyTrue(isequal(para.alignment, []), 'alignment=[] get -> [] (None)');
        end

        function test_clear_returns_self_preserves_ppr(testCase)
            % Nominal + H5 (paragraph.py 61-67, s0028 clear): clear() returns the
            % SAME handle; the w:pPr is PRESERVED and all content removed; text "".
            o = loadOracle().clear;
            d = mat2doc.Document();
            seeded = "<w:pPr><w:jc w:val=""center""/></w:pPr><w:r><w:t>old</w:t></w:r>";
            [p, para] = newParsed(testCase, seeded, d);
            ret = para.clear();
            testCase.verifyTrue(ret == para, 'clear() returns the SAME handle (H5)');
            testCase.verifyEqual(string(o.returns_self), "True", 'oracle records returns_self True');
            testCase.verifyEqual(hx_e(p), string(o.serhex), 'clear serhex vs oracle');
            testCase.verifyEqual(ser(p), pWrap(testCase, "<w:pPr><w:jc w:val=""center""/></w:pPr>"), ...
                'clear() preserves w:pPr, drops content (hard-coded L1)');
            testCase.verifyEqual(string(para.text), "", 'clear() text ""');
            testCase.verifyEqual(numel(para.runs), 0, 'clear() -> 0 runs');
        end

        % =============================================================== %
        % 3. contains_page_break                                          %
        % =============================================================== %

        function test_contains_page_break_true_false(testCase)
            % Nominal + Edge (paragraph.py 69-72, H4 bool(node-list), s0028
            % contains_page_break): false for w:r / w:r+w:t; true for a
            % lastRenderedPageBreak in a hyperlink and for two breaks.
            o = loadOracle().contains_page_break;
            d = mat2doc.Document();
            cases = { ...
                "r_only",          "<w:r/>",                                                     false; ...
                "r_text",          "<w:r><w:t>foobar</w:t></w:r>",                               false; ...
                "hyperlink_break", "<w:hyperlink><w:r><w:t>abc</w:t><w:lastRenderedPageBreak/><w:t>def</w:t></w:r></w:hyperlink>", true; ...
                "two_breaks",      "<w:r><w:lastRenderedPageBreak/><w:lastRenderedPageBreak/></w:r>", true};
            for i = 1:size(cases, 1)
                [~, para] = newParsed(testCase, cases{i, 2}, d);
                testCase.verifyEqual(para.contains_page_break, cases{i, 3}, ...
                    sprintf('contains_page_break %s (hard-coded)', cases{i, 1}));
                testCase.verifyEqual(tf(para.contains_page_break), string(o.(cases{i, 1})), ...
                    sprintf('contains_page_break %s vs oracle', cases{i, 1}));
            end
        end

        % =============================================================== %
        % 4. list-property surface shape (object arrays)                  %
        % =============================================================== %

        function test_runs_hyperlinks_list_shape(testCase)
            % Nominal + Edge + H1 (paragraph.py 74-77, 124-128, s0028 order):
            % runs/hyperlinks are homogeneous 1xN object arrays; H1 first/last;
            % EMPTY -> a TYPED 1x0 array (Class.empty(1,0)), NOT [].
            d = mat2doc.Document();

            % interleaved: runs [R1,R2,R3], hyperlinks [A1,A2]
            interleaved = "<w:r><w:t>R1</w:t></w:r>" + ...
                "<w:hyperlink><w:r><w:t>A1</w:t></w:r></w:hyperlink>" + ...
                "<w:r><w:t>R2</w:t></w:r>" + ...
                "<w:hyperlink><w:r><w:t>A2</w:t></w:r></w:hyperlink>" + ...
                "<w:r><w:t>R3</w:t></w:r>";
            [~, para] = newParsed(testCase, interleaved, d);

            runs = para.runs;
            testCase.verifyClass(runs, 'mat2doc.text.Run', 'runs is a Run object array');
            testCase.verifyEqual(numel(runs), 3, 'runs numel == len(runs) == 3');
            testCase.verifyEqual(string(runs(1).text), "R1", 'runs H1 first == R1');
            testCase.verifyEqual(string(runs(end).text), "R3", 'runs H1 last == R3');
            testCase.verifyEqual(string(runs(2).text), "R2", 'runs(2) == R2');

            hls = para.hyperlinks;
            testCase.verifyClass(hls, 'mat2doc.text.Hyperlink', 'hyperlinks is a Hyperlink object array');
            testCase.verifyEqual(numel(hls), 2, 'hyperlinks numel == 2');
            testCase.verifyEqual(string(hls(1).text), "A1", 'hyperlinks H1 first == A1');
            testCase.verifyEqual(string(hls(end).text), "A2", 'hyperlinks H1 last == A2');

            % EMPTY paragraph -> typed 1x0 (NOT [])
            [~, empt] = newParsed(testCase, "", d);
            er = empt.runs;
            testCase.verifyClass(er, 'mat2doc.text.Run', 'empty runs is a TYPED Run (not [])');
            testCase.verifySize(er, [1 0], 'empty runs -> 1x0');
            eh = empt.hyperlinks;
            testCase.verifyClass(eh, 'mat2doc.text.Hyperlink', 'empty hyperlinks is a TYPED Hyperlink (not [])');
            testCase.verifySize(eh, [1 0], 'empty hyperlinks -> 1x0');

            % single-element run list
            [~, one] = newParsed(testCase, "<w:r><w:t>solo</w:t></w:r>", d);
            or = one.runs;
            testCase.verifyEqual(numel(or), 1, 'single-run paragraph -> 1x1');
            testCase.verifyEqual(string(or(1).text), "solo", 'single run text');
        end

        function test_rendered_page_breaks_list_shape(testCase)
            % Nominal + Edge (paragraph.py 115-122, s0028 rendered_page_breaks):
            % RenderedPageBreak 1xN object array; empty -> typed 1x0 (NOT []).
            o = loadOracle().rendered_page_breaks;
            d = mat2doc.Document();
            cases = { ...
                "none",     "<w:r><w:t>x</w:t></w:r>",                                            0; ...
                "one",      "<w:r><w:lastRenderedPageBreak/><w:t>x</w:t></w:r>",                  1; ...
                "in_hlink", "<w:hyperlink><w:r><w:t>a</w:t><w:lastRenderedPageBreak/><w:t>b</w:t></w:r></w:hyperlink>", 1};
            for i = 1:size(cases, 1)
                [~, para] = newParsed(testCase, cases{i, 2}, d);
                rpbs = para.rendered_page_breaks;
                testCase.verifyClass(rpbs, 'mat2doc.text.RenderedPageBreak', ...
                    sprintf('rendered_page_breaks %s is TYPED', cases{i, 1}));
                testCase.verifyEqual(numel(rpbs), cases{i, 3}, ...
                    sprintf('rendered_page_breaks %s count (hard-coded)', cases{i, 1}));
                testCase.verifyEqual(numel(rpbs), double(o.(cases{i, 1}).count), ...
                    sprintf('rendered_page_breaks %s count vs oracle', cases{i, 1}));
            end
            % empty -> typed 1x0
            [~, none] = newParsed(testCase, "<w:r><w:t>x</w:t></w:r>", d);
            testCase.verifySize(none.rendered_page_breaks, [1 0], 'empty rpbs -> 1x0 (not [])');
        end

        function test_iter_inner_content_document_order(testCase)
            % Nominal + Edge + H1/H9/H10 (paragraph.py 94-107, s0028 order): a 1xN
            % CELL of Run|Hyperlink in EXACT document order; empty -> 1x0 cell.
            d = mat2doc.Document();
            interleaved = "<w:r><w:t>R1</w:t></w:r>" + ...
                "<w:hyperlink><w:r><w:t>A1</w:t></w:r></w:hyperlink>" + ...
                "<w:r><w:t>R2</w:t></w:r>" + ...
                "<w:hyperlink><w:r><w:t>A2</w:t></w:r></w:hyperlink>" + ...
                "<w:r><w:t>R3</w:t></w:r>";
            [~, para] = newParsed(testCase, interleaved, d);
            items = para.iter_inner_content();
            testCase.verifyClass(items, 'cell', 'iter_inner_content -> cell');
            testCase.verifyEqual(numel(items), 5, 'iter_inner_content 5 items');
            expTypes = ["Run","Hyperlink","Run","Hyperlink","Run"];
            expText  = ["R1","A1","R2","A2","R3"];
            for k = 1:5
                it = items{k};
                if isa(it, "mat2doc.text.Run"), ty = "Run"; else, ty = "Hyperlink"; end
                testCase.verifyEqual(ty, expTypes(k), sprintf('inner[%d] type', k));
                testCase.verifyEqual(string(it.text), expText(k), sprintf('inner[%d] text (document order)', k));
            end
            % para.text is the concatenation
            testCase.verifyEqual(string(para.text), "R1A1R2A2R3", 'para.text concatenates inner content');
            % empty -> 1x0 cell
            [~, empt] = newParsed(testCase, "", d);
            ei = empt.iter_inner_content();
            testCase.verifyClass(ei, 'cell', 'empty iter -> cell');
            testCase.verifySize(ei, [1 0], 'empty iter -> 1x0 cell');
        end

        % =============================================================== %
        % 5. text get/set                                                 %
        % =============================================================== %

        function test_text_get_set_roundtrip(testCase)
            % Nominal + Edge + Regression (paragraph.py 149-168, s0028 text_set):
            % set replaces all content with a SINGLE run (\t -> <w:tab/>), pPr
            % preserved; "" -> a single empty run; get maps tab.
            oSet = loadOracle().text_set;
            oEmpty = loadOracle().text_set_empty;
            d = mat2doc.Document();
            seeded = "<w:pPr><w:jc w:val=""center""/></w:pPr><w:r><w:t>old</w:t></w:r>";

            [p, para] = newParsed(testCase, seeded, d);
            para.text = "new" + sprintf('\t') + "text";
            testCase.verifyEqual(hx_e(p), string(oSet.serhex), 'text set serhex vs oracle');
            testCase.verifyEqual(ser(p), pWrap(testCase, ...
                "<w:pPr><w:jc w:val=""center""/></w:pPr>" + ...
                "<w:r><w:t>new</w:t><w:tab/><w:t>text</w:t></w:r>"), ...
                'text set: single run, pPr preserved, \t -> <w:tab/> (hard-coded L1)');
            testCase.verifyEqual(string(para.text), "new" + sprintf('\t') + "text", 'text set round-trip');
            testCase.verifyEqual(numel(para.runs), 1, 'text set -> single run');

            [p, para] = newParsed(testCase, seeded, d);
            para.text = "";
            testCase.verifyEqual(hx_e(p), string(oEmpty.serhex), 'text set "" serhex vs oracle');
            testCase.verifyEqual(ser(p), pWrap(testCase, ...
                "<w:pPr><w:jc w:val=""center""/></w:pPr><w:r/>"), ...
                'text set "" -> single empty run, pPr preserved (hard-coded L1)');
            testCase.verifyEqual(string(para.text), "", 'text set "" round-trip');
            testCase.verifyEqual(numel(para.runs), 1, 'text set "" -> single (empty) run');
        end

        % =============================================================== %
        % 6. insert_paragraph_before                                      %
        % =============================================================== %

        function test_insert_paragraph_before(testCase)
            % Nominal + Edge + Regression (paragraph.py 79-92, 170-173, s0028
            % insert_before): a new paragraph inserted BEFORE this one (body child
            % order [new, orig]); with text -> single run; empty -> empty.
            o = loadOracle().insert_before;
            oe = loadOracle().insert_before_empty;
            d = mat2doc.Document();

            body = rparse(bodyXml(testCase, "<w:p><w:r><w:t>orig</w:t></w:r></w:p>"));
            p0 = body.xpath(".//w:p"); p0 = p0(1);
            para = mat2doc.text.Paragraph(p0, d);
            newp = para.insert_paragraph_before("intro");
            ps = body.xpath("./w:p");
            testCase.verifyClass(newp, 'mat2doc.text.Paragraph', 'insert_paragraph_before returns a Paragraph');
            testCase.verifyEqual(string(newp.text), "intro", 'new paragraph text "intro"');
            testCase.verifyEqual(numel(ps), 2, 'body now has 2 paragraphs');
            testCase.verifyEqual(string(mat2doc.text.Paragraph(ps(1), d).text), "intro", ...
                'first child is the NEW paragraph (inserted before)');
            testCase.verifyEqual(string(mat2doc.text.Paragraph(ps(2), d).text), "orig", ...
                'second child is the original');
            testCase.verifyEqual(hx_e(body), string(o.body_serhex), 'insert_before body serhex vs oracle');

            % empty insert
            body = rparse(bodyXml(testCase, "<w:p><w:r><w:t>orig</w:t></w:r></w:p>"));
            p0 = body.xpath(".//w:p"); p0 = p0(1);
            newp = mat2doc.text.Paragraph(p0, d).insert_paragraph_before();
            testCase.verifyEqual(string(newp.text), "", 'empty insert -> empty paragraph');
            testCase.verifyEqual(numel(body.xpath("./w:p")), 2, 'empty insert -> 2 paragraphs');
            testCase.verifyEqual(string(oe.new_text), "", 'oracle empty insert new_text ""');
        end

        function test_insert_paragraph_before_rootless_typeerror(testCase)
            % Error path (Gate-3 section 6): insert_paragraph_before over a
            % rootless <w:p/> fails at add_p_before (addprevious with no parent) ->
            % mat2doc:TypeError, IDENTICALLY on both sides. Pin the IDENTIFIER +
            % the verbatim lxml message.
            d = mat2doc.Document();
            prootless = rparse("<w:p xmlns:w=""" + testCase.W + """/>");
            para = mat2doc.text.Paragraph(prootless, d);
            ME = captureError(@() para.insert_paragraph_before("x"));
            testCase.verifyEqual(string(ME.identifier), "mat2doc:TypeError", ...
                'rootless insert_paragraph_before -> mat2doc:TypeError identifier');
            testCase.verifyEqual(string(ME.message), ...
                "Only processing instructions and comments can be siblings of the root element", ...
                'rootless insert -> verbatim lxml message');
        end

        % =============================================================== %
        % 7. Paragraph.style -- the P4-7 stub (get AND set raise)         %
        % =============================================================== %

        function test_paragraph_style_stub_notYetPorted(testCase)
            % Error path (paragraph.py 130-147): style get AND set both delegate
            % through part().get_style / get_style_id -- the P4-7 STUBS -> both
            % raise mat2doc:notYetPorted (faithful propagation, NOT a stand-in).
            d = mat2doc.Document();
            [~, para] = newParsed(testCase, "", d);
            MEget = captureError(@() para.style);
            testCase.verifyEqual(string(MEget.identifier), "mat2doc:notYetPorted", ...
                'Paragraph.style GETTER -> mat2doc:notYetPorted (P4-7 stub)');
            MEset = captureError(@() setParaStyle(para, "Normal"));
            testCase.verifyEqual(string(MEset.identifier), "mat2doc:notYetPorted", ...
                'Paragraph.style SETTER -> mat2doc:notYetPorted (P4-7 stub)');
        end

        % =============================================================== %
        % 8. Hyperlink surface (LIVE rels)                                %
        % =============================================================== %

        function test_hyperlink_surface(testCase)
            % Nominal + Edge (hyperlink.py 33-121, s0028 hyperlink): external r:id
            % -> target_ref resolved LIVE via a real relationship; internal jump
            % (w:anchor only) -> address ""; fragment ±; url = address /
            % address#fragment / "" (no address); runs; text; contains_page_break.
            o = loadOracle().hyperlink;
            d = mat2doc.Document();
            rid = d.part().relate_to("https://google.com/", ...
                mat2doc.opc.RELATIONSHIP_TYPE.HYPERLINK, true);

            % external r:id -> target_ref LIVE
            h = hlink(testCase, "r:id=""" + rid + """", "<w:r><w:t>post</w:t></w:r>", d);
            testCase.verifyEqual(string(h.address), "https://google.com/", 'external address (LIVE target_ref)');
            testCase.verifyEqual(string(h.fragment), "", 'external fragment ""');
            testCase.verifyEqual(string(h.url), "https://google.com/", 'external url == address');
            testCase.verifyEqual(string(h.text), "post", 'external text');
            testCase.verifyEqual(numel(h.runs), 1, 'external runs -> 1x1');
            testCase.verifyEqual(string(h.runs(1).text), "post", 'external run text');
            testCase.verifyFalse(h.contains_page_break, 'external cpb False');
            testCase.verifyEqual(string(h.address), string(o.external.address), 'external address vs oracle');
            testCase.verifyEqual(string(h.url), string(o.external.url), 'external url vs oracle');

            % external + fragment -> url = address#fragment
            h = hlink(testCase, "r:id=""" + rid + """ w:anchor=""foo""", "<w:r><w:t>post</w:t></w:r>", d);
            testCase.verifyEqual(string(h.address), "https://google.com/", 'external+frag address');
            testCase.verifyEqual(string(h.fragment), "foo", 'external+frag fragment');
            testCase.verifyEqual(string(h.url), "https://google.com/#foo", 'external+frag url == address#fragment');
            testCase.verifyEqual(string(h.url), string(o.external_fragment.url), 'external+frag url vs oracle');

            % internal jump: w:anchor only, no r:id -> address ""
            h = hlink(testCase, "w:anchor=""_Toc147925734""", "<w:r><w:t>Heading</w:t></w:r>", d);
            testCase.verifyEqual(string(h.address), "", 'internal jump address "" (no r:id)');
            testCase.verifyEqual(string(h.fragment), "_Toc147925734", 'internal jump fragment');
            testCase.verifyEqual(string(h.url), "", 'internal jump url "" (no address)');
            testCase.verifyEqual(string(h.text), "Heading", 'internal jump text');
            testCase.verifyEqual(string(h.address), string(o.internal.address), 'internal address vs oracle');
            testCase.verifyEqual(string(h.fragment), string(o.internal.fragment), 'internal fragment vs oracle');

            % bare hyperlink: no rId, no anchor, no runs
            h = hlink(testCase, "", "", d);
            testCase.verifyEqual(string(h.address), "", 'bare address ""');
            testCase.verifyEqual(string(h.fragment), "", 'bare fragment ""');
            testCase.verifyEqual(string(h.url), "", 'bare url ""');
            testCase.verifyEqual(string(h.text), "", 'bare text ""');
            testCase.verifyEqual(numel(h.runs), 0, 'bare runs -> 1x0');
            testCase.verifySize(h.runs, [1 0], 'bare runs typed 1x0 (not [])');
            testCase.verifyFalse(h.contains_page_break, 'bare cpb False');

            % with break inside
            h = hlink(testCase, "", "<w:r><w:t>a</w:t><w:lastRenderedPageBreak/><w:t>b</w:t></w:r>", d);
            testCase.verifyEqual(string(h.text), "ab", 'with-break text "ab"');
            testCase.verifyTrue(h.contains_page_break, 'with-break cpb True');
        end

        % =============================================================== %
        % 9. RenderedPageBreak fragment split                             %
        % =============================================================== %

        function test_renderedpagebreak_fragment_split(testCase)
            % Nominal + Edge + Regression (pagebreak.py 47-104, s0028 pagebreak):
            % the preceding/following fragment split incl. leading/trailing None
            % and the ATOMIC break-inside-hyperlink (preceding keeps the WHOLE
            % w:hyperlink -> [pPr,hyperlink]; following DROPS it -> [pPr,r]). Every
            % fragment serhex + localnames + text pinned vs the frozen oracle.
            o = loadOracle().pagebreak;
            d = mat2doc.Document();

            % leading: preceding None, following "foobar" [pPr,r]
            checkFrag(testCase, d, o.leading, ...
                "<w:pPr><w:ind/></w:pPr><w:r><w:lastRenderedPageBreak/>" + ...
                "<w:t>foo</w:t><w:t>bar</w:t></w:r>", 1);

            % mid-run: preceding "foo" [pPr,r], following "barbarfoo" [pPr,r,r]
            checkFrag(testCase, d, o.mid, ...
                "<w:pPr><w:ind/></w:pPr>" + ...
                "<w:r><w:t>foo</w:t><w:lastRenderedPageBreak/><w:t>bar</w:t></w:r>" + ...
                "<w:r><w:t>barfoo</w:t></w:r>", 1);

            % ATOMIC break-inside-hyperlink: preceding "foobar" keeps WHOLE
            % w:hyperlink [pPr,hyperlink]; following "barfoo" drops it [pPr,r].
            checkFrag(testCase, d, o.in_hyperlink, ...
                "<w:pPr><w:ind/></w:pPr>" + ...
                "<w:hyperlink><w:r><w:t>foo</w:t><w:lastRenderedPageBreak/>" + ...
                "<w:t>bar</w:t></w:r></w:hyperlink><w:r><w:t>barfoo</w:t></w:r>", 1);

            % trailing: preceding "foo" [r], following None
            checkFrag(testCase, d, o.trailing, ...
                "<w:r><w:t>foo</w:t><w:lastRenderedPageBreak/></w:r>", 1);
        end

        function test_renderedpagebreak_second_break_valueerror(testCase)
            % Error path (pagebreak.py, s0028 pagebreak_guard): the SECOND
            % (non-first) rendered page-break's preceding_paragraph_fragment raises
            % mat2doc:ValueError. Pin the IDENTIFIER + the verbatim message.
            o = loadOracle().pagebreak_guard;
            d = mat2doc.Document();
            p = rparse(pXml(testCase, ...
                "<w:r><w:t>abc</w:t><w:lastRenderedPageBreak/><w:lastRenderedPageBreak/></w:r>"));
            lrpbs = p.xpath(".//w:lastRenderedPageBreak");
            rpb2 = mat2doc.text.RenderedPageBreak(lrpbs(2), d);
            ME = captureError(@() rpb2.preceding_paragraph_fragment);
            testCase.verifyEqual(string(ME.identifier), "mat2doc:ValueError", ...
                'second-break preceding fragment -> mat2doc:ValueError identifier');
            testCase.verifyEqual(string(ME.message), ...
                "only defined on first rendered page-break in paragraph", ...
                'second-break -> verbatim message');
            testCase.verifyEqual(string(ME.message), string(o.msg), 'second-break message vs oracle');
        end

        % =============================================================== %
        % 10. tier / identity (H5) lineage                                %
        % =============================================================== %

        function test_tier_and_identity_h5(testCase)
            % H5 / tier (paragraph.py 23, hyperlink.py 20, pagebreak.py 15): none
            % of the three overrides __eq__ -> two proxies over the SAME element are
            % NOT equal (instance identity == Python default object identity), the
            % OPPOSITE of an ElementProxy subclass. clear() returns the same handle.
            d = mat2doc.Document();
            [p, para] = newParsed(testCase, "<w:r><w:t>x</w:t></w:r>", d);
            testCase.verifyTrue(isa(para, 'mat2doc.shared.StoryChild'), 'Paragraph < StoryChild');
            % two Paragraph proxies over the SAME element -> instance identity (unequal)
            para2 = mat2doc.text.Paragraph(p, d);
            testCase.verifyFalse(para == para2, ...
                'two Paragraph proxies over the same element are NOT equal (no __eq__; H5)');
            testCase.verifyTrue(para == para, 'a Paragraph equals itself (handle identity)');
            testCase.verifyTrue(para.clear() == para, 'clear() returns the SAME handle');

            hp = rparse(hlinkXml(testCase, "", "<w:r><w:t>t</w:t></w:r>"));
            h1 = mat2doc.text.Hyperlink(hp, d);
            h2 = mat2doc.text.Hyperlink(hp, d);
            testCase.verifyTrue(isa(h1, 'mat2doc.shared.Parented'), 'Hyperlink < Parented');
            testCase.verifyFalse(h1 == h2, 'two Hyperlink proxies over the same element are NOT equal (H5)');

            bp = rparse(pXml(testCase, "<w:r><w:lastRenderedPageBreak/></w:r>"));
            lr = bp.xpath(".//w:lastRenderedPageBreak");
            r1 = mat2doc.text.RenderedPageBreak(lr(1), d);
            r2 = mat2doc.text.RenderedPageBreak(lr(1), d);
            testCase.verifyTrue(isa(r1, 'mat2doc.shared.Parented'), 'RenderedPageBreak < Parented');
            testCase.verifyFalse(r1 == r2, 'two RenderedPageBreak proxies over the same element are NOT equal (H5)');
        end

        % =============================================================== %
        % 11. G-scenario document byte pin (the M2 de-risker)             %
        % =============================================================== %

        function test_gscenario_document_byte_pin(testCase)
            % Regression (THE G-scenario byte pin, s0029): a fresh mat2doc.Document()
            % whose three paragraphs are built THROUGH the NEW Paragraph API
            % (buildGScenario, the s0029 .m twin body verbatim) saved -> the
            % extracted word/document.xml is byte-identical (size + SHA-256 +
            % whole-bytes) to the frozen python-docx 1.2.0 reference. Closest guard
            % to M2's add_paragraph path (P4-6).
            docxBytes = buildGScenario();
            [blobs, names] = zipEntryList(docxBytes);
            i = find(names == "word/document.xml", 1);
            testCase.assertNotEmpty(i, 'saved package must contain word/document.xml');
            got = blobs{i};

            testCase.verifyEqual(numel(got), testCase.SIZE_S0029_DOCUMENT, ...
                sprintf('word/document.xml must be exactly %d B', testCase.SIZE_S0029_DOCUMENT));
            testCase.verifyEqual(sha256hex(got), testCase.SHA_S0029_DOCUMENT, ...
                'word/document.xml SHA-256 must equal the frozen s0029 oracle (byte-identical L1)');
            want = loadDocumentFixture();
            verifyByteIdentical(testCase, got, want, ...
                'G-scenario word/document.xml == frozen s0029 reference');
        end

        % =============================================================== %
        % 12. EQUIVALENCE -- full s0028 battery vs the frozen oracle       %
        % =============================================================== %

        function test_equivalence_full_battery_vs_frozen_oracle(testCase)
            % Equivalence: replay the ENTIRE s0028 battery (runProbes -- the .m
            % twin's body VERBATIM: contains_page_break / order (runs, hyperlinks,
            % iter_inner_content, text) / rendered_page_breaks / add_run / clear /
            % text_set / insert_before / hyperlink (LIVE rels) / pagebreak (fragment
            % split) / pagebreak_guard) and flatten-compare EVERY leaf to the frozen
            % python-docx 1.2.0 oracle copied into data\s0028_probe_oracle.json.
            % Gate-3 found ZERO divergences (probe_diff MATCH 170/170), so every
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
            testCase.verifyGreaterThan(numel(oKeys), 120, ...
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

function x = nsWR()
    x = "xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""" + ...
        " xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""";
end

function x = pXml(~, inner)
    x = "<w:p " + nsWR() + ">" + string(inner) + "</w:p>";
end

function x = bodyXml(~, inner)
    x = "<w:body " + nsWR() + ">" + string(inner) + "</w:body>";
end

function x = hlinkXml(~, attrs, inner)
    a = string(attrs);
    if strlength(a) > 0, a = " " + a; end
    x = "<w:hyperlink " + nsWR() + a + ">" + string(inner) + "</w:hyperlink>";
end

function h = hlink(testCase, attrs, inner, d)
    h = mat2doc.text.Hyperlink(rparse(hlinkXml(testCase, attrs, inner)), d);
end

function e = rparse(xml)
    e = mat2doc.oxml.parse_xml(uint8(unicode2native(char(xml), "UTF-8")));
end

function [p, para] = newParsed(testCase, inner, d)
    % Parse a <w:p> (with xmlns:w + xmlns:r, matching the scenario twin) and wrap
    % it in a Paragraph whose parent is the real Document d (a ProvidesStoryPart).
    p = rparse(pXml(testCase, inner));
    para = mat2doc.text.Paragraph(p, d);
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
    % decl + newline + <w:p xmlns:w="..." xmlns:r="...">BODY</w:p>. The paragraph
    % is created via parse_xml of a <w:p> declaring BOTH xmlns:w and xmlns:r (the
    % scenario-twin construction), so both prefixes appear on the serialized root.
    s = decl() + newline + "<w:p xmlns:w=""" + testCase.W + """ xmlns:r=""" + testCase.R + """>" + ...
        string(body) + "</w:p>";
end

function setParaStyle(para, name)
    para.style = name;
end

function s = tf(b)
    if b, s = "True"; else, s = "False"; end
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

function checkFrag(testCase, d, od, pInner, breakIndex)
    % Verify a RenderedPageBreak fragment split against the frozen oracle struct
    % od: preceding/following is-None + text + serhex + child localnames.
    p = rparse(pXml(testCase, pInner));
    lrpbs = p.xpath(".//w:lastRenderedPageBreak");
    lrpb = lrpbs(breakIndex);
    rpb = mat2doc.text.RenderedPageBreak(lrpb, d);
    pre = rpb.preceding_paragraph_fragment;
    fol = rpb.following_paragraph_fragment;

    testCase.verifyEqual(tf(isequal(pre, [])), string(od.preceding_is_none), 'preceding_is_none vs oracle');
    testCase.verifyEqual(tf(isequal(fol, [])), string(od.following_is_none), 'following_is_none vs oracle');

    if isequal(pre, [])
        testCase.verifyEqual("None", string(od.preceding_text), 'preceding None text');
    else
        testCase.verifyEqual(string(pre.text), string(od.preceding_text), 'preceding_text vs oracle');
        e = lrpb.preceding_fragment_p();
        testCase.verifyEqual(hx_e(e), string(od.preceding_serhex), 'preceding fragment serhex vs oracle');
        testCase.verifyEqual(localnames(e), cellstr(od.preceding_localnames)', 'preceding localnames vs oracle');
    end
    if isequal(fol, [])
        testCase.verifyEqual("None", string(od.following_text), 'following None text');
    else
        testCase.verifyEqual(string(fol.text), string(od.following_text), 'following_text vs oracle');
        e = lrpb.following_fragment_p();
        testCase.verifyEqual(hx_e(e), string(od.following_serhex), 'following fragment serhex vs oracle');
        testCase.verifyEqual(localnames(e), cellstr(od.following_localnames)', 'following localnames vs oracle');
    end
end

function C = localnames(e)
    % Ordered child localnames -> 1xN cell of char ({} when none).
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

function o = loadOracle()
    % Read the co-located frozen s0028 oracle in BINARY mode (no CRLF translation)
    % and decode UTF-8 -> struct. jsondecode is whitespace/line-ending agnostic, so
    % no `* binary` pin is needed for this value/serhex fixture (s0024/s0026 precedent).
    here = fileparts(mfilename('fullpath'));
    p = fullfile(here, 'data', 's0028_probe_oracle.json');
    fid = fopen(p, 'r', 'n');
    if fid < 0
        error('mat2doc:test:openFail', 'cannot open frozen oracle %s', p);
    end
    raw = fread(fid, Inf, '*uint8')';
    fclose(fid);
    o = jsondecode(native2unicode(raw, 'UTF-8'));
end

function want = loadDocumentFixture()
    % The frozen s0029 word/document.xml byte fixture (co-located `* binary`
    % .gitattributes so it is checked out byte-for-byte). Read in BINARY mode.
    here = fileparts(mfilename('fullpath'));
    p = fullfile(here, 'data', 's0029_document.xml');
    want = readBytes(p);
end

function docxBytes = buildGScenario()
    % Replay the s0029 G-scenario body VERBATIM (validation\mat2doc\scenarios\
    % s0029_p4_5b_paragraph_gscenario.m): a fresh mat2doc.Document() whose three
    % paragraphs are built THROUGH the NEW Paragraph API (add_run / text= /
    % alignment=), created via body.add_p() (inserted before w:sectPr). Returns
    % the whole-package .save() bytes.
    TAB = sprintf('\t'); NL = sprintf('\n');
    NONASCII = string(char([99 97 102 233 32 9749 32 26085 26412 35486])); % café ☕ 日本語

    d = mat2doc.Document();
    body = d.element().body;

    % ---- P1: two runs + center alignment (via the Paragraph API) ----
    p1 = mat2doc.text.Paragraph(body.add_p(), d);
    p1.add_run("Hello ");
    p1.add_run("World");
    p1.alignment = mat2doc.enum.text.WD_ALIGN_PARAGRAPH.CENTER;

    % ---- P2: text setter (tab + line break) ----
    p2 = mat2doc.text.Paragraph(body.add_p(), d);
    p2.text = "Second" + TAB + "paragraph" + NL + "line2";

    % ---- P3: non-ASCII run ----
    p3 = mat2doc.text.Paragraph(body.add_p(), d);
    p3.add_run(NONASCII);

    tmp = [tempname '.docx'];
    cleanup = onCleanup(@() deleteIfExists(tmp)); %#ok<NASGU>
    d.save(tmp);
    docxBytes = readBytes(tmp);
end

function P = runProbes()
    % Replay the s0028 probe sequence (the .m twin's P-struct body, VERBATIM
    % tags/inputs/order) and return the nested struct of tagged canonical values.
    % Embedded so the Equivalence leg is self-contained (the validation-folder
    % scenario is NOT on the toolbox path and must not be a dependency). Mirrors
    % validation\mat2doc\scenarios\s0028_p4_5b_paragraph_probe.m.
    d = mat2doc.Document();                     % a real ProvidesStoryPart parent
    P = struct();

    TAB = sprintf('\t'); NL = sprintf('\n'); CR = sprintf('\r');
    NONASCII = string(char([99 97 102 233 32 9749 32 26085 26412 35486])); % café ☕ 日本語

    % ===================== 1. contains_page_break (parse trees) ============
    cpb = struct();
    cpbCases = { ...
        "r_only",          "<w:r/>"; ...
        "r_text",          "<w:r><w:t>foobar</w:t></w:r>"; ...
        "hyperlink_break", "<w:hyperlink><w:r><w:t>abc</w:t><w:lastRenderedPageBreak/><w:t>def</w:t></w:r></w:hyperlink>"; ...
        "two_breaks",      "<w:r><w:lastRenderedPageBreak/><w:lastRenderedPageBreak/></w:r>"};
    for i = 1:size(cpbCases, 1)
        para = mat2doc.text.Paragraph(rparse(rpXml(cpbCases{i, 2})), d);
        cpb.(cpbCases{i, 1}) = tf(para.contains_page_break);
    end
    P.contains_page_break = cpb;

    % ===================== 2. runs / hyperlinks / iter / text ==============
    order = struct();
    orderCases = { ...
        "empty",       ""; ...
        "r_h_r",       "<w:r/><w:hyperlink/><w:r/>"; ...
        "h_r_h",       "<w:hyperlink/><w:r/><w:hyperlink/>"; ...
        "interleaved", "<w:r><w:t>R1</w:t></w:r>" + ...
                       "<w:hyperlink><w:r><w:t>A1</w:t></w:r></w:hyperlink>" + ...
                       "<w:r><w:t>R2</w:t></w:r>" + ...
                       "<w:hyperlink><w:r><w:t>A2</w:t></w:r></w:hyperlink>" + ...
                       "<w:r><w:t>R3</w:t></w:r>"};
    for i = 1:size(orderCases, 1)
        para = mat2doc.text.Paragraph(rparse(rpXml(orderCases{i, 2})), d);
        s = struct();
        s.runs = arrInfo(para.runs);
        s.hyperlinks = arrInfo(para.hyperlinks);
        s.inner = innerList(para.iter_inner_content());
        s.text = string(para.text);
        order.(orderCases{i, 1}) = s;
    end
    P.order = order;

    % ===================== 3. rendered_page_breaks list (count) ============
    rpbs = struct();
    rpbCases = { ...
        "none",     "<w:r><w:t>x</w:t></w:r>"; ...
        "one",      "<w:r><w:lastRenderedPageBreak/><w:t>x</w:t></w:r>"; ...
        "in_hlink", "<w:hyperlink><w:r><w:t>a</w:t><w:lastRenderedPageBreak/><w:t>b</w:t></w:r></w:hyperlink>"};
    for i = 1:size(rpbCases, 1)
        para = mat2doc.text.Paragraph(rparse(rpXml(rpbCases{i, 2})), d);
        s = struct(); s.count = numel(para.rendered_page_breaks);
        rpbs.(rpbCases{i, 1}) = s;
    end
    P.rendered_page_breaks = rpbs;

    % ===================== 4. add_run mutation -- serhex parity ============
    ar = struct();
    p = rparse(rpXml("")); para = mat2doc.text.Paragraph(p, d);
    para.add_run();
    s = struct(); s.serhex = hx_e(p); s.text = string(para.text); s.count = numel(para.runs);
    ar.empty = s;

    p = rparse(rpXml("")); para = mat2doc.text.Paragraph(p, d);
    para.add_run("Hello");
    s = struct(); s.serhex = hx_e(p); s.text = string(para.text); s.count = numel(para.runs);
    ar.plain = s;

    p = rparse(rpXml("")); para = mat2doc.text.Paragraph(p, d);
    para.add_run("a" + TAB + "b");
    s = struct(); s.serhex = hx_e(p); s.text = string(para.text);
    ar.tab = s;

    p = rparse(rpXml("")); para = mat2doc.text.Paragraph(p, d);
    para.add_run("x" + NL + "y" + CR + "z");
    s = struct(); s.serhex = hx_e(p); s.text = string(para.text);
    ar.newline_cr = s;

    p = rparse(rpXml("")); para = mat2doc.text.Paragraph(p, d);
    para.add_run(NONASCII);
    s = struct(); s.serhex = hx_e(p); s.text = string(para.text);
    ar.nonascii = s;

    p = rparse(rpXml("")); para = mat2doc.text.Paragraph(p, d);
    para.add_run("Hello ");
    para.add_run("World");
    s = struct(); s.serhex = hx_e(p); s.text = string(para.text); s.count = numel(para.runs);
    ar.multi = s;
    P.add_run = ar;

    % ===================== 5. clear -- returns self + pPr preserved ========
    seeded = "<w:pPr><w:jc w:val=""center""/></w:pPr><w:r><w:t>old</w:t></w:r>";
    p = rparse(rpXml(seeded)); para = mat2doc.text.Paragraph(p, d);
    ret = para.clear();
    s = struct();
    s.returns_self = tf(ret == para);
    s.serhex = hx_e(p);
    s.text = string(para.text);
    P.clear = s;

    % ===================== 6. text setter ==================================
    p = rparse(rpXml(seeded)); para = mat2doc.text.Paragraph(p, d);
    para.text = "new" + TAB + "text";
    s = struct(); s.serhex = hx_e(p); s.text = string(para.text); s.count = numel(para.runs);
    P.text_set = s;

    p = rparse(rpXml(seeded)); para = mat2doc.text.Paragraph(p, d);
    para.text = "";
    s = struct(); s.serhex = hx_e(p); s.text = string(para.text); s.count = numel(para.runs);
    P.text_set_empty = s;

    % ===================== 7. insert_paragraph_before ======================
    body = rparse(rbodyXml("<w:p><w:r><w:t>orig</w:t></w:r></w:p>"));
    p0 = body.xpath(".//w:p"); p0 = p0(1);
    para = mat2doc.text.Paragraph(p0, d);
    newp = para.insert_paragraph_before("intro");
    ps = body.xpath("./w:p");
    s = struct();
    s.new_text = string(newp.text);
    s.child_count = numel(ps);
    s.first_text = string(mat2doc.text.Paragraph(ps(1), d).text);
    s.second_text = string(mat2doc.text.Paragraph(ps(2), d).text);
    s.body_serhex = hx_e(body);
    P.insert_before = s;

    body = rparse(rbodyXml("<w:p><w:r><w:t>orig</w:t></w:r></w:p>"));
    p0 = body.xpath(".//w:p"); p0 = p0(1);
    newp = mat2doc.text.Paragraph(p0, d).insert_paragraph_before();
    s = struct();
    s.new_text = string(newp.text);
    ps = body.xpath("./w:p"); s.child_count = numel(ps);
    P.insert_before_empty = s;

    % ===================== 8. Hyperlink full surface (REAL rels) ===========
    rid_ext = d.part().relate_to("https://google.com/", ...
        mat2doc.opc.RELATIONSHIP_TYPE.HYPERLINK, true);
    hl = struct();

    h = rhlink("r:id=""" + rid_ext + """", "<w:r><w:t>post</w:t></w:r>", d);
    s = struct();
    s.address = string(h.address); s.fragment = string(h.fragment); s.url = string(h.url);
    s.text = string(h.text); s.runs = arrInfo(h.runs);
    s.contains_page_break = tf(h.contains_page_break);
    hl.external = s;

    h = rhlink("r:id=""" + rid_ext + """ w:anchor=""foo""", "<w:r><w:t>post</w:t></w:r>", d);
    s = struct();
    s.address = string(h.address); s.fragment = string(h.fragment);
    s.url = string(h.url); s.text = string(h.text);
    hl.external_fragment = s;

    h = rhlink("w:anchor=""_Toc147925734""", "<w:r><w:t>Heading</w:t></w:r>", d);
    s = struct();
    s.address = string(h.address); s.fragment = string(h.fragment);
    s.url = string(h.url); s.text = string(h.text);
    hl.internal = s;

    h = rhlink("", "", d);
    s = struct();
    s.address = string(h.address); s.fragment = string(h.fragment); s.url = string(h.url);
    s.text = string(h.text); s.runs = arrInfo(h.runs);
    s.contains_page_break = tf(h.contains_page_break);
    hl.bare = s;

    h = rhlink("", "<w:r><w:t>a</w:t><w:lastRenderedPageBreak/><w:t>b</w:t></w:r>", d);
    s = struct();
    s.text = string(h.text); s.contains_page_break = tf(h.contains_page_break);
    hl.with_break = s;
    P.hyperlink = hl;

    % ===================== 9. RenderedPageBreak fragment split =============
    pb = struct();
    pb.leading = fragInfo( ...
        "<w:pPr><w:ind/></w:pPr><w:r><w:lastRenderedPageBreak/>" + ...
        "<w:t>foo</w:t><w:t>bar</w:t></w:r>", 1, d);
    pb.mid = fragInfo( ...
        "<w:pPr><w:ind/></w:pPr>" + ...
        "<w:r><w:t>foo</w:t><w:lastRenderedPageBreak/><w:t>bar</w:t></w:r>" + ...
        "<w:r><w:t>barfoo</w:t></w:r>", 1, d);
    pb.in_hyperlink = fragInfo( ...
        "<w:pPr><w:ind/></w:pPr>" + ...
        "<w:hyperlink><w:r><w:t>foo</w:t><w:lastRenderedPageBreak/>" + ...
        "<w:t>bar</w:t></w:r></w:hyperlink><w:r><w:t>barfoo</w:t></w:r>", 1, d);
    pb.trailing = fragInfo( ...
        "<w:r><w:t>foo</w:t><w:lastRenderedPageBreak/></w:r>", 1, d);
    P.pagebreak = pb;

    p = rparse(rpXml("<w:r><w:t>abc</w:t><w:lastRenderedPageBreak/><w:lastRenderedPageBreak/></w:r>"));
    lrpbs = p.xpath(".//w:lastRenderedPageBreak");
    rpb2 = mat2doc.text.RenderedPageBreak(lrpbs(2), d);
    g = struct();
    try
        rpb2.preceding_paragraph_fragment; %#ok<VUNUS>
        g.raised = "False"; g.msg = "";
    catch ME
        g.raised = "True"; g.msg = string(ME.message);
    end
    P.pagebreak_guard = g;
end

% -------- runProbes-local XML/probe helpers (mirror the s0028 .m twin) --------

function x = rpXml(inner)
    x = "<w:p " + nsWR() + ">" + string(inner) + "</w:p>";
end

function x = rbodyXml(inner)
    x = "<w:body " + nsWR() + ">" + string(inner) + "</w:body>";
end

function x = rhlinkXml(attrs, inner)
    a = string(attrs);
    if strlength(a) > 0, a = " " + a; end
    x = "<w:hyperlink " + nsWR() + a + ">" + string(inner) + "</w:hyperlink>";
end

function h = rhlink(attrs, inner, d)
    h = mat2doc.text.Hyperlink(rparse(rhlinkXml(attrs, inner)), d);
end

function out = arrInfo(arr)
    % Homogeneous object-array surface -> {count, texts, first, last} (rv text).
    n = numel(arr);
    texts = cell(1, n);
    for k = 1:n
        texts{k} = string(arr(k).text);
    end
    out = struct();
    out.count = n;
    out.texts = texts;                              % empty cell -> []
    if n == 0
        out.first = "None"; out.last = "None";
    else
        out.first = string(arr(1).text);
        out.last  = string(arr(n).text);
    end
end

function L = innerList(items)
    % iter_inner_content cell -> ORDERED list of {type, text}.
    n = numel(items);
    L = cell(1, n);
    for k = 1:n
        it = items{k};
        if isa(it, "mat2doc.text.Run"), ty = "Run"; else, ty = "Hyperlink"; end
        s = struct(); s.type = ty; s.text = string(it.text);
        L{k} = s;
    end
end

function C = lnsCell(e)
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

function out = fragInfo(pInner, breakIndex, d)
    p = rparse(rpXml(pInner));
    lrpbs = p.xpath(".//w:lastRenderedPageBreak");
    lrpb = lrpbs(breakIndex);
    rpb = mat2doc.text.RenderedPageBreak(lrpb, d);
    pre = rpb.preceding_paragraph_fragment;
    fol = rpb.following_paragraph_fragment;
    out = struct();
    out.preceding_is_none = tf(isequal(pre, []));
    out.following_is_none = tf(isequal(fol, []));
    if isequal(pre, []), out.preceding_text = "None"; else, out.preceding_text = string(pre.text); end
    if isequal(fol, []), out.following_text = "None"; else, out.following_text = string(fol.text); end
    if isequal(pre, [])
        out.preceding_serhex = "None";
        out.preceding_localnames = "None";
    else
        e = lrpb.preceding_fragment_p();
        out.preceding_serhex = hx_e(e);
        out.preceding_localnames = lnsCell(e);
    end
    if isequal(fol, [])
        out.following_serhex = "None";
        out.following_localnames = "None";
    else
        e = lrpb.following_fragment_p();
        out.following_serhex = hx_e(e);
        out.following_localnames = lnsCell(e);
    end
end

% -------- flatten (robust: struct arrays + cell-of-structs + nesting) --------

function flattenLeaves(s, prefix, map)
    % Recursively flatten a (possibly nested) struct / struct-array / cell / array
    % into map(dotted.path) -> canonical string, so a MATLAB-built probe struct and
    % the jsondecode'd oracle compare leaf-by-leaf regardless of container types.
    % Handles the s0028 shapes: scalar structs, struct ARRAYS (oracle "inner"),
    % cells-of-structs (port "inner"), cells-of-strings ("texts"/"localnames"),
    % empty containers ([] / 0x0 -> ''), and scalar string/num/bool leaves.
    if isstruct(s)
        if numel(s) == 1
            fn = fieldnames(s);
            for i = 1:numel(fn)
                flattenLeaves(s.(fn{i}), keyOf(prefix, fn{i}), map);
            end
        else
            for j = 1:numel(s)
                flattenLeaves(s(j), sprintf('%s[%d]', prefix, j), map);
            end
        end
    elseif iscell(s)
        if isempty(s)
            map(prefix) = ''; %#ok<NASGU>
        else
            for j = 1:numel(s)
                flattenLeaves(s{j}, sprintf('%s[%d]', prefix, j), map);
            end
        end
    elseif ischar(s)
        map(prefix) = s; %#ok<NASGU>
    elseif isstring(s) && isscalar(s)
        map(prefix) = char(s); %#ok<NASGU>
    elseif isstring(s)
        if isempty(s)
            map(prefix) = ''; %#ok<NASGU>
        else
            for j = 1:numel(s)
                map(sprintf('%s[%d]', prefix, j)) = char(s(j)); %#ok<NASGU>
            end
        end
    elseif (isnumeric(s) || islogical(s)) && isscalar(s)
        map(prefix) = char(canonScalar(s)); %#ok<NASGU>
    elseif isnumeric(s) || islogical(s)
        if isempty(s)
            map(prefix) = ''; %#ok<NASGU>
        else
            for j = 1:numel(s)
                map(sprintf('%s[%d]', prefix, j)) = char(canonScalar(s(j))); %#ok<NASGU>
            end
        end
    else
        map(prefix) = char(canonScalar(s)); %#ok<NASGU>
    end
end

function k = keyOf(prefix, name)
    if isempty(prefix), k = name; else, k = [prefix '.' name]; end
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

% -------- zip / byte helpers (file-local, copied from Test_p4_5a) ----

function [blobs, names] = zipEntryList(zipBytes)
    % Enumerate an in-memory zip in stream (write) order -> ordered {names, blobs}.
    bais = java.io.ByteArrayInputStream(int8(typecast(uint8(zipBytes(:)'), 'int8')));
    zis  = java.util.zip.ZipInputStream(bais);
    cleanup = onCleanup(@() zis.close()); %#ok<NASGU>
    copier = com.mathworks.mlwidgets.io.InterruptibleStreamCopier.getInterruptibleStreamCopier;
    names = strings(1, 0);
    blobs = {};
    while true
        ze = zis.getNextEntry();
        if isempty(ze)
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

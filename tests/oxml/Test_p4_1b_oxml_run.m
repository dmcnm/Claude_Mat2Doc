classdef Test_p4_1b_oxml_run < matlab.unittest.TestCase
% TEST_P4_1B_OXML_RUN  Gate-4 permanent unit tests for Mat2Doc P4-1b
%   (src/docx/oxml/text/run.py -> +mat2doc\+oxml\+text\CT_R + the run
%   inner-content classes CT_Br/CT_Cr/CT_NoBreakHyphen/CT_PTab/CT_Text and the
%   module-private _RunContentAppender -> RunContentAppender_).
%
%   P4-1b is the M2 byte-critical run surface: document.xml body content is
%   w:p/w:r/w:t. This class permanently freezes the guarantees the three prior
%   gates established (audit + engine-fix re-audit APPROVE; Gate-3 confirmation
%   PASS after the XmlParser.m fix):
%
%     (1) add_t xml:space="preserve" battery INCLUDING the AUD-1 non-ASCII
%         regression pin. Python emits xml:space="preserve" iff
%         len(text.strip()) < len(text). CPython str.strip() removes Unicode
%         whitespace that MATLAB strip() KEEPS -- U+0085 (NEL), U+00A0 (NBSP),
%         U+2007 (figure space), U+202F (narrow NBSP). Porting add_t with the
%         native strip() dropped xml:space for NBSP/NEL-boundary text (a real
%         byte divergence, Gate-2 audit probes xmlspace_10/11/13). The fix is the
%         CT_R.PY_STRIP_WS 29-code-point constant. A revert to MATLAB strip()
%         goes RED here (test_xmlspace_* + test_py_strip_ws_membership_*).
%
%     (2) The XmlParser.m engine fix (a03 defect-class). Registering w:r->CT_R
%         moved the a03 fixture's child (q:r == {wURI}r, reached via an
%         ANCESTOR-declared non-fixed prefix q) from the parser's generic-element
%         path onto the registered-class path, exposing a latent asymmetry: the
%         registered branch did not forward the parser-resolved URI, so the ctor
%         re-resolved prefix q (not in the element's own decls, not in the fixed
%         nsmap) and threw mat2doc:KeyError 'q'. The fix (XmlParser.m ~line 229)
%         forwards the resolved URI into feval(cls, name, ownDecls, uri) exactly
%         as the generic fallback already did. This class pins the PARSE-CLEAN
%         property (a registered element reached via an ancestor-declared
%         non-fixed prefix parses without KeyError) so a revert goes RED. NOTE:
%         the a03 SERIALIZED deviation bytes are already pinned by
%         Test_p1_2_oxml's a03 tests -- NOT duplicated here.
%
%     (3) _RunContentAppender char mapping (\t->w:tab, \r|\n->w:br, else
%         accumulate w:t with xml:space when boundary-ws), H11 rPr-first via the
%         _insert_rPr override, CT_Text/CT_Br/CT_Cr/CT_NoBreakHyphen/CT_PTab
%         __str__, and the M1 document.xml (1548 B) + styles.xml (349458 B) byte
%         pins (the run-tag registration AND the engine fix both touch the parse
%         path; the full 17/17 sweep stays owned by Test_p1_8).
%
%   Provenance (Gate-1..3, all 2026-07-26/27):
%     * Audit    : validation\mat2doc\audit_P4-1b_oxml_run.md (Porter Gate-1 +
%                  Gate-2 APPROVE-WITH-FIXES [AUD-1 NBSP-strip, fixed inline] +
%                  the XmlParser engine-fix section + engine-fix re-audit APPROVE).
%     * Validate : validation\mat2doc\validate_P4-1b_oxml_run.md (Gate-3 FAIL
%                  diagnosis of the a03 regression -> engine fix -> CONFIRMATION
%                  RE-RUN PASS: regression 550/550, M1 17/17, a03 deviation bytes
%                  unchanged, s0021 surface probe_diff MATCH; 0 new D-numbers).
%     * Scenario : validation\mat2doc\scenarios\s0021_p4_1b_oxml_run.{py,m}
%                  (the probe sequence replayed VERBATIM by runProbes() below).
%     * Frozen ref (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE):
%         references\s0021\probe.json -- copied verbatim (self-contained) into
%           tests\oxml\data\s0021_probe_oracle.json (the Equivalence replay set;
%           value/serhex JSON, jsondecode is line-ending agnostic -> no
%           `* binary` .gitattributes needed, per the s0020 precedent).
%         references\s0001\parts\word\{document,styles}.xml -- the M1 byte
%           references; NOT copied (the byte pins compare SHA-256 of what
%           mat2doc.Document().save() itself emits, so no fixture is needed).
%
%   Coverage taxonomy
%   -----------------
%   * Nominal      -- add_t documented xml:space happy path; _RunContentAppender
%                     "a\tb\nc"; H11 rPr-first; CT_Text "abc".
%   * Edge         -- empty text ("" -> no children / no xml:space), non-ASCII
%                     boundary whitespace (NBSP/NEL/figure-space/narrow-NBSP/
%                     line-sep), astral char (U+1F600, surrogate pair), CR/LF
%                     only, single-element, the a03 error-CLASS (registered
%                     element via ancestor-declared non-fixed prefix -> must NOT
%                     KeyError).
%   * Equivalence  -- test_equivalence_full_battery_vs_frozen_oracle replays the
%                     ENTIRE s0021 battery live and flatten-compares every leaf to
%                     the frozen python-docx 1.2.0 oracle (0 divergences).
%   * Regression   -- hard-coded expected serialized-XML strings / serhex fragments
%                     (all ASCII where compared as strings, so string-equality is a
%                     byte-identical L1 assertion; UTF-8 serhex where the bytes are
%                     non-ASCII) + SHA-256 of the two M1 parts.
%   * Upstream     -- the add_t xml:space rule and the run-content __str__ mappings
%                     are the python-docx run.py surface; the frozen oracle IS the
%                     lxml expected output for this API sequence.
%
%   Byte-level (L1) note: every serialized-XML comparison is either the FULL
%   serialize_part_xml output as UPPERCASE hex of the raw UTF-8 shipping bytes
%   (serhex -- masks nothing, incl. the non-ASCII NBSP/astral bytes) or an ASCII
%   decoded string (string-equality == byte-equality). No D-number granted any L2
%   relaxation in this WP (Gate-3 §R.5: zero new, none exercised at L2), so every
%   pin here is L1. The `nRpr > 100` guard and the equivalence key-count guard are
%   the only looser-than-byte checks and are commented at their site.
%
%   Determinism: no network, no absolute paths. The worktree root and the
%   co-located oracle resolve relative to this file via
%   fileparts(mfilename('fullpath')); saves go to tempname .docx deleted via
%   onCleanup; every file read is binary ('r','n'); no 'wt'. Non-ASCII inputs are
%   built via char(<codepoint>) for source-encoding independence. The +mat2doc
%   package resolves via the MANDATORY PathFixture(worktree-root) in
%   TestClassSetup (WP9-F4 lesson).

    properties (Constant)
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        % --- frozen s0001 M1 byte references (the P4-1b parse-path risk) ---
        DOC_SIZE    = 1548
        DOC_SHA     = "0e4dd503bc095219cd24e8312b372c588ca56bb834aacac53eafb7bf47836327"
        STYLES_SIZE = 349458
        STYLES_SHA  = "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"

        % --- the EXACT CPython str.strip() set (CT_R.PY_STRIP_WS), 29 BMP code
        %     points (Gate-2 audit enumeration). Duplicated here as an INDEPENDENT
        %     regression sentinel: a drift in the port's constant OR a revert to
        %     MATLAB strip() both diverge from this. ---
        PY_STRIP_WS_CODES = [9 10 11 12 13 28 29 30 31 32 133 160 5760 ...
            8192 8193 8194 8195 8196 8197 8198 8199 8200 8201 8202 ...
            8232 8233 8239 8287 12288]

        % --- the four AUD-1 code points MATLAB strip() KEEPS but CPython removes
        %     (audit probes xmlspace_10/11/13): NEL, NBSP, figure-space, narrow-NBSP ---
        AUD1_MATLAB_KEEPS = [133 160 8199 8239]
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied verbatim from tests\oxml\Test_p4_1a_oxml_font.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\oxml
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. add_t xml:space battery incl. the AUD-1 NBSP regression pin   %
        % =============================================================== %

        function test_xmlspace_battery_vs_frozen_oracle(testCase)
            % Nominal + Edge + Regression (BYTE-CRITICAL, s0021 xmlspace): replay
            % add_t over the full 18-case whitespace battery and assert the
            % serialized <w:r> bytes (serhex) EXACTLY equal the frozen python-docx
            % oracle, together with has_xmlspace and has_xmlns_xml. The 5 non-ASCII
            % boundary cases (NBSP/NEL/figure-space/narrow-NBSP/line-sep) are the
            % exact bytes that diverged pre-AUD-1-fix; an astral U+1F600 boundary
            % emits NO xml:space. NO case declares xmlns:xml (reserved prefix).
            % Inputs built via char(<codepoint>) (source-encoding independent).
            oracle = loadOracle();
            XS = xmlspaceBattery();
            for k = 1:size(XS, 1)
                name = XS{k, 1};
                r = mat2doc.oxml.OxmlElement("w:r");
                r.add_t(XS{k, 2});
                raw = mat2doc.oxml.serialize_part_xml(r);
                s   = string(native2unicode(raw, "UTF-8"));
                exp = oracle.xmlspace.(name);
                testCase.verifyEqual(hx(raw), string(exp.serhex), ...
                    sprintf('xmlspace.%s serialized bytes (L1) must match the frozen oracle', name));
                testCase.verifyEqual(ternary(contains(s, 'xml:space="preserve"')), ...
                    string(exp.has_xmlspace), sprintf('xmlspace.%s has_xmlspace', name));
                testCase.verifyEqual(ternary(contains(s, "xmlns:xml")), ...
                    string(exp.has_xmlns_xml), ...
                    sprintf('xmlspace.%s must NOT declare xmlns:xml (reserved prefix)', name));
            end
        end

        function test_xmlspace_aud1_nonascii_boundary_regression(testCase)
            % Edge + Regression (AUD-1, the load-bearing pin): each of the 5 Unicode
            % whitespace chars that CPython str.strip() removes but MATLAB strip()
            % handles differently, placed at a text boundary, MUST make add_t emit
            % xml:space="preserve" (with NO xmlns:xml). An astral char (U+1F600) at
            % a boundary with no other whitespace must NOT. Built via char() codes.
            % Hard-coded expectations -- a regression to native strip() goes RED.
            wsCodes = [160 133 8199 8239 8232];   % NBSP NEL FIGSP NNBSP LSEP
            wsNames = ["NBSP(U+00A0)" "NEL(U+0085)" "FIGSP(U+2007)" ...
                       "NNBSP(U+202F)" "LSEP(U+2028)"];
            for j = 1:numel(wsCodes)
                lead = string(char(wsCodes(j))) + "x";     % boundary at the front
                s = serRun(addTRun(lead));
                testCase.verifyTrue(contains(s, 'xml:space="preserve"'), ...
                    sprintf('leading %s must emit xml:space="preserve"', wsNames(j)));
                testCase.verifyFalse(contains(s, "xmlns:xml"), ...
                    sprintf('leading %s must NOT declare xmlns:xml', wsNames(j)));

                trail = "x" + string(char(wsCodes(j)));    % boundary at the end
                s2 = serRun(addTRun(trail));
                testCase.verifyTrue(contains(s2, 'xml:space="preserve"'), ...
                    sprintf('trailing %s must emit xml:space="preserve"', wsNames(j)));
            end

            % astral U+1F600 (surrogate pair) with NO boundary whitespace -> no preserve.
            grin = string(char([55357 56832]));
            sg = serRun(addTRun(grin));
            testCase.verifyFalse(contains(sg, 'xml:space="preserve"'), ...
                'a lone astral char (U+1F600) has no boundary whitespace -> no xml:space');
            % but astral + trailing space -> preserve (the whitespace, not the astral, drives it).
            sgw = serRun(addTRun(grin + " "));
            testCase.verifyTrue(contains(sgw, 'xml:space="preserve"'), ...
                'astral + trailing space -> xml:space="preserve"');
        end

        function test_py_strip_ws_membership_pins_aud1_fix(testCase)
            % Regression (AUD-1 constant): the port's CT_R.PY_STRIP_WS must equal
            % the exact 29-code-point CPython strip set, and the four AUD-1 code
            % points must be members that MATLAB strip() does NOT remove -- proving
            % the port cannot silently fall back to native strip() without going RED.
            got = double(mat2doc.oxml.text.CT_R.PY_STRIP_WS);
            testCase.verifyEqual(sort(got), sort(testCase.PY_STRIP_WS_CODES), ...
                'CT_R.PY_STRIP_WS must be the exact CPython str.strip() code-point set');
            testCase.verifyEqual(numel(got), 29, 'PY_STRIP_WS must hold exactly 29 code points');

            for c = testCase.AUD1_MATLAB_KEEPS
                testCase.verifyTrue(any(got == c), ...
                    sprintf('U+%04X must be a member of PY_STRIP_WS', c));
                boundaried = string(char(c)) + "x";
                testCase.verifyEqual(strlength(strip(boundaried)), strlength(boundaried), ...
                    sprintf(['MATLAB strip() must KEEP U+%04X (documents why the port ' ...
                             'uses PY_STRIP_WS, not native strip)'], c));
            end
        end

        % =============================================================== %
        % 2. XmlParser engine fix -- registered element via ancestor prefix %
        % =============================================================== %

        function test_enginefix_a03_registered_via_ancestor_prefix_parses_clean(testCase)
            % Edge + Regression (engine fix, the a03 defect-CLASS): parse
            %   <w:p xmlns:w=W xmlns:q=W><q:r q:val="1"/></w:p>
            % where q is a SECOND prefix bound to the SAME URI as w on the PARENT,
            % so q:r == {W}r reaches the REGISTERED (CT_R) parse path via an
            % ancestor-declared, non-fixed prefix. Before the XmlParser.m fix this
            % threw mat2doc:KeyError 'q' (the ctor re-resolved q). The pin: parse is
            % clean AND the child is CT_R with the ancestor-resolved {W} URI. A
            % revert of XmlParser.m's registered branch to 2-arg feval goes RED.
            % (The SERIALIZED a03 deviation bytes stay pinned by Test_p1_2_oxml.)
            p = rparse("<w:p xmlns:w=""" + testCase.W + """ xmlns:q=""" + testCase.W + ...
                       """><q:r q:val=""1""/></w:p>");
            child = firstEl(p.xpath("./*"));
            testCase.verifyEqual(class(child), 'mat2doc.oxml.text.CT_R', ...
                'q:r (== {W}r via ancestor prefix q) must parse to CT_R');
            testCase.verifyEqual(child.nsuri, testCase.W, ...
                'the CT_R child must carry the ancestor-resolved {W} URI (no KeyError)');
            testCase.verifyEqual(string(child.local_part), "r");
        end

        function test_enginefix_multilevel_ancestor_prefix_parses_clean(testCase)
            % Edge + Regression (engine fix, multi-level defect-class closure):
            %   <w:p xmlns:w=W xmlns:z=W><z:r><z:rPr><z:b/></z:rPr><z:t>hi</z:t></z:r></w:p>
            % z bound to the w URI two levels up; EVERY registered descendant reached
            % via z must resolve cleanly to its CT_* class with the {W} URI:
            % z:r->CT_R, z:rPr->CT_RPr, z:b->CT_OnOff, z:t->CT_Text. Pins that the
            % fix closes the WHOLE ancestor-declared-non-fixed-prefix class.
            p = rparse("<w:p xmlns:w=""" + testCase.W + """ xmlns:z=""" + testCase.W + ...
                       """><z:r><z:rPr><z:b/></z:rPr><z:t>hi</z:t></z:r></w:p>");
            r   = firstEl(p.xpath("./w:r"));
            rPr = firstEl(r.xpath("./w:rPr"));
            b   = firstEl(rPr.xpath("./w:b"));
            t   = firstEl(r.xpath("./w:t"));
            testCase.verifyEqual(class(r),   'mat2doc.oxml.text.CT_R',        'z:r -> CT_R');
            testCase.verifyEqual(class(rPr), 'mat2doc.oxml.text.CT_RPr',      'z:rPr -> CT_RPr');
            testCase.verifyEqual(class(b),   'mat2doc.oxml.shared.CT_OnOff',  'z:b -> CT_OnOff');
            testCase.verifyEqual(class(t),   'mat2doc.oxml.text.CT_Text',     'z:t -> CT_Text');
            testCase.verifyEqual([r.nsuri rPr.nsuri b.nsuri t.nsuri], ...
                repmat(testCase.W, 1, 4), 'every descendant must carry the {W} URI (no KeyError)');
            testCase.verifyEqual(t.str_(), "hi", 'the z:t CT_Text must round-trip its char data');
        end

        % =============================================================== %
        % 3. _RunContentAppender char mapping (the run-content FSM)        %
        % =============================================================== %

        function test_appender_char_mapping_vs_oracle(testCase)
            % Nominal + Edge + Regression (byte-critical, s0021 appender): set
            % r.text (the _RunContentAppender FSM) over 7 cases and assert the child
            % localname sequence AND the serialized bytes vs the frozen oracle:
            % \t->w:tab, \r|\n->w:br, else accumulate w:t (with xml:space when
            % boundary-ws). Covers empty ("" -> no children), astral surrogate pair,
            % tab/nl only, CRLF.
            oracle = loadOracle();
            AP = appenderBattery();
            for k = 1:size(AP, 1)
                name = AP{k, 1};
                r = mat2doc.oxml.OxmlElement("w:r");
                r.text = AP{k, 2};
                raw = mat2doc.oxml.serialize_part_xml(r);
                exp = oracle.appender.(name);
                testCase.verifyEqual(lnsCell(r), asCellRow(exp.localnames), ...
                    sprintf('appender.%s child localname sequence', name));
                testCase.verifyEqual(hx(raw), string(exp.serhex), ...
                    sprintf('appender.%s serialized bytes (L1) must match the frozen oracle', name));
            end
        end

        function test_appender_nominal_hardcoded(testCase)
            % Regression (hard-coded): the documented "a\tb\nc" mapping, pinned to a
            % hard-coded expected serialized string (ASCII -> string-equality is L1)
            % independently of the oracle file -- guards the oracle itself.
            r = mat2doc.oxml.OxmlElement("w:r");
            r.text = "a" + string(char(9)) + "b" + string(char(10)) + "c";
            testCase.verifyEqual(serRun(r), decl() + newline + ...
                "<w:r xmlns:w=""" + testCase.W + """>" + ...
                "<w:t>a</w:t><w:tab/><w:t>b</w:t><w:br/><w:t>c</w:t></w:r>", ...
                '"a\tb\nc" -> t,tab,t,br,t (\t->w:tab, \n->w:br)');
        end

        % =============================================================== %
        % 4. H11 rPr-first (_insert_rPr override)                          %
        % =============================================================== %

        function test_h11_rpr_first_ordering(testCase)
            % Nominal + Regression (H11, s0021 h11): get_or_add_rPr on a run that
            % already has content puts rPr FIRST (the _insert_rPr override,
            % insert(1,...) <-> run.py:146 insert(0,...)); content descriptors have
            % successors=() so they append in insertion order. Covers content-then-
            % rPr, interleaved (t,rPr,t), the style-setter route, and get_or_add
            % idempotency (H5). Localnames + bytes vs the frozen oracle; the primary
            % case also pinned hard-coded.
            oracle = loadOracle();

            r = mat2doc.oxml.OxmlElement("w:r"); r.add_t("x"); r.get_or_add_rPr();
            testCase.verifyEqual(lnsCell(r), {'rPr','t'}, 'content-then-rPr -> [rPr, t]');
            testCase.verifyEqual(serRun(r), decl() + newline + ...
                "<w:r xmlns:w=""" + testCase.W + """><w:rPr/><w:t>x</w:t></w:r>", ...
                'content-then-rPr hard-coded serialized bytes (L1)');
            testCase.verifyEqual(hx_e(r), string(oracle.h11.content_then_rpr.serhex));

            r = mat2doc.oxml.OxmlElement("w:r"); r.add_t("a"); r.get_or_add_rPr(); r.add_t("b");
            testCase.verifyEqual(lnsCell(r), {'rPr','t','t'}, 'interleaved -> [rPr, t, t]');
            testCase.verifyEqual(hx_e(r), string(oracle.h11.interleaved.serhex));

            r = mat2doc.oxml.OxmlElement("w:r"); r.add_t("x"); r.style = "Emphasis";
            testCase.verifyEqual(lnsCell(r), {'rPr','t'}, 'style-setter route -> [rPr, t]');
            testCase.verifyEqual(r.style, "Emphasis");
            testCase.verifyEqual(hx_e(r), string(oracle.h11.style_route.serhex));

            r = mat2doc.oxml.OxmlElement("w:r"); a = r.get_or_add_rPr(); b = r.get_or_add_rPr();
            testCase.verifyTrue(a == b, 'get_or_add_rPr must be idempotent (same live handle, H5)');
        end

        % =============================================================== %
        % 5. CT_Text / run-content __str__ + CT_Br tri-state              %
        % =============================================================== %

        function test_ct_text_str_h3(testCase)
            % Nominal + Edge (H3/H4, s0021 ct_text_str): CT_Text.__str__ ->
            % self.text or "" -- empty w:t yields "" (never None/[]), "abc" -> "abc",
            % " a " preserves spaces.
            t = mat2doc.oxml.OxmlElement("w:t");
            testCase.verifyEqual(t.str_(), "", 'empty w:t __str__ -> "" (never None)');
            t = mat2doc.oxml.OxmlElement("w:t"); t.text = "abc";
            testCase.verifyEqual(t.str_(), "abc");
            t = mat2doc.oxml.OxmlElement("w:t"); t.text = " a ";
            testCase.verifyEqual(t.str_(), " a ", 'CT_Text __str__ preserves surrounding spaces');
        end

        function test_run_content_str_and_br_tristate(testCase)
            % Regression (s0021 run_content_str): CT_Br/CT_Cr/CT_NoBreakHyphen/
            % CT_PTab __str__ and CT_Br @type/@clear attr round-trips.
            %   CT_Br: default & type="textWrapping" -> "\n"; page/column -> "";
            %          absent @type reads "textWrapping" (non-None default);
            %          absent @clear reads [] (-> "none").
            %   CT_Cr -> "\n"; CT_NoBreakHyphen -> "-"; CT_PTab -> "\t".
            LF = string(char(10)); TAB = string(char(9));

            br = mat2doc.oxml.OxmlElement("w:br");
            testCase.verifyEqual(br.str_(), LF, 'CT_Br default __str__ -> "\n"');
            testCase.verifyEqual(br.type, "textWrapping", 'CT_Br absent @type -> "textWrapping"');
            testCase.verifyTrue(isequal(br.clear, []), 'CT_Br absent @clear -> [] (-> "none")');

            br = mat2doc.oxml.OxmlElement("w:br"); br.type = "textWrapping";
            testCase.verifyEqual(br.str_(), LF);
            br = mat2doc.oxml.OxmlElement("w:br"); br.type = "page";
            testCase.verifyEqual(br.str_(), "", 'CT_Br page -> ""');
            br = mat2doc.oxml.OxmlElement("w:br"); br.type = "column";
            testCase.verifyEqual(br.str_(), "", 'CT_Br column -> ""');

            cr = mat2doc.oxml.OxmlElement("w:cr");
            testCase.verifyEqual(cr.str_(), LF, 'CT_Cr __str__ -> "\n"');
            nbh = mat2doc.oxml.OxmlElement("w:noBreakHyphen");
            testCase.verifyEqual(nbh.str_(), "-", 'CT_NoBreakHyphen __str__ -> "-"');
            pt = mat2doc.oxml.OxmlElement("w:ptab");
            testCase.verifyEqual(pt.str_(), TAB, 'CT_PTab __str__ -> "\t"');
        end

        function test_ct_r_text_getter_mixed_content(testCase)
            % Regression (s0021 ct_r_text): CT_R.text getter joins str_() of the
            % inner-content children in document order. Over
            % <w:t>x</w:t><w:br/><w:noBreakHyphen/><w:ptab/><w:cr/> -> "x\n-\t\n".
            r = rparse("<w:r xmlns:w=""" + testCase.W + """>" + ...
                "<w:t>x</w:t><w:br/><w:noBreakHyphen/><w:ptab/><w:cr/></w:r>");
            testCase.verifyEqual(r.text, ...
                "x" + string(char(10)) + "-" + string(char(9)) + string(char(10)), ...
                'CT_R.text getter over mixed content -> "x\n-\t\n"');
        end

        % =============================================================== %
        % 6. comment range helpers                                        %
        % =============================================================== %

        function test_comment_range_helpers(testCase)
            % Regression (s0021 comments): insert_comment_range_start_above +
            % insert_comment_range_end_and_reference_below on a run under a <w:p>
            % produce sibling order [commentRangeStart, r, commentRangeEnd, r], with
            % the reference-run styled CommentReference and w:id via pyStr. Localnames
            % + reference style + serialized bytes vs the frozen oracle.
            oracle = loadOracle();
            p = mat2doc.oxml.OxmlElement("w:p");
            r = mat2doc.oxml.OxmlElement("w:r");
            p.append(r);
            r.insert_comment_range_start_above(0);
            r.insert_comment_range_end_and_reference_below(0);
            testCase.verifyEqual(lnsCell(p), ...
                {'commentRangeStart','r','commentRangeEnd','r'}, ...
                'comment helpers sibling order');
            kids = p.xpath("./*");
            testCase.verifyEqual(kids(end).style, "CommentReference", ...
                'reference-run must be styled CommentReference');
            testCase.verifyEqual(hx_e(p), string(oracle.comments.serhex), ...
                'comment helpers serialized bytes (L1) must match the frozen oracle');
        end

        % =============================================================== %
        % 7. M1 byte-neutrality (parse-path risk: registration + engine fix) %
        % =============================================================== %

        function test_m1_document_xml_byte_identical(testCase)
            % Regression (byte-neutrality, L1): mat2doc.Document().save() emits
            % word/document.xml at EXACTLY 1548 B with the frozen s0001 SHA-256 --
            % byte-identical DESPITE the w:r->CT_R / w:t->CT_Text parse-class flip
            % AND the XmlParser engine fix (both touch this part's parse path).
            % SHA-256 equality is a byte-level (L1) assertion. (Test_p1_8 owns the
            % full 17/17 M1 sweep; this class pins the run-carrying part.)
            bytes = testCase.emitPart('document.xml');
            testCase.verifyEqual(numel(bytes), testCase.DOC_SIZE, ...
                sprintf('word/document.xml must be exactly %d B', testCase.DOC_SIZE));
            testCase.verifyEqual(sha256hex(bytes), testCase.DOC_SHA, ...
                'word/document.xml SHA-256 must equal the frozen s0001 oracle (byte-identical L1)');
        end

        function test_m1_styles_xml_byte_identical(testCase)
            % Regression (byte-neutrality, L1): word/styles.xml stays 349458 B with
            % the frozen s0001 SHA-256 -- the run-tag registration + engine fix
            % re-route its parse but must not move a byte.
            bytes = testCase.emitPart('styles.xml');
            testCase.verifyEqual(numel(bytes), testCase.STYLES_SIZE, ...
                sprintf('word/styles.xml must be exactly %d B', testCase.STYLES_SIZE));
            testCase.verifyEqual(sha256hex(bytes), testCase.STYLES_SHA, ...
                'word/styles.xml SHA-256 must equal the frozen s0001 oracle (byte-identical L1)');
        end

        function test_run_bearing_document_parse_serialize_idempotent_L1(testCase)
            % Regression (byte-neutrality, L1): the shipped default document.xml
            % template has an EMPTY body (zero runs), so it does not exercise the
            % CT_R parse path. Instead parse a REALISTIC run-bearing w:document
            % fragment (w:p/w:r with w:rPr>w:b, and a boundary-whitespace w:t) that
            % instantiates CT_R / CT_RPr / CT_OnOff / CT_Text through the
            % engine-fixed registered path, then assert serialize is a STABLE
            % FIXPOINT: serialize(parse(serialize(parse(X)))) == serialize(parse(X))
            % byte-for-byte. A parse/serialize asymmetry on the run surface goes RED.
            frag = "<w:document xmlns:w=""" + testCase.W + """><w:body><w:p><w:r>" + ...
                "<w:rPr><w:b/></w:rPr><w:t xml:space=""preserve""> hi </w:t>" + ...
                "</w:r></w:p></w:body></w:document>";
            root1 = rparse(frag);
            out1  = mat2doc.oxml.serialize_part_xml(root1);
            root2 = mat2doc.oxml.parse_xml(out1);
            out2  = mat2doc.oxml.serialize_part_xml(root2);
            testCase.verifyEqual(uint8(out2(:)'), uint8(out1(:)'), ...
                'run-bearing document parse->serialize must be a byte-identical fixpoint (L1)');
            % Confirm the run surface really was instantiated through the registered
            % CT_* classes (guards a vacuous pass), and that xml:space survives.
            r = firstEl(root1.xpath('.//w:r'));
            testCase.verifyEqual(class(r), 'mat2doc.oxml.text.CT_R', ...
                'the parsed run must be a CT_R (registered parse path)');
            testCase.verifyEqual(class(firstEl(r.xpath('./w:t'))), 'mat2doc.oxml.text.CT_Text');
            testCase.verifyTrue(contains(string(native2unicode(out1, "UTF-8")), ...
                'xml:space="preserve"'), 'the boundary-ws w:t must keep xml:space="preserve"');
        end

        % =============================================================== %
        % 8. EQUIVALENCE -- full s0021 battery vs the frozen oracle        %
        % =============================================================== %

        function test_equivalence_full_battery_vs_frozen_oracle(testCase)
            % Equivalence: replay the ENTIRE s0021 battery (runProbes -- the .m
            % twin's body VERBATIM: xmlspace / appender / h11 / ct_text_str /
            % run_content_str / ct_r_text / comments) and flatten-compare EVERY leaf
            % to the frozen python-docx 1.2.0 oracle copied into
            % data\s0021_probe_oracle.json. Gate-3 found ZERO divergences, so every
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
            testCase.verifyGreaterThan(numel(oKeys), 60, ...
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
            % both cleaned up on exit. (Idiom from Test_p4_1a_oxml_font.m; tempname
            % paths are absolute so no cwd handling is needed.)
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

function r = addTRun(text)
    % A fresh <w:r> with add_t(text) applied.
    r = mat2doc.oxml.OxmlElement("w:r");
    r.add_t(text);
end

function s = serRun(e)
    % Decode serialize_part_xml bytes to a string (ASCII content -> string-equality
    % is a byte-identical L1 assertion; used only where the content is ASCII or the
    % assertion is a contains()).
    s = string(native2unicode(mat2doc.oxml.serialize_part_xml(e), "UTF-8"));
end

function h = hx(raw)
    % UPPERCASE hex of raw UTF-8 bytes (matches Python bytes.hex().upper()).
    h = string(sprintf('%02X', uint8(raw)));
end

function h = hx_e(e)
    h = hx(mat2doc.oxml.serialize_part_xml(e));
end

function e = firstEl(arr)
    % First element of an xpath result array (guarded).
    assert(~isempty(arr), 'xpath returned no elements');
    e = arr(1);
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

function C = asCellRow(v)
    % Normalize a jsondecode'd localnames value (cell column / empty / char) into a
    % 1xN cell row of char, matching lnsCell's shape.
    if isempty(v)
        C = cell(1, 0);
    elseif iscell(v)
        C = reshape(cellfun(@char, v, 'UniformOutput', false), 1, []);
    else
        C = {char(v)};   % a single JSON string decodes to char
    end
end

function s = ternary(tf)
    if tf, s = "true"; else, s = "false"; end
end

function s = tri(v)
    if isequal(v, []), s = "none";
    elseif v, s = "true";
    else, s = "false"; end
end

function o = loadOracle()
    % Read the co-located frozen oracle in BINARY mode (no CRLF translation) and
    % decode UTF-8 -> struct. jsondecode is whitespace/line-ending agnostic, so no
    % `* binary` .gitattributes pin is needed (value-based fixture, s0020 precedent).
    here = fileparts(mfilename('fullpath'));
    p = fullfile(here, 'data', 's0021_probe_oracle.json');
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

function XS = xmlspaceBattery()
    % The 18-case add_t whitespace battery (mirrors s0021 XS), inputs built via
    % char() codepoints for source-encoding independence.
    NBSP  = string(char(160));    NEL   = string(char(133));
    FIGSP = string(char(8199));   NNBSP = string(char(8239));
    LSEP  = string(char(8232));   GRIN  = string(char([55357 56832]));
    LF    = string(char(10));     TAB   = string(char(9));
    XS = { ...
        "sp_lead",          " x"; ...
        "sp_trail",         "x "; ...
        "sp_both",          " x "; ...
        "none",             "x"; ...
        "empty",            ""; ...
        "all_ws",           "   "; ...
        "internal_tab",     "x" + TAB + "y"; ...
        "lead_tab",         TAB + "hi"; ...
        "hello_both",       " hello "; ...
        "trail_nl",         "x" + LF; ...
        "nbsp_lead",        NBSP + "x"; ...
        "nbsp_trail",       "x" + NBSP; ...
        "nel_trail",        "x" + NEL; ...
        "figspace_lead",    FIGSP + "x"; ...
        "narrownbsp_trail", "x" + NNBSP; ...
        "lsep_lead",        LSEP + "x"; ...
        "astral_none",      GRIN; ...
        "astral_ws",        GRIN + " "};
end

function AP = appenderBattery()
    % The 7-case _RunContentAppender battery (mirrors s0021 AP).
    LF = string(char(10)); TAB = string(char(9)); CR = string(char(13));
    GRIN = string(char([55357 56832]));
    AP = { ...
        "tab_nl",      "a" + TAB + "b" + LF + "c"; ...
        "cr",          "x" + CR + "y"; ...
        "sp_both",     " a "; ...
        "tab_nl_only", TAB + LF; ...
        "astral",      "a" + GRIN + "b"; ...
        "empty",       ""; ...
        "crlf",        CR + LF};
end

function P = runProbes()
    % Replay the s0021 probe sequence (the .m twin's P-struct body, VERBATIM
    % tags/inputs/order) and return the nested struct of tagged canonical values.
    % Embedded here so the Equivalence leg is self-contained (the validation-folder
    % scenario is NOT on the toolbox path and must not be a dependency). Mirrors
    % validation\mat2doc\scenarios\s0021_p4_1b_oxml_run.m.
    P = struct();

    % ===================== xml:space battery (BYTE-CRITICAL) =============
    XS = xmlspaceBattery();
    xs = struct();
    for k = 1:size(XS, 1)
        r = mat2doc.oxml.OxmlElement("w:r");
        r.add_t(XS{k, 2});
        raw = mat2doc.oxml.serialize_part_xml(r);
        s = string(native2unicode(raw, "UTF-8"));
        e = struct();
        e.serhex        = hx(raw);
        e.has_xmlspace  = ternary(contains(s, 'xml:space="preserve"'));
        e.has_xmlns_xml = ternary(contains(s, "xmlns:xml"));
        xs.(XS{k, 1}) = e;
    end
    P.xmlspace = xs;

    % ===================== _RunContentAppender char mapping =============
    AP = appenderBattery();
    ap = struct();
    for k = 1:size(AP, 1)
        r = mat2doc.oxml.OxmlElement("w:r");
        r.text = AP{k, 2};
        raw = mat2doc.oxml.serialize_part_xml(r);
        e = struct();
        e.localnames = lnsCell(r);
        e.serhex     = hx(raw);
        ap.(AP{k, 1}) = e;
    end
    P.appender = ap;

    % ===================== H11 rPr-first (_insert_rPr override) =========
    h11 = struct();
    r = mat2doc.oxml.OxmlElement("w:r"); r.add_t("x"); r.get_or_add_rPr();
    h11.content_then_rpr = struct("localnames", {lnsCell(r)}, "serhex", hx_e(r));
    r = mat2doc.oxml.OxmlElement("w:r"); r.add_t("a"); r.get_or_add_rPr(); r.add_t("b");
    h11.interleaved = struct("localnames", {lnsCell(r)}, "serhex", hx_e(r));
    r = mat2doc.oxml.OxmlElement("w:r"); r.add_t("x"); r.style = "Emphasis";
    h11.style_route = struct("localnames", {lnsCell(r)}, "serhex", hx_e(r), "style", r.style);
    r = mat2doc.oxml.OxmlElement("w:r"); a = r.get_or_add_rPr(); b = r.get_or_add_rPr();
    h11.get_or_add_idempotent = ternary(a == b);
    P.h11 = h11;

    % ===================== CT_Text.__str__ (H3/H4) =====================
    cts = struct();
    t = mat2doc.oxml.OxmlElement("w:t");                cts.empty  = t.str_();
    t = mat2doc.oxml.OxmlElement("w:t"); t.text = "abc"; cts.abc    = t.str_();
    t = mat2doc.oxml.OxmlElement("w:t"); t.text = " a "; cts.spaced = t.str_();
    P.ct_text_str = cts;

    % ===================== run-content __str__ + CT_Br tri-state =======
    rc = struct();
    br = mat2doc.oxml.OxmlElement("w:br");                            rc.br_default = br.str_();
    br = mat2doc.oxml.OxmlElement("w:br"); br.type = "textWrapping";  rc.br_textWrapping = br.str_();
    br = mat2doc.oxml.OxmlElement("w:br"); br.type = "page";          rc.br_page = br.str_();
    br = mat2doc.oxml.OxmlElement("w:br"); br.type = "column";        rc.br_column = br.str_();
    cr = mat2doc.oxml.OxmlElement("w:cr");                            rc.cr = cr.str_();
    nbh = mat2doc.oxml.OxmlElement("w:noBreakHyphen");                rc.noBreakHyphen = nbh.str_();
    pt = mat2doc.oxml.OxmlElement("w:ptab");                          rc.ptab = pt.str_();
    br = mat2doc.oxml.OxmlElement("w:br");                            rc.br_type_absent = br.type;
    br = mat2doc.oxml.OxmlElement("w:br");                            rc.br_clear_absent = tri(br.clear);
    P.run_content_str = rc;

    % ===================== CT_R.text getter over mixed content =========
    r = rparse("<w:r xmlns:w=""" + W_() + """>" + ...
        "<w:t>x</w:t><w:br/><w:noBreakHyphen/><w:ptab/><w:cr/></w:r>");
    P.ct_r_text = struct("text", r.text);

    % ===================== comment range helpers =======================
    p = mat2doc.oxml.OxmlElement("w:p");
    r = mat2doc.oxml.OxmlElement("w:r");
    p.append(r);
    r.insert_comment_range_start_above(0);
    r.insert_comment_range_end_and_reference_below(0);
    kids = p.xpath("./*");
    ref_run = kids(end);
    P.comments = struct("localnames", {lnsCell(p)}, ...
                        "ref_run_style", ref_run.style, ...
                        "serhex", hx_e(p));
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
        % Empty cell (e.g. appender.empty.localnames) -> "" ; note join() of an
        % empty string array returns <missing>, so guard it explicitly.
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
        % Empty numeric (jsondecode of a JSON []) -> "" (matches an empty cell).
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

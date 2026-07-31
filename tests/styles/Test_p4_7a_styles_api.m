classdef Test_p4_7a_styles_api < matlab.unittest.TestCase
% TEST_P4_7A_STYLES_API  Gate-4 permanent unit tests for Mat2Doc P4-7a
%   (the styles API/proxy tier): src/docx/styles/style.py ->
%   +mat2doc\+styles\{StyleFactory, BaseStyle, CharacterStyle, ParagraphStyle,
%   TableStyle_, NumberingStyle_}, src/docx/styles/styles.py::Styles ->
%   +mat2doc\+styles\Styles, src/docx/styles/__init__.py::BabelFish ->
%   +mat2doc\+styles\BabelFish, PLUS the delegation UN-STUB across
%   StylesPart/DocumentPart/Document/Run/Paragraph.
%
%   P4-7a is an API/proxy-tier WP: it adds NO oxml registry rows and NO
%   serialization-path change (the byte-critical CT_* layer landed at P4-6). So
%   equivalence is BEHAVIORAL (probe value parity, proven at Gate-3 by probe_diff
%   MATCH exit 0 over the full surface + a re-proven 164-row template dump) PLUS
%   serialized-bytes parity on the two output-visible whole-part paths (delete_,
%   add_style) PLUS the end-to-end style-by-name document byte pin (the M2
%   add_heading precursor). This class permanently FREEZES that surface --
%   byte/value-identical to python-docx 1.2.0.
%
%   THE HIGHEST-VALUE PERMANENT PINS IN THIS WP (a regression goes RED loudly):
%     * delete_() whole-part byte pin (test_delete_byte_pin): a proxy
%       delete_("Heading 1") over a real Document -> the serialized styles part is
%       EXACTLY 348872 B, SHA-256 cc0bb35d...8614, byte-identical to python-docx
%       style.delete(). H17 Gate-4 rule: only the PARENT-side serialized bytes are
%       asserted -- the destroyed proxy handle is NEVER inspected after delete_.
%     * GC-safety pin (test_gc_safety_proxy_layer): minting several transient
%       StyleFactory proxies in a loop and letting them go out of scope leaves the
%       styles part bit-for-bit UNCHANGED (349458 B, 02d71a68...e384). This is the
%       H17 proxy-layer safety property -- because `delete` is NOT overridden on
%       any style proxy, proxy GC has no tree effect. A regression that
%       re-introduces a `delete` override (detaching the live w:style on ordinary
%       iteration) goes RED here.
%     * add_style whole-part byte pin (test_add_style_byte_pin): two add_style
%       calls (custom CHARACTER + builtin PARAGRAPH) -> the serialized styles part
%       is EXACTLY 349670 B, SHA-256 37e2136b...6541, byte-identical to python-docx.
%     * G-scenario document byte pin (test_gscenario_document_byte_pin): a fresh
%       mat2doc.Document() whose paragraphs get their style BY NAME ("Heading 1")
%       and BY Style OBJECT (d.styles["Heading 2"]) through the NOW-UN-STUBBED
%       delegation chain, saved -> word/document.xml is byte-identical (size +
%       SHA-256 + whole-bytes) to the frozen s0032 python-docx reference (1774 B,
%       db047dd9...6aa7); word/styles.xml stays UNCHANGED from M1 (02d71a68...e384).
%       The closest guard yet to M2's add_heading path.
%     * StyleFactory H10 dispatch + the KeyError-on-None edge
%       (test_stylefactory_dispatch): the 4 WD_STYLE_TYPE members -> the correct
%       leaf class; a bare w:style (type []/None) -> mat2doc:KeyError with the
%       Python key-repr message "None".
%
%   Provenance (Gate-1..3, all 2026-07-30):
%     * Audit    : validation\mat2doc\audit_P4-7a_styles_api.md (Porter Gate-1
%                  self-probe 69/0 + Fable/mso-auditor Gate-2 128/0 adversarial
%                  APPROVE; H17 delete_() naming resolution SOUND; 164/164 dump;
%                  add_style + delete_ whole-part byte proofs; zero defects).
%     * Validate : validation\mat2doc\validate_P4-7a_styles_api.md (Gate-3 PASS --
%                  probe_diff s0031 MATCH exit 0 over the full surface + 164-row
%                  dump 164/164; three whole-part byte proofs independently
%                  re-derived; G-scenario s0032 17/17 with word/document.xml
%                  byte-identical; M1 17/17; ZERO new D-numbers).
%     * Scenarios: validation\mat2doc\scenarios\s0031_p4_7a_styles_probe.{py,m}
%                  (the full-surface behavioral + 164-row dump + byteproof probe,
%                  its .m body replayed VERBATIM by runProbes() below) and
%                  s0032_p4_7a_style_byname_gscenario.{py,m} (the style-by-name
%                  document byte pin replayed by buildGScenario() below).
%     * Frozen refs (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE at Gate-3):
%         references\s0031\probe.json -- copied verbatim (self-contained) into
%           tests\styles\data\s0031_probe_oracle.json (value/serhex/byteproof JSON;
%           jsondecode is line-ending agnostic -> no `* binary` pin, s0028/s0030
%           precedent).
%         references\s0032\parts\word\document.xml (SHA-256 db047dd9...6aa7, 1774 B)
%           -- copied byte-for-byte into tests\styles\data\s0032_document.xml
%           (co-located `* binary` .gitattributes) as the G-scenario byte fixture,
%           AND its SHA/size embedded below.
%
%   Coverage taxonomy
%   -----------------
%   * Nominal      -- StyleFactory 4-type dispatch; BabelFish both directions;
%                     BaseStyle get/set props; base_style + next_paragraph_style
%                     resolution; the whole Styles collection surface; add_style;
%                     the un-stubbed Document/Run/Paragraph style paths; the
%                     G-scenario.
%   * Edge         -- H3 None-vs-CT_OnOff-False (bare style); dangling refs -> [];
%                     empty "" name; non-ASCII BabelFish passthrough; H15 case
%                     traps; single-element; error paths (KeyError on None type,
%                     KeyError on missing name, ValueError on type mismatch,
%                     ValueError on duplicate name) -- each verifies the IDENTIFIER
%                     (mat2doc:<PyExceptionName>) AND the verbatim message.
%   * Equivalence  -- test_equivalence_full_battery_vs_frozen_oracle replays the
%                     ENTIRE s0031 battery (runProbes, the .m twin's body verbatim,
%                     incl. babelfish/factory/basestyle/base_chain/para_style/
%                     styles/dump(164 rows)/byteproofs) and flatten-compares every
%                     leaf to the frozen python-docx 1.2.0 oracle (Gate-3 found
%                     ZERO divergences).
%   * Regression   -- hard-coded expected property values + the four whole-part /
%                     document SHA-256 + size byte pins (delete_, add_style, GC,
%                     G-scenario).
%   * Upstream     -- BabelFish's 12-alias table, StyleFactory's 4-way dispatch +
%                     KeyError-on-None, the next_paragraph_style fallback matrix and
%                     the verbatim ValueError/KeyError messages ARE the python-docx
%                     styles.py / __init__.py contract; the frozen oracle IS lxml's
%                     expected output for this API sequence.
%
%   Byte-level (L1) note: every serialized-bytes assertion is a SHA-256 (+ size)
%   pin of the raw UTF-8 shipping bytes (serialize_part_xml for the styles part;
%   the extracted word/document.xml for the G-scenario). SHA-256 equality == byte
%   identity (L1). No D-number granted any L2 relaxation in this WP (Gate-3: zero
%   new; delete_ is a FLAG-3 method-naming resolution, byte-identical to
%   python-docx style.delete()), so every byte pin is L1. The 164-count / leaf-key
%   guards in the equivalence leg are the only looser-than-byte checks and are
%   commented at their site.
%
%   Determinism: no network, no absolute paths. The co-located oracle + byte
%   fixture resolve relative to this file via fileparts(mfilename('fullpath'));
%   saves go to tempname .docx / tempname dirs deleted via onCleanup; every file
%   read is binary ('r','n'). The +mat2doc package resolves via the MANDATORY
%   PathFixture(worktree-root) in TestClassSetup (WP9-F4 lesson).

    properties (Constant)
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        % --- delete_("Heading 1") whole styles part (Gate-3 s0031 byteproof) ---
        SIZE_DELETE = 348872
        SHA_DELETE  = "cc0bb35d40c717a8329a01ea66e8aeaa592aaa89b9624c882e7a9083dbbc8614"

        % --- add_style(custom CHARACTER + builtin PARAGRAPH) whole styles part ---
        SIZE_ADDSTYLE = 349670
        SHA_ADDSTYLE  = "37e2136b9f5995a006c87865dacb6c6ce268027e244e87476bf89c6b0ce46541"

        % --- GC-invariant / M1 styles.xml (a fresh Document, no mutation) ---
        SIZE_STYLES_M1 = 349458
        SHA_STYLES_M1  = "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"

        % --- G-scenario (s0032) style-by-name word/document.xml byte pin ---
        SIZE_GSCENARIO_DOC = 1774
        SHA_GSCENARIO_DOC  = "db047dd90148ff9728a9e863bcb389721b860647674885fa339373c953d46aa7"
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied from tests\text\Test_p4_5b_paragraph_api.m. here is
            % tests\styles; the worktree root is two levels up.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\styles
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. StyleFactory H10 dispatch + KeyError-on-None edge            %
        % =============================================================== %

        function test_stylefactory_dispatch(testCase)
            % Nominal + Edge + Error path (style.py 15-24, H10): the 4 WD_STYLE_TYPE
            % members dispatch to ParagraphStyle/CharacterStyle/TableStyle_/
            % NumberingStyle_; a bare w:style (type []/None) -> mat2doc:KeyError with
            % the Python key-repr message "None". Both hard-coded AND vs the oracle.
            ST = @(n) mat2doc.enum.style.WD_STYLE_TYPE.(n);
            cases = { "PARAGRAPH", 'mat2doc.styles.ParagraphStyle'; ...
                      "CHARACTER", 'mat2doc.styles.CharacterStyle'; ...
                      "TABLE",     'mat2doc.styles.TableStyle_'; ...
                      "LIST",      'mat2doc.styles.NumberingStyle_' };
            for i = 1:size(cases, 1)
                se = mat2doc.oxml.OxmlElement("w:style");
                se.type = ST(cases{i, 1});
                s = mat2doc.styles.StyleFactory(se);
                testCase.verifyClass(s, cases{i, 2}, ...
                    sprintf('StyleFactory(%s) -> %s', cases{i, 1}, cases{i, 2}));
            end

            % Edge: bare w:style (type absent -> []) -> KeyError "None"
            bare = mat2doc.oxml.OxmlElement("w:style");
            ME = captureError(@() mat2doc.styles.StyleFactory(bare));
            testCase.verifyEqual(string(ME.identifier), "mat2doc:KeyError", ...
                'StyleFactory(None-type) -> mat2doc:KeyError identifier');
            testCase.verifyEqual(string(ME.message), "None", ...
                'StyleFactory(None-type) -> verbatim Python key-repr message "None"');

            % vs the frozen oracle
            o = loadOracle().factory;
            testCase.verifyEqual(string(o.paragraph), "ParagraphStyle");
            testCase.verifyEqual(string(o.character), "CharacterStyle");
            testCase.verifyEqual(string(o.table),     "TableStyle");
            testCase.verifyEqual(string(o.numbering), "NumberingStyle");
            testCase.verifyEqual(string(o.none_type_raises), "True");
            testCase.verifyEqual(string(o.none_type_msg),    "None");
        end

        % =============================================================== %
        % 2. BabelFish (H15 case-sensitive alias table)                   %
        % =============================================================== %

        function test_babelfish(testCase)
            % Nominal + Edge (styles/__init__.py 8-40, H15): all 12 alias pairs both
            % directions (pinned) + unknown-name passthrough + the case traps
            % ("heading 1"/"Heading 1" are distinct keys) + empty "" -> "".
            BF = @(m, x) mat2doc.styles.BabelFish.(m)(x);
            pairs = [ ...
                "Caption","caption"; "Footer","footer"; "Header","header"; ...
                "Heading 1","heading 1"; "Heading 2","heading 2"; "Heading 3","heading 3"; ...
                "Heading 4","heading 4"; "Heading 5","heading 5"; "Heading 6","heading 6"; ...
                "Heading 7","heading 7"; "Heading 8","heading 8"; "Heading 9","heading 9"];
            for i = 1:size(pairs, 1)
                ui = pairs(i, 1); internal = pairs(i, 2);
                testCase.verifyEqual(BF("ui2internal", ui), internal, ...
                    sprintf('ui2internal("%s") -> "%s"', ui, internal));
                testCase.verifyEqual(BF("internal2ui", internal), ui, ...
                    sprintf('internal2ui("%s") -> "%s"', internal, ui));
            end

            % passthrough (both directions) -- unknown names returned verbatim
            testCase.verifyEqual(BF("ui2internal", "Totally Unknown"), "Totally Unknown", ...
                'ui2internal unknown -> passthrough');
            testCase.verifyEqual(BF("internal2ui", "totally unknown"), "totally unknown", ...
                'internal2ui unknown -> passthrough');
            % non-ASCII passthrough (é / CJK / emoji survive verbatim)
            testCase.verifyEqual(BF("ui2internal", "Ünïcode 段落 ☕"), "Ünïcode 段落 ☕", ...
                'ui2internal non-ASCII -> passthrough');

            % case traps (H15): "heading 1" is NOT a UI key; "Heading 1" is NOT an
            % internal key -- each falls through to passthrough.
            testCase.verifyEqual(BF("ui2internal", "heading 1"), "heading 1", ...
                'H15: ui2internal("heading 1") is a MISS -> passthrough (not "heading 1"->internal)');
            testCase.verifyEqual(BF("internal2ui", "Heading 1"), "Heading 1", ...
                'H15: internal2ui("Heading 1") is a MISS -> passthrough');

            % empty
            testCase.verifyEqual(BF("ui2internal", ""), "", 'empty "" -> ""');
            testCase.verifyEqual(BF("internal2ui", ""), "", 'empty "" -> ""');

            % vs the frozen oracle
            o = loadOracle().babelfish;
            for i = 0:size(pairs, 1) - 1
                k = "p" + sprintf("%02d", i);
                testCase.verifyEqual(string(o.ui2internal.(k)), pairs(i + 1, 2), ...
                    sprintf('oracle ui2internal.%s', k));
                testCase.verifyEqual(string(o.internal2ui.(k)), pairs(i + 1, 1), ...
                    sprintf('oracle internal2ui.%s', k));
            end
            testCase.verifyEqual(string(o.passthrough_ui), "Totally Unknown");
            testCase.verifyEqual(string(o.case_trap_ui), "heading 1");
            testCase.verifyEqual(string(o.case_trap_internal), "Heading 1");
        end

        % =============================================================== %
        % 3. BaseStyle properties (H3 None vs CT_OnOff False)             %
        % =============================================================== %

        function test_basestyle_props(testCase)
            % Nominal + Edge (style.py 34-161, H3/H4): a BARE w:style shows the H3
            % tri-state -- name/priority/style_id are [] (None), builtin is True
            % (customStyle absent), while hidden/locked/quick_style/unhide_when_used
            % are CT_OnOff-backed logical FALSE (NOT []). type None -> PARAGRAPH.
            % Setters write real values readable back through the getters; the
            % faithful quirk: name="Heading 1" stores "Heading 1" verbatim.
            se = parse("<w:style " + testCase.nsWStr() + "/>");
            b = mat2doc.styles.ParagraphStyle(se);

            % H3: None (via [])
            testCase.verifyTrue(isequal(b.name, []),      'bare name -> [] (None)');
            testCase.verifyTrue(isequal(b.priority, []),  'bare priority -> [] (None)');
            testCase.verifyTrue(isequal(b.style_id, []),  'bare style_id -> [] (None)');
            % H3: CT_OnOff-backed logical FALSE (NOT [])
            testCase.verifyFalse(b.hidden,           'bare hidden -> logical FALSE (not None)');
            testCase.verifyFalse(b.locked,           'bare locked -> logical FALSE');
            testCase.verifyFalse(b.quick_style,      'bare quick_style -> logical FALSE');
            testCase.verifyFalse(b.unhide_when_used, 'bare unhide_when_used -> logical FALSE');
            testCase.verifyTrue(islogical(b.hidden), 'bare hidden is logical (H3 distinguishes false from [])');
            % builtin: not customStyle (H4) -- absent customStyle -> True
            testCase.verifyTrue(b.builtin, 'bare builtin -> true (not customStyle, H4)');
            % type: None -> PARAGRAPH default (style.py 144-147)
            testCase.verifyEqual(string(b.type), "PARAGRAPH", 'bare type None -> PARAGRAPH');

            % setters
            b.name = "Heading 1"; b.style_id = "Heading1"; b.priority = 9;
            b.hidden = true; b.locked = true; b.quick_style = true; b.unhide_when_used = true;
            testCase.verifyEqual(b.name, "Heading 1", 'name getter (internal2ui of stored value)');
            % faithful quirk: the SETTER writes verbatim (no ui2internal), so the
            % stored internal name_val is "Heading 1", not "heading 1".
            testCase.verifyEqual(se.name_val, "Heading 1", 'name setter stores value verbatim (no ui2internal)');
            testCase.verifyEqual(b.style_id, "Heading1", 'style_id get');
            testCase.verifyEqual(b.priority, 9,          'priority get');
            testCase.verifyTrue(b.hidden);   testCase.verifyTrue(b.locked);
            testCase.verifyTrue(b.quick_style); testCase.verifyTrue(b.unhide_when_used);
            testCase.verifyTrue(b.builtin, 'builtin stays true (customStyle never set; not None)');

            % the internal-name round-trip: name="heading 1" stores "heading 1",
            % getter maps internal->UI "Heading 1".
            se2 = parse("<w:style " + testCase.nsWStr() + "/>");
            b2 = mat2doc.styles.ParagraphStyle(se2);
            b2.name = "heading 1";
            testCase.verifyEqual(se2.name_val, "heading 1", 'stored internal name');
            testCase.verifyEqual(b2.name, "Heading 1", 'getter maps internal "heading 1" -> UI "Heading 1"');

            % H3 clear: set-false / set-[] -> back to false/None
            b.hidden = false; b.locked = false; b.quick_style = false;
            b.unhide_when_used = false; b.priority = [];
            testCase.verifyFalse(b.hidden); testCase.verifyFalse(b.locked);
            testCase.verifyFalse(b.quick_style); testCase.verifyFalse(b.unhide_when_used);
            testCase.verifyTrue(isequal(b.priority, []), 'priority=[] -> [] (None)');

            % builtin: customStyle="1" -> not True -> False
            bt = mat2doc.styles.ParagraphStyle( ...
                parse("<w:style w:customStyle=""1"" " + testCase.nsWStr() + "/>"));
            testCase.verifyFalse(bt.builtin, 'customStyle="1" -> builtin false');
        end

        % =============================================================== %
        % 4. base_style chain (resolve / no-basedOn / dangling / set)      %
        % =============================================================== %

        function test_base_style_chain(testCase)
            % Nominal + Edge (style.py 171-183, H3): base_style resolves the sibling
            % via w:basedOn; no basedOn -> [] (None); a DANGLING ref -> [] (None);
            % font is a lazy Font proxy. Setter writes style_id / removes on [].
            styles = parse("<w:styles " + testCase.nsWStr() + ">" + ...
                styleXml("Normal", "paragraph") + ...
                styleXml("Heading1", "paragraph", "<w:basedOn w:val=""Normal""/>") + ...
                styleXml("Dangling", "paragraph", "<w:basedOn w:val=""NoSuchZzz""/>") + ...
                "</w:styles>");
            h1   = mat2doc.styles.StyleFactory(styles.get_by_id("Heading1"));
            dang = mat2doc.styles.StyleFactory(styles.get_by_id("Dangling"));
            norm = mat2doc.styles.StyleFactory(styles.get_by_id("Normal"));

            testCase.verifyClass(h1.base_style, 'mat2doc.styles.ParagraphStyle', ...
                'base_style resolves to a ParagraphStyle');
            testCase.verifyEqual(h1.base_style.style_id, "Normal", 'base_style id == Normal');
            testCase.verifyTrue(isequal(norm.base_style, []), 'no basedOn -> [] (None)');
            testCase.verifyTrue(isequal(dang.base_style, []), 'dangling basedOn ref -> [] (None)');

            % font is a lazy Font proxy over the wrapped w:style
            testCase.verifyClass(h1.font, 'mat2doc.text.Font', 'font -> Font proxy');

            % set base_style to a Style -> writes its style_id
            norm.base_style = h1;
            testCase.verifyEqual(norm.base_style.style_id, "Heading1", 'set base_style writes style_id');
            % set base_style to [] -> removes basedOn
            norm.base_style = [];
            testCase.verifyTrue(isequal(norm.base_style, []), 'set base_style=[] removes basedOn');
        end

        % =============================================================== %
        % 5. next_paragraph_style fallback matrix + paragraph_format       %
        % =============================================================== %

        function test_next_paragraph_style_matrix(testCase)
            % Nominal + Edge (style.py 206-232): next_paragraph_style falls back to
            % SELF when w:next is absent, dangling, or references a NON-paragraph
            % style; a real paragraph next resolves. Setter: None or self -> remove
            % (-> self); another style -> write.
            styles = parse("<w:styles " + testCase.nsWStr() + ">" + ...
                styleXml("Normal", "paragraph") + ...
                styleXml("BodyChar", "character") + ...
                styleXml("H_none", "paragraph") + ...
                styleXml("H_dangling", "paragraph", "<w:next w:val=""NoSuchZzz""/>") + ...
                styleXml("H_nonpara", "paragraph", "<w:next w:val=""BodyChar""/>") + ...
                styleXml("H_real", "paragraph", "<w:next w:val=""Normal""/>") + ...
                "</w:styles>");
            para = @(sid) mat2doc.styles.StyleFactory(styles.get_by_id(sid));

            testCase.verifyEqual(para("H_none").next_paragraph_style.style_id, "H_none", ...
                'no w:next -> self');
            testCase.verifyEqual(para("H_dangling").next_paragraph_style.style_id, "H_dangling", ...
                'dangling w:next -> self');
            testCase.verifyEqual(para("H_nonpara").next_paragraph_style.style_id, "H_nonpara", ...
                'w:next to a non-paragraph style -> self');
            testCase.verifyEqual(para("H_real").next_paragraph_style.style_id, "Normal", ...
                'real paragraph next -> resolves');
            testCase.verifyClass(para("H_real").next_paragraph_style, 'mat2doc.styles.ParagraphStyle', ...
                'resolved next is a ParagraphStyle');

            % setter matrix
            hr = para("H_real");
            hr.next_paragraph_style = [];      % None -> remove -> self
            testCase.verifyEqual(hr.next_paragraph_style.style_id, "H_real", 'set None -> remove -> self');
            hr.next_paragraph_style = para("Normal");   % other -> write
            testCase.verifyEqual(hr.next_paragraph_style.style_id, "Normal", 'set other -> write');
            hr.next_paragraph_style = hr;      % self -> remove -> self
            testCase.verifyEqual(hr.next_paragraph_style.style_id, "H_real", 'set self -> remove -> self');

            % paragraph_format is a lazy ParagraphFormat proxy
            testCase.verifyClass(para("Normal").paragraph_format, 'mat2doc.text.ParagraphFormat', ...
                'paragraph_format -> ParagraphFormat proxy');
        end

        % =============================================================== %
        % 6. Styles collection surface (over a REAL Document)              %
        % =============================================================== %

        function test_styles_collection(testCase)
            % Nominal + Edge (styles.py 15-136): len_/contains_/getitem_(name +
            % deprecated-id + KeyError)/to_array(heterogeneous)/default/get_by_id
            % branches/get_style_id branches/element over a real Document.
            ST = @(n) mat2doc.enum.style.WD_STYLE_TYPE.(n);
            d = mat2doc.Document();
            st = d.styles;
            testCase.verifyClass(st, 'mat2doc.styles.Styles', 'd.styles -> Styles');

            % len_
            testCase.verifyEqual(st.len_(), 164, 'len_ == 164 (default template)');

            % contains_ (by UI name / internal / miss)
            testCase.verifyTrue(st.contains_("Heading 1"), 'contains_ UI name');
            testCase.verifyTrue(st.contains_("Normal"),    'contains_ internal name');
            testCase.verifyFalse(st.contains_("No Such Style Zzz"), 'contains_ miss -> false');

            % getitem_ by UI name
            h1 = st.getitem_("Heading 1");
            testCase.verifyClass(h1, 'mat2doc.styles.ParagraphStyle', 'getitem_ UI name -> ParagraphStyle');
            testCase.verifyEqual(h1.style_id, "Heading1", 'getitem_ UI name id');

            % getitem_ by deprecated style-id -> emits mat2doc:UserWarning (VERIFY-WARN)
            prev = warning('off', 'mat2doc:UserWarning');
            cleanup = onCleanup(@() warning(prev)); %#ok<NASGU>
            lastwarn('', '');
            byId = st.getitem_("Heading1");
            [~, wid] = lastwarn();
            testCase.verifyEqual(string(wid), "mat2doc:UserWarning", ...
                'deprecated by-id getitem_ emits mat2doc:UserWarning (house convention)');
            testCase.verifyEqual(byId.style_id, "Heading1", 'deprecated by-id getitem_ still resolves');

            % getitem_ miss -> KeyError (identifier + verbatim message)
            ME = captureError(@() st.getitem_("No Such Style Zzz"));
            testCase.verifyEqual(string(ME.identifier), "mat2doc:KeyError", 'getitem_ miss -> KeyError id');
            testCase.verifyEqual(string(ME.message), "no style with name 'No Such Style Zzz'", ...
                'getitem_ miss -> verbatim message');

            % to_array: heterogeneous 1xN BaseStyle array, correct leaf classes
            arr = st.to_array();
            testCase.verifyEqual(numel(arr), 164, 'to_array len == 164');
            testCase.verifyClass(arr, 'mat2doc.styles.BaseStyle', 'to_array typed BaseStyle (heterogeneous root)');
            % iteration yields the correct leaf subclasses
            sawPara = false; sawChar = false;
            for s = arr
                if isa(s, 'mat2doc.styles.ParagraphStyle'), sawPara = true; end
                if strcmp(class(s), 'mat2doc.styles.CharacterStyle'), sawChar = true; end
            end
            testCase.verifyTrue(sawPara && sawChar, 'to_array carries mixed leaf classes');

            % default
            testCase.verifyEqual(st.default(ST("PARAGRAPH")).style_id, "Normal", 'default PARAGRAPH -> Normal');
            testCase.verifyClass(st.default(ST("CHARACTER")), 'mat2doc.styles.CharacterStyle', ...
                'default CHARACTER -> a CharacterStyle');

            % get_by_id branches
            testCase.verifyEqual(st.get_by_id("Heading1", ST("PARAGRAPH")).style_id, "Heading1", 'get_by_id found');
            testCase.verifyEqual(st.get_by_id([], ST("PARAGRAPH")).style_id, "Normal", 'get_by_id None -> default');
            testCase.verifyEqual(st.get_by_id("", ST("PARAGRAPH")).style_id, "Normal", 'get_by_id "" (H4 falsy) -> default');
            testCase.verifyEqual(st.get_by_id("Nope", ST("PARAGRAPH")).style_id, "Normal", 'get_by_id miss -> default');
            testCase.verifyEqual(st.get_by_id("Heading1", ST("CHARACTER")).style_id, ...
                st.default(ST("CHARACTER")).style_id, 'get_by_id wrong-type -> type default');

            % get_style_id branches
            testCase.verifyTrue(isequal(st.get_style_id([], ST("PARAGRAPH")), []), 'get_style_id None -> []');
            testCase.verifyEqual(st.get_style_id("Heading 1", ST("PARAGRAPH")), "Heading1", 'get_style_id name');
            testCase.verifyEqual(st.get_style_id(st.getitem_("Heading 1"), ST("PARAGRAPH")), "Heading1", ...
                'get_style_id Style object');
            testCase.verifyTrue(isequal(st.get_style_id("Normal", ST("PARAGRAPH")), []), ...
                'get_style_id default-style name -> [] (None)');
            % wrong-type -> ValueError (identifier + verbatim message)
            MEw = captureError(@() st.get_style_id("Heading 1", ST("CHARACTER")));
            testCase.verifyEqual(string(MEw.identifier), "mat2doc:ValueError", 'get_style_id wrong-type -> ValueError id');
            testCase.verifyEqual(string(MEw.message), ...
                "assigned style is type PARAGRAPH (1), need type CHARACTER (2)", ...
                'get_style_id wrong-type -> verbatim enum-str message');
            % miss -> KeyError
            MEk = captureError(@() st.get_style_id("No Such Style Zzz", ST("PARAGRAPH")));
            testCase.verifyEqual(string(MEk.identifier), "mat2doc:KeyError", 'get_style_id miss -> KeyError id');
            testCase.verifyEqual(string(MEk.message), "no style with name 'No Such Style Zzz'", ...
                'get_style_id miss -> verbatim message');

            % element localname
            testCase.verifyEqual(string(st.element().local_part), "styles", 'element() localname "styles"');
        end

        function test_add_style_and_duplicate(testCase)
            % Nominal + Error path (styles.py 55-65): add_style (custom + builtin) +
            % the duplicate-name ValueError (identifier + verbatim message).
            ST = @(n) mat2doc.enum.style.WD_STYLE_TYPE.(n);
            d = mat2doc.Document();
            st = d.styles;

            cst = st.add_style("My Custom Style", ST("CHARACTER"));
            testCase.verifyClass(cst, 'mat2doc.styles.CharacterStyle', 'add_style custom -> CharacterStyle');
            testCase.verifyEqual(cst.style_id, "MyCustomStyle", 'add_style custom id');
            testCase.verifyFalse(cst.builtin, 'add_style custom -> builtin false');

            bst = st.add_style("My Builtin Style", ST("PARAGRAPH"), true);
            testCase.verifyClass(bst, 'mat2doc.styles.ParagraphStyle', 'add_style builtin -> ParagraphStyle');
            testCase.verifyEqual(bst.style_id, "MyBuiltinStyle", 'add_style builtin id');
            testCase.verifyTrue(bst.builtin, 'add_style builtin=true -> builtin true');

            % duplicate name -> ValueError (verbatim)
            ME = captureError(@() st.add_style("Heading 1", ST("PARAGRAPH")));
            testCase.verifyEqual(string(ME.identifier), "mat2doc:ValueError", 'add_style dup -> ValueError id');
            testCase.verifyEqual(string(ME.message), "document already contains style 'Heading 1'", ...
                'add_style dup -> verbatim message');
        end

        % =============================================================== %
        % 7. H17 delete_() whole-part byte pin (PARENT-side effect only)   %
        % =============================================================== %

        function test_delete_byte_pin(testCase)
            % Regression (H17, style.py 49-57; the whole-part byte proof): a proxy
            % delete_("Heading 1") over a real Document -> the serialized styles part
            % is EXACTLY 348872 B, SHA-256 cc0bb35d...8614, byte-identical to
            % python-docx style.delete(). H17 Gate-4 RULE: assert ONLY the
            % parent-side serialized bytes -- NEVER inspect the handle after delete_.
            d = mat2doc.Document();
            st = d.styles;
            h = st.getitem_("Heading 1");
            h.delete_();   % detach the w:style; do NOT touch h after this line
            proof = shaOfElement(st.element());
            testCase.verifyEqual(proof.size, testCase.SIZE_DELETE, ...
                sprintf('styles part must be exactly %d B after delete_', testCase.SIZE_DELETE));
            testCase.verifyEqual(proof.sha256, testCase.SHA_DELETE, ...
                'styles part SHA-256 == python-docx style.delete() (byte-identical L1)');
            % looser-than-byte cross-check (justified: proves Heading1 is the one gone):
            % 163 styles remain and Heading1 is absent.
            testCase.verifyEqual(st.len_(), 163, 'one style removed (164 -> 163)');
            testCase.verifyFalse(st.contains_("Heading 1"), 'Heading 1 gone from the collection');
        end

        function test_gc_safety_proxy_layer(testCase)
            % Regression (H17 proxy-layer safety property): minting several transient
            % StyleFactory proxies in a loop and letting them go out of scope leaves
            % the styles part bit-for-bit UNCHANGED (349458 B, 02d71a68...e384).
            % Because `delete` is NOT overridden on any style proxy, proxy GC has no
            % tree effect. A regression that re-introduces a `delete` override
            % (detaching the live w:style on ordinary iteration) goes RED here.
            ST = @(n) mat2doc.enum.style.WD_STYLE_TYPE.(n);
            d = mat2doc.Document();
            st = d.styles;
            before = shaOfElement(st.element());
            testCase.verifyEqual(before.size, testCase.SIZE_STYLES_M1, 'precondition: fresh styles.xml size');
            testCase.verifyEqual(before.sha256, testCase.SHA_STYLES_M1, 'precondition: fresh styles.xml SHA (M1)');
            for iter = 1:25
                a  = st.getitem_("Normal");                 %#ok<NASGU>
                bx = st.get_by_id("Heading1", ST("PARAGRAPH")); %#ok<NASGU>
                cc = st.default(ST("PARAGRAPH"));           %#ok<NASGU>
                arr = st.to_array();                        %#ok<NASGU>  164 transient proxies
                ee = st.getitem_("Heading 1");              %#ok<NASGU>
                clear a bx cc arr ee
            end
            after = shaOfElement(st.element());
            testCase.verifyEqual(after.size, testCase.SIZE_STYLES_M1, ...
                'styles part size UNCHANGED after 25x transient proxies (GC-safe)');
            testCase.verifyEqual(after.sha256, testCase.SHA_STYLES_M1, ...
                'styles part SHA-256 UNCHANGED -- proxy GC does not touch the tree (H17 safety)');
        end

        function test_add_style_byte_pin(testCase)
            % Regression (styles.py 55-65; the whole-part byte proof): two add_style
            % calls (custom CHARACTER carrying w:customStyle="1" + builtin PARAGRAPH
            % omitting it) -> the serialized styles part is EXACTLY 349670 B, SHA-256
            % 37e2136b...6541, byte-identical to python-docx.
            ST = @(n) mat2doc.enum.style.WD_STYLE_TYPE.(n);
            d = mat2doc.Document();
            st = d.styles;
            st.add_style("My Fancy Style", ST("CHARACTER"));           % custom
            st.add_style("My Builtin Style", ST("PARAGRAPH"), true);   % builtin
            proof = shaOfElement(st.element());
            testCase.verifyEqual(proof.size, testCase.SIZE_ADDSTYLE, ...
                sprintf('styles part must be exactly %d B after two add_style', testCase.SIZE_ADDSTYLE));
            testCase.verifyEqual(proof.sha256, testCase.SHA_ADDSTYLE, ...
                'styles part SHA-256 == python-docx (byte-identical L1)');
        end

        % =============================================================== %
        % 8. G-scenario style-by-name document byte pin (M2 precursor)     %
        % =============================================================== %

        function test_gscenario_document_byte_pin(testCase)
            % Regression (THE G-scenario byte pin, s0032): a fresh mat2doc.Document()
            % whose paragraphs get their style BY NAME ("Heading 1") and BY Style
            % OBJECT (d.styles["Heading 2"]) through the NOW-UN-STUBBED delegation
            % chain, saved -> word/document.xml is byte-identical (size + SHA-256 +
            % whole-bytes) to the frozen s0032 python-docx reference; word/styles.xml
            % stays UNCHANGED from M1. Closest guard to M2's add_heading path.
            [exdir, cleanup] = buildAndUnzipGScenario(); %#ok<ASGLU>

            doc = readBytes(fullfile(exdir, 'word', 'document.xml'));
            testCase.verifyEqual(numel(doc), testCase.SIZE_GSCENARIO_DOC, ...
                sprintf('word/document.xml must be exactly %d B', testCase.SIZE_GSCENARIO_DOC));
            testCase.verifyEqual(sha256hex(doc), testCase.SHA_GSCENARIO_DOC, ...
                'word/document.xml SHA-256 == frozen s0032 oracle (byte-identical L1)');
            want = loadDocumentFixture();
            testCase.verifyEqual(uint8(doc(:)'), uint8(want(:)'), ...
                'G-scenario word/document.xml == frozen s0032 reference (whole-bytes)');

            % word/styles.xml stays UNCHANGED from M1 (setting a pre-existing style
            % by name touches document.xml only).
            styles = readBytes(fullfile(exdir, 'word', 'styles.xml'));
            testCase.verifyEqual(numel(styles), testCase.SIZE_STYLES_M1, 'styles.xml size unchanged (M1)');
            testCase.verifyEqual(sha256hex(styles), testCase.SHA_STYLES_M1, ...
                'styles.xml SHA-256 unchanged from M1 -- style-by-name touches document.xml only');
        end

        % =============================================================== %
        % 9. Un-stub resolution pins (delegation now LIVE end-to-end)      %
        % =============================================================== %

        function test_unstub_resolution(testCase)
            % Regression (audit section 8 un-stub ledger): every delegated styles
            % path RESOLVES end-to-end (no mat2doc:notYetPorted), while the genuine
            % still-stubbed Styles.latent_styles REMAINS a clean notYetPorted stub
            % (P4-7b owner).
            ST = @(n) mat2doc.enum.style.WD_STYLE_TYPE.(n);
            d = mat2doc.Document();

            % Document.styles / DocumentPart.styles / StylesPart.styles resolve
            testCase.verifyClass(d.styles, 'mat2doc.styles.Styles', 'Document.styles resolves');
            dp = d.part();
            testCase.verifyClass(dp.styles(), 'mat2doc.styles.Styles', 'DocumentPart.styles resolves');
            testCase.verifyClass(dp.styles_part_().styles(), 'mat2doc.styles.Styles', ...
                'StylesPart.styles resolves');

            % DocumentPart.get_style / get_style_id resolve
            testCase.verifyClass(dp.get_style("Heading1", ST("PARAGRAPH")), 'mat2doc.styles.ParagraphStyle', ...
                'DocumentPart.get_style resolves');
            testCase.verifyEqual(dp.get_style_id("Heading 1", ST("PARAGRAPH")), "Heading1", ...
                'DocumentPart.get_style_id resolves');

            % Paragraph.style get + set resolve (style-by-name -> styleId). Capture
            % the CT_P (Paragraph exposes no element() accessor; StoryChild tier).
            body = d.element().body;
            p = body.add_p();
            para = mat2doc.text.Paragraph(p, d);
            testCase.verifyClass(para.style, 'mat2doc.styles.ParagraphStyle', 'Paragraph.style GET resolves (default Normal)');
            testCase.verifyEqual(para.style.style_id, "Normal", 'Paragraph.style GET default -> Normal');
            para.style = "Heading 1";
            testCase.verifyEqual(para.style.style_id, "Heading1", 'Paragraph.style SET by name -> Heading1');
            testCase.verifyEqual(string(p.style), "Heading1", ...
                'Paragraph.style SET writes styleId "Heading1" on ./w:pPr/w:pStyle (CT_P.style)');

            % Run.style get + set resolve; the applied char style is written to the
            % CT_R rStyle (reached via xpath on the CT_P -- Run exposes no element()).
            run = para.add_run("hi");
            testCase.verifyClass(run.style, 'mat2doc.styles.CharacterStyle', 'Run.style GET resolves (default)');
            run.style = "Emphasis";
            testCase.verifyEqual(run.style.style_id, "Emphasis", 'Run.style SET by name -> Emphasis (round-trip)');
            rr = p.xpath('.//w:r');
            testCase.verifyEqual(string(rr(end).style), "Emphasis", ...
                'Run.style SET writes w:rPr/w:rStyle w:val="Emphasis" on the CT_R (byte-level)');

            % latent_styles STILL a clean notYetPorted stub (P4-7b owner)
            ME = captureError(@() d.styles.latent_styles());
            testCase.verifyEqual(string(ME.identifier), "mat2doc:notYetPorted", ...
                'Styles.latent_styles STAYS a clean mat2doc:notYetPorted stub (P4-7b)');
        end

        % =============================================================== %
        % 10. EQUIVALENCE -- full s0031 battery vs the frozen oracle       %
        % =============================================================== %

        function test_equivalence_full_battery_vs_frozen_oracle(testCase)
            % Equivalence: replay the ENTIRE s0031 battery (runProbes -- the .m twin's
            % body VERBATIM: babelfish / factory / basestyle / base_chain /
            % para_style / styles (over a real Document) / dump (164 rows) /
            % byteproofs (delete_ + add_style SHA)) and flatten-compare EVERY leaf to
            % the frozen python-docx 1.2.0 oracle copied into
            % data\s0031_probe_oracle.json. Gate-3 found ZERO divergences (probe_diff
            % MATCH exit 0), so every leaf must be byte/value-identical.
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
            % Non-trivial size guard (guards against a silent-empty replay). The
            % 164-row dump alone yields > 1900 leaves; the whole battery > 500.
            testCase.verifyGreaterThan(numel(oKeys), 500, ...
                'the flattened oracle must expose the full battery of leaves');
            for i = 1:numel(oKeys)
                k = oKeys{i};
                testCase.verifyTrue(isKey(pMap, k), sprintf('port is missing leaf %s', k));
                testCase.verifyEqual(pMap(k), oMap(k), ...
                    sprintf('leaf %s must be byte/value-identical to the frozen oracle', k));
            end
        end

    end

    methods
        function s = nsWStr(testCase)
            s = "xmlns:w=""" + testCase.W + """";
        end
    end
end

% ===================== file-local helpers ============================== %

function s = nsW()
    s = "xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""";
end

function e = parse(xml)
    e = mat2doc.oxml.parse_xml(uint8(unicode2native(char(xml), "UTF-8")));
end

function x = styleXml(varargin)
    % <w:style w:type=".." w:styleId=".."><w:name w:val=".."/>EXTRA</w:style>
    sid = varargin{1}; wtype = varargin{2};
    if nargin >= 3, extra = varargin{3}; else, extra = ""; end
    x = "<w:style w:type=""" + wtype + """ w:styleId=""" + sid + """>" + ...
        "<w:name w:val=""" + sid + """/>" + extra + "</w:style>";
end

function out = shaOfElement(elm)
    % size + lowercase-hex SHA-256 of the raw serialize_part_xml shipping bytes
    % (mirrors the s0031 sha_elm; MATLAB Java MessageDigest, no toolbox).
    raw = mat2doc.oxml.serialize_part_xml(elm);
    out = struct('size', numel(raw), 'sha256', sha256hex(raw));
end

function h = sha256hex(bytes)
    % Lowercase hex SHA-256 of a uint8 vector (base MATLAB / Java MessageDigest).
    md  = java.security.MessageDigest.getInstance('SHA-256');
    dig = md.digest(typecast(uint8(bytes(:)'), 'int8'));
    b   = typecast(dig, 'uint8');
    h   = lower(string(sprintf('%02x', b)));
end

function [exdir, cleanup] = buildAndUnzipGScenario()
    % Replay the s0032 G-scenario body VERBATIM (validation\mat2doc\scenarios\
    % s0032_p4_7a_style_byname_gscenario.m): a fresh mat2doc.Document() whose
    % paragraphs get their style BY NAME + BY Style OBJECT through the un-stubbed
    % delegation chain, then .save(), unzip, and return the extraction dir. The
    % onCleanup handle removes the temp .docx and the extraction dir.
    d = mat2doc.Document();
    body = d.element().body;

    % ---- P1: style by NAME (the M2 add_heading precursor) ----
    p1 = mat2doc.text.Paragraph(body.add_p(), d);
    p1.add_run("Chapter Title");
    p1.style = "Heading 1";                        % -> "Heading1" on w:pStyle

    % ---- P2: style by Style OBJECT ----
    p2 = mat2doc.text.Paragraph(body.add_p(), d);
    p2.add_run("Subtitle");
    p2.style = d.styles.getitem_("Heading 2");     % -> "Heading2" on w:pStyle

    % ---- P3: no style ----
    p3 = mat2doc.text.Paragraph(body.add_p(), d);
    p3.add_run("Body paragraph text.");

    tmp = [tempname '.docx'];
    exdir = tempname;
    cleanup = onCleanup(@() cleanupGScenario(tmp, exdir));
    d.save(tmp);
    unzip(tmp, exdir);
end

function cleanupGScenario(tmp, exdir)
    if isfile(tmp),   delete(tmp);        end
    if isfolder(exdir), rmdir(exdir, 's'); end
end

function want = loadDocumentFixture()
    % The frozen s0032 word/document.xml byte fixture (co-located `* binary`
    % .gitattributes so it is checked out byte-for-byte). Read in BINARY mode.
    here = fileparts(mfilename('fullpath'));
    p = fullfile(here, 'data', 's0032_document.xml');
    want = readBytes(p);
end

function b = readBytes(p)
    f = fopen(p, 'r', 'n');            % binary read (no CRLF translation)
    assert(f >= 0, 'could not open for read: %s', p);
    b = fread(f, Inf, '*uint8')';
    fclose(f);
end

function o = loadOracle()
    % Read the co-located frozen s0031 oracle in BINARY mode (no CRLF translation)
    % and decode UTF-8 -> struct. jsondecode is whitespace/line-ending agnostic, so
    % no `* binary` pin is needed for this value/serhex/byteproof fixture (s0028/
    % s0030 precedent).
    here = fileparts(mfilename('fullpath'));
    p = fullfile(here, 'data', 's0031_probe_oracle.json');
    fid = fopen(p, 'r', 'n');
    if fid < 0
        error('mat2doc:test:openFail', 'cannot open frozen oracle %s', p);
    end
    raw = fread(fid, Inf, '*uint8')';
    fclose(fid);
    o = jsondecode(native2unicode(raw, 'UTF-8'));
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

% ---- s0031 probe replay (for the Equivalence leg) --------------------- %

function P = runProbes()
    % Replay the s0031 probe sequence (the .m twin's P-struct body, VERBATIM
    % tags/inputs/order) and return the nested struct of tagged canonical values.
    % Embedded so the Equivalence leg is self-contained (the validation-folder
    % scenario is NOT on the toolbox path and must not be a dependency). Mirrors
    % validation\mat2doc\scenarios\s0031_p4_7a_styles_probe.m lines 19-226.
    P = struct();
    ST = @(name) mat2doc.enum.style.WD_STYLE_TYPE.(name);

    % ===================== babelfish: H15 case-sensitive ===================
    bf = struct();
    pairs = [ ...
        "Caption","caption"; "Footer","footer"; "Header","header"; ...
        "Heading 1","heading 1"; "Heading 2","heading 2"; "Heading 3","heading 3"; ...
        "Heading 4","heading 4"; "Heading 5","heading 5"; "Heading 6","heading 6"; ...
        "Heading 7","heading 7"; "Heading 8","heading 8"; "Heading 9","heading 9"];
    u2i = struct(); i2u = struct();
    for i = 0:size(pairs, 1) - 1
        k = sprintf("p%02d", i);
        u2i.(k) = mat2doc.styles.BabelFish.ui2internal(pairs(i + 1, 1));
        i2u.(k) = mat2doc.styles.BabelFish.internal2ui(pairs(i + 1, 2));
    end
    bf.ui2internal = u2i;
    bf.internal2ui = i2u;
    bf.passthrough_ui = mat2doc.styles.BabelFish.ui2internal("Totally Unknown");
    bf.passthrough_internal = mat2doc.styles.BabelFish.internal2ui("totally unknown");
    bf.case_trap_ui = mat2doc.styles.BabelFish.ui2internal("heading 1");
    bf.case_trap_internal = mat2doc.styles.BabelFish.internal2ui("Heading 1");
    bf.empty_ui = mat2doc.styles.BabelFish.ui2internal("");
    P.babelfish = bf;

    % ===================== factory: StyleFactory dispatch ==================
    fa = struct();
    fcases = {"paragraph","PARAGRAPH"; "character","CHARACTER"; ...
              "table","TABLE"; "numbering","LIST"};
    for k = 1:size(fcases, 1)
        se = mat2doc.oxml.OxmlElement("w:style");
        se.type = ST(fcases{k, 2});
        fa.(fcases{k, 1}) = clsTok(mat2doc.styles.StyleFactory(se));
    end
    bare = mat2doc.oxml.OxmlElement("w:style");
    fa.none_type_raises = raisesTok(@() mat2doc.styles.StyleFactory(bare));
    fa.none_type_msg = errmsgTok(@() mat2doc.styles.StyleFactory(bare));
    P.factory = fa;

    % ===================== basestyle: every property =======================
    bs = struct();
    se = parse("<w:style " + nsW() + "/>");
    b = mat2doc.styles.ParagraphStyle(se);
    bs.bare = struct("cls", clsTok(b), ...
        "builtin", rvTok(b.builtin), "hidden", rvTok(b.hidden), "locked", rvTok(b.locked), ...
        "name", rvTok(b.name), "priority", rvTok(b.priority), ...
        "quick_style", rvTok(b.quick_style), "style_id", rvTok(b.style_id), ...
        "type", rvTok(b.type), "unhide_when_used", rvTok(b.unhide_when_used));
    b.name = "Heading 1"; b.style_id = "Heading1"; b.priority = 9;
    b.hidden = true; b.locked = true; b.quick_style = true; b.unhide_when_used = true;
    bs.set = struct("name", rvTok(b.name), "name_internal", rvTok(se.name_val), ...
        "style_id", rvTok(b.style_id), "priority", rvTok(b.priority), ...
        "hidden", rvTok(b.hidden), "locked", rvTok(b.locked), ...
        "quick_style", rvTok(b.quick_style), "unhide_when_used", rvTok(b.unhide_when_used), ...
        "builtin", rvTok(b.builtin));
    b.hidden = false; b.locked = false; b.quick_style = false;
    b.unhide_when_used = false; b.priority = [];
    bs.cleared = struct("hidden", rvTok(b.hidden), "locked", rvTok(b.locked), ...
        "quick_style", rvTok(b.quick_style), "unhide_when_used", rvTok(b.unhide_when_used), ...
        "priority", rvTok(b.priority));
    se2 = parse("<w:style " + nsW() + "/>");
    b2 = mat2doc.styles.ParagraphStyle(se2);
    b2.name = "heading 1";
    bs.name_internal_form = struct("getter", rvTok(b2.name), "stored", rvTok(se2.name_val));
    bt = mat2doc.styles.ParagraphStyle(parse("<w:style w:customStyle=""1"" " + nsW() + "/>"));
    bs.builtin_customstyle_true = rvTok(bt.builtin);
    P.basestyle = bs;

    % ===================== base_chain: base_style + font ===================
    styles = parse("<w:styles " + nsW() + ">" + ...
        styleXml("Normal", "paragraph") + ...
        styleXml("Heading1", "paragraph", "<w:basedOn w:val=""Normal""/>") + ...
        styleXml("Dangling", "paragraph", "<w:basedOn w:val=""NoSuchZzz""/>") + ...
        "</w:styles>");
    h1 = mat2doc.styles.StyleFactory(styles.get_by_id("Heading1"));
    dang = mat2doc.styles.StyleFactory(styles.get_by_id("Dangling"));
    norm = mat2doc.styles.StyleFactory(styles.get_by_id("Normal"));
    bc = struct("base_style_cls", clsTok(h1.base_style), ...
        "base_style_id", rvTok(h1.base_style.style_id), ...
        "base_style_none_no_basedon", clsTok(norm.base_style), ...
        "base_style_dangling_none", clsTok(dang.base_style), ...
        "font_cls", clsTok(h1.font));
    norm.base_style = h1;
    bc.set_base_style_id = rvTok(norm.base_style.style_id);
    norm.base_style = [];
    bc.set_base_style_none = clsTok(norm.base_style);
    P.base_chain = bc;

    % ===================== para_style: next_paragraph_style ================
    ps = struct();
    styles = parse("<w:styles " + nsW() + ">" + ...
        styleXml("Normal", "paragraph") + ...
        styleXml("BodyChar", "character") + ...
        styleXml("H_none", "paragraph") + ...
        styleXml("H_dangling", "paragraph", "<w:next w:val=""NoSuchZzz""/>") + ...
        styleXml("H_nonpara", "paragraph", "<w:next w:val=""BodyChar""/>") + ...
        styleXml("H_real", "paragraph", "<w:next w:val=""Normal""/>") + ...
        "</w:styles>");
    para = @(sid) mat2doc.styles.StyleFactory(styles.get_by_id(sid));
    ps.none_self = rvTok(para("H_none").next_paragraph_style.style_id);
    ps.dangling_self = rvTok(para("H_dangling").next_paragraph_style.style_id);
    ps.nonpara_self = rvTok(para("H_nonpara").next_paragraph_style.style_id);
    ps.real = rvTok(para("H_real").next_paragraph_style.style_id);
    ps.real_cls = clsTok(para("H_real").next_paragraph_style);
    hr = para("H_real");
    hr.next_paragraph_style = [];
    ps.set_none_then_get = rvTok(hr.next_paragraph_style.style_id);
    hr.next_paragraph_style = para("Normal");
    ps.set_other_then_get = rvTok(hr.next_paragraph_style.style_id);
    hr.next_paragraph_style = hr;
    ps.set_self_then_get = rvTok(hr.next_paragraph_style.style_id);
    ps.pfmt_cls = clsTok(para("Normal").paragraph_format);
    P.para_style = ps;

    % ===================== styles: over a REAL Document() ==================
    d = mat2doc.Document();
    st = d.styles;
    sd = struct();
    sd.len = rvTok(st.len_());
    sd.contains_heading1_ui = rvTok(st.contains_("Heading 1"));
    sd.contains_normal = rvTok(st.contains_("Normal"));
    sd.contains_missing = rvTok(st.contains_("No Such Style Zzz"));
    sd.getitem_ui_cls = clsTok(st.getitem_("Heading 1"));
    sd.getitem_ui_id = rvTok(st.getitem_("Heading 1").style_id);
    ow = warning("off", "mat2doc:UserWarning");
    sd.getitem_id_deprecated_id = rvTok(st.getitem_("Heading1").style_id);
    warning(ow);
    sd.getitem_missing_raises = raisesTok(@() st.getitem_("No Such Style Zzz"));
    sd.getitem_missing_msg = errmsgTok(@() st.getitem_("No Such Style Zzz"));
    lst = st.to_array();
    sd.iter_len = rvTok(numel(lst));
    sd.iter_first_cls = clsTok(lst(1));
    sd.iter_first_id = rvTok(lst(1).style_id);
    sd.iter_last_cls = clsTok(lst(end));
    sd.iter_last_id = rvTok(lst(end).style_id);
    sd.default_para_cls = clsTok(st.default(ST("PARAGRAPH")));
    sd.default_para_id = rvTok(st.default(ST("PARAGRAPH")).style_id);
    sd.default_char_cls = clsTok(st.default(ST("CHARACTER")));
    sd.get_by_id_found_id = rvTok(st.get_by_id("Heading1", ST("PARAGRAPH")).style_id);
    sd.get_by_id_none_default = rvTok(st.get_by_id([], ST("PARAGRAPH")).style_id);
    sd.get_by_id_empty_default = rvTok(st.get_by_id("", ST("PARAGRAPH")).style_id);
    sd.get_by_id_missing_default = rvTok(st.get_by_id("Nope", ST("PARAGRAPH")).style_id);
    sd.get_by_id_wrongtype_default = rvTok(st.get_by_id("Heading1", ST("CHARACTER")).style_id);
    sd.get_style_id_none = rvTok(st.get_style_id([], ST("PARAGRAPH")));
    sd.get_style_id_name = rvTok(st.get_style_id("Heading 1", ST("PARAGRAPH")));
    sd.get_style_id_style = rvTok(st.get_style_id(st.getitem_("Heading 1"), ST("PARAGRAPH")));
    sd.get_style_id_default_name = rvTok(st.get_style_id("Normal", ST("PARAGRAPH")));
    sd.get_style_id_wrongtype_raises = raisesTok(@() st.get_style_id("Heading 1", ST("CHARACTER")));
    sd.get_style_id_wrongtype_msg = errmsgTok(@() st.get_style_id("Heading 1", ST("CHARACTER")));
    sd.get_style_id_missing_raises = raisesTok(@() st.get_style_id("No Such Style Zzz", ST("PARAGRAPH")));
    sd.get_style_id_missing_msg = errmsgTok(@() st.get_style_id("No Such Style Zzz", ST("PARAGRAPH")));
    d2 = mat2doc.Document();
    cst = d2.styles.add_style("My Custom Style", ST("CHARACTER"));
    bst = d2.styles.add_style("My Builtin Style", ST("PARAGRAPH"), true);
    sd.add_custom_cls = clsTok(cst);
    sd.add_custom_id = rvTok(cst.style_id);
    sd.add_custom_builtin = rvTok(cst.builtin);
    sd.add_builtin_cls = clsTok(bst);
    sd.add_builtin_id = rvTok(bst.style_id);
    sd.add_builtin_builtin = rvTok(bst.builtin);
    sd.add_dup_raises = raisesTok(@() d2.styles.add_style("Heading 1", ST("PARAGRAPH")));
    sd.add_dup_msg = errmsgTok(@() d2.styles.add_style("Heading 1", ST("PARAGRAPH")));
    sd.element_localname = string(st.element().local_part);
    P.styles = sd;

    % ===================== dump: 164-row template style dump ===============
    d = mat2doc.Document();
    arr = d.styles.to_array();
    rows = struct();
    for i = 0:numel(arr) - 1
        s = arr(i + 1);
        row = struct("cls", clsTok(s), "style_id", rvTok(s.style_id), ...
            "name", rvTok(s.name), "type", rvTok(s.type), "builtin", rvTok(s.builtin), ...
            "hidden", rvTok(s.hidden), "locked", rvTok(s.locked), ...
            "priority", rvTok(s.priority), "quick_style", rvTok(s.quick_style), ...
            "unhide_when_used", rvTok(s.unhide_when_used));
        if isa(s, "mat2doc.styles.CharacterStyle")
            bso = s.base_style;
            if isequal(bso, [])
                row.base_style_id = "None";
            else
                row.base_style_id = rvTok(bso.style_id);
            end
        else
            row.base_style_id = "N/A";
        end
        if isa(s, "mat2doc.styles.ParagraphStyle")
            row.next_para_style_id = rvTok(s.next_paragraph_style.style_id);
        else
            row.next_para_style_id = "N/A";
        end
        rows.(sprintf("s%03d", i)) = row;
    end
    P.dump = rows;

    % ===================== byteproofs: delete_ + add_style =================
    bp = struct();
    d = mat2doc.Document();
    st = d.styles;
    h = st.getitem_("Heading 1");
    h.delete_();
    bp.delete_heading1 = shaOfElement(st.element());
    d = mat2doc.Document();
    st = d.styles;
    st.add_style("My Fancy Style", ST("CHARACTER"));
    st.add_style("My Builtin Style", ST("PARAGRAPH"), true);
    bp.add_style_custom_builtin = shaOfElement(st.element());
    P.byteproofs = bp;
end

function s = clsTok(obj)
    % Canonical class token: leaf class name, package prefix stripped, trailing
    % '_' stripped (TableStyle_ -> 'TableStyle'). [] (None) -> 'None'. Mirrors the
    % s0031 cls().
    if isequal(obj, [])
        s = "None"; return
    end
    c = string(class(obj));
    parts = split(c, ".");
    s = parts(end);
    if endsWith(s, "_")
        s = extractBefore(s, strlength(s));
    end
end

function s = rvTok(x)
    % Uniform accessor repr mirroring the s0031 rv(): None->"None", enum->member
    % name, bool->"True"/"False", int->decimal string.
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

function s = raisesTok(fn)
    try
        fn();
        s = "False";
    catch
        s = "True";
    end
end

function s = errmsgTok(fn)
    try
        fn();
        s = "NO-RAISE";
    catch ME
        s = string(ME.message);
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

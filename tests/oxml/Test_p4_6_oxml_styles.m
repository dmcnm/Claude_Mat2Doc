classdef Test_p4_6_oxml_styles < matlab.unittest.TestCase
% TEST_P4_6_OXML_STYLES  Gate-4 permanent unit tests for Mat2Doc P4-6
%   (the oxml layer for STYLES): src/docx/oxml/styles.py ->
%   +mat2doc\+oxml\+styles\{CT_Styles, CT_Style, CT_LatentStyles,
%   CT_LsdException, styleId_from_name}, src/docx/oxml/shared.py::CT_DecimalNumber
%   -> +mat2doc\+oxml\+shared\CT_DecimalNumber, plus the 12 styles-block registry
%   rows + the closed w:outlineLvl deferral in +mat2doc\+oxml\registry.m.
%
%   P4-6 is a REGISTRY-ADDING, byte-critical parse-path WP: registering the 12
%   styles tags flips the PARSE CLASS of every real w:style (164) / w:latentStyles
%   / w:lsdException (137) / w:uiPriority in styles.xml from generic XmlElement to
%   the new CT_* classes. All CT_* exit through the identical +oxml\
%   serialize_part_xml walk, so the flip is byte-neutral -- word/styles.xml stays
%   349458 B, SHA-256 02d71a68... This class permanently freezes the guarantees
%   the prior gates established:
%     * Gate-1 Porter  : audit_P4-6_oxml_styles.md (self-probe 75/75).
%     * Gate-2 Auditor (Fable): audit_P4-6_oxml_styles.md GATE-2 section --
%       REVISE on ONE H4 defect F-1 (set_bool_prop(None) must WRITE w:def..="0",
%       not remove) -> Porter one-line fix -> re-verified. Everything else
%       APPROVED, incl. the delete()/H17 ruling and the 12-row registry scope.
%     * Gate-3 Validator: validate_P4-6_oxml_styles.md -- PASS, ZERO new
%       D-numbers: probe_diff s0030 MATCH exit 0, M1 17/17 (styles.xml 349458 B
%       L1), H11 scrambled-order parity byte-identical for all 3 container
%       classes, the 349 KB styles.xml parse-path round-trip byte-identical,
%       F-1 fix oracle-proven, targeted regression 237/237.
%
%   THE HIGHEST-VALUE PERMANENT PINS IN THIS WP (a regression goes RED loudly):
%     * F-1 (test_ct_latentstyles_bool_prop_truth_table): set_bool_prop(attr,[])
%       -- Python bool(None) is False -- MUST serialize @w:defQFormat="0"
%       (attribute PRESENT, not removed) and round-trip bool_prop -> false. This
%       is the exact defect Gate-2 caught; the pin is deliberately loud + commented.
%     * H11 child-ORDER pins (test_h11_child_order_pins): CT_Style 10-descriptor
%       reverse-scramble -> canonical name..rPr; CT_Styles latentStyles sorts
%       BEFORE an already-present style; CT_LatentStyles lsdException appends. Add
%       children OUT OF ORDER, assert the serialized BYTES are canonical schema
%       order. Wrong order -> Word repair. Byte-pinned (serhex) vs the oracle.
%     * delete_* PARENT-SIDE-EFFECT ONLY (test_h17_delete_parent_side_only):
%       CT_Style.delete_style() / CT_LsdException.delete_lsd_exception() remove
%       the element from its parent -- assert the PARENT's child count / serialized
%       bytes only. These methods are named by the kind of thing they remove (not
%       `delete`), so MATLAB's handle destructor is never involved (H17 dissolved).
%       The test still never inspects the removed handle afterward.
%     * M1 styles.xml SHA pin (test_m1_styles_xml_byte_identical): the
%       registry-adding parse-path regression guard -- mat2doc.Document().save()
%       -> word/styles.xml == 349458 B, SHA-256 02d71a68...e384.
%
%   Provenance (all 2026-07-30):
%     * Audit    : validation\mat2doc\audit_P4-6_oxml_styles.md
%     * Validate : validation\mat2doc\validate_P4-6_oxml_styles.md
%     * Scenario : validation\mat2doc\scenarios\s0030_p4_6_styles_probe.{py,m}
%                  (its probe body is replayed VERBATIM by runProbes() below).
%     * Frozen ref (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE at Gate-3):
%         references\s0030\probe.json -- copied verbatim (self-contained) into
%           tests\oxml\data\s0030_probe_oracle.json (the Equivalence replay set;
%           value/serhex JSON, jsondecode is line-ending agnostic -> no
%           `* binary` .gitattributes needed, per the s0022/s0023 precedent).
%         references\s0001\parts\word\styles.xml -- the M1 byte reference; NOT
%           copied (the byte pin compares SHA-256 of what Document().save() itself
%           emits, so no fixture is needed).
%
%   Coverage taxonomy
%   -----------------
%   * Nominal      -- styleId_from_name happy path, CT_DecimalNumber.new, CT_Style
%                     attrs/val accessors, CT_Styles.add_style_of_type builtin,
%                     CT_LatentStyles set_bool_prop true, CT_LsdException set.
%   * Edge         -- empty text "" -> "" (styleId), non-ASCII round-trip, bare
%                     w:style tri-state (None vs False), val setter remove, H15
%                     case-sensitive "Heading 1"->"Heading1", missing REQUIRED attr
%                     error paths (mat2doc:InvalidXmlError), default_for last-wins /
%                     default="0"-excluded / None, F-1 None case.
%   * Equivalence  -- test_equivalence_full_battery_vs_frozen_oracle replays the
%                     ENTIRE s0030 battery (runProbes, the .m twin's body verbatim,
%                     incl. the real styles.xml parse->serialize leg) and
%                     flatten-compares every leaf to the frozen oracle (0 diffs).
%   * Regression   -- hard-coded expected serialized-XML strings (ASCII, so
%                     string-equality == byte-identical L1) + UPPERCASE serhex of
%                     the raw UTF-8 bytes vs the frozen oracle + SHA-256 of the M1
%                     styles.xml part.
%   * Upstream     -- the H11 successor-slice ordering, styleId_from_name's
%                     special-case table, and the add_style_of_type attr order are
%                     the python-docx styles.py surface; the frozen oracle IS
%                     lxml's expected output for this API sequence.
%
%   Byte-level (L1) note: every serialized-XML comparison is either the FULL
%   serialize_part_xml output decoded as an ASCII string (string-equality ==
%   byte-equality L1) or its UPPERCASE hex (serhex) vs the frozen oracle. No
%   D-number granted any L2 relaxation in this WP (Gate-3: zero new, none at L2),
%   so every pin here is L1. The parse-count guards (n_style==164 etc.) and the
%   equivalence key-count guard are the only looser-than-byte checks and are
%   commented at their site.
%
%   Determinism: no network, no absolute paths. The worktree root and the
%   co-located oracle resolve relative to this file via
%   fileparts(mfilename('fullpath')); saves go to tempname .docx deleted via
%   onCleanup; every file read is binary ('r','n'). The +mat2doc package resolves
%   via the MANDATORY PathFixture(worktree-root) in TestClassSetup (WP9-F4 lesson).

    properties (Constant)
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        % --- frozen s0001 M1 byte reference (the P4-6 parse-path risk) ---
        STYLES_SIZE = 349458
        STYLES_SHA  = "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"

        % --- canonical child order of CT_Style's 10-descriptor reverse scramble
        %     (H11): any scrambled add-order must converge to exactly this
        %     sequence, driven by the NON-CONTIGUOUS _tag_seq successor slices. ---
        STYLE_SCHEMA_ORDER = ["name" "basedOn" "next" "uiPriority" "semiHidden" ...
            "unhideWhenUsed" "qFormat" "locked" "pPr" "rPr"]
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
        % 1. styleId_from_name (H15 case-sensitive 11-vector)             %
        % =============================================================== %

        function test_styleId_from_name_h15_vector(testCase)
            % Nominal + Edge + Regression (H15, s0030 styleid): the 11-vector,
            % including the H15 crux "Heading 1"->"Heading1" (capital-H misses the
            % lowercase table key, falls to the space-strip default), the empty
            % "" -> "", and multi-space "  double  spaced  " -> "doublespaced".
            % Hard-coded expected results AND vs the frozen oracle.
            fn = @(s) mat2doc.oxml.styles.styleId_from_name(s);

            testCase.verifyEqual(fn("caption"),   "Caption",  'table hit "caption"');
            testCase.verifyEqual(fn("heading 1"), "Heading1", 'table hit "heading 1"');
            testCase.verifyEqual(fn("heading 9"), "Heading9", 'table hit "heading 9"');
            testCase.verifyEqual(fn("Heading 1"), "Heading1", ...
                'H15: capital-H MISSES the lowercase table -> space-strip default (same result)');
            testCase.verifyEqual(fn("Caption"),   "Caption",  'H15: "Caption" misses table -> strrep default');
            testCase.verifyEqual(fn("Normal"),    "Normal",   'not in table -> strrep (no spaces)');
            testCase.verifyEqual(fn("heading 10"),"heading10",'"heading 10" NOT in table -> strrep');
            testCase.verifyEqual(fn(""),          "",         'empty text -> empty');
            testCase.verifyEqual(fn("  double  spaced  "), "doublespaced", 'all spaces stripped');
            testCase.verifyEqual(fn("My Custom Style"),    "MyCustomStyle");
            testCase.verifyEqual(fn("heading 5"), "Heading5");

            % vs the frozen oracle (k00..k10, same order as s0030)
            oracle = loadOracle();
            expect = ["Caption" "Heading1" "Heading9" "Heading1" "Caption" "Normal" ...
                "heading10" "" "doublespaced" "MyCustomStyle" "Heading5"];
            for i = 0:numel(expect)-1
                key = "k" + sprintf("%02d", i);
                testCase.verifyEqual(string(oracle.styleid.(key)), expect(i+1), ...
                    sprintf('oracle styleid.%s', key));
            end

            % Edge: non-ASCII default branch (é / CJK / emoji) survives strrep.
            testCase.verifyEqual(fn("Héading Ünïcode"), "HéadingÜnïcode");
            testCase.verifyEqual(fn("段落 样式"), "段落样式");
        end

        % =============================================================== %
        % 2. CT_DecimalNumber.new (raw @w:val bytes + missing-val raise)  %
        % =============================================================== %

        function test_ct_decimalnumber_new(testCase)
            % Nominal + Edge + Regression (H14/H6, s0030 decnum): CT_DecimalNumber.new
            % writes @w:val = str(val) DIRECTLY (no range-validate); raw bytes for
            % 9 / -3 / 0 hard-coded AND serhex vs the frozen oracle; val round-trips;
            % missing @w:val on a bare w:uiPriority RAISES mat2doc:InvalidXmlError.
            oracle = loadOracle();

            nine = mat2doc.oxml.shared.CT_DecimalNumber.new("w:uiPriority", 9);
            testCase.verifyEqual(ser(nine), decl() + newline + ...
                "<w:uiPriority xmlns:w=""" + testCase.W + """ w:val=""9""/>", ...
                'CT_DecimalNumber.new(9) raw @w:val bytes (L1) hard-coded');
            testCase.verifyEqual(hx_e(nine), string(oracle.decnum.nine.serhex));
            testCase.verifyEqual(nine.val, 9);

            neg = mat2doc.oxml.shared.CT_DecimalNumber.new("w:uiPriority", -3);
            testCase.verifyEqual(ser(neg), decl() + newline + ...
                "<w:uiPriority xmlns:w=""" + testCase.W + """ w:val=""-3""/>", ...
                'signed -3 written verbatim');
            testCase.verifyEqual(hx_e(neg), string(oracle.decnum.negthree.serhex));
            testCase.verifyEqual(neg.val, -3);

            zero = mat2doc.oxml.shared.CT_DecimalNumber.new("w:uiPriority", 0);
            testCase.verifyEqual(hx_e(zero), string(oracle.decnum.zero.serhex));
            testCase.verifyEqual(zero.val, 0);

            % Error path: bare w:uiPriority (CT_DecimalNumber) -> get.val raises.
            bare = mat2doc.oxml.OxmlElement("w:uiPriority");
            testCase.verifyError(@() bare.val, 'mat2doc:InvalidXmlError', ...
                'missing REQUIRED @w:val -> mat2doc:InvalidXmlError (identifier pinned)');
        end

        % =============================================================== %
        % 3. CT_Style attributes + H3 val tri-state                       %
        % =============================================================== %

        function test_ct_style_attrs_and_val_tristate(testCase)
            % Nominal + Edge + Regression (H3, s0030 ct_style): bare w:style has all
            % four attrs None and the val quartet split None (basedOn/name/uiPriority)
            % vs False (locked/qFormat/semiHidden/unhideWhenUsed); attrs_set order
            % type->styleId->default->customStyle; the _val setters ADD then REMOVE
            % (localnames []); the bool-ish quartet True->bare child / False->removed.
            oracle = loadOracle();

            % -- bare: tri-state None vs the CT_OnOff-derived False (H3) --
            b = mat2doc.oxml.OxmlElement("w:style");
            testCase.verifyTrue(isequal(b.type, []),        'bare type -> [] (None)');
            testCase.verifyTrue(isequal(b.styleId, []),     'bare styleId -> [] (None)');
            testCase.verifyTrue(isequal(b.default, []),     'bare default -> [] (None)');
            testCase.verifyTrue(isequal(b.customStyle, []), 'bare customStyle -> [] (None)');
            testCase.verifyTrue(isequal(b.basedOn_val, []), 'bare basedOn_val -> [] (None)');
            testCase.verifyTrue(isequal(b.name_val, []),    'bare name_val -> [] (None)');
            testCase.verifyTrue(isequal(b.uiPriority_val, []), 'bare uiPriority_val -> [] (None)');
            testCase.verifyFalse(b.locked_val,         'bare locked_val -> logical FALSE (not None)');
            testCase.verifyFalse(b.qFormat_val,        'bare qFormat_val -> FALSE');
            testCase.verifyFalse(b.semiHidden_val,     'bare semiHidden_val -> FALSE');
            testCase.verifyFalse(b.unhideWhenUsed_val, 'bare unhideWhenUsed_val -> FALSE');

            % -- attrs_set: type -> styleId -> default -> customStyle --
            s = mat2doc.oxml.OxmlElement("w:style");
            s.type = mat2doc.enum.style.WD_STYLE_TYPE.PARAGRAPH;
            s.styleId = "MyStyle"; s.default = true; s.customStyle = true;
            testCase.verifyEqual(ser(s), decl() + newline + ...
                "<w:style xmlns:w=""" + testCase.W + """ w:type=""paragraph"" " + ...
                "w:styleId=""MyStyle"" w:default=""1"" w:customStyle=""1""/>", ...
                'CT_Style attrs_set serialized bytes (L1): type->styleId->default->customStyle');
            testCase.verifyEqual(hx_e(s), string(oracle.ct_style.attrs_set.serhex));
            testCase.verifyEqual(string(s.type), "PARAGRAPH");
            testCase.verifyEqual(s.styleId, "MyStyle");
            testCase.verifyTrue(s.default); testCase.verifyTrue(s.customStyle);

            % -- _val setters: ADD (name/basedOn/uiPriority, H11-ordered) --
            s = mat2doc.oxml.OxmlElement("w:style");
            s.name_val = "heading 1"; s.uiPriority_val = 9; s.basedOn_val = "Normal";
            testCase.verifyEqual(childLocalnames(s), ["name" "basedOn" "uiPriority"], ...
                'val setters add children in canonical H11 order');
            testCase.verifyEqual(hx_e(s), string(oracle.ct_style.val_set.serhex));
            testCase.verifyEqual(s.name_val, "heading 1");
            testCase.verifyEqual(s.uiPriority_val, 9);
            testCase.verifyEqual(s.basedOn_val, "Normal");
            % -- ... then REMOVE (set-to-[] removes the child) --
            s.name_val = []; s.uiPriority_val = []; s.basedOn_val = [];
            testCase.verifyEqual(childLocalnames(s), strings(1,0), 'val=[] removes each child');
            testCase.verifyEqual(hx_e(s), string(oracle.ct_style.val_cleared.serhex));

            % -- bool-ish quartet: True -> bare child (H11 order), False -> removed --
            s = mat2doc.oxml.OxmlElement("w:style");
            s.locked_val = true; s.semiHidden_val = true; s.qFormat_val = true; s.unhideWhenUsed_val = true;
            testCase.verifyEqual(childLocalnames(s), ["semiHidden" "unhideWhenUsed" "qFormat" "locked"], ...
                'bool-ish True adds bare children in canonical order');
            testCase.verifyEqual(hx_e(s), string(oracle.ct_style.bool_true.serhex));
            testCase.verifyTrue(s.locked_val); testCase.verifyTrue(s.qFormat_val);
            testCase.verifyTrue(s.semiHidden_val); testCase.verifyTrue(s.unhideWhenUsed_val);
            s.locked_val = false; s.semiHidden_val = false; s.unhideWhenUsed_val = false;
            testCase.verifyEqual(childLocalnames(s), "qFormat", ...
                'False removes the child; qFormat (left True) remains');
            testCase.verifyEqual(hx_e(s), string(oracle.ct_style.bool_false_removed.serhex));
        end

        % =============================================================== %
        % 4. CT_Style chains + CT_Styles lookups (BOTH xpath forms)       %
        % =============================================================== %

        function test_ct_style_chain_and_lookups(testCase)
            % Nominal + Edge (s0030 ct_style.chain): over a built styles tree,
            % basedOn_val / uiPriority_val / name_val read the child .val;
            % base_style + next_style resolve a sibling via get_by_id
            % ([@w:styleId=] xpath form); get_by_name uses the [w:name/@w:val=] form.
            % Misses -> [] (None).
            styles = parse("<w:styles " + nsW() + ">" + ...
                "<w:style w:type=""paragraph"" w:styleId=""Normal""><w:name w:val=""Normal""/></w:style>" + ...
                "<w:style w:type=""paragraph"" w:styleId=""Heading1"">" + ...
                "<w:name w:val=""heading 1""/><w:basedOn w:val=""Normal""/>" + ...
                "<w:next w:val=""Normal""/><w:uiPriority w:val=""9""/></w:style>" + ...
                "</w:styles>");
            h1 = styles.get_by_id("Heading1");   % [@w:styleId="Heading1"] form
            testCase.verifyEqual(h1.basedOn_val, "Normal");
            testCase.verifyEqual(h1.uiPriority_val, 9);
            testCase.verifyEqual(h1.name_val, "heading 1");
            testCase.verifyEqual(h1.base_style.styleId, "Normal", 'base_style resolves via get_by_id');
            testCase.verifyEqual(h1.next_style.styleId, "Normal", 'next_style resolves via get_by_id');
            % get_by_name uses the [w:name/@w:val="..."] predicate sub-path form
            testCase.verifyEqual(styles.get_by_name("heading 1").styleId, "Heading1", ...
                'get_by_name uses the [w:name/@w:val=] xpath form');
            testCase.verifyTrue(isequal(styles.get_by_id("Nope"), []),   'get_by_id miss -> []');
            testCase.verifyTrue(isequal(styles.get_by_name("nope"), []), 'get_by_name miss -> []');
        end

        % =============================================================== %
        % 5. CT_Styles.add_style_of_type + default_for                    %
        % =============================================================== %

        function test_ct_styles_add_style_of_type_and_default_for(testCase)
            % Nominal + Edge + Regression (H1/H4, s0030 styles): add_style_of_type
            % builtin=true OMITS @w:customStyle; builtin=false writes
            % @w:customStyle="1" with attr order type->customStyle->styleId (the
            % Gate-2 VERIFY-2 pin). default_for returns the LAST default in document
            % order (H4 last-wins), EXCLUDES @w:default="0", and returns [] (None)
            % when no default of that type exists. Bytes hard-coded AND vs oracle.
            oracle = loadOracle();
            PARA = mat2doc.enum.style.WD_STYLE_TYPE.PARAGRAPH;
            CHAR = mat2doc.enum.style.WD_STYLE_TYPE.CHARACTER;
            TABLE = mat2doc.enum.style.WD_STYLE_TYPE.TABLE;

            % builtin=true -> customStyle ABSENT
            root = mat2doc.oxml.OxmlElement("w:styles");
            builtin = root.add_style_of_type("heading 1", PARA, true);
            testCase.verifyEqual(ser(builtin), decl() + newline + ...
                "<w:style xmlns:w=""" + testCase.W + """ w:type=""paragraph"" " + ...
                "w:styleId=""Heading1""><w:name w:val=""heading 1""/></w:style>", ...
                'builtin=true OMITS w:customStyle (L1)');
            testCase.verifyEqual(hx_e(builtin), string(oracle.styles.add_builtin.serhex));
            testCase.verifyEqual(builtin.styleId, "Heading1");
            testCase.verifyTrue(isequal(builtin.customStyle, []), 'builtin customStyle -> [] (None)');
            testCase.verifyEqual(string(builtin.type), "PARAGRAPH");
            testCase.verifyEqual(builtin.name_val, "heading 1");

            % builtin=false -> w:customStyle="1", attr order type->customStyle->styleId
            root2 = mat2doc.oxml.OxmlElement("w:styles");
            custom = root2.add_style_of_type("My Custom Style", CHAR, false);
            testCase.verifyEqual(ser(custom), decl() + newline + ...
                "<w:style xmlns:w=""" + testCase.W + """ w:type=""character"" " + ...
                "w:customStyle=""1"" w:styleId=""MyCustomStyle"">" + ...
                "<w:name w:val=""My Custom Style""/></w:style>", ...
                'builtin=false attr order type->customStyle->styleId (L1, Gate-2 VERIFY-2)');
            testCase.verifyEqual(hx_e(custom), string(oracle.styles.add_custom.serhex));
            testCase.verifyEqual(custom.styleId, "MyCustomStyle");
            testCase.verifyTrue(custom.customStyle);
            testCase.verifyEqual(string(custom.type), "CHARACTER");

            % default_for: last-wins (B beats A), default="0" excluded (D), None for TABLE
            styles = parse("<w:styles " + nsW() + ">" + ...
                "<w:style w:type=""paragraph"" w:styleId=""A"" w:default=""1""><w:name w:val=""A""/></w:style>" + ...
                "<w:style w:type=""paragraph"" w:styleId=""B"" w:default=""1""><w:name w:val=""B""/></w:style>" + ...
                "<w:style w:type=""paragraph"" w:styleId=""D"" w:default=""0""><w:name w:val=""D""/></w:style>" + ...
                "<w:style w:type=""character"" w:styleId=""C"" w:default=""1""><w:name w:val=""C""/></w:style>" + ...
                "</w:styles>");
            testCase.verifyEqual(styles.default_for(PARA).styleId, "B", ...
                'default_for LAST-in-order (B), with @w:default="0" (D) excluded (H4)');
            testCase.verifyEqual(styles.default_for(CHAR).styleId, "C");
            testCase.verifyTrue(isequal(styles.default_for(TABLE), []), ...
                'no TABLE default -> [] (None)');
        end

        % =============================================================== %
        % 6. CT_LatentStyles bool-prop truth table (incl. the F-1 pin)    %
        % =============================================================== %

        function test_ct_latentstyles_bool_prop_truth_table(testCase)
            % Regression (H4/H3, s0030 latent) -- THE F-1 REGRESSION PIN.
            % set_bool_prop truth table: true->"1", false->"0", 1->"1", 0->"0",
            % absent->false. The [] (None) case is the DEFECT Gate-2 caught: Python
            % bool(None) is False, so set_bool_prop(attr,[]) MUST WRITE
            % @w:defQFormat="0" (attribute PRESENT, NOT removed) and round-trip
            % bool_prop -> false. A naive logical([]) empty-logical would REMOVE the
            % attribute (wrong). If this pin goes RED, the F-1 fix has regressed.
            oracle = loadOracle();

            % bool_prop on an absent attribute -> false (effective default)
            testCase.verifyFalse(mat2doc.oxml.OxmlElement("w:latentStyles").bool_prop("defQFormat"), ...
                'bool_prop absent -> false');

            % ---- the truth table (fresh element each row) ----
            cases = {true, "1"; false, "0"; 1, "1"; 0, "0"};
            oKeys = {'true','false','one','zero'};
            for i = 1:size(cases,1)
                ls = mat2doc.oxml.OxmlElement("w:latentStyles");
                ls.set_bool_prop("defQFormat", cases{i,1});
                expected = decl() + newline + "<w:latentStyles xmlns:w=""" + testCase.W + ...
                    """ w:defQFormat=""" + cases{i,2} + """/>";
                testCase.verifyEqual(ser(ls), expected, ...
                    sprintf('set_bool_prop truth-table row %s serialized (L1)', oKeys{i}));
                testCase.verifyEqual(hx_e(ls), string(oracle.latent.set_bool_prop.(oKeys{i}).serhex));
            end

            % ===================================================================== %
            %  *** F-1 REGRESSION PIN -- set_bool_prop([]) MUST WRITE "0" ***        %
            %  Python `setattr(self, attr, bool(value))` with bool(None) is False,   %
            %  so styles.py WRITES w:defQFormat="0"; it does NOT remove the attr.    %
            %  Gate-2 (Fable) caught the original port removing it. This pin makes    %
            %  the fix permanent -- if it reverts, THIS goes RED loudly.             %
            % ===================================================================== %
            ls = mat2doc.oxml.OxmlElement("w:latentStyles");
            ls.set_bool_prop("defQFormat", []);   % [] == Python None
            f1Expected = decl() + newline + "<w:latentStyles xmlns:w=""" + testCase.W + ...
                """ w:defQFormat=""0""/>";
            testCase.verifyEqual(ser(ls), f1Expected, ...
                'F-1: set_bool_prop([]) MUST WRITE @w:defQFormat="0" (PRESENT, not removed)');
            % the attribute is genuinely PRESENT (not merely bytes-equal by accident):
            % the typed getter returns logical FALSE when present-as-"0", but [] when
            % ABSENT -- so this distinguishes "written 0" from "removed".
            testCase.verifyTrue(islogical(ls.defQFormat) && ls.defQFormat == false, ...
                'F-1: @w:defQFormat is PRESENT as logical false (not [] absent)');
            testCase.verifyFalse(ls.bool_prop("defQFormat"), ...
                'F-1: bool_prop round-trips the written "0" back to false');
            testCase.verifyEqual(hx_e(ls), string(oracle.latent.set_bool_prop.none_F1.serhex), ...
                'F-1: serhex byte-identical to the frozen oracle none_F1 case');

            % ints + defaults: count/defUIPriority (ST_DecimalNumber) + on/off defaults
            ls = mat2doc.oxml.OxmlElement("w:latentStyles");
            ls.count = 5; ls.defUIPriority = 99;
            ls.set_bool_prop("defSemiHidden", true); ls.set_bool_prop("defUnhideWhenUsed", false);
            testCase.verifyEqual(hx_e(ls), string(oracle.latent.ints_and_defaults.serhex));
            testCase.verifyEqual(ls.count, 5);
            testCase.verifyEqual(ls.defUIPriority, 99);
            testCase.verifyTrue(ls.bool_prop("defSemiHidden"));
            testCase.verifyFalse(ls.bool_prop("defUnhideWhenUsed"));

            % add_lsdException + get_by_name (miss -> [])
            ls = mat2doc.oxml.OxmlElement("w:latentStyles");
            e = ls.add_lsdException(); e.name = "Normal";
            testCase.verifyEqual(hx_e(ls), string(oracle.latent.add_get.serhex));
            testCase.verifyEqual(ls.get_by_name("Normal").name, "Normal");
            testCase.verifyTrue(isequal(ls.get_by_name("Nope"), []), 'get_by_name miss -> []');
        end

        % =============================================================== %
        % 7. CT_LsdException (required attr + on/off dispatch)            %
        % =============================================================== %

        function test_ct_lsdexception(testCase)
            % Nominal + Edge + Regression (H3, s0030 lsd): REQUIRED @w:name absent ->
            % mat2doc:InvalidXmlError; set name/semiHidden/uiPriority (attr order
            % name->semiHidden->uiPriority); on_off_prop present->true, absent->[];
            % set_on_off_prop qFormat true. Bytes hard-coded AND vs oracle.
            oracle = loadOracle();

            % Error path: missing REQUIRED @w:name -> mat2doc:InvalidXmlError
            le = mat2doc.oxml.OxmlElement("w:lsdException");
            testCase.verifyError(@() le.name, 'mat2doc:InvalidXmlError', ...
                'missing REQUIRED @w:name -> mat2doc:InvalidXmlError (identifier pinned)');

            e = mat2doc.oxml.OxmlElement("w:lsdException");
            e.name = "Heading 1"; e.semiHidden = true; e.uiPriority = 9;
            testCase.verifyEqual(ser(e), decl() + newline + ...
                "<w:lsdException xmlns:w=""" + testCase.W + """ w:name=""Heading 1"" " + ...
                "w:semiHidden=""1"" w:uiPriority=""9""/>", ...
                'CT_LsdException set: attr order name->semiHidden->uiPriority (L1)');
            testCase.verifyEqual(hx_e(e), string(oracle.lsd.set.serhex));
            testCase.verifyEqual(e.name, "Heading 1");
            testCase.verifyTrue(e.semiHidden);
            testCase.verifyEqual(e.uiPriority, 9);
            testCase.verifyTrue(e.on_off_prop("semiHidden"), 'on_off_prop present -> true');
            testCase.verifyTrue(isequal(e.on_off_prop("locked"), []), 'on_off_prop absent -> [] (None)');

            % set_on_off_prop dispatches by attr name
            e.set_on_off_prop("qFormat", true);
            testCase.verifyEqual(hx_e(e), string(oracle.lsd.after_set_on_off.serhex));
            testCase.verifyTrue(e.on_off_prop("qFormat"));
        end

        % =============================================================== %
        % 8. H11 child-ORDER pins (CRITICAL -- wrong order = Word repair)  %
        % =============================================================== %

        function test_h11_child_order_pins(testCase)
            % Regression (H11, the highest-value permanent pins, s0030 h11): add
            % children OUT OF canonical order and assert the SERIALIZED BYTES come
            % out in canonical schema order for all three container classes. Wrong
            % order -> Word repair prompt. Bytes hard-coded AND serhex vs the oracle.
            oracle = loadOracle();

            % -- CT_Style: add all 10 ZeroOrOne descriptors in REVERSE order --
            s = mat2doc.oxml.OxmlElement("w:style");
            s.get_or_add_rPr(); s.get_or_add_pPr(); s.get_or_add_locked(); s.get_or_add_qFormat();
            s.get_or_add_unhideWhenUsed(); s.get_or_add_semiHidden(); s.get_or_add_uiPriority();
            s.get_or_add_next(); s.get_or_add_basedOn(); s.get_or_add_name();
            testCase.verifyEqual(childLocalnames(s), testCase.STYLE_SCHEMA_ORDER, ...
                'CT_Style reverse scramble must re-sort to canonical name..rPr (H11)');
            testCase.verifyEqual(ser(s), decl() + newline + ...
                "<w:style xmlns:w=""" + testCase.W + """>" + ...
                "<w:name/><w:basedOn/><w:next/><w:uiPriority/><w:semiHidden/>" + ...
                "<w:unhideWhenUsed/><w:qFormat/><w:locked/><w:pPr/><w:rPr/></w:style>", ...
                'CT_Style canonical child order serialized bytes (L1) hard-coded');
            testCase.verifyEqual(hx_e(s), string(oracle.h11.ct_style_reverse.serhex));

            % -- CT_Styles: add_style FIRST, then latentStyles -> latentStyles SORTS FIRST --
            root = mat2doc.oxml.OxmlElement("w:styles");
            root.add_style(); root.get_or_add_latentStyles();
            testCase.verifyEqual(childLocalnames(root), ["latentStyles" "style"], ...
                'CT_Styles: latentStyles sorts BEFORE the already-present style (H11)');
            testCase.verifyEqual(ser(root), decl() + newline + ...
                "<w:styles xmlns:w=""" + testCase.W + """><w:latentStyles/><w:style/></w:styles>", ...
                'CT_Styles latentStyles-first serialized bytes (L1) hard-coded');
            testCase.verifyEqual(hx_e(root), string(oracle.h11.ct_styles_latent_after_style.serhex));

            % -- CT_LatentStyles: lsdException is ZeroOrMore -> APPEND (order preserved) --
            ls = mat2doc.oxml.OxmlElement("w:latentStyles");
            nms = ["N1" "N2" "N3"];
            for k = 1:numel(nms)
                ex = ls.add_lsdException(); ex.name = nms(k);
            end
            testCase.verifyEqual(childLocalnames(ls), ["lsdException" "lsdException" "lsdException"], ...
                'CT_LatentStyles: lsdException appends (ZeroOrMore, no re-sort)');
            testCase.verifyEqual(ser(ls), decl() + newline + ...
                "<w:latentStyles xmlns:w=""" + testCase.W + """>" + ...
                "<w:lsdException w:name=""N1""/><w:lsdException w:name=""N2""/>" + ...
                "<w:lsdException w:name=""N3""/></w:latentStyles>", ...
                'CT_LatentStyles append-order serialized bytes (L1) hard-coded');
            testCase.verifyEqual(hx_e(ls), string(oracle.h11.ct_latent_lsd_append.serhex));
        end

        % =============================================================== %
        % 9. H17 delete() -- PARENT-SIDE EFFECT ONLY (design.md sec 9)     %
        % =============================================================== %

        function test_h17_delete_parent_side_only(testCase)
            % Regression (binding Gate-4 rule): CT_Style.delete_style() and
            % CT_LsdException.delete_lsd_exception() remove the element from its
            % PARENT. Assert the PARENT's child count / serialized bytes ONLY, and
            % never inspect the removed handle afterward. These methods are named by
            % the kind of thing they remove (not `delete`), so MATLAB's handle
            % destructor is never involved (H17 dissolved -- no `delete` override).
            % We do NOT port any "removed element still has tag X".

            % -- CT_Style.delete_style(): delete the MIDDLE of three styles --
            styles = parse("<w:styles " + nsW() + ">" + ...
                "<w:style w:styleId=""A""/><w:style w:styleId=""B""/><w:style w:styleId=""C""/>" + ...
                "</w:styles>");
            kids = styles.xpath("w:style");
            testCase.verifyEqual(numel(kids), 3, 'precondition: 3 styles');
            kids(2).delete_style();   % remove B; do NOT touch the handle after this
            after = styles.xpath("w:style");
            testCase.verifyEqual(numel(after), 2, 'parent has 2 styles after delete (parent-side effect)');
            remaining = strings(1, numel(after));
            for k = 1:numel(after); remaining(k) = after(k).styleId; end
            testCase.verifyEqual(remaining, ["A" "C"], 'siblings + order intact; B removed');
            testCase.verifyEqual(ser(styles), decl() + newline + ...
                "<w:styles xmlns:w=""" + testCase.W + """>" + ...
                "<w:style w:styleId=""A""/><w:style w:styleId=""C""/></w:styles>", ...
                'parent serialized bytes reflect B removed (L1)');

            % -- CT_LsdException.delete_lsd_exception(): delete N2 of three --
            ls = parse("<w:latentStyles " + nsW() + ">" + ...
                "<w:lsdException w:name=""N1""/><w:lsdException w:name=""N2""/>" + ...
                "<w:lsdException w:name=""N3""/></w:latentStyles>");
            lex = ls.lsdException_lst;
            testCase.verifyEqual(numel(lex), 3, 'precondition: 3 lsdExceptions');
            lex(2).delete_lsd_exception();   % remove N2; do NOT touch the handle after this
            names = strings(1, numel(ls.lsdException_lst));
            afterLex = ls.lsdException_lst;
            for k = 1:numel(afterLex); names(k) = afterLex(k).name; end
            testCase.verifyEqual(names, ["N1" "N3"], 'lsdException N2 removed from parent; N1,N3 intact');
        end

        % =============================================================== %
        % 10. M1 styles.xml byte-pin (registry-adding parse-path guard)    %
        % =============================================================== %

        function test_m1_styles_xml_byte_identical(testCase)
            % Regression (byte-neutrality, L1): mat2doc.Document().save() emits
            % word/styles.xml at EXACTLY 349458 B with the frozen s0001 SHA-256 --
            % byte-identical DESPITE the 164 w:style / 137 w:lsdException / w:latentStyles
            % now parsing to the new CT_* on the save path (P4-6's specific parse-path
            % risk). SHA-256 equality is a byte-level (L1) assertion. (Test_p1_8 owns
            % the full 17/17 M1 sweep; this pins the ONE part P4-6 could break.)
            bytes = emitDocPart('styles.xml');
            testCase.verifyEqual(numel(bytes), testCase.STYLES_SIZE, ...
                sprintf('word/styles.xml must be exactly %d B after the styles registry rows', ...
                    testCase.STYLES_SIZE));
            testCase.verifyEqual(sha256hex(bytes), testCase.STYLES_SHA, ...
                'word/styles.xml SHA-256 must equal the frozen s0001 oracle (byte-identical L1)');
        end

        function test_styles_xml_parse_serialize_roundtrip_L1(testCase)
            % Regression (byte-neutrality, L1): parse the emitted styles.xml back
            % through mat2doc.oxml.parse_xml (instantiating 164 CT_Style, 137
            % CT_LsdException, CT_LatentStyles, CT_DecimalNumber and their children)
            % and re-serialize -- must be byte-identical to the input. Directly
            % exercises the new CT_* parse->serialize path at part scale. numId/ilvl
            % now parse onto CT_DecimalNumber (P8-1 registered w:numId/w:ilvl ->
            % CT_DecimalNumber); the byte roundtrip stays L1 regardless.
            inBytes  = emitDocPart('styles.xml');
            root     = mat2doc.oxml.parse_xml(inBytes);
            outBytes = mat2doc.oxml.serialize_part_xml(root);
            testCase.verifyEqual(uint8(outBytes(:)'), uint8(inBytes(:)'), ...
                'styles.xml parse->serialize must be byte-identical (new CT_* path is byte-neutral)');
            % Count guards (looser-than-byte, justified): the frozen default template
            % ships EXACTLY 164 w:style / 137 w:lsdException; pinning proves the new
            % CT_* path was exercised at scale, not that a short part slipped through.
            % numId (6) / ilvl (1) DO appear (inside w:pPr/w:numPr); P8-1 registered
            % w:numId/w:ilvl -> CT_DecimalNumber, so they now parse onto a real CT_*.
            testCase.verifyEqual(numel(root.xpath('.//w:style')), 164, ...
                'styles.xml must parse exactly 164 w:style (CT_Style path exercised at scale)');
            testCase.verifyEqual(numel(root.xpath('.//w:lsdException')), 137, ...
                'styles.xml must parse exactly 137 w:lsdException (CT_LsdException at scale)');
            numIds = root.xpath('.//w:numId');
            ilvls  = root.xpath('.//w:ilvl');
            testCase.verifyEqual(numel(numIds), 6, 'w:numId present (CT_DecimalNumber, P8-1)');
            testCase.verifyEqual(numel(ilvls), 1,  'w:ilvl present (CT_DecimalNumber, P8-1)');
            % REGISTRY-FLIP RE-PIN (P8-1 Gate-4): P8-1 registered w:numId/w:ilvl ->
            % CT_DecimalNumber, so class(numIds(1)) flips generic XmlElement ->
            % mat2doc.oxml.shared.CT_DecimalNumber. The :583 byte roundtrip PASSES
            % unchanged (styles.xml byte-neutral: 349458 B / 02d71a68...); only this
            % class assertion moves. (registry-flip stale-pins lesson.)
            testCase.verifyEqual(class(numIds(1)), 'mat2doc.oxml.shared.CT_DecimalNumber', ...
                'w:numId parses onto CT_DecimalNumber (P8-1 registered w:numId/w:ilvl)');
            % a live get_by_id transits the new CT_Style parse class
            h1 = root.get_by_id("Heading1");
            testCase.verifyEqual(h1.name_val, "heading 1");
            testCase.verifyEqual(h1.uiPriority_val, 9);
            testCase.verifyEqual(root.default_for(mat2doc.enum.style.WD_STYLE_TYPE.PARAGRAPH).styleId, "Normal");
        end

        % =============================================================== %
        % 11. EQUIVALENCE -- full s0030 battery vs the frozen oracle       %
        % =============================================================== %

        function test_equivalence_full_battery_vs_frozen_oracle(testCase)
            % Equivalence: replay the ENTIRE s0030 battery (runProbes -- the .m twin's
            % body VERBATIM: styleid / decnum / ct_style / latent / lsd / styles /
            % h11 + the real styles.xml parse->serialize leg) and flatten-compare
            % EVERY leaf to the frozen python-docx 1.2.0 oracle copied into
            % data\s0030_probe_oracle.json. Gate-3 found ZERO divergences, so every
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
end

% ===================== file-local helpers ============================== %

function s = decl()
    s = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>";
end

function s = W_()
    s = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
end

function s = nsW()
    s = "xmlns:w=""" + W_() + """";
end

function e = parse(xml)
    e = mat2doc.oxml.parse_xml(uint8(unicode2native(char(xml), "UTF-8")));
end

function s = ser(e)
    % Decode serialize_part_xml bytes to a string (ASCII content -> string-equality
    % is a byte-identical L1 assertion; non-ASCII round-trips via UTF-8).
    s = string(native2unicode(mat2doc.oxml.serialize_part_xml(e), "UTF-8"));
end

function h = hx(raw)
    % UPPERCASE hex of raw UTF-8 bytes (matches Python bytes.hex().upper()).
    h = string(sprintf('%02X', uint8(raw)));
end

function h = hx_e(e)
    h = hx(mat2doc.oxml.serialize_part_xml(e));
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
    % Uniform accessor repr mirroring the s0030 rv(): None->"None", enum->member
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

function s = raises(fn)
    try
        fn();
        s = "False";
    catch
        s = "True";
    end
end

function s = tf(b)
    if b, s = "True"; else, s = "False"; end
end

function o = loadOracle()
    % Read the co-located frozen oracle in BINARY mode (no CRLF translation) and
    % decode UTF-8 -> struct. jsondecode is whitespace/line-ending agnostic, so no
    % `* binary` .gitattributes pin is needed (value-based fixture, s0022/s0023
    % precedent).
    here = fileparts(mfilename('fullpath'));
    p = fullfile(here, 'data', 's0030_probe_oracle.json');
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
    % Base-MATLAB unzip (no toolbox) into a temp dir, both cleaned up on exit.
    % (Idiom from Test_p4_2_oxml_parfmt.m; tempname paths are absolute.)
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
    % Replay the s0030 probe sequence (the .m twin's P-struct body, VERBATIM
    % tags/inputs/order) and return the nested struct of tagged canonical values.
    % Embedded here so the Equivalence leg is self-contained (the validation-folder
    % scenario is NOT on the toolbox path and must not be a dependency). Mirrors
    % validation\mat2doc\scenarios\s0030_p4_6_styles_probe.m lines 19-196. The
    % parse-path leg reads the styles.xml that Document().save() itself emits
    % (rather than an external file arg), so no fixture path is needed.
    ST = @(name) mat2doc.enum.style.WD_STYLE_TYPE.(name);
    P = struct();

    % ===================== styleid: H15 case-sensitive vector ===============
    names = ["caption", "heading 1", "heading 9", "Heading 1", "Caption", ...
        "Normal", "heading 10", "", "  double  spaced  ", "My Custom Style", ...
        "heading 5"];
    sid = struct();
    for i = 0:numel(names) - 1
        sid.("k" + sprintf("%02d", i)) = mat2doc.oxml.styles.styleId_from_name(names(i + 1));
    end
    P.styleid = sid;

    % ===================== decnum: CT_DecimalNumber =========================
    dn = struct();
    cases = {"nine", 9; "negthree", -3; "zero", 0};
    for k = 1:size(cases, 1)
        elm = mat2doc.oxml.shared.CT_DecimalNumber.new("w:uiPriority", cases{k, 2});
        dn.(cases{k, 1}) = struct("serhex", hx_e(elm), "val", rv(elm.val));
    end
    bare = mat2doc.oxml.OxmlElement("w:uiPriority");
    dn.missing_val_raises = raises(@() bare.val);
    P.decnum = dn;

    % ===================== ct_style: attrs + val tri-state ==================
    cs = struct();
    b = mat2doc.oxml.OxmlElement("w:style");
    cs.bare = struct("type", rv(b.type), "styleId", rv(b.styleId), ...
        "default", rv(b.default), "customStyle", rv(b.customStyle), ...
        "basedOn_val", rv(b.basedOn_val), "name_val", rv(b.name_val), ...
        "uiPriority_val", rv(b.uiPriority_val), "locked_val", rv(b.locked_val), ...
        "qFormat_val", rv(b.qFormat_val), "semiHidden_val", rv(b.semiHidden_val), ...
        "unhideWhenUsed_val", rv(b.unhideWhenUsed_val));

    s = mat2doc.oxml.OxmlElement("w:style");
    s.type = ST("PARAGRAPH"); s.styleId = "MyStyle"; s.default = true; s.customStyle = true;
    cs.attrs_set = struct("serhex", hx_e(s), "type", rv(s.type), ...
        "styleId", rv(s.styleId), "default", rv(s.default), "customStyle", rv(s.customStyle));

    s = mat2doc.oxml.OxmlElement("w:style");
    s.name_val = "heading 1"; s.uiPriority_val = 9; s.basedOn_val = "Normal";
    cs.val_set = struct("serhex", hx_e(s), "localnames", {lnsCell(s)}, ...
        "name_val", rv(s.name_val), "uiPriority_val", rv(s.uiPriority_val), ...
        "basedOn_val", rv(s.basedOn_val));
    s.name_val = []; s.uiPriority_val = []; s.basedOn_val = [];
    cs.val_cleared = struct("serhex", hx_e(s), "localnames", {lnsCell(s)});

    s = mat2doc.oxml.OxmlElement("w:style");
    s.locked_val = true; s.semiHidden_val = true; s.qFormat_val = true; s.unhideWhenUsed_val = true;
    cs.bool_true = struct("serhex", hx_e(s), "localnames", {lnsCell(s)}, ...
        "locked_val", rv(s.locked_val), "qFormat_val", rv(s.qFormat_val), ...
        "semiHidden_val", rv(s.semiHidden_val), "unhideWhenUsed_val", rv(s.unhideWhenUsed_val));
    s.locked_val = false; s.semiHidden_val = false; s.unhideWhenUsed_val = false;
    cs.bool_false_removed = struct("serhex", hx_e(s), "localnames", {lnsCell(s)});

    styles = parse("<w:styles " + nsW() + ">" + ...
        "<w:style w:type=""paragraph"" w:styleId=""Normal""><w:name w:val=""Normal""/></w:style>" + ...
        "<w:style w:type=""paragraph"" w:styleId=""Heading1"">" + ...
        "<w:name w:val=""heading 1""/><w:basedOn w:val=""Normal""/>" + ...
        "<w:next w:val=""Normal""/><w:uiPriority w:val=""9""/></w:style>" + ...
        "</w:styles>");
    h1 = styles.get_by_id("Heading1");
    cs.chain = struct("basedOn_val", rv(h1.basedOn_val), ...
        "uiPriority_val", rv(h1.uiPriority_val), "name_val", rv(h1.name_val), ...
        "base_style_styleId", rv(h1.base_style.styleId), ...
        "next_style_styleId", rv(h1.next_style.styleId), ...
        "get_by_name_styleId", rv(styles.get_by_name("heading 1").styleId), ...
        "get_by_id_missing", rv(styles.get_by_id("Nope")), ...
        "get_by_name_missing", rv(styles.get_by_name("nope")));
    P.ct_style = cs;

    % ===================== latent: CT_LatentStyles ==========================
    lt = struct();
    lt.bool_prop_absent = rv(mat2doc.oxml.OxmlElement("w:latentStyles").bool_prop("defQFormat"));

    tt = struct();
    ttc = {"true", true; "false", false; "one", 1; "zero", 0; "none_F1", []};
    for k = 1:size(ttc, 1)
        ls = mat2doc.oxml.OxmlElement("w:latentStyles");
        ls.set_bool_prop("defQFormat", ttc{k, 2});
        tt.(ttc{k, 1}) = struct("serhex", hx_e(ls), "bool_prop", rv(ls.bool_prop("defQFormat")));
    end
    lt.set_bool_prop = tt;

    ls = mat2doc.oxml.OxmlElement("w:latentStyles");
    ls.count = 5; ls.defUIPriority = 99;
    ls.set_bool_prop("defSemiHidden", true); ls.set_bool_prop("defUnhideWhenUsed", false);
    lt.ints_and_defaults = struct("serhex", hx_e(ls), "count", rv(ls.count), ...
        "defUIPriority", rv(ls.defUIPriority), ...
        "defSemiHidden", rv(ls.bool_prop("defSemiHidden")), ...
        "defUnhideWhenUsed", rv(ls.bool_prop("defUnhideWhenUsed")));

    ls = mat2doc.oxml.OxmlElement("w:latentStyles");
    e = ls.add_lsdException(); e.name = "Normal";
    lt.add_get = struct("serhex", hx_e(ls), ...
        "get_by_name_name", rv(ls.get_by_name("Normal").name), ...
        "get_by_name_missing", rv(ls.get_by_name("Nope")));
    P.latent = lt;

    % ===================== lsd: CT_LsdException =============================
    ld = struct();
    ld.name_required_raises = raises(@() mat2doc.oxml.OxmlElement("w:lsdException").name);

    e = mat2doc.oxml.OxmlElement("w:lsdException");
    e.name = "Heading 1"; e.semiHidden = true; e.uiPriority = 9;
    ld.set = struct("serhex", hx_e(e), "name", rv(e.name), ...
        "semiHidden", rv(e.semiHidden), "uiPriority", rv(e.uiPriority), ...
        "on_off_semiHidden", rv(e.on_off_prop("semiHidden")), ...
        "on_off_locked_absent", rv(e.on_off_prop("locked")));
    e.set_on_off_prop("qFormat", true);
    ld.after_set_on_off = struct("serhex", hx_e(e), "qFormat", rv(e.on_off_prop("qFormat")));
    P.lsd = ld;

    % ===================== styles: add_style_of_type + lookups ==============
    st = struct();
    root = mat2doc.oxml.OxmlElement("w:styles");
    builtin = root.add_style_of_type("heading 1", ST("PARAGRAPH"), true);
    st.add_builtin = struct("serhex", hx_e(builtin), "styleId", rv(builtin.styleId), ...
        "customStyle", rv(builtin.customStyle), "type", rv(builtin.type), ...
        "name_val", rv(builtin.name_val));
    root2 = mat2doc.oxml.OxmlElement("w:styles");
    custom = root2.add_style_of_type("My Custom Style", ST("CHARACTER"), false);
    st.add_custom = struct("serhex", hx_e(custom), "styleId", rv(custom.styleId), ...
        "customStyle", rv(custom.customStyle), "type", rv(custom.type), ...
        "name_val", rv(custom.name_val));

    styles = parse("<w:styles " + nsW() + ">" + ...
        "<w:style w:type=""paragraph"" w:styleId=""A"" w:default=""1""><w:name w:val=""A""/></w:style>" + ...
        "<w:style w:type=""paragraph"" w:styleId=""B"" w:default=""1""><w:name w:val=""B""/></w:style>" + ...
        "<w:style w:type=""paragraph"" w:styleId=""D"" w:default=""0""><w:name w:val=""D""/></w:style>" + ...
        "<w:style w:type=""character"" w:styleId=""C"" w:default=""1""><w:name w:val=""C""/></w:style>" + ...
        "</w:styles>");
    st.default_for = struct("paragraph", rv(styles.default_for(ST("PARAGRAPH")).styleId), ...
        "character", rv(styles.default_for(ST("CHARACTER")).styleId), ...
        "table_none", rv(styles.default_for(ST("TABLE"))));
    P.styles = st;

    % ===================== h11: scrambled child adds (CRUX) =================
    h11 = struct();
    s = mat2doc.oxml.OxmlElement("w:style");
    s.get_or_add_rPr(); s.get_or_add_pPr(); s.get_or_add_locked(); s.get_or_add_qFormat();
    s.get_or_add_unhideWhenUsed(); s.get_or_add_semiHidden(); s.get_or_add_uiPriority();
    s.get_or_add_next(); s.get_or_add_basedOn(); s.get_or_add_name();
    h11.ct_style_reverse = struct("localnames", {lnsCell(s)}, "serhex", hx_e(s));

    root = mat2doc.oxml.OxmlElement("w:styles");
    root.add_style(); root.get_or_add_latentStyles();
    h11.ct_styles_latent_after_style = struct("localnames", {lnsCell(root)}, "serhex", hx_e(root));

    ls = mat2doc.oxml.OxmlElement("w:latentStyles");
    nms = ["N1", "N2", "N3"];
    for k = 1:numel(nms)
        ex = ls.add_lsdException(); ex.name = nms(k);
    end
    lsl = ls.lsdException_lst; nmout = cell(1, numel(lsl));
    for k = 1:numel(lsl); nmout{k} = rv(lsl(k).name); end
    h11.ct_latent_lsd_append = struct("localnames", {lnsCell(ls)}, ...
        "names", {nmout}, "serhex", hx_e(ls));
    P.h11 = h11;

    % ===================== parsepath: real styles.xml L1 ====================
    pp = struct();
    sb = emitDocPart('styles.xml');
    elm = mat2doc.oxml.parse_xml(sb);
    out = mat2doc.oxml.serialize_part_xml(elm);
    pp.roundtrip_byte_identical = tf(isequal(uint8(out(:)'), uint8(sb(:)')));
    pp.n_style        = numel(elm.xpath(".//w:style"));
    pp.n_lsdException = numel(elm.xpath(".//w:lsdException"));
    pp.n_numId        = numel(elm.xpath(".//w:numId"));
    pp.n_ilvl         = numel(elm.xpath(".//w:ilvl"));
    h1 = elm.get_by_id("Heading1");
    pp.heading1_name_val       = rv(h1.name_val);
    pp.heading1_uiPriority_val = rv(h1.uiPriority_val);
    pp.default_paragraph_styleId = rv(elm.default_for(ST("PARAGRAPH")).styleId);
    P.parsepath = pp;
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

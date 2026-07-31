classdef Test_p6_2_table_props < matlab.unittest.TestCase
% TEST_P6_2_TABLE_PROPS  Gate-4 permanent unit tests for Mat2Doc P6-2
%   (oxml table-PROPERTIES + row classes): src/docx/oxml/table.py ->
%   +mat2doc\+oxml\+table\ CT_TblPr / CT_TblPrEx / CT_TrPr / CT_Row, plus the
%   SEVEN P6-2 registry rows in +mat2doc\+oxml\registry.m (w:gridAfter ->
%   CT_DecimalNumber; w:gridBefore -> CT_DecimalNumber; w:tblPr -> CT_TblPr;
%   w:tblPrEx -> CT_TblPrEx; w:tblStyle -> CT_String; w:tr -> CT_Row; w:trPr ->
%   CT_TrPr).
%
%   P6-2 is the SECOND Phase-6 (tables) WP. Unlike P6-1 it is NOT M1-neutral by
%   absence: word/styles.xml + word/stylesWithEffects.xml each carry 100 <w:tblPr>
%   nodes (inside table styles), so registering w:tblPr -> CT_TblPr makes those
%   100 nodes transit CT_TblPr on the M1 parse path. Byte-neutral by the
%   CT-registration precedent (registering changes only a parsed node's CLASS,
%   never its content/order; CT_TblPr adds no parse-time behavior) -- proven by
%   the styles.xml SHA staying byte-identical (test_m1_styles_and_document). The
%   other six new tags have zero occurrences in default.docx.
%
%   This class permanently freezes the guarantees the prior gates established:
%     * Gate-1 Porter  : audit_P6-2_table_props.md (self-probe).
%     * Gate-2 Auditor (Fable): APPROVE -- the A2 cross-enum ruled (i) no-D
%       (the WD_PARAGRAPH_ALIGNMENT vs WD_TABLE_ALIGNMENT cross-class == gap is
%       non-byte / non-output; compare by NAME + int VALUE only).
%     * Gate-3 Validator: validate_P6-2_table_props.md -- PASS, ZERO new
%       D-numbers, NO re-pin list (0 flips). M1 17/17 (styles.xml 349458 B /
%       02d71a68... AND document.xml 1548 B / 0e4dd503... byte-identical);
%       probe_diff s0061 MATCH (exit 0) over the full props+row surface (A2 by
%       name/value, incl. the w:jc val="both" -> JUSTIFY read-back); the s0062
%       table-props round-trip 6/6 byte-identical (incl. the both edge).
%
%   THE HIGHEST-VALUE PERMANENT PINS IN THIS WP (a regression goes RED loudly):
%     * (A2) CT_TblPr.alignment CROSS-ENUM (test_ct_tblpr_alignment): the KEY
%       DISCIPLINE. The getter faithfully returns a WD_PARAGRAPH_ALIGNMENT member
%       (exactly python-docx's runtime object), NEVER converted to
%       WD_TABLE_ALIGNMENT. Every assertion is by enum NAME ("CENTER"/"JUSTIFY")
%       and int VALUE (.value), NEVER a cross-class == (design.md section 2 A2
%       rule). The star edge: parse <w:jc w:val="both"/> -> the getter returns
%       WD_PARAGRAPH_ALIGNMENT.JUSTIFY (name "JUSTIFY", value 3) -- the read-back
%       a converting getter would CRASH on (WD_TABLE_ALIGNMENT has no "both").
%       The SETTER takes a WD_TABLE_ALIGNMENT value and writes correct bytes via
%       int-value cross-enum to_xml: CENTER(1) -> <w:jc w:val="center"/>. serhex
%       byte-pins vs the frozen oracle; set None removes w:jc.
%     * (R) s0062 TABLE-PROPS round-trip byte-pins (test_s0062_roundtrip_byte_pins):
%       6 frozen python-docx fixtures (loose tblPr/trPr/tr + the both edge + two
%       real-nsmap add_table subtrees). Each fixture's bytes parse through the CT
%       classes and re-serialize BYTE-IDENTICAL (SHA-256 == the frozen manifest).
%     * (M) M1 styles.xml + document.xml byte-pins (test_m1_styles_and_document):
%       the registry-neutrality guard -- mat2doc.Document().save() -> styles.xml
%       == 349458 B / 02d71a68... (the 100-node CT_TblPr parse-path neutrality
%       guard) AND document.xml == 1548 B / 0e4dd503... SHA equality is L1.
%     * (H11) CT_TblPr + CT_TrPr scrambled-build canonical order
%       (test_h11_canonical_order): out-of-order adds land in schema (_tag_seq)
%       order; the built subtree is byte-identical to the frozen loose fixtures
%       (tblPr b5de419f, trPr c1a760cc).
%     * (H1) CT_Row.tr_idx (test_ct_row): 0-based index among sibling <w:tr>
%       (a parsed 3-row parent -> [0 1 2]); the custom inserters force tblPrEx to
%       the front and trPr immediately after.
%     * (P6-3a guards) CT_Row _new_tc / add_tc / tc_at_grid_offset grid_span
%       (test_ct_row_tc_boundary): raise mat2doc:notYetPorted at the CT_Tc
%       dependency (never silently mis-walk); the fast paths (offset==grid_before
%       -> first cell; empty row -> mat2doc:ValueError) work now.
%
%   Provenance (all Gate-3 frozen 2026-07-31):
%     * Audit    : validation\mat2doc\audit_P6-2_table_props.md
%     * Validate : validation\mat2doc\validate_P6-2_table_props.md
%     * Scenarios: validation\mat2doc\scenarios\s0061_p6_2_table_props_probe.{py,m}
%                  (its probe body is replayed VERBATIM by runProbes() below);
%                  s0062_p6_2_table_roundtrip.{py,m} (the 6 byte fixtures).
%     * Frozen refs (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE at Gate-3):
%         references\s0061\probe.json -- copied verbatim (self-contained) into
%           tests\oxml\data\s0061_probe_oracle.json (the Equivalence replay set;
%           value/serhex JSON, jsondecode is line-ending agnostic -> no `* binary`
%           .gitattributes pin needed, per the s0059/s0030 precedent).
%         references\s0062\ (6 byte fixtures + manifest.json) -- copied verbatim
%           into tests\oxml\data\s0062\ WITH a co-located `.gitattributes`
%           `* binary` pin (frozen-byte fixtures must not be line-ending mangled
%           on the master checkout -- the Gate-4 byte-fixture lesson).
%         references\s0001\parts\word\{styles,document}.xml -- the M1 byte
%           references (SHA of what Document().save() emits); NOT copied.
%
%   Coverage taxonomy
%   -----------------
%   * Nominal      -- alignment set CENTER/LEFT/RIGHT; autofit set/parse; style
%                     get/set/parse; grid_before/after set; trHeight_val/hRule
%                     set; row trHeight delegation; tr_idx; inserters; registry.
%   * Edge         -- the w:jc val="both" -> JUSTIFY read-back (A2 star); set None
%                     removes (alignment/style); the None-guard (assigning None to
%                     a bare trPr creates NO <w:trHeight>); bare defaults 0/0/None;
%                     the CT_Tc P6-3a boundary error paths (mat2doc:notYetPorted +
%                     mat2doc:ValueError verbatim); CT_TblPrEx bare pass-through.
%   * Equivalence  -- test_equivalence_full_battery_vs_frozen_oracle replays the
%                     ENTIRE s0061 battery (runProbes, the .m twin's body verbatim)
%                     and flatten-compares every leaf to the frozen oracle (0 diffs).
%   * Regression   -- hard-coded expected serialized-XML strings (ASCII ==
%                     byte-identical L1) + UPPERCASE serhex vs the frozen oracle +
%                     SHA-256 of the 6 s0062 fixtures + the M1 styles/document parts.
%   * Upstream     -- the two s0062 real-nsmap subtrees (realtblPr/realtrPr) are
%                     REAL python-docx add_table output; the frozen fixtures ARE
%                     lxml's expected serialization for them.
%
%   Byte-level (L1) note: every serialized-XML comparison is either the FULL
%   serialize_part_xml output as an ASCII string (string-equality == byte-equality
%   L1), or its UPPERCASE hex (serhex) vs the frozen oracle, or the SHA-256 of a
%   frozen fixture / emitted M1 part. NO D-number granted any L2 relaxation in this
%   WP (Gate-3: zero new, none at L2), so every pin here is L1. The equivalence
%   leaf-key-count guard is the only looser-than-byte check and is commented at its
%   site. The A2 cross-enum compares are by NAME + int VALUE (never cross-class ==)
%   -- this is FAITHFULNESS, not a relaxation (design.md section 2, ratified no-D).
%
%   Determinism: no network, no absolute paths. The worktree root, the co-located
%   s0061 oracle and the s0062 byte fixtures resolve relative to this file via
%   fileparts(mfilename('fullpath')); saves go to tempname .docx deleted via
%   onCleanup; every file read is binary ('r','n'). The +mat2doc package resolves
%   via the MANDATORY PathFixture(worktree-root) in TestClassSetup (WP9-F4 lesson).

    properties (Constant)
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        % --- registered class names (P6-2 rows, +oxml\registry.m 306-317) ---
        CT_TBLPR   = 'mat2doc.oxml.table.CT_TblPr'
        CT_TBLPREX = 'mat2doc.oxml.table.CT_TblPrEx'
        CT_TRPR    = 'mat2doc.oxml.table.CT_TrPr'
        CT_ROW     = 'mat2doc.oxml.table.CT_Row'
        CT_STRING  = 'mat2doc.oxml.shared.CT_String'
        CT_DECNUM  = 'mat2doc.oxml.shared.CT_DecimalNumber'
        CT_JC      = 'mat2doc.oxml.text.CT_Jc'         % A2: shared, NOT re-registered

        % --- frozen s0001 M1 byte references (the P6-2 registry-adding
        %     neutrality guard: styles.xml carries 100 <w:tblPr> nodes) ---
        STYLES_SIZE_M1 = 349458
        STYLES_SHA_M1  = "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"
        DOC_SIZE_M1    = 1548
        DOC_SHA_M1     = "0e4dd503bc095219cd24e8312b372c588ca56bb834aacac53eafb7bf47836327"
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied verbatim from tests\oxml\Test_p6_1_table_leaves.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\oxml
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. registry rows (the 7 P6-2 flips + A2 w:jc NOT re-registered)   %
        % =============================================================== %

        function test_registry_resolves_seven_rows(testCase)
            % Nominal / Regression (registry.m 306-317): the 7 P6-2 table-props
            % rows. REGISTRATION is what flips the parse class of a table subtree.
            pairs = { ...
                "w:gridAfter",  testCase.CT_DECNUM; ...
                "w:gridBefore", testCase.CT_DECNUM; ...
                "w:tblPr",      testCase.CT_TBLPR; ...
                "w:tblPrEx",    testCase.CT_TBLPREX; ...
                "w:tblStyle",   testCase.CT_STRING; ...
                "w:tr",         testCase.CT_ROW; ...
                "w:trPr",       testCase.CT_TRPR };
            for i = 1:size(pairs, 1)
                tag = pairs{i, 1}; cls = pairs{i, 2};
                r = mat2doc.oxml.registry(mat2doc.oxml.qn(tag));
                testCase.verifyEqual(char(r), cls, ...
                    sprintf('registry must resolve %s -> %s (P6-2 row)', tag, cls));
                testCase.verifyEqual(class(mat2doc.oxml.OxmlElement(tag)), cls, ...
                    sprintf('OxmlElement(%s) must be a %s', tag, cls));
            end
            % A2: w:jc is the SINGLE P4-2 CT_Jc row -- NOT re-registered to a table
            % class. CT_TblPr.alignment REUSES that one CT_Jc (one element class,
            % two context enums). A silent re-registration would be a divergence.
            testCase.verifyEqual(char(mat2doc.oxml.registry(mat2doc.oxml.qn("w:jc"))), ...
                testCase.CT_JC, 'A2: w:jc stays the single P4-2 CT_Jc row');
            testCase.verifyEqual(class(mat2doc.oxml.OxmlElement("w:jc")), testCase.CT_JC, ...
                'A2: OxmlElement(w:jc) is CT_Jc (paragraph consumer shared)');
        end

        % =============================================================== %
        % 2. CT_TblPr.alignment -- the A2 CROSS-ENUM (the key discipline)   %
        % =============================================================== %

        function test_ct_tblpr_alignment(testCase)
            % (A2) Nominal + Edge + Regression (s0061 tblpr align*): assert by enum
            % NAME + int VALUE + simple class NAME -- NEVER a cross-class ==. The
            % getter returns a WD_PARAGRAPH_ALIGNMENT member verbatim (the faithful
            % port); the setter cross-enum-writes correct bytes. serhex vs oracle.
            o = loadOracle();

            % -- parse center: getter returns WD_PARAGRAPH_ALIGNMENT.CENTER --
            e = parse("<w:tblPr " + nsW() + "><w:jc w:val=""center""/></w:tblPr>");
            testCase.verifyEqual(class(e), testCase.CT_TBLPR);
            a = e.alignment;
            testCase.verifyEqual(string(a), "CENTER", 'A2: parse center -> name CENTER (by NAME, not ==)');
            testCase.verifyEqual(double(a.value), 1, 'A2: CENTER int value 1');
            testCase.verifyTrue(endsWith(class(a), 'WD_PARAGRAPH_ALIGNMENT'), ...
                'A2: getter returns the PARAGRAPH enum (faithful, not WD_TABLE_ALIGNMENT)');
            % tie to the frozen oracle field-wise (jsondecode yields char -> string())
            testCase.verifyEqual(string(o.tblpr.align_parse_center.name), "CENTER", ...
                'A2: oracle parse-center name');
            testCase.verifyEqual(string(o.tblpr.align_parse_center.value), "1", ...
                'A2: oracle parse-center value');
            testCase.verifyEqual(string(o.tblpr.align_parse_center.ecls), "WD_PARAGRAPH_ALIGNMENT", ...
                'A2: oracle parse-center ecls');

            % -- ★ parse both: the JUSTIFY read-back a converting getter would CRASH
            %    on. WD_TABLE_ALIGNMENT has NO "both"; the faithful port returns
            %    WD_PARAGRAPH_ALIGNMENT.JUSTIFY (name JUSTIFY, value 3). --
            e = parse("<w:tblPr " + nsW() + "><w:jc w:val=""both""/></w:tblPr>");
            a = e.alignment;
            testCase.verifyEqual(string(a), "JUSTIFY", ...
                'A2 STAR: <w:jc val="both"> -> name JUSTIFY (a converting getter crashes here)');
            testCase.verifyEqual(double(a.value), 3, 'A2 STAR: JUSTIFY int value 3');
            testCase.verifyTrue(endsWith(class(a), 'WD_PARAGRAPH_ALIGNMENT'), ...
                'A2 STAR: JUSTIFY is a WD_PARAGRAPH_ALIGNMENT member');
            % tie to the frozen oracle field-wise (jsondecode yields char -> string())
            testCase.verifyEqual(string(o.tblpr.align_parse_both.name), "JUSTIFY", ...
                'A2 STAR: oracle parse-both name JUSTIFY');
            testCase.verifyEqual(string(o.tblpr.align_parse_both.value), "3", ...
                'A2 STAR: oracle parse-both value 3');
            testCase.verifyEqual(string(o.tblpr.align_parse_both.ecls), "WD_PARAGRAPH_ALIGNMENT", ...
                'A2 STAR: oracle parse-both ecls');

            % -- set WD_TABLE_ALIGNMENT LEFT/CENTER/RIGHT: correct bytes via
            %    int-value cross-enum to_xml; read-back is the PARAGRAPH member --
            names = ["LEFT" "CENTER" "RIGHT"];
            xmlv  = struct('LEFT', "left", 'CENTER', "center", 'RIGHT', "right");
            valn  = struct('LEFT', 0, 'CENTER', 1, 'RIGHT', 2);
            for k = 1:numel(names)
                nm = names(k);
                e = mat2doc.oxml.OxmlElement("w:tblPr");
                e.alignment = mat2doc.enum.table.WD_TABLE_ALIGNMENT.(nm);
                a = e.alignment;
                testCase.verifyEqual(string(a), nm, ...
                    sprintf('A2: set TABLE %s -> read-back name %s', nm, nm));
                testCase.verifyEqual(double(a.value), valn.(nm), ...
                    sprintf('A2: %s int value', nm));
                testCase.verifyTrue(endsWith(class(a), 'WD_PARAGRAPH_ALIGNMENT'), ...
                    'A2: read-back is ALWAYS the paragraph enum');
                % hard-coded serialized bytes (L1) + serhex vs oracle
                testCase.verifyEqual(ser(e), decl() + newline + ...
                    "<w:tblPr xmlns:w=""" + testCase.W + """><w:jc w:val=""" + ...
                    xmlv.(nm) + """/></w:tblPr>", ...
                    sprintf('A2: set %s serialized bytes (L1) hard-coded', nm));
                testCase.verifyEqual(hx_e(e), string(o.tblpr.("align_set_" + nm).serhex), ...
                    sprintf('A2: set %s serhex (L1) vs frozen oracle', nm));
            end

            % -- set None removes w:jc --
            e = parse("<w:tblPr " + nsW() + "><w:jc w:val=""center""/></w:tblPr>");
            e.alignment = [];
            testCase.verifyTrue(isequal(e.find(mat2doc.oxml.qn("w:jc")), []), ...
                'set alignment None removes the w:jc child');
            testCase.verifyEqual(hx_e(e), string(o.tblpr.align_set_none_removes.serhex), ...
                'set alignment None serhex (L1) vs frozen oracle (<w:tblPr/>)');

            % -- bare tblPr alignment -> [] (None) --
            testCase.verifyTrue(isequal(mat2doc.oxml.OxmlElement("w:tblPr").alignment, []), ...
                'bare tblPr alignment -> [] (None, w:jc absent)');
        end

        % =============================================================== %
        % 3. CT_TblPr.autofit -- fixed layout semantics (H3/H4)             %
        % =============================================================== %

        function test_ct_tblpr_autofit(testCase)
            % Nominal + Edge + Regression (s0061 autofit*): False IFF <w:tblLayout>
            % @type="fixed"; True when tblLayout absent (H3) or @type="autofit".
            % Setter writes @type="autofit" if truthy else "fixed" (H4).
            o = loadOracle();

            % bare -> True (tblLayout absent, H3)
            testCase.verifyTrue(mat2doc.oxml.OxmlElement("w:tblPr").autofit, ...
                'bare tblPr autofit -> True (tblLayout absent)');

            % parse fixed -> False ; parse autofit -> True
            e = parse("<w:tblPr " + nsW() + "><w:tblLayout w:type=""fixed""/></w:tblPr>");
            testCase.verifyFalse(e.autofit, 'parse type="fixed" -> autofit False');
            e = parse("<w:tblPr " + nsW() + "><w:tblLayout w:type=""autofit""/></w:tblPr>");
            testCase.verifyTrue(e.autofit, 'parse type="autofit" -> autofit True');

            % set false -> <w:tblLayout w:type="fixed"/>
            e = mat2doc.oxml.OxmlElement("w:tblPr");
            e.autofit = false;
            testCase.verifyFalse(e.autofit, 'read-back after set false');
            testCase.verifyEqual(ser(e), decl() + newline + ...
                "<w:tblPr xmlns:w=""" + testCase.W + """><w:tblLayout w:type=""fixed""/></w:tblPr>", ...
                'set autofit false serialized bytes (L1) hard-coded');
            testCase.verifyEqual(hx_e(e), string(o.tblpr.autofit_set_false.serhex), ...
                'set autofit false serhex (L1) vs frozen oracle');

            % set true -> <w:tblLayout w:type="autofit"/>
            e = mat2doc.oxml.OxmlElement("w:tblPr");
            e.autofit = true;
            testCase.verifyTrue(e.autofit, 'read-back after set true');
            testCase.verifyEqual(hx_e(e), string(o.tblpr.autofit_set_true.serhex), ...
                'set autofit true serhex (L1) vs frozen oracle');
        end

        % =============================================================== %
        % 4. CT_TblPr.style -- ./w:tblStyle/@val (CT_String)                %
        % =============================================================== %

        function test_ct_tblpr_style(testCase)
            % Nominal + Edge + Regression (s0061 style*): bare None; set writes
            % <w:tblStyle w:val="..."/> (via the PRIVATE _add adder); set None
            % removes; parse reads @val. Requires w:tblStyle -> CT_String (the
            % HARD functional dependency this WP registered).
            o = loadOracle();

            % bare -> [] (None)
            testCase.verifyTrue(isequal(mat2doc.oxml.OxmlElement("w:tblPr").style, []), ...
                'bare tblPr style -> [] (None, w:tblStyle absent)');

            % set "TestTableStyle"
            e = mat2doc.oxml.OxmlElement("w:tblPr");
            e.style = "TestTableStyle";
            testCase.verifyEqual(string(e.style), "TestTableStyle", 'style read-back');
            testCase.verifyEqual(ser(e), decl() + newline + ...
                "<w:tblPr xmlns:w=""" + testCase.W + """><w:tblStyle w:val=""TestTableStyle""/></w:tblPr>", ...
                'set style serialized bytes (L1) hard-coded');
            testCase.verifyEqual(hx_e(e), string(o.tblpr.style_set.serhex), ...
                'set style serhex (L1) vs frozen oracle');

            % set None removes w:tblStyle
            e = mat2doc.oxml.OxmlElement("w:tblPr");
            e.style = "TestTableStyle"; e.style = [];
            testCase.verifyTrue(isequal(e.style, []), 'style -> [] after None');
            testCase.verifyEqual(hx_e(e), string(o.tblpr.style_set_none_removes.serhex), ...
                'set style None removes w:tblStyle (L1)');

            % parse reads @val -- proves w:tblStyle transits CT_String (.val works)
            e = parse("<w:tblPr " + nsW() + "><w:tblStyle w:val=""TableGrid""/></w:tblPr>");
            testCase.verifyEqual(class(e.tblStyle), testCase.CT_STRING, ...
                'w:tblStyle parses as CT_String (functional dep)');
            testCase.verifyEqual(string(e.style), "TableGrid", 'parse style -> TableGrid');
        end

        % =============================================================== %
        % 5. CT_TrPr -- grid_before/after + trHeight + the None-guard        %
        % =============================================================== %

        function test_ct_trpr(testCase)
            % Nominal + Edge + Regression (s0061 trpr): grid_before/after default 0
            % (absent child, H3), set via CT_DecimalNumber .val; trHeight_val/hRule
            % None when trHeight absent; set both; the None-guard (assigning None to
            % a bare trPr must NOT create an empty <w:trHeight>). serhex vs oracle.
            o = loadOracle();

            % bare grid_before/after -> 0/0 (H3)
            b = mat2doc.oxml.OxmlElement("w:trPr");
            testCase.verifyEqual(b.grid_before, 0, 'bare trPr grid_before -> 0');
            testCase.verifyEqual(b.grid_after, 0,  'bare trPr grid_after -> 0');

            % set via get_or_add_gridBefore/After().val -> CT_DecimalNumber
            e = mat2doc.oxml.OxmlElement("w:trPr");
            gb = e.get_or_add_gridBefore(); gb.val = 1;
            ga = e.get_or_add_gridAfter();  ga.val = 2;
            testCase.verifyEqual(class(gb), testCase.CT_DECNUM, 'gridBefore is CT_DecimalNumber');
            testCase.verifyEqual(class(ga), testCase.CT_DECNUM, 'gridAfter is CT_DecimalNumber');
            testCase.verifyEqual(e.grid_before, 1, 'grid_before reads child .val');
            testCase.verifyEqual(e.grid_after, 2,  'grid_after reads child .val');
            % H11: gridBefore (@3) serializes BEFORE gridAfter (@4) despite the add order
            testCase.verifyEqual(childLocalnames(e), ["gridBefore" "gridAfter"], ...
                'H11: gridBefore before gridAfter (schema order)');
            testCase.verifyEqual(hx_e(e), string(o.trpr.grid_set.serhex), ...
                'trPr grid_set serhex (L1) vs frozen oracle');

            % trHeight_val / trHeight_hRule bare -> None
            b = mat2doc.oxml.OxmlElement("w:trPr");
            testCase.verifyTrue(isequal(b.trHeight_val, []),   'bare trHeight_val -> [] (None)');
            testCase.verifyTrue(isequal(b.trHeight_hRule, []), 'bare trHeight_hRule -> [] (None)');

            % set both (val Twips(360)=228600 EMU; hRule AT_LEAST -> "atLeast")
            e = mat2doc.oxml.OxmlElement("w:trPr");
            e.trHeight_val = mat2doc.shared.Twips(360);
            e.trHeight_hRule = mat2doc.enum.table.WD_ROW_HEIGHT_RULE.AT_LEAST;
            testCase.verifyEqual(double(e.trHeight_val), 228600, 'trHeight_val EMU exact (Twips(360))');
            testCase.verifyEqual(e.trHeight_hRule, mat2doc.enum.table.WD_ROW_HEIGHT_RULE.AT_LEAST);
            testCase.verifyEqual(hx_e(e), string(o.trpr.trheight_set_both.serhex), ...
                'trPr trHeight set-both serhex (L1) vs frozen oracle');

            % ★ None-guard: assigning None to a bare trPr must NOT create <w:trHeight>
            e = mat2doc.oxml.OxmlElement("w:trPr");
            e.trHeight_val = [];
            testCase.verifyTrue(isequal(e.trHeight, []), ...
                'None-guard: trHeight_val=None on bare trPr creates NO <w:trHeight>');
            testCase.verifyEqual(hx_e(e), string(o.trpr.trheight_val_none_guard.serhex), ...
                'trPr trHeight_val None-guard serhex (L1) -> bare <w:trPr/>');
            e = mat2doc.oxml.OxmlElement("w:trPr");
            e.trHeight_hRule = [];
            testCase.verifyTrue(isequal(e.trHeight, []), ...
                'None-guard: trHeight_hRule=None on bare trPr creates NO <w:trHeight>');
            testCase.verifyEqual(hx_e(e), string(o.trpr.trheight_hrule_none_guard.serhex), ...
                'trPr trHeight_hRule None-guard serhex (L1) -> bare <w:trPr/>');
        end

        % =============================================================== %
        % 6. CT_Row -- delegation, tr_idx (H1), custom inserters (H11)       %
        % =============================================================== %

        function test_ct_row(testCase)
            % Nominal + Edge + Regression (s0061 row): grid_before/after delegate to
            % trPr (0 when trPr absent); trHeight set creates <w:trPr><w:trHeight>;
            % tr_idx 0-based among sibling <w:tr>; custom inserters force tblPrEx to
            % the front and trPr immediately after. serhex vs frozen oracle.
            o = loadOracle();

            % bare grid_before/after -> 0/0 (trPr absent, H3)
            b = mat2doc.oxml.OxmlElement("w:tr");
            testCase.verifyEqual(class(b), testCase.CT_ROW);
            testCase.verifyEqual(b.grid_before, 0, 'bare tr grid_before -> 0 (no trPr)');
            testCase.verifyEqual(b.grid_after, 0,  'bare tr grid_after -> 0 (no trPr)');

            % trHeight set delegates: creates <w:trPr><w:trHeight .../></w:trPr>
            e = mat2doc.oxml.OxmlElement("w:tr");
            e.trHeight_val = mat2doc.shared.Twips(360);
            e.trHeight_hRule = mat2doc.enum.table.WD_ROW_HEIGHT_RULE.AT_LEAST;
            testCase.verifyEqual(double(e.trHeight_val), 228600, 'row trHeight_val via trPr');
            testCase.verifyEqual(e.trHeight_hRule, mat2doc.enum.table.WD_ROW_HEIGHT_RULE.AT_LEAST);
            testCase.verifyEqual(childLocalnames(e), "trPr", 'row created a single <w:trPr>');
            testCase.verifyEqual(hx_e(e), string(o.row.trheight_set.serhex), ...
                'row trHeight_set serhex (L1) vs frozen oracle');

            % ★ tr_idx (H1): 0-based index among sibling <w:tr> off a parsed parent
            tbl = parse("<w:tbl " + nsW() + "><w:tr/><w:tr/><w:tr/></w:tbl>");
            trs = tbl.findall(mat2doc.oxml.qn("w:tr"));
            testCase.verifyEqual(numel(trs), 3, 'parsed tbl has 3 rows');
            testCase.verifyEqual([trs(1).tr_idx trs(2).tr_idx trs(3).tr_idx], [0 1 2], ...
                'H1: tr_idx is 0-based [0 1 2]');

            % ★ custom inserters: tblPrEx forced to front, trPr immediately after.
            %   Add trPr FIRST (no tblPrEx yet -> insert(0)), then tblPrEx (forced
            %   to front) -> order [tblPrEx, trPr].
            e = mat2doc.oxml.OxmlElement("w:tr");
            e.get_or_add_trPr();
            e.get_or_add_tblPrEx();
            testCase.verifyEqual(childLocalnames(e), ["tblPrEx" "trPr"], ...
                'H11 inserters: [tblPrEx, trPr] (tblPrEx front, trPr after)');
            testCase.verifyEqual(class(e.tblPrEx), testCase.CT_TBLPREX, 'tblPrEx is CT_TblPrEx');
            testCase.verifyEqual(class(e.trPr), testCase.CT_TRPR, 'trPr is CT_TrPr');
            testCase.verifyEqual(hx_e(e), string(o.row.inserters_no_tc.serhex), ...
                'row inserters-no-tc serhex (L1) vs frozen oracle');

            % with a present <w:tc>: trPr before tc; tblPrEx forced front -> [tblPrEx, trPr, tc]
            e = parse("<w:tr " + nsW() + "><w:tc/></w:tr>");
            e.get_or_add_trPr();
            e.get_or_add_tblPrEx();
            testCase.verifyEqual(childLocalnames(e), ["tblPrEx" "trPr" "tc"], ...
                'H11 inserters with tc: [tblPrEx, trPr, tc]');
            testCase.verifyEqual(hx_e(e), string(o.row.inserters_with_tc.serhex), ...
                'row inserters-with-tc serhex (L1) vs frozen oracle');
        end

        % =============================================================== %
        % 7. CT_Row -- the CT_Tc P6-3a BOUNDARY (guards, not mis-walks)      %
        % =============================================================== %

        function test_ct_row_tc_boundary(testCase)
            % Edge (P6-3a boundary): _new_tc / add_tc / add_tc_ raise
            % mat2doc:notYetPorted (CT_Tc.new is P6-3a); tc_at_grid_offset raises
            % mat2doc:notYetPorted at the grid_span dependency (never mis-walks);
            % the fast paths (offset==grid_before -> first cell; empty row ->
            % mat2doc:ValueError verbatim) work now.
            r = mat2doc.oxml.OxmlElement("w:tr");

            % _new_tc / add_tc / add_tc_ -> notYetPorted (P6-3a)
            testCase.verifyError(@() r.new_tc_(), 'mat2doc:notYetPorted', ...
                'CT_Row.new_tc_ raises notYetPorted (CT_Tc.new is P6-3a)');
            testCase.verifyError(@() r.add_tc(), 'mat2doc:notYetPorted', ...
                'CT_Row.add_tc raises notYetPorted (routes through _new_tc)');
            testCase.verifyError(@() r.add_tc_(), 'mat2doc:notYetPorted', ...
                'CT_Row.add_tc_ raises notYetPorted');

            % tc_at_grid_offset fast path: offset==grid_before(0) returns the FIRST
            % tc WITHOUT touching grid_span (works now, generic tc ok)
            row1 = parse("<w:tr " + nsW() + "><w:tc/></w:tr>");
            tc0 = row1.tc_at_grid_offset(0);
            testCase.verifyEqual(string(tc0.local_part), "tc", ...
                'tc_at_grid_offset(0) returns the first tc (fast path, no grid_span)');

            % tc_at_grid_offset that MUST walk grid_span (offset>grid_before, tc not
            % yet CT_Tc) -> notYetPorted at exactly the dependency point (P6-3a)
            testCase.verifyError(@() row1.tc_at_grid_offset(1), 'mat2doc:notYetPorted', ...
                'tc_at_grid_offset(1) raises notYetPorted at grid_span (P6-3a)');

            % empty row, any offset -> mat2doc:ValueError (fast path, no grid_span)
            emptyRow = mat2doc.oxml.OxmlElement("w:tr");
            testCase.verifyError(@() emptyRow.tc_at_grid_offset(5), 'mat2doc:ValueError', ...
                'tc_at_grid_offset on an empty row raises mat2doc:ValueError');
            testCase.verifyEqual(errmsg(@() emptyRow.tc_at_grid_offset(5)), ...
                "no `tc` element at grid_offset=5", ...
                'ValueError message verbatim (grid_offset=5)');
        end

        % =============================================================== %
        % 8. CT_TblPrEx -- bare pass-through container (byte-neutral)        %
        % =============================================================== %

        function test_ct_tblprex_passthrough(testCase)
            % Nominal + Edge: CT_TblPrEx declares NO descriptors/attrs/@property --
            % a pure pass-through so a parsed <w:tblPrEx> resolves to the named
            % class. Serialize is byte-identical to the generic element (no
            % accessors run on parse), so registering it is byte-neutral.
            e = mat2doc.oxml.OxmlElement("w:tblPrEx");
            testCase.verifyEqual(class(e), testCase.CT_TBLPREX, 'OxmlElement(w:tblPrEx) is CT_TblPrEx');
            testCase.verifyEqual(ser(e), decl() + newline + ...
                "<w:tblPrEx xmlns:w=""" + testCase.W + """/>", ...
                'bare CT_TblPrEx serialized bytes (L1) hard-coded');
            % parse round-trips byte-identical (pass-through neutrality)
            xml = "<w:tblPrEx " + nsW() + "/>";
            p = parse(xml);
            testCase.verifyEqual(class(p), testCase.CT_TBLPREX, 'parsed <w:tblPrEx> is CT_TblPrEx');
        end

        % =============================================================== %
        % 9. (R) s0062 TABLE-PROPS round-trip byte-pins (the freeze)         %
        % =============================================================== %

        function test_s0062_roundtrip_byte_pins(testCase)
            % (R) Regression + Upstream (byte-identical L1): each of the 6 frozen
            % python-docx fixtures is read, SHA-checked against the manifest, parsed
            % through the CT classes and RE-SERIALIZED byte-identical (SHA of the
            % re-serialized output == the frozen SHA). Includes the w:jc val="both"
            % edge + two REAL add_table subtrees (full 18-decl nsmap).
            tbl = fixtureTable();          % {name, size, sha}
            for i = 1:size(tbl, 1)
                name = tbl{i, 1};
                inBytes = readFixture(name);
                testCase.verifyEqual(numel(inBytes), tbl{i, 2}, ...
                    sprintf('%s fixture size', name));
                testCase.verifyEqual(sha256hex(inBytes), tbl{i, 3}, ...
                    sprintf('%s frozen fixture SHA-256 (intact)', name));
                root     = mat2doc.oxml.parse_xml(inBytes);
                outBytes = mat2doc.oxml.serialize_part_xml(root);
                testCase.verifyEqual(uint8(outBytes(:)'), uint8(inBytes(:)'), ...
                    sprintf('%s parse->serialize must be byte-identical', name));
                testCase.verifyEqual(sha256hex(outBytes), tbl{i, 3}, ...
                    sprintf('%s re-serialized SHA-256 == frozen oracle (L1)', name));
            end

            % ---- structural corroboration through the CT accessors ----
            % tblPr (loose): style TestTableStyle, alignment CENTER (JUSTIFY-family
            % paragraph enum), autofit False (tblLayout fixed)
            tp = mat2doc.oxml.parse_xml(readFixture('tblPr'));
            testCase.verifyEqual(class(tp), testCase.CT_TBLPR, 'tblPr parses as CT_TblPr');
            testCase.verifyEqual(string(tp.style), "TestTableStyle", 'tblPr style');
            testCase.verifyEqual(string(tp.alignment), "CENTER", 'tblPr alignment name CENTER');
            testCase.verifyFalse(tp.autofit, 'tblPr autofit False (tblLayout fixed)');

            % ★ tblPr_both edge: alignment -> JUSTIFY (the converting-getter crash)
            tpb = mat2doc.oxml.parse_xml(readFixture('tblPr_both'));
            testCase.verifyEqual(string(tpb.alignment), "JUSTIFY", ...
                'tblPr_both alignment -> JUSTIFY (A2 star edge)');
            testCase.verifyEqual(double(tpb.alignment.value), 3, 'tblPr_both JUSTIFY value 3');

            % trPr (loose): grid_before 1, grid_after 2, trHeight 360/atLeast
            trp = mat2doc.oxml.parse_xml(readFixture('trPr'));
            testCase.verifyEqual(class(trp), testCase.CT_TRPR, 'trPr parses as CT_TrPr');
            testCase.verifyEqual(trp.grid_before, 1, 'trPr grid_before 1');
            testCase.verifyEqual(trp.grid_after, 2,  'trPr grid_after 2');
            testCase.verifyEqual(double(trp.trHeight_val), 228600, 'trPr trHeight_val 360tw');
            testCase.verifyEqual(trp.trHeight_hRule, mat2doc.enum.table.WD_ROW_HEIGHT_RULE.AT_LEAST, ...
                'trPr trHeight_hRule AT_LEAST');

            % tr (loose CT_Row): tblPrEx + trPr via the custom inserters
            tr = mat2doc.oxml.parse_xml(readFixture('tr'));
            testCase.verifyEqual(class(tr), testCase.CT_ROW, 'tr parses as CT_Row');
            testCase.verifyEqual(childLocalnames(tr), ["tblPrEx" "trPr"], 'tr children [tblPrEx, trPr]');

            % realtblPr (real add_table, full nsmap): alignment CENTER, autofit
            % False, style "TableGrid" -- the real-nsmap parse path
            rtp = mat2doc.oxml.parse_xml(readFixture('realtblPr'));
            testCase.verifyEqual(class(rtp), testCase.CT_TBLPR, 'realtblPr parses as CT_TblPr');
            testCase.verifyEqual(string(rtp.alignment), "CENTER", 'realtblPr alignment CENTER');
            testCase.verifyFalse(rtp.autofit, 'realtblPr autofit False');
            testCase.verifyEqual(string(rtp.style), "TableGrid", 'realtblPr style TableGrid');

            % realtrPr (real add_table row-0, full nsmap): trHeight Twips(500) EXACTLY
            rtr = mat2doc.oxml.parse_xml(readFixture('realtrPr'));
            testCase.verifyEqual(class(rtr), testCase.CT_TRPR, 'realtrPr parses as CT_TrPr');
            testCase.verifyEqual(double(rtr.trHeight_val), 317500, 'realtrPr trHeight_val 500tw');
            testCase.verifyEqual(rtr.trHeight_hRule, mat2doc.enum.table.WD_ROW_HEIGHT_RULE.EXACTLY, ...
                'realtrPr trHeight_hRule EXACTLY');
        end

        function test_s0062_manifest_matches_pinned(testCase)
            % Regression (fixture-drift guard): the shipped data\s0062\manifest.json
            % SHAs must equal the hard-coded pin table in this class. A silent
            % re-freeze would flip the manifest but not this constant.
            here = fileparts(mfilename('fullpath'));
            p = fullfile(here, 'data', 's0062', 'manifest.json');
            fid = fopen(p, 'r', 'n');
            testCase.assertGreaterThanOrEqual(fid, 0, 'cannot open s0062 manifest');
            raw = fread(fid, Inf, '*uint8')';
            fclose(fid);
            man = jsondecode(native2unicode(raw, 'UTF-8'));
            tbl = fixtureTable();
            testCase.verifyEqual(numel(man.fixtures), size(tbl, 1), '6 manifest fixtures');
            for i = 1:numel(man.fixtures)
                fx = man.fixtures(i);
                row = tbl(strcmp(tbl(:, 1), fx.name), :);
                testCase.verifyEqual(size(row, 1), 1, sprintf('manifest fixture %s pinned', fx.name));
                testCase.verifyEqual(fx.size, row{1, 2}, sprintf('%s size matches pin', fx.name));
                testCase.verifyEqual(string(fx.sha256), row{1, 3}, sprintf('%s SHA matches pin', fx.name));
            end
        end

        % =============================================================== %
        % 10. (M) M1 styles.xml + document.xml byte-pins                    %
        % =============================================================== %

        function test_m1_styles_and_document(testCase)
            % (M) Regression (byte-neutrality, L1): mat2doc.Document().save() emits
            % word/styles.xml at EXACTLY 349458 B (the frozen s0001 SHA) even though
            % its 100 <w:tblPr> nodes now transit CT_TblPr on the M1 parse path --
            % the registration is byte-neutral. word/document.xml stays 1548 B (no
            % table tag transits on M1). SHA-256 equality is an L1 assertion.
            % A single save() emits both parts.
            [styBytes, docBytes] = emitTwoParts('styles.xml', 'document.xml');

            testCase.verifyEqual(numel(styBytes), testCase.STYLES_SIZE_M1, ...
                sprintf('word/styles.xml must be exactly %d B (100 w:tblPr nodes transit CT_TblPr)', ...
                    testCase.STYLES_SIZE_M1));
            testCase.verifyEqual(sha256hex(styBytes), testCase.STYLES_SHA_M1, ...
                'word/styles.xml SHA-256 == frozen s0001 oracle (CT_TblPr parse-path neutral, L1)');

            testCase.verifyEqual(numel(docBytes), testCase.DOC_SIZE_M1, ...
                sprintf('word/document.xml must be exactly %d B', testCase.DOC_SIZE_M1));
            testCase.verifyEqual(sha256hex(docBytes), testCase.DOC_SHA_M1, ...
                'word/document.xml SHA-256 == frozen s0001 oracle (byte-identical L1)');
        end

        % =============================================================== %
        % 11. (H11) scrambled-build -> canonical order (byte-identical)      %
        % =============================================================== %

        function test_h11_canonical_order(testCase)
            % (H11) Regression: out-of-order API adds land in schema (_tag_seq)
            % order. Build the loose tblPr / trPr fixtures via the API with the
            % children added in a SCRAMBLED order and assert byte-identical to the
            % frozen s0062 fixtures (SHA b5de419f / c1a760cc).
            tbl = fixtureTable();
            shaTblPr = tbl{strcmp(tbl(:, 1), 'tblPr'), 3};
            shaTrPr  = tbl{strcmp(tbl(:, 1), 'trPr'), 3};

            % tblPr target: <w:tblStyle val="TestTableStyle"/><w:jc val="center"/>
            %               <w:tblLayout type="fixed"/>  (schema idx 1 / 8 / 13)
            % Build SCRAMBLED: autofit(false)->tblLayout, then alignment CENTER->jc
            % (must land BEFORE tblLayout), then style->tblStyle (must land FIRST).
            e = mat2doc.oxml.OxmlElement("w:tblPr");
            e.autofit = false;                                        % adds tblLayout
            e.alignment = mat2doc.enum.table.WD_TABLE_ALIGNMENT.CENTER;% adds jc before tblLayout
            e.style = "TestTableStyle";                                % adds tblStyle at front
            testCase.verifyEqual(childLocalnames(e), ["tblStyle" "jc" "tblLayout"], ...
                'H11: scrambled tblPr adds land in schema order');
            testCase.verifyEqual(sha256hex(mat2doc.oxml.serialize_part_xml(e)), shaTblPr, ...
                'H11: scrambled-built tblPr == frozen tblPr fixture (b5de419f, byte-identical)');

            % trPr target: <w:gridBefore val="1"/><w:gridAfter val="2"/>
            %              <w:trHeight val="360" hRule="atLeast"/> (schema idx 3/4/8)
            % Build SCRAMBLED: gridAfter(2) first, then gridBefore(1) (must land
            % BEFORE gridAfter), then trHeight.
            e = mat2doc.oxml.OxmlElement("w:trPr");
            ga = e.get_or_add_gridAfter();  ga.val = 2;               % gridAfter first
            gb = e.get_or_add_gridBefore(); gb.val = 1;               % gridBefore -> before gridAfter
            e.trHeight_val = mat2doc.shared.Twips(360);
            e.trHeight_hRule = mat2doc.enum.table.WD_ROW_HEIGHT_RULE.AT_LEAST;
            testCase.verifyEqual(childLocalnames(e), ["gridBefore" "gridAfter" "trHeight"], ...
                'H11: scrambled trPr adds land in schema order (gridBefore before gridAfter)');
            testCase.verifyEqual(sha256hex(mat2doc.oxml.serialize_part_xml(e)), shaTrPr, ...
                'H11: scrambled-built trPr == frozen trPr fixture (c1a760cc, byte-identical)');
        end

        % =============================================================== %
        % 12. EQUIVALENCE -- full s0061 battery vs the frozen oracle          %
        % =============================================================== %

        function test_equivalence_full_battery_vs_frozen_oracle(testCase)
            % Equivalence: replay the ENTIRE s0061 battery (runProbes -- the .m
            % twin's body VERBATIM: tblpr / trpr / row) and flatten-compare EVERY
            % leaf to the frozen python-docx 1.2.0 oracle copied into
            % data\s0061_probe_oracle.json. Gate-3 found ZERO divergences (probe_diff
            % exit 0), so every leaf must be byte/value-identical.
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
            % Non-trivial size guard (guards a silent-empty replay). The only looser-
            % than-byte assertion in this class; justified as a floor on leaf count.
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

function r = enum_read(x)
    % Capture an enum read by NAME + int VALUE + simple class NAME (the A2 idiom;
    % NEVER a cross-class ==). None ([]) -> all "None". Mirrors the s0061 .m twin.
    if isequal(x, [])
        r = struct("name", "None", "value", "None", "ecls", "None");
        return
    end
    cparts = split(string(class(x)), ".");
    r = struct("name", string(x), ...
               "value", string(sprintf('%.0f', double(x.value))), ...
               "ecls", cparts(end));
end

function s = rv(x)
    % Uniform accessor repr mirroring the s0061 rv(): None->"None",
    % enum->member NAME, bool->"True"/"False", Length/int->EMU/int decimal.
    if isequal(x, [])
        s = "None";
    elseif isenum(x)
        s = string(x);
    elseif islogical(x)
        if x, s = "True"; else, s = "False"; end
    elseif isa(x, 'mat2doc.shared.Length')
        s = string(sprintf('%.0f', double(x)));
    elseif isnumeric(x)
        s = string(sprintf('%.0f', double(x)));
    else
        s = string(x);
    end
end

function s = errmsg(fn)
    try
        fn();
        s = "NOERR";
    catch e
        s = string(e.message);
    end
end

function o = loadOracle()
    % Read the co-located frozen s0061 oracle in BINARY mode (no CRLF translation)
    % and decode UTF-8 -> struct. jsondecode is whitespace/line-ending agnostic, so
    % no `* binary` .gitattributes pin is needed (value-based fixture; s0059/s0030
    % precedent).
    here = fileparts(mfilename('fullpath'));
    p = fullfile(here, 'data', 's0061_probe_oracle.json');
    fid = fopen(p, 'r', 'n');
    if fid < 0
        error('mat2doc:test:openFail', 'cannot open frozen oracle %s', p);
    end
    raw = fread(fid, Inf, '*uint8')';
    fclose(fid);
    o = jsondecode(native2unicode(raw, 'UTF-8'));
end

% ---- s0062 frozen byte fixtures ----

function tbl = fixtureTable()
    % Hard-coded pin table {name, size, sha256} for the 6 frozen s0062 fixtures
    % (validate report section 3 / manifest.json). SHAs are of the python-docx
    % serialize_part_xml bytes == the MATLAB parse->serialize output (round-trip).
    tbl = { ...
        'tblPr',      233,  "b5de419fd6fe69ee35eebda6cae2b808d9366a01c9815f300cf24db4c2d42dc6"; ...
        'tblPr_both', 166,  "620dea7a2cafb72478572b8da770bda450fc5e170d53c7b739195743a9b74490"; ...
        'trPr',       236,  "c1a760cc58e92660a8fec10f19d3d4616cd655e3ea574b32955e603990f90997"; ...
        'tr',         261,  "6bcc2d38cb1716b16f44cc15686579e4ccaf396ff0e82f92f51df16ae4282531"; ...
        'realtblPr',  1451, "de2f695c9549d58826bb9dae181af5c950a85bda68b03425a2d4728cf9fc5a42"; ...
        'realtrPr',   1260, "8e0b1fc36ba356b5f489b716734a6c6965a2dba58f40612ecdea2bc57c640022" };
end

function b = readFixture(name)
    here = fileparts(mfilename('fullpath'));
    p = fullfile(here, 'data', 's0062', 'parts', [name '.xml']);
    b = readBytes(p);
end

% ---- M1 styles.xml + document.xml emit (one save, two parts) ----

function [aBytes, bBytes] = emitTwoParts(leafA, leafB)
    % Document().save() to a temp .docx, unzip once, return two word/<leaf> parts.
    d = mat2doc.Document();
    tmp = [tempname '.docx'];
    cleanTmp = onCleanup(@() deleteIfExists(tmp)); %#ok<NASGU>
    d.save(tmp);
    exdir = tempname;
    cleanEx = onCleanup(@() rmdirIfExists(exdir)); %#ok<NASGU>
    unzip(tmp, exdir);
    aBytes = readBytes(fullfile(exdir, 'word', leafA));
    bBytes = readBytes(fullfile(exdir, 'word', leafB));
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

% ---- s0061 equivalence replay (the .m twin body, VERBATIM) ----

function P = runProbes()
    % Replay the s0061 probe sequence (the .m twin's P-struct body, VERBATIM
    % tags/inputs/order) and return the nested struct of tagged canonical values.
    % Embedded here so the Equivalence leg is self-contained (the validation-folder
    % scenario is NOT on the toolbox path and must not be a dependency). Mirrors
    % validation\mat2doc\scenarios\s0061_p6_2_table_props_probe.m lines 22-113.
    TA = @(n) mat2doc.enum.table.WD_TABLE_ALIGNMENT.(n);
    HR = @(n) mat2doc.enum.table.WD_ROW_HEIGHT_RULE.(n);
    TW = @(v) mat2doc.shared.Twips(v);

    P = struct();

    % ============================== tblpr =============================
    tp = struct();
    e = parse("<w:tblPr " + nsW() + "><w:jc w:val=""center""/></w:tblPr>");
    tp.align_parse_center = enum_read(e.alignment);
    e = parse("<w:tblPr " + nsW() + "><w:jc w:val=""both""/></w:tblPr>");
    tp.align_parse_both = enum_read(e.alignment);      % -> JUSTIFY (int 3)
    for nm = ["LEFT", "CENTER", "RIGHT"]
        e = nw("w:tblPr"); e.alignment = TA(nm);
        rec = enum_read(e.alignment);
        rec.serhex = hx_e(e);
        tp.("align_set_" + nm) = rec;
    end
    e = parse("<w:tblPr " + nsW() + "><w:jc w:val=""center""/></w:tblPr>");
    e.alignment = [];
    tp.align_set_none_removes = struct("serhex", hx_e(e), ...
        "jc_present", rv(~isequal(e.find(mat2doc.oxml.qn("w:jc")), [])));
    tp.align_bare_none = enum_read(nw("w:tblPr").alignment);
    tp.autofit_bare = struct("read", rv(nw("w:tblPr").autofit));
    e = parse("<w:tblPr " + nsW() + "><w:tblLayout w:type=""fixed""/></w:tblPr>");
    tp.autofit_parse_fixed = struct("read", rv(e.autofit));
    e = parse("<w:tblPr " + nsW() + "><w:tblLayout w:type=""autofit""/></w:tblPr>");
    tp.autofit_parse_autofit = struct("read", rv(e.autofit));
    e = nw("w:tblPr"); e.autofit = false;
    tp.autofit_set_false = struct("serhex", hx_e(e), "read", rv(e.autofit));
    e = nw("w:tblPr"); e.autofit = true;
    tp.autofit_set_true = struct("serhex", hx_e(e), "read", rv(e.autofit));
    tp.style_bare = struct("read", rv(nw("w:tblPr").style));
    e = nw("w:tblPr"); e.style = "TestTableStyle";
    tp.style_set = struct("serhex", hx_e(e), "read", rv(e.style));
    e = nw("w:tblPr"); e.style = "TestTableStyle"; e.style = [];
    tp.style_set_none_removes = struct("serhex", hx_e(e), "read", rv(e.style));
    e = parse("<w:tblPr " + nsW() + "><w:tblStyle w:val=""TableGrid""/></w:tblPr>");
    tp.style_parse_read = struct("read", rv(e.style));
    P.tblpr = tp;

    % ============================== trpr =============================
    tr = struct();
    b = nw("w:trPr");
    tr.grid_bare = struct("before", rv(b.grid_before), "after", rv(b.grid_after));
    e = nw("w:trPr");
    gb = e.get_or_add_gridBefore(); gb.val = 1;
    ga = e.get_or_add_gridAfter();  ga.val = 2;
    tr.grid_set = struct("before", rv(e.grid_before), "after", rv(e.grid_after), ...
        "serhex", hx_e(e));
    b = nw("w:trPr");
    tr.trheight_bare = struct("val", rv(b.trHeight_val), "hRule", rv(b.trHeight_hRule));
    e = nw("w:trPr"); e.trHeight_val = TW(360); e.trHeight_hRule = HR("AT_LEAST");
    tr.trheight_set_both = struct("serhex", hx_e(e), "val", rv(e.trHeight_val), ...
        "hRule", rv(e.trHeight_hRule));
    e = nw("w:trPr"); e.trHeight_val = [];
    tr.trheight_val_none_guard = struct("serhex", hx_e(e), ...
        "trHeight_present", rv(~isequal(e.trHeight, [])));
    e = nw("w:trPr"); e.trHeight_hRule = [];
    tr.trheight_hrule_none_guard = struct("serhex", hx_e(e), ...
        "trHeight_present", rv(~isequal(e.trHeight, [])));
    P.trpr = tr;

    % ============================== row ==============================
    rw = struct();
    b = nw("w:tr");
    rw.grid_bare = struct("before", rv(b.grid_before), "after", rv(b.grid_after));
    e = nw("w:tr"); e.trHeight_val = TW(360); e.trHeight_hRule = HR("AT_LEAST");
    rw.trheight_set = struct("serhex", hx_e(e), "val", rv(e.trHeight_val), ...
        "hRule", rv(e.trHeight_hRule));
    tbl = parse("<w:tbl " + nsW() + "><w:tr/><w:tr/><w:tr/></w:tbl>");
    trs = tbl.findall(mat2doc.oxml.qn("w:tr"));
    idx = cell(1, numel(trs));
    for k = 1:numel(trs); idx{k} = rv(trs(k).tr_idx); end
    rw.tr_idx = idx;
    e = nw("w:tr");
    e.get_or_add_trPr();        % no tblPrEx yet -> insert(0)
    e.get_or_add_tblPrEx();     % forced to front -> insert(0)
    rw.inserters_no_tc = struct("localnames", {lnsCell(e)}, "serhex", hx_e(e));
    e = parse("<w:tr " + nsW() + "><w:tc/></w:tr>");   % a present <w:tc>
    e.get_or_add_trPr();        % trPr before tc
    e.get_or_add_tblPrEx();     % tblPrEx to front
    rw.inserters_with_tc = struct("localnames", {lnsCell(e)}, "serhex", hx_e(e));
    P.row = rw;
end

function e = nw(tag)
    e = mat2doc.oxml.OxmlElement(tag);
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

classdef Test_p6_1_table_leaves < matlab.unittest.TestCase
% TEST_P6_1_TABLE_LEAVES  Gate-4 permanent unit tests for Mat2Doc P6-1
%   (oxml table width/height/grid LEAF classes): src/docx/oxml/table.py ->
%   +mat2doc\+oxml\+table\ CT_Height / CT_TblWidth / CT_TblGrid / CT_TblGridCol /
%   CT_TblLayoutType / CT_VerticalJc / CT_VMerge, and the SEVEN table-block
%   registry rows in +mat2doc\+oxml\registry.m (w:gridCol->CT_TblGridCol;
%   w:tblGrid->CT_TblGrid; w:tblLayout->CT_TblLayoutType; w:tcW->CT_TblWidth;
%   w:trHeight->CT_Height; w:vAlign->CT_VerticalJc; w:vMerge->CT_VMerge).
%
%   P6-1 is the FIRST Phase-6 (tables) WP and is REGISTRY-ADDING but
%   M1-NEUTRAL + flip-neutral: default.docx carries no table, so NONE of the 7
%   new tags is present in any of its 17 parts. Registering the 7 CT leaf classes
%   therefore moves ZERO bytes of any M1 part -- word/document.xml stays 1548 B,
%   SHA-256 0e4dd503... This class permanently freezes the guarantees the prior
%   gates established:
%     * Gate-1 Porter  : audit_P6-1_table_leaves.md (self-probe).
%     * Gate-2 Auditor (Fable): audit_P6-1_table_leaves.md GATE-2 -- APPROVE;
%       the width UNION proven exact.
%     * Gate-3 Validator: validate_P6-1_table_leaves.md -- PASS, ZERO new
%       D-numbers, NO re-pin list (M1/flip-neutral). probe_diff s0059 MATCH
%       (exit 0) over the full leaf surface (value + serhex + verbatim error
%       messages); M1 17/17 (document.xml 1548 B / 0e4dd503 byte-identical);
%       the s0060 table-leaf round-trip 10/10 byte-identical; the width union
%       byte-proven two ways; registry C4 confirmed (w:tblW deliberately
%       UNregistered).
%
%   THE HIGHEST-VALUE PERMANENT PINS IN THIS WP (a regression goes RED loudly):
%     * (U) CT_TblWidth WIDTH UNION (test_ct_tblwidth_width_union): the H6 crux.
%       dxa -> width is Twips(@w:w) (dxa 2880 -> 1828800 EMU, a Length); EVERY
%       other @w:type (pct/auto/nil) -> width is [] (None); the width SETTER
%       forces @w:type="dxa" and @w:w=Emu(value).twips (Twips(1440) -> dxa/1440,
%       width 914400). The three RequiredAttribute-absent reads (.w, .type,
%       .width -- which reads .type FIRST) all raise mat2doc:InvalidXmlError with
%       the verbatim "required 'w:...' attribute not present on element {...}tcW".
%       Pinned by build-from-scratch serhex (s0059) AND the frozen s0060 loose-tcW
%       byte fixtures.
%     * (R) s0060 TABLE-LEAF ROUND-TRIP byte-pins (test_s0060_roundtrip_byte_pins):
%       10 frozen python-docx fixtures from a REAL 3x2 table (grid + tblLayout +
%       trPr trHeight + cell tcW/vAlign/vMerge + 4 loose width-union tcW). Each
%       fixture's bytes parse through the CT leaf classes and re-serialize
%       BYTE-IDENTICAL (SHA-256 == the frozen manifest). The whole-grid +
%       representative-cell SHAs are the permanent table-leaf guards.
%     * (M) M1 document.xml byte-pin (test_m1_document_xml_byte_identical): the
%       registry-neutrality guard -- mat2doc.Document().save() -> word/document.xml
%       == 1548 B, SHA-256 0e4dd503... (frozen s0001 oracle). The 7 table rows must
%       NOT perturb M1. SHA equality is an L1 assertion.
%     * (H11) CT_TblGrid.add_gridCol SUCCESSOR insertion (test_ct_tblgrid): a new
%       gridCol lands BEFORE the first <w:tblGridChange> (successors) -> new_idx=1;
%       on a plain grid every add APPENDS. serhex vs the frozen oracle.
%     * (D) CT_VMerge non-None DEFAULT (test_ct_vmerge): the sole non-None default
%       in this WP -- absent @w:val reads "continue"; set "restart" writes; set
%       "continue" (== default) OR [] (None) REMOVES @w:val; set "" is REJECTED
%       (mat2doc:ValueError "must be one of ('continue', 'restart'), got ''").
%     * (E) enum BREADTH: CT_Height.hRule over ALL WD_ROW_HEIGHT_RULE members;
%       CT_VerticalJc.val over ALL WD_CELL_VERTICAL_ALIGNMENT members.
%
%   Provenance (all Gate-3 frozen 2026-07-31):
%     * Audit    : validation\mat2doc\audit_P6-1_table_leaves.md
%     * Validate : validation\mat2doc\validate_P6-1_table_leaves.md
%     * Scenarios: validation\mat2doc\scenarios\s0059_p6_1_table_leaves_probe.{py,m}
%                  (its probe body is replayed VERBATIM by runProbes() below);
%                  s0060_p6_1_table_roundtrip.{py,m} (the 10 byte fixtures).
%     * Frozen refs (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE at Gate-3):
%         references\s0059\probe.json -- copied verbatim (self-contained) into
%           tests\oxml\data\s0059_probe_oracle.json (the Equivalence replay set;
%           value/serhex JSON, jsondecode is line-ending agnostic -> no `* binary`
%           .gitattributes pin needed, per the s0030/s0036/s0039 precedent).
%         references\s0060\ (10 real-table byte fixtures + manifest.json) -- copied
%           verbatim into tests\oxml\data\s0060\ WITH a co-located `.gitattributes`
%           `* binary` pin (frozen-byte fixtures must not be line-ending mangled on
%           the master checkout -- the Gate-4 byte-fixture lesson).
%         references\s0001\parts\word\document.xml -- the M1 byte reference
%           (1548 / 0e4dd503); NOT copied (SHA of what Document().save() emits).
%
%   Coverage taxonomy
%   -----------------
%   * Nominal      -- width union dxa/setter; CT_Height val+hRule; CT_TblGridCol w
%                     + gridCol_idx; CT_TblGrid add_gridCol; CT_TblLayoutType type;
%                     CT_VerticalJc val (all members); CT_VMerge val; registry.
%   * Edge         -- width union pct/auto/nil -> None; None ([]) removes on every
%                     OptionalAttribute; the RequiredAttribute-absent error paths
%                     (mat2doc:InvalidXmlError, verbatim message); H11 insertion
%                     before tblGridChange; CT_VMerge "" rejection (mat2doc:
%                     ValueError, verbatim tuple-repr message); non-ASCII is N/A
%                     (these leaves carry only measures/enums/fixed simple-types).
%   * Equivalence  -- test_equivalence_full_battery_vs_frozen_oracle replays the
%                     ENTIRE s0059 battery (runProbes, the .m twin's body verbatim)
%                     and flatten-compares every leaf to the frozen oracle (0 diffs).
%   * Regression   -- hard-coded expected serialized-XML strings (ASCII == byte-
%                     identical L1) + UPPERCASE serhex vs the frozen oracle + SHA-256
%                     of the 10 s0060 fixtures + the M1 document.xml part.
%   * Upstream     -- the s0060 subtrees are REAL python-docx table output; the
%                     frozen oracle IS lxml's expected serialization for them.
%
%   Byte-level (L1) note: every serialized-XML comparison is either the FULL
%   serialize_part_xml output decoded as an ASCII string (string-equality ==
%   byte-equality L1), or its UPPERCASE hex (serhex) vs the frozen oracle, or the
%   SHA-256 of a frozen fixture / emitted document.xml part. NO D-number granted
%   any L2 relaxation in this WP (Gate-3: zero new, none at L2), so every pin here
%   is L1. The equivalence leaf-key-count guard is the only looser-than-byte check
%   and is commented at its site.
%
%   Determinism: no network, no absolute paths. The worktree root, the co-located
%   s0059 oracle and the s0060 byte fixtures resolve relative to this file via
%   fileparts(mfilename('fullpath')); saves go to tempname .docx deleted via
%   onCleanup; every file read is binary ('r','n'). The +mat2doc package resolves
%   via the MANDATORY PathFixture(worktree-root) in TestClassSetup (WP9-F4 lesson).

    properties (Constant)
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        % --- registered class names (P6-1, +oxml\registry.m 281-287) ---
        CT_TBLGRIDCOL = 'mat2doc.oxml.table.CT_TblGridCol'
        CT_TBLGRID    = 'mat2doc.oxml.table.CT_TblGrid'
        CT_TBLLAYOUT  = 'mat2doc.oxml.table.CT_TblLayoutType'
        CT_TBLWIDTH   = 'mat2doc.oxml.table.CT_TblWidth'
        CT_HEIGHT     = 'mat2doc.oxml.table.CT_Height'
        CT_VERTICALJC = 'mat2doc.oxml.table.CT_VerticalJc'
        CT_VMERGE     = 'mat2doc.oxml.table.CT_VMerge'

        % --- frozen s0001 M1 word/document.xml byte reference (the P6-1
        %     registry-adding neutrality guard on the CENTRAL part) ---
        DOC_SIZE_M1 = 1548
        DOC_SHA_M1  = "0e4dd503bc095219cd24e8312b372c588ca56bb834aacac53eafb7bf47836327"
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied verbatim from tests\oxml\Test_p5_2a_sectpr.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\oxml
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. registry rows (the 7 P6-1 flips + C4 w:tblW UNregistered)      %
        % =============================================================== %

        function test_registry_resolves_seven_rows(testCase)
            % Nominal / Regression (registry.m 281-287): the 7 table-leaf rows.
            % REGISTRATION is what flips the parse class of a table subtree.
            pairs = { ...
                "w:gridCol",   testCase.CT_TBLGRIDCOL; ...
                "w:tblGrid",   testCase.CT_TBLGRID; ...
                "w:tblLayout", testCase.CT_TBLLAYOUT; ...
                "w:tcW",       testCase.CT_TBLWIDTH; ...
                "w:trHeight",  testCase.CT_HEIGHT; ...
                "w:vAlign",    testCase.CT_VERTICALJC; ...
                "w:vMerge",    testCase.CT_VMERGE };
            for i = 1:size(pairs, 1)
                tag = pairs{i, 1}; cls = pairs{i, 2};
                r = mat2doc.oxml.registry(mat2doc.oxml.qn(tag));
                testCase.verifyEqual(char(r), cls, ...
                    sprintf('registry must resolve %s -> %s (P6-1 row)', tag, cls));
                % a real OxmlElement(tag) is the exact-class CT
                testCase.verifyEqual(class(mat2doc.oxml.OxmlElement(tag)), cls, ...
                    sprintf('OxmlElement(%s) must be a %s', tag, cls));
            end
            % C4: w:tblW is deliberately NOT registered (only w:tcW is), so the
            % registry returns "" and OxmlElement(w:tblW) falls back to generic.
            testCase.verifyEqual(mat2doc.oxml.registry(mat2doc.oxml.qn("w:tblW")), "", ...
                'C4: w:tblW must be UNREGISTERED (only w:tcW registered to CT_TblWidth)');
            testCase.verifyNotEqual(class(mat2doc.oxml.OxmlElement("w:tblW")), testCase.CT_TBLWIDTH, ...
                'C4: OxmlElement(w:tblW) must NOT be a CT_TblWidth');
        end

        % =============================================================== %
        % 2. CT_TblWidth -- the WIDTH UNION (H6 crux)                       %
        % =============================================================== %

        function test_ct_tblwidth_width_union(testCase)
            % (U) Nominal + Edge + Regression (s0059 tblwidth): the width union.
            % dxa -> Twips(@w:w) EMU Length; pct/auto/nil -> [] (None); the setter
            % forces type="dxa"; the RequiredAttribute-absent reads raise
            % mat2doc:InvalidXmlError (verbatim message). serhex vs the frozen oracle.
            o = loadOracle();

            % -- dxa: width is a Length (Twips(2880) = 1828800 EMU) --
            e = parse("<w:tcW " + nsW() + " w:type=""dxa"" w:w=""2880""/>");
            testCase.verifyEqual(class(e), testCase.CT_TBLWIDTH);
            testCase.verifyEqual(string(e.type), "dxa");
            testCase.verifyEqual(double(e.w), 2880, 'dxa @w:w read as int');
            testCase.verifyTrue(isa(e.width, 'mat2doc.shared.Length'), 'dxa width is a Length');
            testCase.verifyEqual(double(e.width), 1828800, 'dxa width EMU exact (Twips(2880))');
            testCase.verifyEqual(hx_e(e), string(o.tblwidth.dxa.serhex), ...
                'tcW dxa serhex (L1) vs frozen oracle');

            % -- pct / auto / nil: width is [] (None), NOT a divergence (the union) --
            unionNone = { ...
                "<w:tcW " + nsW() + " w:type=""pct"" w:w=""5000""/>", o.tblwidth.pct.serhex, 'pct'; ...
                "<w:tcW " + nsW() + " w:type=""auto"" w:w=""0""/>",   o.tblwidth.auto.serhex, 'auto'; ...
                "<w:tcW " + nsW() + " w:type=""nil"" w:w=""0""/>",    o.tblwidth.nil.serhex,  'nil' };
            for i = 1:size(unionNone, 1)
                e = parse(unionNone{i, 1});
                testCase.verifyTrue(isequal(e.width, []), ...
                    sprintf('%s width -> [] (None) -- the faithful union', unionNone{i, 3}));
                testCase.verifyEqual(hx_e(e), string(unionNone{i, 2}), ...
                    sprintf('tcW %s serhex (L1) vs frozen oracle', unionNone{i, 3}));
            end

            % -- width SETTER forces type="dxa" + w = Emu(value).twips --
            e = mat2doc.oxml.OxmlElement("w:tcW");
            e.width = mat2doc.shared.Twips(1440);   % Length -> stored as twips
            testCase.verifyEqual(string(e.type), "dxa", 'setter forces @w:type="dxa"');
            testCase.verifyEqual(double(e.w), 1440, 'setter writes @w:w = twips count');
            testCase.verifyEqual(double(e.width), 914400, 'setter width round-trips (Twips(1440))');
            testCase.verifyEqual(hx_e(e), string(o.tblwidth.setter_forces_dxa.serhex), ...
                'tcW width-setter serhex (L1) vs frozen oracle');

            % -- RequiredAttribute-absent: .w / .type / .width all raise --
            bare = mat2doc.oxml.OxmlElement("w:tcW");
            testCase.verifyError(@() bare.w, 'mat2doc:InvalidXmlError', ...
                'absent @w:w raises mat2doc:InvalidXmlError');
            testCase.verifyError(@() bare.type, 'mat2doc:InvalidXmlError', ...
                'absent @w:type raises mat2doc:InvalidXmlError');
            % width reads type FIRST, so it raises the @w:type error, not @w:w
            testCase.verifyError(@() bare.width, 'mat2doc:InvalidXmlError', ...
                'absent @w:type via .width raises mat2doc:InvalidXmlError (reads type first)');
            testCase.verifyEqual(errmsg(@() bare.w), string(o.tblwidth.err_w_absent), ...
                'absent @w:w message verbatim');
            testCase.verifyEqual(errmsg(@() bare.type), string(o.tblwidth.err_type_absent), ...
                'absent @w:type message verbatim');
            testCase.verifyEqual(errmsg(@() bare.width), string(o.tblwidth.err_width_absent), ...
                'absent @w:type-via-width message verbatim (== err_type_absent)');
        end

        % =============================================================== %
        % 3. CT_Height -- val (Length) + hRule (enum, ALL members)          %
        % =============================================================== %

        function test_ct_height(testCase)
            % Nominal + Edge + Regression (s0059 height): bare val/hRule None; set
            % both -> serhex; set-one-only; None removes; hRule over EVERY
            % WD_ROW_HEIGHT_RULE member. serhex vs the frozen oracle.
            o = loadOracle();

            b = mat2doc.oxml.OxmlElement("w:trHeight");
            testCase.verifyEqual(class(b), testCase.CT_HEIGHT);
            testCase.verifyTrue(isequal(b.val, []), 'bare trHeight val -> [] (None, H3)');
            testCase.verifyTrue(isequal(b.hRule, []), 'bare trHeight hRule -> [] (None, H3)');

            % set both (val Twips(500) = 317500 EMU; hRule EXACTLY -> "exact")
            e = mat2doc.oxml.OxmlElement("w:trHeight");
            e.val = mat2doc.shared.Twips(500);
            e.hRule = mat2doc.enum.table.WD_ROW_HEIGHT_RULE.EXACTLY;
            testCase.verifyEqual(double(e.val), 317500, 'val EMU exact (Twips(500))');
            testCase.verifyEqual(e.hRule, mat2doc.enum.table.WD_ROW_HEIGHT_RULE.EXACTLY);
            testCase.verifyEqual(hx_e(e), string(o.height.set_both.serhex), ...
                'trHeight set-both serhex (L1) vs frozen oracle');

            % set val only (Twips(360) = 228600 EMU)
            e = mat2doc.oxml.OxmlElement("w:trHeight");
            e.val = mat2doc.shared.Twips(360);
            testCase.verifyEqual(double(e.val), 228600);
            testCase.verifyEqual(hx_e(e), string(o.height.set_val_only.serhex), ...
                'trHeight set-val-only serhex (L1)');

            % set hRule only (AT_LEAST -> "atLeast")
            e = mat2doc.oxml.OxmlElement("w:trHeight");
            e.hRule = mat2doc.enum.table.WD_ROW_HEIGHT_RULE.AT_LEAST;
            testCase.verifyEqual(hx_e(e), string(o.height.set_hrule_only.serhex), ...
                'trHeight set-hRule-only serhex (L1)');

            % None removes each attribute
            e = mat2doc.oxml.OxmlElement("w:trHeight");
            e.val = mat2doc.shared.Twips(500); e.val = [];
            testCase.verifyTrue(isequal(e.val, []), 'val -> [] after None');
            testCase.verifyEqual(hx_e(e), string(o.height.val_none_removes.serhex), ...
                'set val None removes @w:val (L1)');
            e = mat2doc.oxml.OxmlElement("w:trHeight");
            e.hRule = mat2doc.enum.table.WD_ROW_HEIGHT_RULE.EXACTLY; e.hRule = [];
            testCase.verifyTrue(isequal(e.hRule, []), 'hRule -> [] after None');
            testCase.verifyEqual(hx_e(e), string(o.height.hrule_none_removes.serhex), ...
                'set hRule None removes @w:hRule (L1)');

            % ENUM BREADTH: every WD_ROW_HEIGHT_RULE member round-trips + serializes
            % its xml_value (auto/atLeast/exact). L1 hard-coded serialized bytes.
            members = enumeration('mat2doc.enum.table.WD_ROW_HEIGHT_RULE');
            xmlVals = struct('AUTO', "auto", 'AT_LEAST', "atLeast", 'EXACTLY', "exact");
            for k = 1:numel(members)
                name = char(members(k));
                e = mat2doc.oxml.OxmlElement("w:trHeight");
                e.hRule = members(k);
                testCase.verifyEqual(e.hRule, members(k), ...
                    sprintf('hRule round-trips %s', name));
                testCase.verifyEqual(ser(e), decl() + newline + ...
                    "<w:trHeight xmlns:w=""" + testCase.W + """ w:hRule=""" + ...
                    xmlVals.(name) + """/>", ...
                    sprintf('hRule %s serialized bytes (L1) hard-coded', name));
            end
        end

        % =============================================================== %
        % 4. CT_TblGridCol -- w + 0-based gridCol_idx                       %
        % =============================================================== %

        function test_ct_tblgridcol(testCase)
            % Nominal + Edge + Regression (s0059 gridcol): bare w None; set w ->
            % serhex; None removes; the 0-BASED gridCol_idx off a parsed 2-col grid.
            o = loadOracle();

            b = mat2doc.oxml.OxmlElement("w:gridCol");
            testCase.verifyEqual(class(b), testCase.CT_TBLGRIDCOL);
            testCase.verifyTrue(isequal(b.w, []), 'bare gridCol w -> [] (None, H3)');

            e = mat2doc.oxml.OxmlElement("w:gridCol");
            e.w = mat2doc.shared.Twips(2880);
            testCase.verifyEqual(double(e.w), 1828800, 'gridCol w EMU exact (Twips(2880))');
            testCase.verifyEqual(hx_e(e), string(o.gridcol.set_w.serhex), ...
                'gridCol set-w serhex (L1) vs frozen oracle');

            e.w = [];
            testCase.verifyTrue(isequal(e.w, []), 'w -> [] after None');
            testCase.verifyEqual(hx_e(e), string(o.gridcol.w_none_removes.serhex), ...
                'set w None removes @w:w (L1)');

            % 0-based gridCol_idx (H1): parse a real 2-col grid, index each col
            g = parse("<w:tblGrid " + nsW() + "><w:gridCol w:w=""2880""/><w:gridCol w:w=""1440""/></w:tblGrid>");
            cols = g.gridCol_lst;
            testCase.verifyEqual(numel(cols), 2, 'parsed grid has 2 gridCol');
            testCase.verifyEqual(cols(1).gridCol_idx, 0, 'first gridCol_idx is 0 (0-based, H1)');
            testCase.verifyEqual(double(cols(1).w), 1828800, 'col0 w (Twips(2880))');
            testCase.verifyEqual(cols(2).gridCol_idx, 1, 'second gridCol_idx is 1 (0-based)');
            testCase.verifyEqual(double(cols(2).w), 914400, 'col1 w (Twips(1440))');
        end

        % =============================================================== %
        % 5. CT_TblGrid -- add_gridCol + H11 successor insertion            %
        % =============================================================== %

        function test_ct_tblgrid(testCase)
            % Nominal + Regression (s0059 tblgrid): two public add_gridCol appends
            % preserve order; H11 -- add_gridCol on a grid holding <w:tblGridChange>
            % inserts BEFORE it (new_idx=1). serhex vs the frozen oracle.
            o = loadOracle();

            g = mat2doc.oxml.OxmlElement("w:tblGrid");
            testCase.verifyEqual(class(g), testCase.CT_TBLGRID);
            c1 = g.add_gridCol(); c1.w = mat2doc.shared.Twips(2880);
            c2 = g.add_gridCol(); c2.w = mat2doc.shared.Twips(1440);
            testCase.verifyEqual(class(c1), testCase.CT_TBLGRIDCOL, 'add_gridCol returns CT_TblGridCol');
            testCase.verifyEqual(numel(g.gridCol_lst), 2, 'two appends -> 2 gridCol');
            testCase.verifyEqual(childLocalnames(g), ["gridCol" "gridCol"], 'appends preserve order');
            testCase.verifyEqual(hx_e(g), string(o.tblgrid.two_appends.serhex), ...
                'tblGrid two-appends serhex (L1) vs frozen oracle');

            % H11: a new gridCol lands BEFORE the first <w:tblGridChange> (successor)
            g = parse("<w:tblGrid " + nsW() + "><w:gridCol w:w=""100""/><w:tblGridChange w:id=""1""/></w:tblGrid>");
            ncol = g.add_gridCol(); ncol.w = mat2doc.shared.Twips(200);
            testCase.verifyEqual(childLocalnames(g), ["gridCol" "gridCol" "tblGridChange"], ...
                'H11: new gridCol inserted BEFORE tblGridChange');
            testCase.verifyEqual(ncol.gridCol_idx, 1, 'H11: new gridCol_idx is 1 (before change)');
            testCase.verifyEqual(hx_e(g), string(o.tblgrid.h11_before_change.serhex), ...
                'tblGrid H11 insertion serhex (L1) vs frozen oracle');
        end

        % =============================================================== %
        % 6. CT_TblLayoutType -- type (fixed / autofit simple-type)         %
        % =============================================================== %

        function test_ct_tbllayout(testCase)
            % Nominal + Edge + Regression (s0059 tbllayout): bare None; set
            % "fixed"/"autofit" (ST_TblLayoutType STRING simple-type); None removes.
            o = loadOracle();

            b = mat2doc.oxml.OxmlElement("w:tblLayout");
            testCase.verifyEqual(class(b), testCase.CT_TBLLAYOUT);
            testCase.verifyTrue(isequal(b.type, []), 'bare tblLayout type -> [] (None, H3)');

            e = mat2doc.oxml.OxmlElement("w:tblLayout");
            e.type = "fixed";
            testCase.verifyEqual(string(e.type), "fixed");
            testCase.verifyEqual(hx_e(e), string(o.tbllayout.set_fixed.serhex), ...
                'tblLayout type="fixed" serhex (L1) vs frozen oracle');

            e = mat2doc.oxml.OxmlElement("w:tblLayout");
            e.type = "autofit";
            testCase.verifyEqual(string(e.type), "autofit");
            testCase.verifyEqual(hx_e(e), string(o.tbllayout.set_autofit.serhex), ...
                'tblLayout type="autofit" serhex (L1)');

            e.type = [];
            testCase.verifyTrue(isequal(e.type, []), 'type -> [] after None');
            testCase.verifyEqual(hx_e(e), string(o.tbllayout.type_none_removes.serhex), ...
                'set type None removes @w:type (L1)');
        end

        % =============================================================== %
        % 7. CT_VerticalJc -- val (enum, ALL 4 members) + Required error    %
        % =============================================================== %

        function test_ct_verticaljc(testCase)
            % Nominal + Edge + Regression (s0059 verticaljc): every
            % WD_CELL_VERTICAL_ALIGNMENT member serializes <w:vAlign w:val="..."/>;
            % the RequiredAttribute-absent read raises mat2doc:InvalidXmlError.
            o = loadOracle();

            members = enumeration('mat2doc.enum.table.WD_CELL_VERTICAL_ALIGNMENT');
            testCase.verifyEqual(numel(members), 4, 'WD_CELL_VERTICAL_ALIGNMENT has 4 members');
            for k = 1:numel(members)
                name = char(members(k));
                e = mat2doc.oxml.OxmlElement("w:vAlign");
                testCase.verifyEqual(class(e), testCase.CT_VERTICALJC);
                e.val = members(k);
                testCase.verifyEqual(e.val, members(k), ...
                    sprintf('vAlign val round-trips %s', name));
                testCase.verifyEqual(hx_e(e), string(o.verticaljc.members.(name).serhex), ...
                    sprintf('vAlign %s serhex (L1) vs frozen oracle', name));
            end

            % RequiredAttribute-absent read raises + verbatim message
            bare = mat2doc.oxml.OxmlElement("w:vAlign");
            testCase.verifyError(@() bare.val, 'mat2doc:InvalidXmlError', ...
                'absent @w:val raises mat2doc:InvalidXmlError');
            testCase.verifyEqual(errmsg(@() bare.val), string(o.verticaljc.err_val_absent), ...
                'absent @w:val message verbatim');
        end

        % =============================================================== %
        % 8. CT_VMerge -- the sole non-None DEFAULT ("continue")            %
        % =============================================================== %

        function test_ct_vmerge(testCase)
            % (D) Nominal + Edge + Regression (s0059 vmerge): the ONE non-None
            % default. absent @w:val -> "continue"; set "restart" writes; set
            % "continue" (== default) OR [] (None) REMOVES @w:val; set "" is
            % REJECTED (mat2doc:ValueError, verbatim tuple-repr message).
            o = loadOracle();

            % bare (parsed) reads the "continue" default
            b = parse("<w:vMerge " + nsW() + "/>");
            testCase.verifyEqual(class(b), testCase.CT_VMERGE);
            testCase.verifyEqual(string(b.val), "continue", 'absent @w:val -> "continue" (default)');
            testCase.verifyEqual(hx_e(b), string(o.vmerge.bare_continue.serhex), ...
                'bare <w:vMerge/> serhex (L1) vs frozen oracle');
            % a freshly-created vMerge (attr absent) reads the same default
            n = mat2doc.oxml.OxmlElement("w:vMerge");
            testCase.verifyEqual(string(n.val), "continue", 'fresh vMerge -> "continue" default');

            % set "restart" -> @w:val="restart"
            e = mat2doc.oxml.OxmlElement("w:vMerge");
            e.val = "restart";
            testCase.verifyEqual(string(e.val), "restart");
            testCase.verifyEqual(hx_e(e), string(o.vmerge.set_restart.serhex), ...
                'vMerge set "restart" serhex (L1)');

            % set "continue" (== the default) REMOVES @w:val -> bare <w:vMerge/>
            e = mat2doc.oxml.OxmlElement("w:vMerge");
            e.val = "restart"; e.val = "continue";
            testCase.verifyEqual(string(e.val), "continue", 'reads "continue" after removal');
            testCase.verifyEqual(childAttrCount(e), 0, 'no @w:val remains after == default');
            testCase.verifyEqual(hx_e(e), string(o.vmerge.set_continue_removes.serhex), ...
                'set "continue" (== default) removes @w:val (L1)');

            % set [] (None) also removes
            e = mat2doc.oxml.OxmlElement("w:vMerge");
            e.val = "restart"; e.val = [];
            testCase.verifyEqual(string(e.val), "continue", 'reads "continue" after None removal');
            testCase.verifyEqual(hx_e(e), string(o.vmerge.set_none_removes.serhex), ...
                'set None ([]) removes @w:val (L1)');

            % set "" is REJECTED -- "" is a real string, NOT the default (H4)
            e = mat2doc.oxml.OxmlElement("w:vMerge");
            testCase.verifyError(@() setVmergeEmpty(e), 'mat2doc:ValueError', ...
                'set @w:val = "" raises mat2doc:ValueError (ST_Merge reject)');
            testCase.verifyEqual(errmsg(@() setVmergeEmpty(e)), string(o.vmerge.err_empty_string), ...
                '"" rejection message verbatim ("must be one of (''continue'', ''restart''), got '''')');
        end

        % =============================================================== %
        % 9. (R) s0060 TABLE-LEAF round-trip byte-pins (the freeze)         %
        % =============================================================== %

        function test_s0060_roundtrip_byte_pins(testCase)
            % (R) Regression + Upstream (byte-identical L1): each of the 10 frozen
            % python-docx fixtures (a REAL 3x2 table's subtrees + 4 loose tcW) is
            % read, SHA-checked against the manifest, parsed through the CT leaf
            % classes and RE-SERIALIZED byte-identical (SHA of the re-serialized
            % output == the frozen SHA). The whole-grid + representative-cell SHAs
            % are the permanent table-leaf guards.
            tbl = fixtureTable();          % {name, size, sha}
            for i = 1:size(tbl, 1)
                name = tbl{i, 1};
                inBytes = readFixture(name);
                % 1) the frozen fixture bytes match the manifest SHA (fixture intact)
                testCase.verifyEqual(numel(inBytes), tbl{i, 2}, ...
                    sprintf('%s fixture size', name));
                testCase.verifyEqual(sha256hex(inBytes), tbl{i, 3}, ...
                    sprintf('%s frozen fixture SHA-256 (intact)', name));
                % 2) parse -> serialize is byte-identical to the input (the guard)
                root     = mat2doc.oxml.parse_xml(inBytes);
                outBytes = mat2doc.oxml.serialize_part_xml(root);
                testCase.verifyEqual(uint8(outBytes(:)'), uint8(inBytes(:)'), ...
                    sprintf('%s parse->serialize must be byte-identical', name));
                testCase.verifyEqual(sha256hex(outBytes), tbl{i, 3}, ...
                    sprintf('%s re-serialized SHA-256 == frozen oracle (L1)', name));
            end

            % ---- structural corroboration through the CT accessors ----
            % grid: 2 gridCol, 0-based idx, dxa widths
            grid = mat2doc.oxml.parse_xml(readFixture('grid'));
            testCase.verifyEqual(class(grid), testCase.CT_TBLGRID, 'grid parses as CT_TblGrid');
            cols = grid.gridCol_lst;
            testCase.verifyEqual(numel(cols), 2);
            testCase.verifyEqual([cols(1).gridCol_idx cols(2).gridCol_idx], [0 1], 'grid col idx 0,1');
            testCase.verifyEqual(double(cols(1).w), 1828800, 'grid col0 w=2880tw');
            testCase.verifyEqual(double(cols(2).w), 914400,  'grid col1 w=1440tw');

            % tblLayout type="fixed"
            lay = mat2doc.oxml.parse_xml(readFixture('tblLayout'));
            testCase.verifyEqual(class(lay), testCase.CT_TBLLAYOUT);
            testCase.verifyEqual(string(lay.type), "fixed", 'tblLayout type=fixed');

            % trPr -> trHeight val=Twips(500)=317500, hRule EXACTLY
            trPr = mat2doc.oxml.parse_xml(readFixture('trPr'));
            th = trPr.xpath("./w:trHeight");
            testCase.verifyEqual(numel(th), 1);
            testCase.verifyEqual(class(th(1)), testCase.CT_HEIGHT, 'trHeight parses as CT_Height');
            testCase.verifyEqual(double(th(1).val), 317500, 'trHeight val=500tw');
            testCase.verifyEqual(th(1).hRule, mat2doc.enum.table.WD_ROW_HEIGHT_RULE.EXACTLY, ...
                'trHeight hRule EXACTLY');

            % tcPr_00 -> tcW dxa width (Twips(4320)=2743200) + vAlign CENTER
            tc0 = mat2doc.oxml.parse_xml(readFixture('tcPr_00'));
            w0 = tc0.xpath("./w:tcW");  va0 = tc0.xpath("./w:vAlign");
            testCase.verifyEqual(class(w0(1)), testCase.CT_TBLWIDTH, 'tcW parses as CT_TblWidth');
            testCase.verifyEqual(double(w0(1).width), 2743200, 'cell(0,0) tcW dxa width=4320tw');
            testCase.verifyEqual(class(va0(1)), testCase.CT_VERTICALJC, 'vAlign parses as CT_VerticalJc');
            testCase.verifyEqual(va0(1).val, mat2doc.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER, ...
                'cell(0,0) vAlign CENTER');

            % tcPr_01 -> vMerge restart
            tc1 = mat2doc.oxml.parse_xml(readFixture('tcPr_01'));
            vm1 = tc1.xpath("./w:vMerge");
            testCase.verifyEqual(class(vm1(1)), testCase.CT_VMERGE, 'vMerge parses as CT_VMerge');
            testCase.verifyEqual(string(vm1(1).val), "restart", 'cell(0,1) vMerge restart');

            % tcPr_11 -> vMerge (bare) reads the "continue" default
            tc11 = mat2doc.oxml.parse_xml(readFixture('tcPr_11'));
            vm11 = tc11.xpath("./w:vMerge");
            testCase.verifyEqual(string(vm11(1).val), "continue", 'cell(1,1) bare vMerge -> "continue"');
        end

        function test_s0060_manifest_matches_pinned(testCase)
            % Regression (fixture-drift guard): the shipped data\s0060\manifest.json
            % SHAs must equal the hard-coded pin table in this class. A silent
            % re-freeze of a fixture would flip the manifest but not this constant.
            here = fileparts(mfilename('fullpath'));
            p = fullfile(here, 'data', 's0060', 'manifest.json');
            fid = fopen(p, 'r', 'n');
            testCase.assertGreaterThanOrEqual(fid, 0, 'cannot open s0060 manifest');
            raw = fread(fid, Inf, '*uint8')';
            fclose(fid);
            man = jsondecode(native2unicode(raw, 'UTF-8'));
            tbl = fixtureTable();
            testCase.verifyEqual(numel(man.fixtures), size(tbl, 1), '10 manifest fixtures');
            for i = 1:numel(man.fixtures)
                fx = man.fixtures(i);
                row = tbl(strcmp(tbl(:, 1), fx.name), :);
                testCase.verifyEqual(size(row, 1), 1, sprintf('manifest fixture %s pinned', fx.name));
                testCase.verifyEqual(fx.size, row{1, 2}, sprintf('%s size matches pin', fx.name));
                testCase.verifyEqual(string(fx.sha256), row{1, 3}, sprintf('%s SHA matches pin', fx.name));
            end
        end

        % =============================================================== %
        % 10. (M) M1 document.xml byte-pin (registry-neutrality guard)      %
        % =============================================================== %

        function test_m1_document_xml_byte_identical(testCase)
            % (M) Regression (byte-neutrality, L1): mat2doc.Document().save() emits
            % word/document.xml at EXACTLY 1548 B with the frozen s0001 SHA-256 --
            % byte-identical because default.docx has no table, so NONE of the 7
            % new table tags transits the new CT classes on the M1 save path.
            % SHA-256 equality is an L1 assertion.
            bytes = emitDocPart('document.xml');
            testCase.verifyEqual(numel(bytes), testCase.DOC_SIZE_M1, ...
                sprintf('word/document.xml must be exactly %d B after the table registry rows', ...
                    testCase.DOC_SIZE_M1));
            testCase.verifyEqual(sha256hex(bytes), testCase.DOC_SHA_M1, ...
                'word/document.xml SHA-256 must equal the frozen s0001 oracle (byte-identical L1)');
        end

        % =============================================================== %
        % 11. EQUIVALENCE -- full s0059 battery vs the frozen oracle         %
        % =============================================================== %

        function test_equivalence_full_battery_vs_frozen_oracle(testCase)
            % Equivalence: replay the ENTIRE s0059 battery (runProbes -- the .m
            % twin's body VERBATIM: tblwidth/height/gridcol/tblgrid/tbllayout/
            % verticaljc/vmerge) and flatten-compare EVERY leaf to the frozen
            % python-docx 1.2.0 oracle copied into data\s0059_probe_oracle.json.
            % Gate-3 found ZERO divergences (probe_diff exit 0), so every leaf must
            % be byte/value-identical. Ties the suite to the Gate-3 output.
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

function n = childAttrCount(e)
    % Number of attributes on the element (via attrib_names, mirrors the sectPr
    % template's attrib_names() usage). Used to prove @w:val removal.
    n = numel(e.attrib_names());
end

function s = rv(x)
    % Uniform accessor repr mirroring the s0059 rv(): None->"None",
    % enum->member NAME, bool->"True"/"False", Length/int->EMU decimal.
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

function setVmergeEmpty(e)
    % Assigning "" from within a function so verifyError can capture the throw
    % (a bare `e.val = ""` at the call site would be a statement, not a handle).
    e.val = "";
end

function o = loadOracle()
    % Read the co-located frozen s0059 oracle in BINARY mode (no CRLF translation)
    % and decode UTF-8 -> struct. jsondecode is whitespace/line-ending agnostic, so
    % no `* binary` .gitattributes pin is needed (value-based fixture; s0030/s0036/
    % s0039 precedent).
    here = fileparts(mfilename('fullpath'));
    p = fullfile(here, 'data', 's0059_probe_oracle.json');
    fid = fopen(p, 'r', 'n');
    if fid < 0
        error('mat2doc:test:openFail', 'cannot open frozen oracle %s', p);
    end
    raw = fread(fid, Inf, '*uint8')';
    fclose(fid);
    o = jsondecode(native2unicode(raw, 'UTF-8'));
end

% ---- s0060 frozen byte fixtures ----

function tbl = fixtureTable()
    % Hard-coded pin table {name, size, sha256} for the 10 frozen s0060 fixtures
    % (validate report section 3 / manifest.json). SHAs are of the python-docx
    % serialize_part_xml bytes == the MATLAB parse->serialize output (round-trip).
    tbl = { ...
        'grid',       1271, "bc491330ac0ddb8484d8934c13a825a524f5c578ceda3a44b1f9ec05536d9d81"; ...
        'tblLayout',  1231, "a437756a1e8b8e78781b57c2e0d463f4c845e116e404bb7f4e6a3702dfb0a09f"; ...
        'trPr',       1260, "8e0b1fc36ba356b5f489b716734a6c6965a2dba58f40612ecdea2bc57c640022"; ...
        'tcPr_00',    1277, "c8a9709e5a8a2c8f65c42f7afa7d88a3ff657dfec2fef8f8987d48026816f539"; ...
        'tcPr_01',    1278, "ef33d893482736808a954ba98cca3ee6ebd53b8029445658de6f3f6286462a33"; ...
        'tcPr_11',    1262, "3b2c20907cd959997f447cfeb334585584fb64b06a102e754ba3224240f9ad88"; ...
        'tcW_dxa',     159, "99d255fe4cade8f8d8916871e78923f11d1806746afd17e03724e8f8177c6700"; ...
        'tcW_pct',     159, "288496dd137b936c0c73bae38d4cd4652779608d105d72aa7661884fe7bd8ee1"; ...
        'tcW_auto',    157, "59cff78bdee07e3204e8a42d9a10c07d9d3902eb5edb062c8bf1de043259e1f4"; ...
        'tcW_setter',  159, "9e429ebbe06a2cbc53a1d71c07bd04016e8d04cc9c3b3aa43a866481d3f6bcb9" };
end

function b = readFixture(name)
    here = fileparts(mfilename('fullpath'));
    p = fullfile(here, 'data', 's0060', 'parts', [name '.xml']);
    b = readBytes(p);
end

% ---- M1 document.xml emit ----

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

% ---- s0059 equivalence replay (the .m twin body, VERBATIM) ----

function P = runProbes()
    % Replay the s0059 probe sequence (the .m twin's P-struct body, VERBATIM
    % tags/inputs/order) and return the nested struct of tagged canonical values.
    % Embedded here so the Equivalence leg is self-contained (the validation-folder
    % scenario is NOT on the toolbox path and must not be a dependency). Mirrors
    % validation\mat2doc\scenarios\s0059_p6_1_table_leaves_probe.m lines 20-122.
    HR = @(n) mat2doc.enum.table.WD_ROW_HEIGHT_RULE.(n);
    TW = @(v) mat2doc.shared.Twips(v);

    P = struct();

    % ---- tblwidth ----
    tw = struct();
    specs = {"dxa","dxa","2880"; "pct","pct","5000"; "auto","auto","0"; "nil","nil","0"};
    for k = 1:size(specs,1)
        key = specs{k,1};
        e = parse("<w:tcW " + nsW() + " w:type=""" + specs{k,2} + """ w:w=""" + specs{k,3} + """/>");
        rec = struct("serhex", hx_e(e), "width", rv(e.width), "type", rv(e.type));
        if key == "dxa" || key == "pct"
            rec.w = rv(e.w);
        end
        tw.(key) = rec;
    end
    e = nw("w:tcW"); e.width = TW(1440);
    tw.setter_forces_dxa = struct("serhex", hx_e(e), "type", rv(e.type), ...
        "w", rv(e.w), "width", rv(e.width));
    bare = nw("w:tcW");
    tw.err_w_absent = errmsg(@() bare.w);
    tw.err_type_absent = errmsg(@() bare.type);
    tw.err_width_absent = errmsg(@() bare.width);   % width reads type first
    P.tblwidth = tw;

    % ---- height ----
    h = struct();
    b = nw("w:trHeight");
    h.bare = struct("val", rv(b.val), "hRule", rv(b.hRule));
    e = nw("w:trHeight"); e.val = TW(500); e.hRule = HR("EXACTLY");
    h.set_both = struct("serhex", hx_e(e), "val", rv(e.val), "hRule", rv(e.hRule));
    e = nw("w:trHeight"); e.val = TW(360);
    h.set_val_only = struct("serhex", hx_e(e), "val", rv(e.val));
    e = nw("w:trHeight"); e.hRule = HR("AT_LEAST");
    h.set_hrule_only = struct("serhex", hx_e(e), "hRule", rv(e.hRule));
    e = nw("w:trHeight"); e.val = TW(500); e.val = [];
    h.val_none_removes = struct("serhex", hx_e(e), "val", rv(e.val));
    e = nw("w:trHeight"); e.hRule = HR("EXACTLY"); e.hRule = [];
    h.hrule_none_removes = struct("serhex", hx_e(e), "hRule", rv(e.hRule));
    P.height = h;

    % ---- gridcol ----
    gc = struct();
    b = nw("w:gridCol");
    gc.bare = struct("w", rv(b.w));
    e = nw("w:gridCol"); e.w = TW(2880);
    gc.set_w = struct("serhex", hx_e(e), "w", rv(e.w));
    e = nw("w:gridCol"); e.w = TW(2880); e.w = [];
    gc.w_none_removes = struct("serhex", hx_e(e), "w", rv(e.w));
    g = parse("<w:tblGrid " + nsW() + "><w:gridCol w:w=""2880""/><w:gridCol w:w=""1440""/></w:tblGrid>");
    cols = g.gridCol_lst;
    gc.idx = struct("idx0", rv(cols(1).gridCol_idx), "w0", rv(cols(1).w), ...
        "idx1", rv(cols(2).gridCol_idx), "w1", rv(cols(2).w));
    P.gridcol = gc;

    % ---- tblgrid ----
    tg = struct();
    g = nw("w:tblGrid");
    c1 = g.add_gridCol(); c1.w = TW(2880);
    c2 = g.add_gridCol(); c2.w = TW(1440); %#ok<NASGU>
    tg.two_appends = struct("localnames", {lnsCell(g)}, "count", rv(numel(g.gridCol_lst)), ...
        "serhex", hx_e(g));
    g = parse("<w:tblGrid " + nsW() + "><w:gridCol w:w=""100""/><w:tblGridChange w:id=""1""/></w:tblGrid>");
    ncol = g.add_gridCol(); ncol.w = TW(200);
    tg.h11_before_change = struct("localnames", {lnsCell(g)}, "new_idx", rv(ncol.gridCol_idx), ...
        "serhex", hx_e(g));
    P.tblgrid = tg;

    % ---- tbllayout ----
    tl = struct();
    b = nw("w:tblLayout");
    tl.bare = struct("type", rv(b.type));
    e = nw("w:tblLayout"); e.type = "fixed";
    tl.set_fixed = struct("serhex", hx_e(e), "type", rv(e.type));
    e = nw("w:tblLayout"); e.type = "autofit";
    tl.set_autofit = struct("serhex", hx_e(e), "type", rv(e.type));
    e = nw("w:tblLayout"); e.type = "fixed"; e.type = [];
    tl.type_none_removes = struct("serhex", hx_e(e), "type", rv(e.type));
    P.tbllayout = tl;

    % ---- verticaljc ----
    vj = struct();
    members = struct();
    marr = enumeration('mat2doc.enum.table.WD_CELL_VERTICAL_ALIGNMENT');
    for k = 1:numel(marr)
        e = nw("w:vAlign"); e.val = marr(k);
        members.(char(marr(k))) = struct("serhex", hx_e(e), "val", rv(e.val));
    end
    vj.members = members;
    bare = nw("w:vAlign");
    vj.err_val_absent = errmsg(@() bare.val);
    P.verticaljc = vj;

    % ---- vmerge ----
    vm = struct();
    b = parse("<w:vMerge " + nsW() + "/>");
    vm.bare_continue = struct("serhex", hx_e(b), "val", rv(b.val));
    e = nw("w:vMerge"); e.val = "restart";
    vm.set_restart = struct("serhex", hx_e(e), "val", rv(e.val));
    e = nw("w:vMerge"); e.val = "restart"; e.val = "continue";
    vm.set_continue_removes = struct("serhex", hx_e(e), "val", rv(e.val));
    e = nw("w:vMerge"); e.val = "restart"; e.val = [];
    vm.set_none_removes = struct("serhex", hx_e(e), "val", rv(e.val));
    vm.err_empty_string = errmsg(@() set_vmerge_empty_local());
    P.vmerge = vm;
end

function e = nw(tag)
    e = mat2doc.oxml.OxmlElement(tag);
end

function set_vmerge_empty_local()
    e = mat2doc.oxml.OxmlElement("w:vMerge");
    e.val = "";
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

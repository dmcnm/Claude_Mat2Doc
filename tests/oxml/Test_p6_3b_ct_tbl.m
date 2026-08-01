classdef Test_p6_3b_ct_tbl < matlab.unittest.TestCase
% TEST_P6_3B_CT_TBL  Gate-4 permanent unit tests for Mat2Doc P6-3b [N] -- the
%   oxml CT_Tbl class + the w:tbl un-defer sweep. src/docx/oxml/table.py::CT_Tbl
%   (146-256) -> NEW +mat2doc\+oxml\+table\CT_Tbl.m; +2 registry rows (w:tbl ->
%   CT_Tbl; w:bidiVisual -> CT_OnOff); and the V2 shim swap in CT_Tc (the P6-3a
%   generic-ancestor xpath shim trLstOfTbl_ is REPLACED by the faithful
%   self._tbl.tr_lst, now that w:tbl resolves to a real CT_Tbl). Completes the
%   table oxml layer + registry.
%
%   This is THE table-authoring foundation: CT_Tbl.new_tbl(rows, cols, width) is
%   the constructor the API tier (BlockItemContainer.add_table, P6-4a) calls to
%   emit the initial <w:tbl> XML. Its byte-exactness is load-bearing for every
%   table the toolbox will ever create -- hence the 7-size byte matrix below is
%   the headline permanent pin.
%
%   This class permanently freezes what the prior gates established:
%     * Gate-1 Porter  : audit_P6-3b_ct_tbl.md (self-probe; new_tbl 7 sizes).
%     * Gate-2 Auditor (Fable): APPROVE -- new_tbl 7 sizes independently byte-checked;
%       ZERO new D-numbers.
%     * Gate-3 Validator: validate_P6-3b_ct_tbl.md -- PASS, ZERO new D-numbers. FROZE
%       the table-authoring byte oracle references\s0065\ (7 new_*.xml + manifest)
%       and the CT_Tbl-surface probe references\s0066\ (probe.json). Re-proved the
%       P6-3a merge byte-matrix (references\s0063\) 10/10 byte-identical THROUGH the
%       V2 shim->CT_Tbl swap. M1 17/17 (styles.xml 349458 B / 02d71a68...;
%       document.xml 1548 B / 0e4dd503...). Targeted regression 83 total, 81 pass,
%       2 EXPECTED registry-flip stale-pins (re-pinned at Gate-4: Test_p2_3:291,
%       Test_p5_2b:201/211).
%
%   THE HIGHEST-VALUE PERMANENT PINS IN THIS WP (a regression goes RED loudly):
%     * ★ (NEW_TBL) THE 7-SIZE new_tbl BYTE MATRIX (test_new_tbl_byte_matrix) -- THE
%       TABLE-AUTHORING FOUNDATION. For each frozen s0065 size, CT_Tbl.new_tbl(rows,
%       cols, width) is CONSTRUCTED through the port, serialized, and the whole
%       <w:tbl> is asserted BYTE-IDENTICAL (raw uint8 AND SHA-256) to the frozen
%       python-docx oracle. Covers the headline 2x3/Inches(6) (835 B / 10211e87...),
%       the non-even 2x7/Inches(6.5) EMU-floor col-rounding (5943600//7=849085,
%       1471 B / b42b0032...), 1x1, 3x2, 5x4, cols==0 -> Emu(0) branch (c47e76ed...),
%       and the negative-floor edge (floor(-1000/3)=-334, bae4f89b...). If a future
%       change perturbs a single byte of the table constructor, this goes RED.
%     * ★ (MERGE-AFTER-SWAP) test_merge_after_swap_pin -- guards the V2 shim swap.
%       Re-runs the nested (m10), horizontal (m01) and vertical-bare-vMerge (m02)
%       cases from the frozen s0063 merge oracle THROUGH the swapped CT_Tc (whose
%       self._tbl.tr_lst now resolves a REAL CT_Tbl instead of the removed
%       trLstOfTbl_ xpath shim) -> byte-identical. m10 (nested) proves the nearest-
%       ancestor CT_Tbl resolves the INNER table's rows exactly as the shim did.
%       A future regression that breaks tbl resolution goes RED here.
%     * ★ (bidiVisual TRI-STATE) test_bidiVisual_val_tristate -- the D-delta-1 shape:
%       RTL (WD_TABLE_DIRECTION, bool(1)=True == CT_OnOff True default) -> DELETES
%       @w:val -> bare <w:bidiVisual/>; LTR (bool(0)=False) -> <w:bidiVisual
%       w:val="0"/>; None -> element removed. The tri-state byte shape is pinned.
%     * ★ (M1) test_m1_styles_and_document -- mat2doc.Document().save() ->
%       word/styles.xml == 349458 B / 02d71a68... AND word/document.xml == 1548 B /
%       0e4dd503... default.docx has NO <w:tbl> root, so registering w:tbl/bidiVisual
%       is byte-neutral (registering a CT changes only a parsed node's CLASS). L1.
%     * (VERIFY-1) fresh new_tbl has tblStyle_val == None and emits NO <w:tblStyle>
%       (test_tblStyle_val) -- "TableGrid" is API-tier (P6-4a), NOT baked into new_tbl.
%     * (ROW-MAJOR) test_iter_tcs_row_major -- fresh 2x3 -> 6 CT_Tc, grid_offset
%       flatten [0 1 2 0 1 2] (H9 materialized order == the Python generator's).
%
%   Provenance (all Gate-3 frozen 2026-07-31):
%     * Audit    : validation\mat2doc\audit_P6-3b_ct_tbl.md
%     * Validate : validation\mat2doc\validate_P6-3b_ct_tbl.md
%     * Scenarios: validation\mat2doc\scenarios\s0065_p6_3b_new_tbl_bytes.{py,m}
%                  (the 7-size new_tbl byte twin), s0066_p6_3b_ct_tbl_probe.{py,m}
%                  (the CT_Tbl-surface probe; runProbe() below replays its body
%                  VERBATIM).
%     * Frozen refs (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE at Gate-3):
%         references\s0065\ (7 new_*.xml + manifest.json) copied verbatim into
%           tests\oxml\data\s0065\ WITH a co-located `.gitattributes` `* binary`
%           pin (frozen-byte fixtures must not be line-ending mangled on the master
%           checkout -- the Gate-4 byte-fixture lesson).
%         references\s0066\ (probe.json + manifest.json) copied into
%           tests\oxml\data\s0066\ WITH the same `* binary` pin.
%         references\s0063\ (the P6-3a merge oracle) already present at
%           tests\oxml\data\s0063\ (copied by the P6-3a class) -- REUSED here for
%           the merge-after-swap guard.
%         references\s0001\parts\word\{styles,document}.xml -- the M1 byte
%           references (SHA of what Document().save() emits); NOT copied.
%
%   Coverage taxonomy
%   -----------------
%   * Nominal      -- registry rows; new_tbl at 7 sizes; tblPr/tblGrid OneAndOnlyOne
%                     getters; tr_lst/add_tr; col_count; iter_tcs; bidiVisual/tblStyle
%                     get+set.
%   * Edge         -- cols==0 -> Emu(0) branch (empty tblGrid); negative-width floor;
%                     the OneAndOnlyOne raise on a missing required child
%                     (mat2doc:InvalidXmlError, verbatim); bidiVisual None-removal;
%                     tblStyle None-removal + fresh-none (VERIFY-1); the m02 bare
%                     <w:vMerge/> continuation through the swapped CT_Tc.
%   * Equivalence  -- test_equivalence_probe_vs_frozen_oracle replays the ENTIRE
%                     s0066 probe (runProbe, the .m twin's body verbatim) and
%                     isequaln-compares each section to the frozen probe.json oracle
%                     (Gate-3 probe_diff was exit 0).
%   * Regression   -- hard-coded expected serialized-XML strings (ASCII ==
%                     byte-identical L1) + the 7 new_tbl SHA-256 pins + the 3
%                     merge-after-swap SHA pins + the M1 styles/document part SHAs.
%   * Upstream     -- the s0065 new_*.xml ARE python-docx CT_Tbl.new_tbl output (full
%                     Word nsmap); the s0063 merged fixtures ARE lxml's expected
%                     serialization of the merged trees.
%
%   Byte-level (L1) note: every new_tbl comparison and every merge-after-swap
%   comparison is BOTH a raw uint8 byte-equality AND a SHA-256 equality vs the frozen
%   oracle -- the ladder demanded L1 and Gate-3 delivered 7/7 (new_tbl) + 10/10
%   (merge) byte-identical with ZERO new D-numbers, so every pin here is L1. The
%   only looser-than-byte assertion is the Equivalence section isequaln (values, not
%   bytes) and its non-triviality floor -- commented at its site. The
%   mat2doc:InvalidXmlError identifier segment equals the Python class name and the
%   message text is verbatim (the signed exception-model mapping, design.md
%   section 2 -- non-byte, non-output, NOT a D-number).
%
%   Determinism: no network, no absolute paths. The worktree root and the co-located
%   s0063/s0065/s0066 fixtures resolve relative to this file via
%   fileparts(mfilename('fullpath')); the M1 save goes to a tempname .docx deleted
%   via onCleanup; every fixture read is binary ('r','n'). The +mat2doc package
%   resolves via the MANDATORY PathFixture(worktree-root) in TestClassSetup
%   (WP9-F4 lesson).

    properties (Constant)
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        % --- registered class names (P6-3b rows + prior table classes) ---
        CT_TBL     = 'mat2doc.oxml.table.CT_Tbl'
        CT_TC      = 'mat2doc.oxml.table.CT_Tc'
        CT_ROW     = 'mat2doc.oxml.table.CT_Row'
        CT_TBLPR   = 'mat2doc.oxml.table.CT_TblPr'
        CT_TBLGRID = 'mat2doc.oxml.table.CT_TblGrid'
        CT_ONOFF   = 'mat2doc.oxml.shared.CT_OnOff'

        % --- frozen s0001 M1 byte references (registering w:tbl/bidiVisual neutral) ---
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
            % Idiom copied verbatim from tests\oxml\Test_p6_3a_ct_tc.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\oxml
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. registry rows (the 2 P6-3b flips)                             %
        % =============================================================== %

        function test_registry_resolves_two_rows(testCase)
            % Nominal / Regression: REGISTRATION is what flips the parse class of a
            % <w:tbl> subtree. w:tbl -> CT_Tbl is the table root; w:bidiVisual ->
            % CT_OnOff backs CT_Tbl.bidiVisual_val (the RTL/LTR/None direction).
            pairs = { ...
                "w:tbl",        testCase.CT_TBL; ...
                "w:bidiVisual", testCase.CT_ONOFF };
            for i = 1:size(pairs, 1)
                tag = pairs{i, 1}; cls = pairs{i, 2};
                r = mat2doc.oxml.registry(mat2doc.oxml.qn(tag));
                testCase.verifyEqual(char(r), cls, ...
                    sprintf('registry must resolve %s -> %s (P6-3b row)', tag, cls));
                testCase.verifyEqual(class(mat2doc.oxml.OxmlElement(tag)), cls, ...
                    sprintf('OxmlElement(%s) must be a %s', tag, cls));
            end
        end

        % =============================================================== %
        % 2. ★ THE new_tbl BYTE MATRIX -- the table-authoring foundation     %
        % =============================================================== %

        function test_new_tbl_byte_matrix(testCase)
            % ★ (NEW_TBL) Regression + Upstream (byte-identical L1) -- THE headline
            % permanent pin and THE FOUNDATION for every table the toolbox creates.
            % For each of the 7 frozen s0065 sizes: CONSTRUCT CT_Tbl.new_tbl(rows,
            % cols, width) through the port, serialize the <w:tbl>, and assert the
            % result is BYTE-IDENTICAL (raw uint8 AND SHA-256) to the frozen
            % python-docx oracle part. Also cross-check the shipped fixture file is
            % intact (size+SHA) and verify col_count + the CT_Tbl class. Gate-3 froze
            % 7/7 byte-identical with ZERO new D-numbers.
            here = fileparts(mfilename('fullpath'));
            tbl = newTblFixtureTable();
            for i = 1:size(tbl, 1)
                name    = tbl{i, 1};
                rows    = tbl{i, 2};
                cols    = tbl{i, 3};
                widthFn = tbl{i, 4};
                colCnt  = tbl{i, 5};
                sz      = tbl{i, 6};
                sha     = tbl{i, 7};

                % -- the shipped frozen fixture is intact --
                frozen = readBytes(fullfile(here, 'data', 's0065', 'parts', [name '.xml']));
                testCase.verifyEqual(numel(frozen), sz, ...
                    sprintf('%s frozen fixture size intact (%d B)', name, sz));
                testCase.verifyEqual(sha256hex(frozen), sha, ...
                    sprintf('%s frozen fixture SHA-256 intact', name));

                % -- CONSTRUCT through the port and byte-compare to the oracle --
                elm = mat2doc.oxml.table.CT_Tbl.new_tbl(rows, cols, widthFn());
                testCase.verifyClass(elm, 'mat2doc.oxml.table.CT_Tbl', ...
                    sprintf('%s new_tbl returns a CT_Tbl', name));
                testCase.verifyEqual(elm.col_count, colCnt, ...
                    sprintf('%s col_count == %d', name, colCnt));

                out = mat2doc.oxml.serialize_part_xml(elm);
                testCase.verifyEqual(numel(out), sz, ...
                    sprintf('%s new_tbl serialized size == frozen oracle (%d B)', name, sz));
                testCase.verifyEqual(uint8(out(:)'), uint8(frozen(:)'), ...
                    sprintf('%s CT_Tbl.new_tbl -> serialize must be BYTE-IDENTICAL to the frozen s0065 oracle', name));
                testCase.verifyEqual(sha256hex(out), sha, ...
                    sprintf('%s new_tbl SHA-256 == frozen manifest (L1)', name));
            end
        end

        function test_new_tbl_manifest_matches_pinned(testCase)
            % Regression (fixture-drift guard): the shipped data\s0065\manifest.json
            % sizes/SHAs must equal the hard-coded pin table in this class. A silent
            % re-freeze would flip the manifest but not the constant.
            here = fileparts(mfilename('fullpath'));
            raw = readBytes(fullfile(here, 'data', 's0065', 'manifest.json'));
            man = jsondecode(native2unicode(raw, 'UTF-8'));
            tbl = newTblFixtureTable();
            testCase.verifyEqual(numel(man.fixtures), size(tbl, 1), '7 manifest fixtures');
            for i = 1:numel(man.fixtures)
                fx = man.fixtures(i);
                row = tbl(strcmp(tbl(:, 1), fx.name), :);
                testCase.verifyEqual(size(row, 1), 1, sprintf('manifest fixture %s pinned', fx.name));
                testCase.verifyEqual(fx.rows, row{1, 2}, sprintf('%s rows', fx.name));
                testCase.verifyEqual(fx.cols, row{1, 3}, sprintf('%s cols', fx.name));
                testCase.verifyEqual(fx.col_count, row{1, 5}, sprintf('%s col_count', fx.name));
                testCase.verifyEqual(fx.size, row{1, 6}, sprintf('%s size', fx.name));
                testCase.verifyEqual(string(fx.sha256), row{1, 7}, sprintf('%s sha', fx.name));
            end
        end

        function test_new_tbl_cols_zero_edge(testCase)
            % Edge (cols==0 -> Emu(0) branch): a 1x0 table has an EMPTY <w:tblGrid>
            % (no <w:gridCol>) but still one <w:tr> with no cells. col_count 0; the
            % byte oracle (c47e76ed...) already pins the exact shape in the matrix --
            % this asserts the structural read tier over the degenerate table.
            elm = mat2doc.oxml.table.CT_Tbl.new_tbl(1, 0, mat2doc.shared.Inches(1));
            testCase.verifyEqual(elm.col_count, 0, 'cols==0 -> col_count 0 (empty tblGrid)');
            testCase.verifyEqual(numel(elm.tblGrid.gridCol_lst), 0, 'no <w:gridCol> children');
            testCase.verifyEqual(numel(elm.tr_lst), 1, 'still one <w:tr>');
            testCase.verifyEqual(numel(elm.iter_tcs), 0, 'no cells (cols==0)');
        end

        % =============================================================== %
        % 3. CT_Tbl descriptors -- OneAndOnlyOne + ZeroOrMore + col_count    %
        % =============================================================== %

        function test_ct_tbl_descriptors(testCase)
            % Nominal + Edge: tblPr/tblGrid OneAndOnlyOne getters return the required
            % children; tr_lst is the ZeroOrMore CT_Row list; add_tr appends a CT_Row;
            % col_count == numel(tblGrid.gridCol_lst). Plus the OneAndOnlyOne RAISE on
            % a missing required child (mat2doc:InvalidXmlError, verbatim message).
            tbl = mat2doc.oxml.table.CT_Tbl.new_tbl(2, 3, mat2doc.shared.Inches(6));
            testCase.verifyEqual(class(tbl.tblPr), testCase.CT_TBLPR, 'tblPr is CT_TblPr (OneAndOnlyOne)');
            testCase.verifyEqual(class(tbl.tblGrid), testCase.CT_TBLGRID, 'tblGrid is CT_TblGrid (OneAndOnlyOne)');
            testCase.verifyEqual(tbl.col_count, 3, 'col_count 3 (len tblGrid.gridCol_lst)');

            trs = tbl.tr_lst;
            testCase.verifyEqual(numel(trs), 2, 'fresh 2x3 has 2 rows');
            testCase.verifyEqual(class(trs(1)), testCase.CT_ROW, 'tr_lst entries are CT_Row');
            added = tbl.add_tr();
            testCase.verifyClass(added, 'mat2doc.oxml.table.CT_Row', 'add_tr returns a CT_Row');
            testCase.verifyEqual(numel(tbl.tr_lst), 3, 'add_tr appended (2 -> 3)');

            % ★ OneAndOnlyOne raise: a bare <w:tbl> has NO required tblPr/tblGrid.
            bare = mat2doc.oxml.OxmlElement("w:tbl");
            testCase.verifyClass(bare, 'mat2doc.oxml.table.CT_Tbl', 'OxmlElement(w:tbl) is CT_Tbl');
            fn = @() bare.tblPr;
            testCase.verifyError(fn, 'mat2doc:InvalidXmlError', ...
                'missing required w:tblPr raises mat2doc:InvalidXmlError (OneAndOnlyOne getter)');
            testCase.verifyEqual(errmsg(fn), "required ``<w:tblPr>`` child element not present", ...
                'OneAndOnlyOne raise message verbatim (RST double-backticks preserved)');
            testCase.verifyError(@() bare.tblGrid, 'mat2doc:InvalidXmlError', ...
                'missing required w:tblGrid raises mat2doc:InvalidXmlError');
        end

        % =============================================================== %
        % 4. iter_tcs -- row-major materialized order (H9)                  %
        % =============================================================== %

        function test_iter_tcs_row_major(testCase)
            % (ROW-MAJOR) Nominal: iter_tcs over a fresh 2x3 yields 6 CT_Tc cells,
            % left-to-right top-to-bottom. The Python generator is materialized (H9)
            % into a 1xN array; the grid_offset flatten must be [0 1 2 0 1 2] (each
            % row restarts at offset 0), proving row-major order.
            tbl = mat2doc.oxml.table.CT_Tbl.new_tbl(2, 3, mat2doc.shared.Inches(6));
            cells = tbl.iter_tcs();
            testCase.verifyEqual(numel(cells), 6, 'fresh 2x3 -> 6 cells');
            for k = 1:numel(cells)
                testCase.verifyClass(cells(k), 'mat2doc.oxml.table.CT_Tc', ...
                    sprintf('iter_tcs(%d) is a CT_Tc', k));
            end
            offsets = arrayfun(@(c) c.grid_offset, cells);
            testCase.verifyEqual(offsets, [0 1 2 0 1 2], ...
                'H9: iter_tcs grid_offset flatten is row-major [0 1 2 0 1 2]');
        end

        % =============================================================== %
        % 5. ★ bidiVisual_val tri-state (D-delta-1 byte shape)               %
        % =============================================================== %

        function test_bidiVisual_val_tristate(testCase)
            % ★ (bidiVisual TRI-STATE) Nominal + Edge (D-delta-1 / H10 / H4 / H3):
            % absent -> None; RTL (WD_TABLE_DIRECTION, bool(1)=True == CT_OnOff's True
            % default) -> @w:val DELETED -> BARE <w:bidiVisual/>; LTR (bool(0)=False)
            % -> <w:bidiVisual w:val="0"/>; None -> element removed. The tri-state
            % byte shape is the sharpest new-D guard in this WP.
            tbl = mat2doc.oxml.table.CT_Tbl.new_tbl(1, 1, mat2doc.shared.Inches(1));
            testCase.verifyTrue(isequal(tbl.bidiVisual_val, []), 'absent bidiVisual -> [] (None)');

            % RTL -> get True + BARE <w:bidiVisual/> (no @w:val)
            tbl.bidiVisual_val = mat2doc.enum.table.WD_TABLE_DIRECTION.RTL;
            testCase.verifyEqual(tbl.bidiVisual_val, true, 'RTL -> bidiVisual_val True');
            sRTL = ser(tbl.tblPr);
            testCase.verifyTrue(contains(sRTL, "<w:bidiVisual/>"), ...
                'RTL -> BARE <w:bidiVisual/> (D-delta-1 deletes @w:val at the True default)');
            testCase.verifyFalse(contains(sRTL, "w:bidiVisual w:val"), ...
                'RTL bidiVisual carries NO @w:val');

            % LTR -> get False + <w:bidiVisual w:val="0"/>
            tbl.bidiVisual_val = mat2doc.enum.table.WD_TABLE_DIRECTION.LTR;
            testCase.verifyEqual(tbl.bidiVisual_val, false, 'LTR -> bidiVisual_val False');
            testCase.verifyTrue(contains(ser(tbl.tblPr), '<w:bidiVisual w:val="0"/>'), ...
                'LTR -> <w:bidiVisual w:val="0"/>');

            % None -> get None + element removed
            tbl.bidiVisual_val = [];
            testCase.verifyTrue(isequal(tbl.bidiVisual_val, []), 'set None -> bidiVisual_val [] (None)');
            testCase.verifyFalse(contains(ser(tbl.tblPr), "<w:bidiVisual"), ...
                'set None removes the <w:bidiVisual> element');
        end

        % =============================================================== %
        % 6. tblStyle_val get/set/None + VERIFY-1 (fresh has no style)       %
        % =============================================================== %

        function test_tblStyle_val(testCase)
            % Nominal + Edge (VERIFY-1): a fresh new_tbl has tblStyle_val == None and
            % emits NO <w:tblStyle> ("TableGrid" is applied at the API tier, P6-4a,
            % NOT baked into new_tbl). set writes <w:tblStyle w:val="..."/> via the
            % PRIVATE _add adder; re-set exercises the remove-then-add path; None
            % removes.
            tbl = mat2doc.oxml.table.CT_Tbl.new_tbl(1, 1, mat2doc.shared.Inches(1));
            testCase.verifyTrue(isequal(tbl.tblStyle_val, []), ...
                'VERIFY-1: fresh new_tbl tblStyle_val == None');
            testCase.verifyFalse(contains(ser(tbl.tblPr), "<w:tblStyle"), ...
                'VERIFY-1: fresh new_tbl emits NO <w:tblStyle>');

            % set "TableGrid"
            tbl.tblStyle_val = "TableGrid";
            testCase.verifyEqual(string(tbl.tblStyle_val), "TableGrid", 'tblStyle_val read-back TableGrid');
            testCase.verifyTrue(contains(ser(tbl.tblPr), '<w:tblStyle w:val="TableGrid"/>'), ...
                'set -> <w:tblStyle w:val="TableGrid"/>');

            % re-set "LightShading" (remove-then-add path -> exactly one tblStyle)
            tbl.tblStyle_val = "LightShading";
            testCase.verifyEqual(string(tbl.tblStyle_val), "LightShading", 'tblStyle_val re-set LightShading');
            s = ser(tbl.tblPr);
            testCase.verifyTrue(contains(s, '<w:tblStyle w:val="LightShading"/>'), ...
                're-set -> <w:tblStyle w:val="LightShading"/>');
            testCase.verifyEqual(count(s, '<w:tblStyle'), 1, ...
                're-set is remove-then-add: exactly one <w:tblStyle> (not doubled)');

            % None removes
            tbl.tblStyle_val = [];
            testCase.verifyTrue(isequal(tbl.tblStyle_val, []), 'set None -> tblStyle_val [] (None)');
            testCase.verifyFalse(contains(ser(tbl.tblPr), "<w:tblStyle"), ...
                'set None removes the <w:tblStyle> element');
        end

        % =============================================================== %
        % 7. ★ MERGE-AFTER-SWAP -- the V2 shim->CT_Tbl guard                 %
        % =============================================================== %

        function test_merge_after_swap_pin(testCase)
            % ★ (MERGE-AFTER-SWAP) Regression (byte-identical L1) -- guards the V2
            % swap. The P6-3a CT_Tc.merge now resolves rows via self._tbl.tr_lst on a
            % REAL CT_Tbl (the removed trLstOfTbl_ xpath shim). Re-run the nested
            % (m10), horizontal (m01) and vertical-bare-vMerge (m02) cases from the
            % frozen s0063 merge oracle THROUGH the swapped CT_Tc and assert each is
            % BYTE-IDENTICAL (raw uint8 AND SHA-256) to the frozen MERGED oracle. If a
            % future change breaks tbl resolution, this goes RED. Gate-3 re-proved
            % 10/10; these 3 are the load-bearing subset (nested nearest-ancestor +
            % horizontal gridSpan + vertical bare-vMerge).
            here = fileparts(mfilename('fullpath'));
            cases = { ...
                'm01_horiz_1x2', "0,0",     "0,1",     2340, ...
                    "425401a74699931664bcf64d9bfba399b129f23d580dbefe9c464e386951bf1c"; ...
                'm02_vert_2x1',  "0,0",     "1,0",     2423, ...
                    "2a78bdb88109eaba7ea9be25e8a84f5a9338f8bce9facf4cee0d19bea903fd57"; ...
                'm10_nested',    "1,1;0,0", "1,1;0,1", 2503, ...
                    "58686bdc68e7d359b2d8c7b29ae95784dcff97c290d5ed30ebe390c5dcca0c38" };
            for i = 1:size(cases, 1)
                name = cases{i, 1};
                out  = doMerge(name, cases{i, 2}, cases{i, 3});
                frozen = readBytes(fullfile(here, 'data', 's0063', 'parts', ['merged_' name '.xml']));
                testCase.verifyEqual(numel(out), cases{i, 4}, ...
                    sprintf('%s merged size == frozen oracle (%d B)', name, cases{i, 4}));
                testCase.verifyEqual(uint8(out(:)'), uint8(frozen(:)'), ...
                    sprintf('%s merge THROUGH the swapped CT_Tc must be BYTE-IDENTICAL to the frozen s0063 oracle', name));
                testCase.verifyEqual(sha256hex(out), cases{i, 5}, ...
                    sprintf('%s merged SHA-256 == frozen manifest (L1, swap-neutral)', name));
            end

            % m02 detail: the continuation cell is a BARE <w:vMerge/> (no @w:val),
            % the top cell <w:vMerge w:val="restart"/>; "continue" is ABSENT
            % (D-delta-1) -- the swap must not perturb the vertical-merge byte shape.
            s02 = string(native2unicode(doMerge('m02_vert_2x1', "0,0", "1,0"), "UTF-8"));
            testCase.verifyTrue(contains(s02, '<w:vMerge w:val="restart"/>'), ...
                'm02 top cell <w:vMerge w:val="restart"/> (swap-neutral)');
            testCase.verifyTrue(contains(s02, '<w:vMerge/>'), ...
                'm02 continuation BARE <w:vMerge/> (swap-neutral)');
            testCase.verifyFalse(contains(s02, 'continue'), ...
                'm02 never contains "continue" (D-delta-1)');

            % m10 detail: exactly one outer + one inner <w:tbl>; inner merge gridSpan=2
            % (nearest-ancestor CT_Tbl resolution via self._tbl.tr_lst).
            s10 = string(native2unicode(doMerge('m10_nested', "1,1;0,0", "1,1;0,1"), "UTF-8"));
            testCase.verifyEqual(count(s10, '<w:tbl>') + count(s10, '<w:tbl '), 2, ...
                'm10 exactly one outer + one inner <w:tbl> (structure intact after swap)');
            testCase.verifyTrue(contains(s10, '<w:gridSpan w:val="2"/>'), ...
                'm10 inner merge gridSpan=2 (nearest-ancestor CT_Tbl.tr_lst)');
        end

        % =============================================================== %
        % 8. ★ M1 styles.xml + document.xml byte-pins (neutrality guard)     %
        % =============================================================== %

        function test_m1_styles_and_document(testCase)
            % ★ (M1) Regression (byte-neutrality, L1): mat2doc.Document().save()
            % emits word/styles.xml at EXACTLY 349458 B and word/document.xml at 1548
            % B (frozen s0001 SHAs). default.docx has NO <w:tbl> root, so registering
            % w:tbl -> CT_Tbl and w:bidiVisual -> CT_OnOff is byte-neutral (registering
            % a CT changes only a parsed node's CLASS, never content/order -- the
            % P4-6/P6-2/P6-3a precedent). A single save() emits both parts. SHA == L1.
            [styBytes, docBytes] = emitTwoParts('styles.xml', 'document.xml');

            testCase.verifyEqual(numel(styBytes), testCase.STYLES_SIZE_M1, ...
                sprintf('word/styles.xml must be exactly %d B', testCase.STYLES_SIZE_M1));
            testCase.verifyEqual(sha256hex(styBytes), testCase.STYLES_SHA_M1, ...
                'word/styles.xml SHA-256 == frozen s0001 oracle (w:tbl/bidiVisual registration neutral, L1)');

            testCase.verifyEqual(numel(docBytes), testCase.DOC_SIZE_M1, ...
                sprintf('word/document.xml must be exactly %d B', testCase.DOC_SIZE_M1));
            testCase.verifyEqual(sha256hex(docBytes), testCase.DOC_SHA_M1, ...
                'word/document.xml SHA-256 == frozen s0001 oracle (byte-identical L1)');
        end

        % =============================================================== %
        % 9. EQUIVALENCE -- the full s0066 probe vs the frozen oracle         %
        % =============================================================== %

        function test_equivalence_probe_vs_frozen_oracle(testCase)
            % Equivalence: replay the ENTIRE s0066 probe (runProbe -- the .m twin's
            % body VERBATIM: descriptors / add_tr / iter_tcs / bidiVisual / tblStyle)
            % and compare EACH section to the frozen python-docx 1.2.0 oracle copied
            % into data\s0066\probe.json. Gate-3 probe_diff was exit 0, so every
            % section must be value-identical.
            %
            % Both sides are normalized through jsondecode(jsonencode(...)) so that
            % struct vs cell-of-string container shapes collapse the SAME way (the .m
            % twin writes the port with jsonencode; the oracle came from the identical
            % JSON shape) -- the only looser-than-byte comparison in this class
            % (values, not bytes), justified because probe_diff already proved
            % value equivalence at Gate-3.
            here = fileparts(mfilename('fullpath'));
            port    = runProbe();
            portN   = jsondecode(jsonencode(port));               % normalize shapes
            oracleN = jsondecode(native2unicode( ...
                readBytes(fullfile(here, 'data', 's0066', 'probe.json')), 'UTF-8'));

            % Non-triviality floor (guards a silent-empty replay).
            testCase.verifyEqual(sort(fieldnames(portN)), sort(fieldnames(oracleN)), ...
                'the replayed probe and the frozen oracle expose the same top-level sections');
            testCase.verifyGreaterThanOrEqual(numel(fieldnames(oracleN)), 5, ...
                'the oracle must expose all 5 probe sections');

            % Section-by-section isequaln (localizes any regression to a section).
            secs = fieldnames(oracleN);
            for i = 1:numel(secs)
                s = secs{i};
                testCase.verifyTrue(isfield(portN, s), sprintf('port is missing section %s', s));
                testCase.verifyTrue(isequaln(portN.(s), oracleN.(s)), ...
                    sprintf('section "%s" must be value-identical to the frozen s0066 oracle', s));
            end
        end

    end
end

% ===================== file-local helpers ============================== %

function s = ser(e)
    % Decode serialize_part_xml bytes to a string (ASCII content -> string-equality
    % is a byte-identical L1 assertion; non-ASCII round-trips via UTF-8).
    s = string(native2unicode(mat2doc.oxml.serialize_part_xml(e), "UTF-8"));
end

function s = errmsg(fn)
    try
        fn();
        s = "NOERR";
    catch e
        s = string(e.message);
    end
end

function b = readBytes(p)
    f = fopen(p, 'r', 'n');            % binary read (no CRLF translation)
    assert(f >= 0, 'could not open for read: %s', p);
    b = fread(f, Inf, '*uint8')';
    fclose(f);
end

function h = sha256hex(bytes)
    % Lowercase hex SHA-256 of a uint8 vector (base MATLAB / Java MessageDigest;
    % matches the python hashlib manifest SHAs).
    md  = java.security.MessageDigest.getInstance('SHA-256');
    dig = md.digest(typecast(uint8(bytes(:)'), 'int8'));
    b   = typecast(dig, 'uint8');
    h   = lower(string(sprintf('%02x', b)));
end

% ---- the 7-size new_tbl pin table (s0065 manifest, VERBATIM literals) ----

function tbl = newTblFixtureTable()
    % Hard-coded pin table for the 7 frozen s0065 new_tbl cases (manifest.json):
    % {name, rows, cols, widthFn, col_count, size, sha256}. The width constructors
    % are the EXACT literals from the s0065 twin (Inches(6) == 5486400 EMU, etc.);
    % new_tbl reads only double(width), so the frozen bytes are reproduced exactly.
    tbl = { ...
        'new_1x1_in3',    1, 1, @() mat2doc.shared.Inches(3),       1, 436, ...
            "174f948738f4d3022a5d651d65feda5e6b84d07533950889fa89e201bac957a7"; ...
        'new_2x3_in6',    2, 3, @() mat2doc.shared.Inches(6),       3, 835, ...
            "10211e8749896d41feffb3ce6c0739ea42c1e304e7fd7a0d72a0e8023a918258"; ...
        'new_3x2_emu1e6', 3, 2, @() mat2doc.shared.Emu(1000000),    2, 817, ...
            "757af432b791f78279a186b92d574e1d950a65235e755258588d235e66811720"; ...
        'new_5x4_in6p5',  5, 4, @() mat2doc.shared.Inches(6.5),     4, 1849, ...
            "f579d563d866932cb822c528fa9875a2d6594daf8c461defb84af52e5bf04b3b"; ...
        'new_2x7_in6p5',  2, 7, @() mat2doc.shared.Inches(6.5),     7, 1471, ...
            "b42b003268431e87aadac4c3ca6c4fcdc9aa15651daaa4916d77c0869014f274"; ...
        'new_1x0_in1',    1, 0, @() mat2doc.shared.Inches(1),       0, 351, ...
            "c47e76edbcbc1097dbc2287083ed27929a396e336a76569522fc6252ea4d07bc"; ...
        'new_2x3_emuneg', 2, 3, @() mat2doc.shared.Emu(-1000),      3, 817, ...
            "bae4f89be4c40075878f0fe6808400668acb4c324d661ac0aca34e6fed38dda5" };
end

% ---- s0063 merge navigation (VERBATIM from the P6-3a class / s0063 twin) ----

function out = doMerge(name, mergeA, mergeB)
    % Read the frozen s0063 SOURCE, parse, navigate to the two corners, merge, and
    % return the serialized <w:tbl> bytes (uint8).
    here = fileparts(mfilename('fullpath'));
    src = readBytes(fullfile(here, 'data', 's0063', 'parts', ['src_' char(name) '.xml']));
    root = mat2doc.oxml.parse_xml(src);
    a = resolve_corner(root, mergeA);
    b = resolve_corner(root, mergeB);
    a.merge(b);
    out = mat2doc.oxml.serialize_part_xml(root);
end

function tc = resolve_corner(tbl, spec)
    % Walk the descent path "r,c" / "or,oc;ir,ic" to the corner CT_Tc (VERBATIM from
    % s0063_p6_3a_merge_matrix.m). rows(r+1).tc_at_grid_offset(c) at each level
    % (H1: +1 row index; c RAW 0-based grid data). Multi-step descends into the
    % nested w:tbl inside the outer cell.
    steps = strsplit(string(spec), ";");
    node = tbl;
    tc = [];
    for i = 1:numel(steps)
        rc = double(strsplit(steps(i), ","));   % [r c]
        rows = node.xpath("./w:tr");             % CT_Row handles
        tc = rows(rc(1) + 1).tc_at_grid_offset(rc(2));
        if i < numel(steps)
            node = tc.tbl_lst(1);                % descend into nested w:tbl
        end
    end
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

% ---- s0066 probe replay (the .m twin body, VERBATIM) ----

function probe = runProbe()
    % Replay the s0066 CT_Tbl-surface probe and return the nested struct of tagged
    % canonical values (all leaves -> strings via rvp). Embedded here so the
    % Equivalence leg is self-contained (the validation-folder scenario is NOT on
    % the toolbox path and must not be a dependency). Mirrors
    % validation\mat2doc\scenarios\s0066_p6_3b_ct_tbl_probe.m lines 12-74 VERBATIM.
    probe = struct();

    % ---- descriptors (OneAndOnlyOne) + col_count + tr_lst ----
    tbl = mat2doc.oxml.table.CT_Tbl.new_tbl(2, 3, mat2doc.shared.Inches(6));
    trs = tbl.tr_lst;
    d = struct();
    d.tblPr_class   = simple_name(tbl.tblPr);
    d.tblGrid_class = simple_name(tbl.tblGrid);
    d.col_count     = rvp(tbl.col_count);
    d.tr_count      = rvp(numel(trs));
    d.tr_classes    = uniq_sorted_names(trs);
    probe.descriptors = d;

    % ---- add_tr appends a CT_Row ----
    a = struct();
    a.count_before = rvp(numel(tbl.tr_lst));
    added = tbl.add_tr();
    a.added_class = simple_name(added);
    a.count_after = rvp(numel(tbl.tr_lst));
    probe.add_tr = a;

    % ---- iter_tcs: fresh 2x3, row-major ----
    t2 = mat2doc.oxml.table.CT_Tbl.new_tbl(2, 3, mat2doc.shared.Inches(6));
    cells = t2.iter_tcs();
    it = struct();
    it.count   = rvp(numel(cells));
    it.classes = uniq_sorted_names(cells);
    go = cell(1, numel(cells));
    for i = 1:numel(cells); go{i} = rvp(cells(i).grid_offset); end
    it.grid_offset_flatten = go;
    probe.iter_tcs = it;

    % ---- bidiVisual_val get/set/None ----
    tb = mat2doc.oxml.table.CT_Tbl.new_tbl(1, 1, mat2doc.shared.Inches(1));
    bv = struct();
    bv.absent = rvp(tb.bidiVisual_val);
    tb.bidiVisual_val = mat2doc.enum.table.WD_TABLE_DIRECTION.RTL;
    bv.rtl_get     = rvp(tb.bidiVisual_val);
    bv.rtl_bare    = rvp(contains(ser(tb.tblPr), "<w:bidiVisual/>"));
    bv.rtl_has_val = rvp(contains(ser(tb.tblPr), "w:bidiVisual w:val"));
    tb.bidiVisual_val = mat2doc.enum.table.WD_TABLE_DIRECTION.LTR;
    bv.ltr_get  = rvp(tb.bidiVisual_val);
    bv.ltr_val0 = rvp(contains(ser(tb.tblPr), '<w:bidiVisual w:val="0"/>'));
    tb.bidiVisual_val = [];
    bv.none_get     = rvp(tb.bidiVisual_val);
    bv.none_removed = rvp(~contains(ser(tb.tblPr), "<w:bidiVisual"));
    probe.bidiVisual = bv;

    % ---- tblStyle_val get/set/None ----
    ts_tbl = mat2doc.oxml.table.CT_Tbl.new_tbl(1, 1, mat2doc.shared.Inches(1));
    ts = struct();
    ts.default        = rvp(ts_tbl.tblStyle_val);
    ts.default_absent = rvp(~contains(ser(ts_tbl.tblPr), "<w:tblStyle"));
    ts_tbl.tblStyle_val = "TableGrid";
    ts.set_get   = rvp(ts_tbl.tblStyle_val);
    ts.set_bytes = rvp(contains(ser(ts_tbl.tblPr), '<w:tblStyle w:val="TableGrid"/>'));
    ts_tbl.tblStyle_val = "LightShading";   % remove-then-add path
    ts.reset_get   = rvp(ts_tbl.tblStyle_val);
    ts.reset_bytes = rvp(contains(ser(ts_tbl.tblPr), '<w:tblStyle w:val="LightShading"/>'));
    ts_tbl.tblStyle_val = [];
    ts.none_get     = rvp(ts_tbl.tblStyle_val);
    ts.none_removed = rvp(~contains(ser(ts_tbl.tblPr), "<w:tblStyle"));
    probe.tblStyle = ts;
end

function n = simple_name(x)
    parts = split(string(class(x)), ".");
    n = parts(end);
end

function c = uniq_sorted_names(arr)
    names = strings(1, numel(arr));
    for i = 1:numel(arr); names(i) = simple_name(arr(i)); end
    names = unique(names);        % unique sorts ascending (matches Python sorted(set))
    c = cell(1, numel(names));
    for i = 1:numel(names); c{i} = names(i); end
end

function s = rvp(x)
    % Uniform accessor repr mirroring the s0066 rv(): None->"None",
    % bool->"True"/"False", enum->NAME, Length/int->EMU/int decimal.
    if isequal(x, [])
        s = "None";
    elseif islogical(x)
        if x, s = "True"; else, s = "False"; end
    elseif isenum(x)
        s = string(x);
    elseif isa(x, 'mat2doc.shared.Length')
        s = string(sprintf('%.0f', double(x)));
    elseif isnumeric(x)
        s = string(sprintf('%.0f', double(x)));
    else
        s = string(x);
    end
end

classdef Test_p6_4b_cell_merge < matlab.unittest.TestCase
% TEST_P6_4B_CELL_MERGE  Gate-4 permanent unit tests for Mat2Doc P6-4b [N] -- the
%   FINAL Phase-6 WP: _Cell (+ the cell MERGE engine reached end-to-end),
%   Table.add_row / add_column / cell / row_cells / column_cells / _cells, and
%   _Row.cells / _Column.cells, closing the table tier.
%   src/docx/table.py::{_Cell, Table.add_row/add_column/cell/row_cells/
%   column_cells/_cells, _Row.cells, _Column.cells} + blkcntnr.py::
%   BlockItemContainer.tables -> +mat2doc\+table\{Cell_,Table,Row_,Column_}.m +
%   +mat2doc\BlockItemContainer.m.
%
%   THIS IS THE P6 HEADLINE -- an end-to-end MERGED-CELL .docx the toolbox can
%   author from the public API. Cell content, add_row/add_column population, nested
%   tables, and every cell-MERGE geometry (horizontal / vertical / 2x2 block /
%   merge-then-set-text / 3x3 mixed) are pinned as FULL-PACKAGE byte oracles --
%   word/document.xml raw uint8 byte-identical to python-docx 1.2.0 AND all 17 parts
%   size+SHA-256 == the frozen manifest. The 3x3 mixed-merge grid-walk span-identity
%   (the H5 anti-regression guard) is re-proven at the handle level. If a single byte
%   of the cell-authoring or cell-merge path drifts, this class goes RED.
%
%   THE HIGHEST-VALUE PERMANENT PINS IN THIS WP:
%     * ★ (CELL AUTHORING FULL-PACKAGE) test_cell_authoring_full_package_byte_pins --
%       for each of s0073 (cell content: multi-paragraph + Cell_.text setter +
%       non-ASCII café/CJK/emoji), s0074 (add_row + populate), s0075 (add_column +
%       populate), s0076 (nested table in a cell) the MATLAB Document is built through
%       the IDENTICAL public API, saved, unzipped, and compared to the frozen
%       python-docx package: word/document.xml raw byte-identical AND 17/17 parts
%       size+SHA. document.xml SHAs s0073 505d69b0.../s0074 3f7009e1.../s0075
%       52b20778.../s0076 e91d9512...
%     * ★★ (CELL MERGE FULL-PACKAGE -- the P6 headline) test_cell_merge_full_package_-
%       byte_pins -- t.cell(...).merge(t.cell(...)) reached through the full public
%       API produces the EXACT python-docx package for every merge geometry:
%       HORIZONTAL s0077 9f626e2b...60b44737 (the permanent horizontal guard),
%       VERTICAL s0078 e833fac8...f083b055, 2x2 BLOCK s0079 ac155531...21cf73ad (the
%       permanent block guard), merge-then-set-text s0080 c96a3351...e8898a8e. This
%       is _Cell.merge -> CT_Tc.merge (gridSpan/vMerge) end-to-end in a real package,
%       byte-identical across all four geometries. RED on any merge-serialization drift.
%     * ★ (3x3 MIXED-MERGE BYTE + GRID-WALK SPAN-IDENTITY) test_grid_span_bytes_pin +
%       test_grid_walk_span_identity -- s0082 (63b25b1a...c3f1bf2e6) is the full-
%       package byte companion; text written through NON-ROOT grid positions lands in
%       the span-root tc (byte-proven). test_grid_walk_span_identity re-proves H5: a
%       single Table.cells_() access yields the 0-based first-identity partition
%       [0,0,2,3,3,5,3,3,5] EXACTLY (the spanned cell is the SAME handle at every
%       covered grid position), the per-cell texts match, and Table.cell(r,c) at every
%       (r,c) resolves to the correct span-root cell.
%     * ★ (M1) test_m1_byte_pins -- P6-4b un-stub NEUTRALITY: a bare
%       mat2doc.Document().save() STILL emits word/styles.xml at EXACTLY 349458 B /
%       02d71a68... AND word/document.xml at 1548 B / 0e4dd503... (un-stubbing the 8
%       table members + BlockItemContainer.tables touches no default-save byte). L1.
%     * (CELL SURFACE) test_cell_* -- _Cell.text get(newline-join)/set(clear+one
%       paragraph); grid_span; vertical_alignment get/set/None (WD_CELL_VERTICAL_-
%       ALIGNMENT); width get/set (getter returns [] when unset; the SETTER is NOT
%       Optional -- assigning [] raises mat2doc:TypeError, faithful to python-docx);
%       paragraphs/tables; add_paragraph; nested add_table (non-default branch; the
%       Inches(1) DEFAULT-WIDTH branch is not publicly reachable -- see the NOTE at
%       test_cell_merge_then_set_text -- and is auditor-byte-proven as E2). merge
%       returns a Cell_ with grid_span 2.
%     * (COLLECTION SURFACE) test_add_row_returns_row / test_add_column_returns_column
%       / test_row_column_cells -- add_row -> Row_ (index_ 2, cells len 2, +1 w:tr);
%       add_column -> Column_ (index_ 2, cells len 2, +1 gridCol); _Row.cells /
%       _Column.cells (2x2 -> 4 _cells, 2 per row/col; merged -> span-shared).
%     * (BOUNDARY / F-1) test_out_of_range_matlab_behavior -- pins the CURRENT MATLAB
%       out-of-range behavior of the deprecated row_cells / cell (throws
%       MATLAB:badsubscript where Python's deprecated slice clamps to []). This is a
%       KNOWN, non-output-visible divergence in an invalid-usage path -- NOT a
%       D-number (no explicit raise in the Python source to mirror). column_cells OOR
%       matches Python exactly (empty).
%     * (EQUIVALENCE) test_equivalence_probe_vs_frozen_oracle -- replays the ENTIRE
%       s0081 cell/grid-walk probe and value-compares every section to the frozen
%       python-docx 1.2.0 oracle (Gate-3 probe_diff was exit 0).
%
%   Provenance (Gate-1..3, all 2026-08-01):
%     * Audit    : validation\mat2doc\audit_P6-4b_cell_merge.md (Porter Gate-1
%                  self-probe 34/34 + mso-auditor Gate-2 APPROVE -- 10/10 end-to-end
%                  byte scenarios all full-package L1 incl. the full H/V/block/merge-
%                  then-text matrix + 3x3 mixed-merge + the E2 default-width branch;
%                  ZERO new D-numbers).
%     * Validate : Gate-3 PASS -- M1 17/17; each of s0073-s0080 + s0082 frozen 17/17
%                  L1; probe_diff MATCH (exit 0) over the full _Cell / cells-grid-walk
%                  / _Row.cells / _Column.cells surface; the 3x3 span-identity
%                  [0,0,2,3,3,5,3,3,5] re-proven at handle + byte levels; ZERO new
%                  D-numbers. Only the 3 anticipated stub-flips went RED (re-pinned in
%                  Test_p6_4a_table_api + Test_p2_3_document_shell at Gate-4).
%     * Scenarios: validation\mat2doc\scenarios\s0073..s0080,s0082_*_gscenario.m (the
%                  byte twins -- build bodies replayed VERBATIM by the helpers below);
%                  s0081_p6_4b_cell_probe.m (the full cell/grid-walk probe -- runProbe()
%                  below replays its body VERBATIM for the Equivalence leg).
%     * Frozen refs (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE at Gate-3):
%         references\s0073..s0080,s0082\ (manifest.json + parts\word\document.xml)
%           copied into tests\table\data\s00XX\{document.xml,manifest.json} WITH a
%           co-located `.gitattributes` `* binary` pin (the parent data\.gitattributes
%           also pins `* binary` recursively -- belt and suspenders for the master
%           checkout, per the Gate-4 byte-fixture lesson). The manifest SHAs ARE the
%           python-docx package part SHAs.
%         references\s0081\probe.json copied into tests\table\data\
%           s0081_probe_oracle.json (value JSON; jsondecode is line-ending agnostic).
%         references\s0001\parts\word\{styles,document}.xml -- the M1 byte references
%           (SHA of what Document().save() emits); NOT copied (SHAs pinned as Constants).
%
%   Coverage taxonomy
%   -----------------
%   * Nominal      -- cell text get/set; grid_span; add_paragraph; nested add_table;
%                     merge -> Cell_ grid_span 2; add_row -> Row_; add_column ->
%                     Column_; _Row.cells / _Column.cells; Table.cell(r,c) addressing.
%   * Edge         -- non-ASCII (café / CJK 中 / emoji 🎉 astral pair, H2) cell text;
%                     vertical_alignment / width None ([]) round-trips (H3); the
%                     Inches(1) DEFAULT-WIDTH branch of nested add_table; empty cell
%                     default text ""; multi-paragraph newline-join (H16, no strip);
%                     the deprecated row_cells / cell out-of-range MATLAB behavior.
%   * Equivalence  -- test_equivalence_probe_vs_frozen_oracle replays the ENTIRE s0081
%                     probe (cell_props / merge_props / return_types / grid_walk) and
%                     section-compares every value to the frozen python-docx oracle
%                     (Gate-3 probe_diff was exit 0).
%   * Regression   -- the 4 cell-authoring + 4 cell-merge + 1 mixed-merge full-package
%                     (document.xml raw uint8 + 17-part SHA) pins; the M1 styles/
%                     document part SHAs; grid_span / index_ / cells-len values.
%   * Upstream     -- the frozen document.xml parts ARE python-docx _Cell / merge /
%                     add_row / add_column output; the grid-walk partition
%                     [0,0,2,3,3,5,3,3,5] and the "\n".join (no strip) text getter ARE
%                     the python-docx table.py contract.
%
%   Byte-level (L1) note: every full-package document.xml comparison is BOTH a raw
%   uint8 byte-equality AND a SHA-256 equality vs the frozen oracle, and every one of
%   the 17 package parts is SHA-256-pinned to the frozen manifest -- the ladder
%   demanded L1 and Gate-3 delivered 17/17 with ZERO new D-numbers, so every byte pin
%   here is L1. The Equivalence section's jsondecode(jsonencode(...)) value comparison
%   (and its non-triviality floor) is the only looser-than-byte check and is commented
%   at its site (justified: probe_diff already proved value equivalence at Gate-3). The
%   mat2doc:notYetPorted mapping is gone for the 8 flipped members (they are LIVE); the
%   MATLAB:badsubscript out-of-range identifier is pinned as the CURRENT MATLAB
%   behavior of a deprecated invalid-usage path (F-1, a known non-output-visible
%   divergence -- NOT a D-number).
%
%   Determinism: no network, no absolute paths -- the worktree root and the co-located
%   s0073..s0082 fixtures resolve relative to this file via
%   fileparts(mfilename('fullpath')); saves go to tempname .docx / tempname dirs
%   deleted via onCleanup; every fixture read is binary ('r','n'). The +mat2doc package
%   resolves via the MANDATORY PathFixture(worktree-root) in TestClassSetup (WP9-F4
%   lesson).

    properties (Constant)
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        % --- proxy classes under test ---
        CELL_      = 'mat2doc.table.Cell_'
        TABLE      = 'mat2doc.table.Table'
        ROW_       = 'mat2doc.table.Row_'
        COLUMN_    = 'mat2doc.table.Column_'
        PARAGRAPH  = 'mat2doc.text.Paragraph'

        % --- frozen s0001 M1 byte references (P6-4b un-stub neutrality guard) ---
        STYLES_SIZE_M1 = 349458
        STYLES_SHA_M1  = "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"
        DOC_SIZE_M1    = 1548
        DOC_SHA_M1     = "0e4dd503bc095219cd24e8312b372c588ca56bb834aacac53eafb7bf47836327"

        % --- headline document.xml SHAs (loudest single guards) ---
        S0077_HORIZONTAL_SHA = "9f626e2be47cd25db5ddc681ab34286e585ccb0a42019844f7e02f3460b44737"
        S0079_BLOCK_SHA      = "ac1555310ae599cd8396882f692b2e8d985bcd44349f5eb91c06435421cf73ad"
        S0082_MIXED_SHA      = "63b25b1a5a043cd4201109eab15a28111dfa31eff0b43b0e9412686c3f1bf2e6"

        % --- the frozen 3x3 mixed-merge grid-walk oracle (references\s0081) ---
        GRID_PARTITION = [0 0 2 3 3 5 3 3 5]   % 0-based first-identity indices (H5)
        GRID_TEXTS     = ["TL" "TL" "TR" "BLK" "BLK" "V" "BLK" "BLK" "V"]

        % --- regression cell-width values (frozen s0081 probe, EMU) ---
        WIDTH_DEFAULT_2x2 = 2743200   % add_table(2,2) default cell width (_block_width/2)
        WIDTH_IN2         = 1828800   % Inches(2)
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run cannot
            % resolve the +mat2doc package (MATLAB:undefinedVarOrClass). Idiom copied
            % from tests\table\Test_p6_4a_table_api.m. here is tests\table; the
            % worktree root is two levels up.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\table
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. ★ cell-authoring FULL-PACKAGE byte pins (s0073..s0076)        %
        % =============================================================== %

        function test_cell_authoring_full_package_byte_pins(testCase)
            % ★ (CELL AUTHORING FULL-PACKAGE) Regression + Upstream (byte-identical L1).
            % For each of the 4 frozen cell-content scenarios: build a mat2doc.Document
            % through the IDENTICAL public API, save(), unzip, and assert (a)
            % word/document.xml is raw-uint8 byte-identical to the shipped frozen part,
            % AND (b) all 17 package parts are size+SHA-256 == the frozen python-docx
            % manifest. Covers the Cell_.text setter (single-run replace), non-ASCII
            % (café / 中 / 🎉 astral pair, H2), multi-paragraph add_paragraph, add_row +
            % Row_.cells populate, add_column + Column_.cells populate, and a nested
            % table (Cell_.add_table). Gate-3 froze 17/17 byte-identical per scenario.
            scenarios = { ...
                's0073', @buildS0073; ...   % cell content (text setter + non-ASCII + multi-para)
                's0074', @buildS0074; ...   % add_row + populate (trailing space preserved)
                's0075', @buildS0075; ...   % add_column + populate
                's0076', @buildS0076 };     % nested table in a cell
            testCase.assertFullPackageByteIdentical(scenarios);
        end

        % =============================================================== %
        % 2. ★★ CELL MERGE FULL-PACKAGE byte pins (s0077..s0080) --        %
        %    THE P6 HEADLINE: _Cell.merge -> CT_Tc.merge end-to-end        %
        % =============================================================== %

        function test_cell_merge_full_package_byte_pins(testCase)
            % ★★ (CELL MERGE FULL-PACKAGE -- THE P6 HEADLINE) Regression + Upstream
            % (byte-identical L1). This is the loudest guard in Phase 6: for every
            % merge geometry, t.cell(a,b).merge(t.cell(c,d)) -- reached through the FULL
            % public authoring API and serialized into a real package -- produces
            % word/document.xml BYTE-IDENTICAL to python-docx 1.2.0, AND all 17 parts
            % match the frozen manifest. This exercises _Cell.merge (self._tc.merge(
            % other._tc)) driving the CT_Tc.merge gridSpan/vMerge engine (byte-proven
            % standalone at P6-3a) end-to-end in a package:
            %   HORIZONTAL  s0077 cell(0,0).merge(cell(0,1))  -> gridSpan=2   (guard)
            %   VERTICAL    s0078 cell(0,0).merge(cell(1,0))  -> vMerge col0
            %   2x2 BLOCK   s0079 cell(0,0).merge(cell(1,1))  -> gridSpan+vMerge (guard)
            %   MERGE-THEN- s0080 m=cell(0,0).merge(cell(0,1)); m.text="merged"
            %     SET-TEXT        (grid_span 2, content written into the merged tc)
            % If any byte of the merge serialization drifts, this goes RED across four
            % independent geometries at once.
            scenarios = { ...
                's0077', @buildS0077; ...   % HORIZONTAL merge (permanent guard)
                's0078', @buildS0078; ...   % VERTICAL merge
                's0079', @buildS0079; ...   % 2x2 BLOCK merge (permanent guard)
                's0080', @buildS0080 };     % merge then set text
            testCase.assertFullPackageByteIdentical(scenarios);

            % headline hard-pins on the two designated permanent merge guards
            hp = saveAndUnzipParts(buildS0077()); hp = hp('word/document.xml');
            testCase.verifyEqual(sha256hex(hp), testCase.S0077_HORIZONTAL_SHA, ...
                'headline: HORIZONTAL merge document.xml SHA-256 == frozen s0077 oracle 9f626e2b... (L1)');
            bp = saveAndUnzipParts(buildS0079()); bp = bp('word/document.xml');
            testCase.verifyEqual(sha256hex(bp), testCase.S0079_BLOCK_SHA, ...
                'headline: 2x2 BLOCK merge document.xml SHA-256 == frozen s0079 oracle ac155531... (L1)');
        end

        % =============================================================== %
        % 3. ★ 3x3 mixed-merge FULL-PACKAGE byte pin (s0082)               %
        % =============================================================== %

        function test_grid_span_bytes_pin(testCase)
            % ★ (3x3 MIXED-MERGE FULL-PACKAGE, byte-identical L1) The byte companion of
            % the grid-walk span-identity: a 3x3 with row0-horizontal + rows1-2/cols0-1
            % block + col2/rows1-2 vertical merges, with text written through NON-ROOT
            % grid positions (0,1)/(0,2)/(2,1)/(2,2). Byte-identity proves that text
            % written through a continuation/non-root position LANDS in the span-root
            % tc. word/document.xml == frozen s0082 (63b25b1a...) + 17/17 parts.
            testCase.assertFullPackageByteIdentical({'s0082', @buildS0082});
            mp = saveAndUnzipParts(buildS0082()); mp = mp('word/document.xml');
            testCase.verifyEqual(sha256hex(mp), testCase.S0082_MIXED_SHA, ...
                'headline: 3x3 mixed-merge document.xml SHA-256 == frozen s0082 oracle 63b25b1a... (L1)');
        end

        % =============================================================== %
        % 4. ★ M1 byte-pins (P6-4b un-stub neutrality guard)              %
        % =============================================================== %

        function test_m1_byte_pins(testCase)
            % ★ (M1) Regression (byte-neutrality, L1): un-stubbing the 8 table members
            % + BlockItemContainer.tables touches NO default-save byte -- a bare
            % mat2doc.Document().save() STILL emits word/styles.xml at EXACTLY 349458 B
            % / 02d71a68... AND word/document.xml at 1548 B / 0e4dd503... (frozen s0001
            % SHAs). The un-stubs add no registry rows and are not on the open->save
            % path. SHA == L1.
            parts = saveAndUnzipParts(mat2doc.Document());

            sty = parts('word/styles.xml');
            testCase.verifyEqual(numel(sty), testCase.STYLES_SIZE_M1, ...
                sprintf('bare save word/styles.xml must be exactly %d B', testCase.STYLES_SIZE_M1));
            testCase.verifyEqual(sha256hex(sty), testCase.STYLES_SHA_M1, ...
                'M1 word/styles.xml SHA-256 unchanged (P6-4b un-stub neutrality, L1)');

            doc = parts('word/document.xml');
            testCase.verifyEqual(numel(doc), testCase.DOC_SIZE_M1, ...
                sprintf('bare save word/document.xml must be exactly %d B', testCase.DOC_SIZE_M1));
            testCase.verifyEqual(sha256hex(doc), testCase.DOC_SHA_M1, ...
                'M1 word/document.xml SHA-256 unchanged (P6-4b un-stub neutrality, L1)');
        end

        % =============================================================== %
        % 5. ★ cells-grid-walk span-identity (H5) -- s0082 handle partition %
        % =============================================================== %

        function test_grid_walk_span_identity(testCase)
            % ★ (H5 SPAN-IDENTITY) The anti-regression guard for the grid walk. Build
            % the s0082 3x3 mixed-merge table, take ONE Table.cells_() access, and
            % assert the 0-based first-identity partition is EXACTLY [0,0,2,3,3,5,3,3,5]
            % -- i.e. the spanned cell is the SAME Cell_ handle (handle ==) at every
            % grid position it covers (positions 0/1 share, 3/4/6/7 share, 5/8 share).
            % Then: the per-position texts match the frozen oracle, and Table.cell(r,c)
            % at every (r,c) resolves to the correct span-root cell (its .text matches
            % the _cells text at that position). Text written through a continuation
            % position landing in the span-root tc is additionally BYTE-proven by
            % test_grid_span_bytes_pin (s0082). H5 identity holds WITHIN one access;
            % across two accesses fresh wrappers are minted (VERIFY-2 -- faithful to
            % python-docx's recomputed _cells @property), pinned below.
            t = buildS0082Table();

            % --- ONE cells_() access: handle-identity partition ---
            cells = t.cells_();
            testCase.verifyEqual(numel(cells), 9, '3x3 -> 9 grid cells');
            testCase.verifyEqual(partition0(cells), testCase.GRID_PARTITION, ...
                'H5: the _cells handle partition is [0,0,2,3,3,5,3,3,5] (spanned cell = SAME handle at each covered position)');
            % explicit same-handle assertions on the three spans (loud)
            testCase.verifyTrue(cells(1) == cells(2), ...
                'H5: the horizontal span (0,0)/(0,1) is the SAME Cell_ handle');
            testCase.verifyTrue(cells(4) == cells(5) && cells(4) == cells(7) && cells(4) == cells(8), ...
                'H5: the 2x2 block (1,0)/(1,1)/(2,0)/(2,1) is ONE Cell_ handle at all four positions');
            testCase.verifyTrue(cells(6) == cells(9), ...
                'H5: the vertical span (1,2)/(2,2) is the SAME Cell_ handle');
            testCase.verifyFalse(cells(3) == cells(1), ...
                'the unmerged (0,2) cell is a DISTINCT handle');

            % --- per-position texts (read through the single _cells access) ---
            texts = arrayfun(@(x) string(x.text), cells);
            testCase.verifyEqual(texts, testCase.GRID_TEXTS, ...
                '_cells texts == frozen oracle (continuation positions read the span-root content)');

            % --- Table.cell(r,c) resolves to the correct span-root cell at every (r,c) ---
            k = 0; cellTexts = strings(1, 9);
            for r = 0:2
                for c = 0:2
                    k = k + 1;
                    cellTexts(k) = string(t.cell(r, c).text);
                end
            end
            testCase.verifyEqual(cellTexts, testCase.GRID_TEXTS, ...
                'Table.cell(r,c) matches _cells at EVERY grid position (span-root resolution)');

            % --- VERIFY-2: fresh wrappers across two accesses (== Python `is` False) ---
            a = t.cells_(); b = t.cells_();
            testCase.verifyFalse(a(1) == b(1), ...
                'VERIFY-2: separate cells_() accesses mint FRESH Cell_ wrappers (cross-access ~=, faithful; NOT a D-number)');
        end

        % =============================================================== %
        % 6. _Cell surface -- text get/set (+ non-ASCII, multi-para H16)   %
        % =============================================================== %

        function test_cell_text_get_set(testCase)
            % Nominal + Edge (table.py 264-284): the text getter is "\n".join over the
            % cell's paragraph texts (H16 -- NO strip); the setter clears all content
            % and writes a single <w:p>/<w:r>. A fresh 2x2 cell reads "" (one empty
            % paragraph). Non-ASCII (café / 中 / 🎉 astral pair, H2) round-trips through
            % the setter. Multi-paragraph join preserves the newline separator.
            t = mat2doc.Document().add_table(2, 2);
            c = t.cell(0, 0);
            testCase.verifyEqual(string(c.text), "", 'fresh cell text is "" (one empty paragraph, H16 join)');

            c.text = "A";
            testCase.verifyEqual(string(c.text), "A", 'text setter replaces content with "A"');

            nonascii = string(native2unicode( ...
                uint8([99 97 102 195 169 32 228 184 173 32 240 159 142 137]), 'UTF-8')); % "café 中 🎉"
            c.text = nonascii;
            testCase.verifyEqual(string(c.text), nonascii, ...
                'text setter round-trips non-ASCII (café / CJK 中 / emoji 🎉 astral pair, H2)');

            % multi-paragraph newline join (H16, no strip): a fresh cell, add p2/p3
            c2 = t.cell(1, 0);
            c2.text = "hi";
            c2.add_paragraph("p2");
            c2.add_paragraph("p3");
            testCase.verifyEqual(string(c2.text), "hi" + newline + "p2" + newline + "p3", ...
                'text getter is a bare newline-join over paragraph texts (H16, no strip)');
        end

        function test_cell_grid_span(testCase)
            % Nominal + Edge (table.py 228-235): grid_span is 1 for an unmerged cell
            % and grows to the number of grid columns spanned after a horizontal merge.
            t = mat2doc.Document().add_table(2, 2);
            testCase.verifyEqual(t.cell(0, 0).grid_span, 1, 'unmerged cell grid_span == 1');
            m = t.cell(0, 0).merge(t.cell(0, 1));
            testCase.verifyClass(m, testCase.CELL_, 'merge returns a Cell_');
            testCase.verifyEqual(m.grid_span, 2, 'horizontal merge -> merged cell grid_span == 2');
        end

        function test_cell_vertical_alignment(testCase)
            % Nominal + Edge (table.py 286-302, H3/H10): default [] (None); set CENTER
            % (a WD_CELL_VERTICAL_ALIGNMENT member) reads back by NAME "CENTER"; set []
            % removes the explicit alignment (round-trips to []).
            WCV = @(n) mat2doc.enum.table.WD_CELL_VERTICAL_ALIGNMENT.(n);
            c = mat2doc.Document().add_table(2, 2).cell(0, 0);
            testCase.verifyTrue(isequal(c.vertical_alignment, []), 'default vertical_alignment -> [] (None)');

            c.vertical_alignment = WCV("CENTER");
            testCase.verifyEqual(string(c.vertical_alignment), "CENTER", ...
                'set CENTER -> vertical_alignment reads back by name "CENTER"');

            c.vertical_alignment = [];
            testCase.verifyTrue(isequal(c.vertical_alignment, []), ...
                'vertical_alignment = [] removes the explicit alignment (H3)');
        end

        function test_cell_width(testCase)
            % Nominal + Edge (table.py 304-311, H6): a fresh 2x2 cell carries an
            % explicit width (2743200 EMU per the frozen probe); set Inches(2) ->
            % 1828800 readback. The GETTER returns [] (None) when no explicit width is
            % set, but the SETTER is typed `Length` -- NOT Optional (python-docx
            % table.py 309-311 `def width(self, value: Length)`; CT_Tc.width setter is
            % likewise `value: Length`). Assigning [] is therefore an INVALID cell-width
            % operation on BOTH sides: python-docx would raise (Emu(None) -> TypeError)
            % and this port raises the mapped mat2doc:TypeError -- faithful, NOT a
            % D-number. (Column_.width, by contrast, IS Optional -- see Test_p6_4a.)
            IN = @(v) mat2doc.shared.Inches(v);
            c = mat2doc.Document().add_table(2, 2).cell(0, 0);
            testCase.verifyEqual(double(c.width), testCase.WIDTH_DEFAULT_2x2, ...
                'fresh 2x2 cell width == 2743200 EMU (frozen probe)');

            c.width = IN(2);
            testCase.verifyEqual(double(c.width), testCase.WIDTH_IN2, ...
                'set width Inches(2) -> 1828800 EMU readback');

            % _Cell.width setter is NOT Optional: assigning [] raises (faithful to
            % python-docx's non-Optional Length setter -> Emu(None) TypeError).
            testCase.verifyError(@() setCellWidth(c, []), 'mat2doc:TypeError', ...
                'the _Cell.width setter is NOT Optional -> assigning [] raises mat2doc:TypeError (faithful to python-docx, NOT a D-number)');
        end

        function test_cell_paragraphs_and_add_paragraph(testCase)
            % Nominal (table.py 200-254): a fresh cell has one paragraph and no tables;
            % add_paragraph returns a Paragraph and grows the paragraph list.
            c = mat2doc.Document().add_table(2, 2).cell(0, 0);
            testCase.verifyEqual(numel(c.paragraphs), 1, 'fresh cell has one paragraph');
            testCase.verifyEqual(numel(c.tables), 0, 'fresh cell has no tables');

            p = c.add_paragraph("added");
            testCase.verifyClass(p, testCase.PARAGRAPH, 'add_paragraph returns a Paragraph');
            testCase.verifyEqual(string(p.text), "added", 'the added paragraph carries its text');
            testCase.verifyEqual(numel(c.paragraphs), 2, 'add_paragraph grew the paragraph list to 2');
        end

        function test_cell_add_table_nested(testCase)
            % Nominal (table.py 213-226): Cell_.add_table nests a Table inside the
            % cell, returns the nested Table, and grows cell.tables to 1. The cell has
            % an explicit width, so the nested table inherits it (NON-default branch).
            c = mat2doc.Document().add_table(2, 2).cell(0, 0);
            inner = c.add_table(2, 2);
            testCase.verifyClass(inner, testCase.TABLE, 'nested add_table returns a Table');
            testCase.verifyEqual(numel(c.tables), 1, 'cell.tables grows to 1 after nesting');
            testCase.verifyEqual(inner.rows.len_(), 2, 'nested table has 2 rows');
            testCase.verifyEqual(inner.columns.len_(), 2, 'nested table has 2 columns');
        end

        % NOTE (table.py 223, the Inches(1) DEFAULT-width branch of Cell_.add_table):
        % this branch is taken only when self.width is None, but EVERY cell produced by
        % the public add_table authoring path carries an explicit tcW width (a fresh
        % 2x2 cell reads 2743200 EMU; a merged cell reads the summed width), and the
        % _Cell.width setter is not Optional (test_cell_width) so the width cannot be
        % cleared through the public API. The default branch is therefore reachable
        % only by internal tcPr stripping -- exactly what the Gate-2 auditor did, and
        % it is byte-proven there as scenario E2 (document.xml f1cfe2ce...7708787c,
        % 17/17 L1). No reachable public-API test exists; the non-default branch is
        % covered by test_cell_add_table_nested.

        function test_cell_merge_then_set_text(testCase)
            % Nominal (table.py 237-284): merge returns a Cell_ over the merged tc;
            % setting .text on the merged cell writes content into the span-root tc.
            % (The full package byte form is s0080, pinned in section 2.)
            t = mat2doc.Document().add_table(2, 2);
            m = t.cell(0, 0).merge(t.cell(0, 1));
            m.text = "merged";
            testCase.verifyEqual(m.grid_span, 2, 'merged cell grid_span == 2');
            testCase.verifyEqual(string(m.text), "merged", 'merged cell text set/read == "merged"');
        end

        % =============================================================== %
        % 7. add_row / add_column return types + collection cells          %
        % =============================================================== %

        function test_add_row_returns_row(testCase)
            % Nominal (table.py 47-55): add_row appends a <w:tr> and returns a Row_ with
            % 0-based index_ 2 (the 3rd row of a 2-row table) and cells len == the
            % column count (2). rows.len_ grows 2 -> 3 (the +1 w:tr).
            t = mat2doc.Document().add_table(2, 2);
            testCase.verifyEqual(t.rows.len_(), 2, 'table starts with 2 rows');
            row = t.add_row();
            testCase.verifyClass(row, testCase.ROW_, 'add_row returns a Row_');
            testCase.verifyEqual(row.index_(), 2, 'the new row is at 0-based index_ 2');
            testCase.verifyEqual(numel(row.cells), 2, 'the new row has 2 cells (column count)');
            testCase.verifyEqual(t.rows.len_(), 3, 'rows.len_ grew to 3 (the +1 <w:tr>)');
        end

        function test_add_column_returns_column(testCase)
            % Nominal (table.py 37-45): add_column appends a <w:gridCol> (+ a <w:tc>
            % per row) and returns a Column_ with 0-based index_ 2 and cells len == the
            % row count (2). columns.len_ grows 2 -> 3 (the +1 gridCol).
            IN = @(v) mat2doc.shared.Inches(v);
            t = mat2doc.Document().add_table(2, 2);
            testCase.verifyEqual(t.columns.len_(), 2, 'table starts with 2 columns');
            col = t.add_column(IN(1));
            testCase.verifyClass(col, testCase.COLUMN_, 'add_column returns a Column_');
            testCase.verifyEqual(col.index_(), 2, 'the new column is at 0-based index_ 2');
            testCase.verifyEqual(numel(col.cells), 2, 'the new column has 2 cells (row count)');
            testCase.verifyEqual(t.columns.len_(), 3, 'columns.len_ grew to 3 (the +1 <w:gridCol>)');
        end

        function test_row_column_cells_and_underscore_cells(testCase)
            % Nominal (table.py 163-180, 322-325, 395-438): a plain 2x2 table has 4
            % _cells; each Row_.cells / _Column.cells / row_cells / column_cells returns
            % 2 Cell_. A plain row's two cells are DISTINCT handles (H5 -- no span).
            t = mat2doc.Document().add_table(2, 2);
            testCase.verifyEqual(numel(t.cells_()), 4, 'plain 2x2 _cells count == 4');

            r0 = t.rows.getitem_(0);
            rc = r0.cells;
            testCase.verifyEqual(numel(rc), 2, '_Row.cells count == 2');
            testCase.verifyClass(rc(1), testCase.CELL_, '_Row.cells entries are Cell_');
            testCase.verifyFalse(rc(1) == rc(2), 'a plain row''s two cells are DISTINCT handles (no span)');
            testCase.verifyEqual(numel(t.row_cells(0)), 2, 'Table.row_cells(0) count == 2');

            c0 = t.columns.getitem_(0);
            cc = c0.cells;
            testCase.verifyEqual(numel(cc), 2, '_Column.cells count == 2');
            testCase.verifyClass(cc(1), testCase.CELL_, '_Column.cells entries are Cell_');
            testCase.verifyEqual(numel(t.column_cells(0)), 2, 'Table.column_cells(0) count == 2');
        end

        function test_table_cell_addressing(testCase)
            % Nominal / Regression (table.py 85-91, H1): Table.cell(row,col) addresses
            % the (0-based) grid position; text written through cell(r,c) persists in
            % the underlying tc and reads back at the same (r,c). (The byte form is
            % s0073.) Fresh wrappers per access, but the tc content is stable.
            t = mat2doc.Document().add_table(2, 2);
            t.cell(0, 0).text = "A";
            t.cell(0, 1).text = "B";
            t.cell(1, 0).text = "C";
            t.cell(1, 1).text = "D";
            testCase.verifyEqual(string(t.cell(0, 0).text), "A", 'cell(0,0) addresses top-left');
            testCase.verifyEqual(string(t.cell(0, 1).text), "B", 'cell(0,1)');
            testCase.verifyEqual(string(t.cell(1, 0).text), "C", 'cell(1,0)');
            testCase.verifyEqual(string(t.cell(1, 1).text), "D", 'cell(1,1) addresses bottom-right');
        end

        % =============================================================== %
        % 8. F-1 boundary -- deprecated OOR row_cells/cell MATLAB behavior  %
        % =============================================================== %

        function test_out_of_range_matlab_behavior(testCase)
            % Edge / BOUNDARY (Gate-2 F-1 -- NOT a D-number). The deprecated Python
            % row_cells uses a slice cells[start:end] that CLAMPS an out-of-range index
            % to [] (empty); the port's explicit index arithmetic instead throws
            % MATLAB:badsubscript. Likewise cell(0,5) is an implicit IndexError in
            % Python and MATLAB:badsubscript here. There is NO explicit `raise` in the
            % Python source to mirror, so no mat2doc: identifier is owed and this is a
            % known, non-output-visible divergence in an INVALID-usage path (no D-number
            % -- see validate_P6-4b §5 / audit §F-1). This test PINS THE CURRENT MATLAB
            % behavior so the divergence is documented and any change is caught. NOTE:
            % Python would return [] for row_cells(5) / raise IndexError for cell(0,5).
            t = mat2doc.Document().add_table(2, 2);   % 2 rows x 2 cols -> 4 _cells

            testCase.verifyError(@() t.row_cells(5), 'MATLAB:badsubscript', ...
                'row_cells(5) throws MATLAB:badsubscript (Python clamps to []; known divergence, NOT a D-number)');
            testCase.verifyError(@() t.cell(0, 5), 'MATLAB:badsubscript', ...
                'cell(0,5) throws MATLAB:badsubscript (Python implicit IndexError; known divergence, NOT a D-number)');

            % column_cells OOR MATCHES Python exactly: the strided range is empty, so an
            % out-of-range column index returns an EMPTY Cell_ array (no error).
            oor = t.column_cells(5);
            testCase.verifyEqual(numel(oor), 0, ...
                'column_cells(5) returns an EMPTY Cell_ array (matches Python exactly)');
        end

        % =============================================================== %
        % 9. EQUIVALENCE -- full s0081 cell/grid-walk probe vs the oracle   %
        % =============================================================== %

        function test_equivalence_probe_vs_frozen_oracle(testCase)
            % Equivalence: replay the ENTIRE s0081 probe (runProbe -- the .m twin's body
            % VERBATIM: cell_props / merge_props / return_types / grid_walk) and compare
            % EACH section to the frozen python-docx 1.2.0 oracle copied into
            % data\s0081_probe_oracle.json. Gate-3 probe_diff was exit 0, so every
            % section must be value-identical.
            %
            % Both sides are normalized through jsondecode(jsonencode(...)) so struct vs
            % cell-of-string container shapes and row/column vector orientations
            % collapse the SAME way -- the only looser-than-byte comparison in this
            % class (values, not bytes), justified because probe_diff already proved
            % value equivalence at Gate-3.
            here = fileparts(mfilename('fullpath'));
            port    = runProbe();
            portN   = jsondecode(jsonencode(port));               % normalize shapes
            oracleN = jsondecode(native2unicode( ...
                readBytes(fullfile(here, 'data', 's0081_probe_oracle.json')), 'UTF-8'));

            % Non-triviality floor (guards a silent-empty replay).
            testCase.verifyEqual(sort(fieldnames(portN)), sort(fieldnames(oracleN)), ...
                'the replayed probe and the frozen oracle expose the same top-level sections');
            testCase.verifyGreaterThanOrEqual(numel(fieldnames(oracleN)), 4, ...
                'the oracle must expose all 4 probe sections (cell_props/merge_props/return_types/grid_walk)');

            secs = fieldnames(oracleN);
            for i = 1:numel(secs)
                s = secs{i};
                testCase.verifyTrue(isfield(portN, s), sprintf('port is missing section %s', s));
                testCase.verifyTrue(isequaln(portN.(s), oracleN.(s)), ...
                    sprintf('section "%s" must be value-identical to the frozen s0081 oracle', s));
            end
        end

    end

    % ===================== instance helpers ============================ %
    methods (Access = private)
        function assertFullPackageByteIdentical(testCase, scenarios)
            % For each {name, buildFn} row: build through the public API, save, unzip,
            % and assert (a) word/document.xml raw uint8 == the frozen part AND (b) all
            % 17 parts size+SHA-256 == the frozen manifest. Shared by the authoring,
            % merge, and mixed-merge byte-pin tests. L1.
            here = fileparts(mfilename('fullpath'));
            for i = 1:size(scenarios, 1)
                name  = scenarios{i, 1};
                build = scenarios{i, 2};
                dataDir = fullfile(here, 'data', name);

                d = build();
                parts = saveAndUnzipParts(d);

                % (a) raw-uint8 byte-identity on word/document.xml
                emittedDoc = parts('word/document.xml');
                frozenDoc  = readBytes(fullfile(dataDir, 'document.xml'));
                testCase.verifyEqual(uint8(emittedDoc(:)'), uint8(frozenDoc(:)'), ...
                    sprintf('%s: word/document.xml must be BYTE-IDENTICAL to the frozen python-docx part', name));

                % (b) full-package 17/17 size+SHA vs the frozen manifest
                man = jsondecode(native2unicode(readBytes(fullfile(dataDir, 'manifest.json')), 'UTF-8'));
                testCase.verifyEqual(numel(man.parts), 17, ...
                    sprintf('%s manifest must list exactly 17 parts', name));
                testCase.verifyEqual(double(parts.Count), 17, ...
                    sprintf('%s emitted package must contain exactly 17 parts', name));
                for k = 1:numel(man.parts)
                    pn = man.parts(k).name;
                    testCase.verifyTrue(isKey(parts, pn), ...
                        sprintf('%s: emitted package is missing part %s', name, pn));
                    b = parts(pn);
                    testCase.verifyEqual(numel(b), man.parts(k).size, ...
                        sprintf('%s part %s size == frozen manifest', name, pn));
                    testCase.verifyEqual(sha256hex(b), string(man.parts(k).sha256), ...
                        sprintf('%s part %s SHA-256 == frozen python-docx oracle (L1)', name, pn));
                end
            end
        end
    end
end

% ===================== byte-scenario build twins (VERBATIM) ============ %
% Each replays its validation\mat2doc\scenarios\sXXXX_*_gscenario.m body VERBATIM
% (same public API sequence) and returns the built Document for save+unzip.

function d = buildS0073()
    % s0073: cell content -- text setter, non-ASCII (H2), multi-paragraph. VERBATIM.
    d = mat2doc.Document();
    t = d.add_table(2, 2);
    c00 = t.cell(0, 0); c00.text = "A";
    nonascii = string(native2unicode( ...
        uint8([99 97 102 195 169 32 228 184 173 32 240 159 142 137]), 'UTF-8'));   % "café 中 🎉"
    c01 = t.cell(0, 1); c01.text = nonascii;
    c = t.cell(1, 0);
    c.add_paragraph("p1");
    c.add_paragraph("p2");
    % cell(1,1) left default-empty
end

function d = buildS0074()
    % s0074: add_row + populate via Row_.cells (trailing space preserved). VERBATIM.
    d = mat2doc.Document();
    t = d.add_table(2, 2);
    row = t.add_row();
    rc = row.cells;
    c1 = rc(1); c1.text = "X";
    c2 = rc(2); c2.text = "Y ";
end

function d = buildS0075()
    % s0075: add_column(Inches(1)) + populate via Column_.cells. VERBATIM.
    d = mat2doc.Document();
    t = d.add_table(2, 2);
    col = t.add_column(mat2doc.shared.Inches(1));
    cc = col.cells;
    c1 = cc(1); c1.text = "N1";
    c2 = cc(2); c2.text = "N2";
end

function d = buildS0076()
    % s0076: nested 2x2 table inside cell(0,0) via Cell_.add_table. VERBATIM.
    d = mat2doc.Document();
    t = d.add_table(2, 2);
    outer = t.cell(0, 0);
    inner = outer.add_table(2, 2);
    ic = inner.cell(0, 0); ic.text = "inner";
end

function d = buildS0077()
    % s0077: HORIZONTAL merge cell(0,0)+(0,1) -> gridSpan 2 (guard). VERBATIM.
    d = mat2doc.Document();
    t = d.add_table(2, 2);
    t.cell(0, 0).merge(t.cell(0, 1));
end

function d = buildS0078()
    % s0078: VERTICAL merge cell(0,0)+(1,0) -> vMerge col0. VERBATIM.
    d = mat2doc.Document();
    t = d.add_table(2, 2);
    t.cell(0, 0).merge(t.cell(1, 0));
end

function d = buildS0079()
    % s0079: 2x2 BLOCK merge cell(0,0)+(1,1) (guard). VERBATIM.
    d = mat2doc.Document();
    t = d.add_table(2, 2);
    t.cell(0, 0).merge(t.cell(1, 1));
end

function d = buildS0080()
    % s0080: merge then set text on the merged cell. VERBATIM.
    d = mat2doc.Document();
    t = d.add_table(2, 2);
    m = t.cell(0, 0).merge(t.cell(0, 1));
    m.text = "merged";
end

function d = buildS0082()
    % s0082: 3x3 mixed merges + text through NON-ROOT positions. VERBATIM.
    d = buildS0082Doc();
end

function d = buildS0082Doc()
    d = mat2doc.Document();
    t = d.add_table(3, 3);
    t.cell(0, 0).merge(t.cell(0, 1));   % row0 horizontal cols0-1
    t.cell(1, 0).merge(t.cell(2, 1));   % block rows1-2 cols0-1
    t.cell(1, 2).merge(t.cell(2, 2));   % v-span col2 rows1-2
    c01 = t.cell(0, 1); c01.text = "TL";
    c02 = t.cell(0, 2); c02.text = "TR";
    c21 = t.cell(2, 1); c21.text = "BLK";
    c22 = t.cell(2, 2); c22.text = "V";
end

function t = buildS0082Table()
    % The s0082 3x3 mixed-merge table (for the grid-walk span-identity leg). Same
    % VERBATIM sequence as buildS0082; returns the live Table handle (the Document
    % stays alive via the table's parent chain).
    d = mat2doc.Document();
    t = d.add_table(3, 3);
    t.cell(0, 0).merge(t.cell(0, 1));
    t.cell(1, 0).merge(t.cell(2, 1));
    t.cell(1, 2).merge(t.cell(2, 2));
    c01 = t.cell(0, 1); c01.text = "TL";
    c02 = t.cell(0, 2); c02.text = "TR";
    c21 = t.cell(2, 1); c21.text = "BLK";
    c22 = t.cell(2, 2); c22.text = "V";
end

% ===================== file-local helpers ============================== %

function parts = saveAndUnzipParts(d)
    % Document.save() to a temp .docx, unzip once, return a containers.Map of every
    % package part: relative POSIX path -> raw uint8 bytes. Both temp artifacts
    % cleaned on exit.
    tmp = [tempname '.docx'];
    cleanTmp = onCleanup(@() deleteIfExists(tmp)); %#ok<NASGU>
    d.save(tmp);
    exdir = tempname;
    cleanEx = onCleanup(@() rmdirIfExists(exdir)); %#ok<NASGU>
    unzip(tmp, exdir);
    listing = dir(fullfile(exdir, '**', '*'));
    parts = containers.Map('KeyType', 'char', 'ValueType', 'any');
    for i = 1:numel(listing)
        if listing(i).isdir, continue; end
        full = fullfile(listing(i).folder, listing(i).name);
        rel  = strrep(full(numel(exdir) + 2 : end), filesep, '/');   % POSIX rel path
        parts(rel) = readBytes(full); %#ok<NASGU>
    end
end

function p = partition0(cells)
    % 0-based first-identity partition of a Cell_ handle array (Python `is`-partition;
    % mirrors s0081_p6_4b_cell_probe.m::partition). p(i) is the 0-based index of the
    % FIRST array position holding the same handle as position i (== handle identity).
    n = numel(cells);
    p = zeros(1, n);
    for i = 1:n
        for j = 1:i
            if cells(j) == cells(i)
                p(i) = j - 1;
                break
            end
        end
    end
end

function b = readBytes(pth)
    f = fopen(pth, 'r', 'n');            % binary read (no CRLF translation)
    assert(f >= 0, 'could not open for read: %s', pth);
    b = fread(f, Inf, '*uint8')';
    fclose(f);
end

function setCellWidth(c, v)
    % Assign a cell width (for verifyError on the not-Optional setter; property
    % assignment is illegal inside an anonymous function).
    c.width = v;
end

function deleteIfExists(pth)
    if isfile(pth), delete(pth); end
end

function rmdirIfExists(pth)
    if isfolder(pth), rmdir(pth, 's'); end
end

function h = sha256hex(bytes)
    % Lowercase hex SHA-256 of a uint8 vector (base MATLAB / Java MessageDigest;
    % matches the python hashlib manifest SHAs).
    md  = java.security.MessageDigest.getInstance('SHA-256');
    dig = md.digest(typecast(uint8(bytes(:)'), 'int8'));
    b   = typecast(dig, 'uint8');
    h   = lower(string(sprintf('%02x', b)));
end

% ---- s0081 probe replay (the .m twin body, VERBATIM) ------------------ %

function P = runProbe()
    % Replay the s0081 cell/grid-walk probe (the .m twin's P-struct body, VERBATIM
    % tags/inputs/order) and return the nested struct of tagged canonical values.
    % Embedded so the Equivalence leg is self-contained (the validation-folder
    % scenario is NOT on the toolbox path and must not be a dependency). Mirrors
    % validation\mat2doc\scenarios\s0081_p6_4b_cell_probe.m lines 16-106.
    IN  = @(v) mat2doc.shared.Inches(v);
    WCV = @(n) mat2doc.enum.table.WD_CELL_VERTICAL_ALIGNMENT.(n);

    P = struct();

    % ============================ cell_props ============================
    d = mat2doc.Document();
    t = d.add_table(2, 2);
    c = t.cell(0, 0);
    cp = struct();
    cp.default_text   = rv(c.text);
    cp.grid_span      = rv(c.grid_span);
    cp.valign_default = rv(c.vertical_alignment);
    cp.width_default  = rv(c.width);
    cp.paragraphs_len = rv(numel(c.paragraphs));
    cp.tables_len     = rv(numel(c.tables));
    c.text = "hi";
    cp.after_set_text = rv(c.text);
    c.vertical_alignment = WCV("CENTER");
    cp.after_valign_center = rv(c.vertical_alignment);
    c.vertical_alignment = [];
    cp.after_valign_none = rv(c.vertical_alignment);
    c.width = IN(2);
    cp.after_width_inches2 = rv(c.width);
    p2 = c.add_paragraph("p2");
    cp.add_paragraph_ret_text = rv(p2.text);
    cp.paragraphs_len_after_add = rv(numel(c.paragraphs));
    cp.multi_para_text = rv(c.text);
    c.add_table(2, 2);
    cp.tables_len_after_nested = rv(numel(c.tables));
    P.cell_props = cp;

    % ============================ merge_props ==========================
    d2 = mat2doc.Document();
    t2 = d2.add_table(2, 2);
    src = t2.cell(0, 0);
    m = src.merge(t2.cell(0, 1));
    mp = struct();
    mp.grid_span = rv(m.grid_span);
    m.text = "merged";
    mp.merged_text = rv(m.text);
    P.merge_props = mp;

    % ============================ return_types =========================
    rt = struct();
    da = mat2doc.Document();
    ta = da.add_table(2, 2);
    ra = ta.add_row();
    rt.add_row_index    = rv(ra.index_());
    rt.add_row_cells_len = rv(numel(ra.cells));
    db = mat2doc.Document();
    tb = db.add_table(2, 2);
    cb = tb.add_column(IN(1));
    rt.add_col_index    = rv(cb.index_());
    rt.add_col_cells_len = rv(numel(cb.cells));
    P.return_types = rt;

    % ============================ grid_walk ============================
    d3 = mat2doc.Document();
    t3 = d3.add_table(3, 3);
    t3.cell(0, 0).merge(t3.cell(0, 1));   % row0 horizontal cols0-1
    t3.cell(1, 0).merge(t3.cell(2, 1));   % block rows1-2 cols0-1
    t3.cell(1, 2).merge(t3.cell(2, 2));   % v-span col2 rows1-2
    e01 = t3.cell(0, 1); e01.text = "TL";
    e02 = t3.cell(0, 2); e02.text = "TR";
    e21 = t3.cell(2, 1); e21.text = "BLK";
    e22 = t3.cell(2, 2); e22.text = "V";
    gw = struct();
    cells = t3.cells_();                               % ONE access
    gw.partition = partition0(cells);
    gw.texts = arrayfun(@(x) rv(x.text), cells, 'UniformOutput', false);
    ct = cell(1, 9); k = 0;
    for r = 0:2
        for cc = 0:2
            k = k + 1;
            cellrc = t3.cell(r, cc);
            ct{k} = rv(cellrc.text);
        end
    end
    gw.cell_texts = ct;
    rows = t3.rows;
    nrows = rows.len_();
    rlens = zeros(1, nrows);
    rparts = cell(1, nrows);
    for i = 1:nrows
        rc = rows.getitem_(i - 1).cells;
        rlens(i) = numel(rc);
        rparts{i} = partition0(rc);
    end
    gw.row_lens = rlens;
    gw.row_parts = rparts;
    a = t3.cells_(); b = t3.cells_();
    gw.cross_access_same = rv(a(1) == b(1));
    P.grid_walk = gw;
end

function s = rv(x)
    % Uniform accessor repr mirroring the s0081 rv(): None->"None", enum->member NAME,
    % bool->"True"/"False", Length/int->EMU decimal.
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

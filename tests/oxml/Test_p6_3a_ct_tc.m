classdef Test_p6_3a_ct_tc < matlab.unittest.TestCase
% TEST_P6_3A_CT_TC  Gate-4 permanent unit tests for Mat2Doc P6-3a -- the SINGLE
%   HARDEST WP of the port (risk register #2): the table-cell MERGE engine.
%   src/docx/oxml/table.py::CT_Tc (421-783) + CT_TcPr (784-891) ->
%   +mat2doc\+oxml\+table\{CT_Tc,CT_TcPr}.m, plus the THREE P6-3a registry rows
%   (w:gridSpan -> CT_DecimalNumber [:172]; w:tc -> CT_Tc [:179]; w:tcPr -> CT_TcPr
%   [:180]) and the CT_Row tc un-stub (new_tc_ -> CT_Tc.new(); tc_at_grid_offset
%   grid_span walk now live).
%
%   This class permanently freezes what the prior gates established:
%     * Gate-1 Porter  : audit_P6-3a_ct_tc.md (self-probe 31/31; merge byte matrix,
%       bare-vMerge, H4 width, styles.xml SHA).
%     * Gate-2 Auditor (Fable): APPROVE -- INDEPENDENT 16/16 merge byte-matrix on
%       disjoint auditor-minted fixtures (incl. nested n01/n02), handle-identity,
%       every H1 site, 10 verbatim raises, M1 17/17. ZERO new D-numbers.
%     * Gate-3 Validator: validate_P6-3a_ct_tc.md -- PASS, ZERO new D-numbers. FROZE
%       the merge oracle: references\s0063\ (10 src + 10 merged, manifest+SHAs) and
%       references\s0064\ (probe.json + 8 input parts). Targeted regression 69/69,
%       0 flips. M1 17/17 (styles.xml 349458 B / 02d71a68...; document.xml 1548 B /
%       0e4dd503...; stylesWithEffects 438131 B / 463ae092...).
%
%   THE HIGHEST-VALUE PERMANENT PINS IN THIS WP (a regression goes RED loudly):
%     * (MATRIX) THE 10-CASE MERGE BYTE-MATRIX (test_merge_byte_matrix): for each
%       frozen s0063 case, read the SOURCE bytes, parse, navigate to the two merge-
%       corner cells by the frozen coord path, CT_Tc.merge, serialize the whole
%       <w:tbl>, and assert the result is BYTE-IDENTICAL (uint8 + SHA-256) to the
%       frozen MERGED oracle. All 10: horizontal 1x2, vertical 2x1, block 2x2,
%       row-span gridSpan=3, merge-into-existing-h-span, merge-into-existing-v-span,
%       width-sum 2880, H4 falsy-zero (0-width skip), sdt-passenger consolidation,
%       nested inner-merge. This IS the permanent table-merge byte oracle.
%     * (BARE vMerge) test_merge_m02_bare_vmerge_guard -- the SHARPEST NEW-D guard:
%       the m02 vertical merge emits the continuation cell as a BARE <w:vMerge/>
%       (NO @w:val) and the top cell as <w:vMerge w:val="restart"/>; the substring
%       "continue" is ABSENT. The port NEVER writes w:val="continue" (D-delta-1:
%       CT_VMerge OptionalAttribute default "continue" DELETES @w:val).
%     * (H4 SKIP) test_merge_width_sum_and_h4_skip -- m07 sums 1440+1440 -> 2880
%       (tcW BEFORE gridSpan, H11); m08's 0-EMU next-cell width is FALSY, so
%       _add_width_of SKIPS the add and w STAYS 1440 (NOT 2880). Falsy-zero.
%     * (NESTED) test_merge_m10_nested -- the inner 2x2 table's inner merge emits
%       gridSpan=2; the outer host table (exactly one outer + one inner <w:tbl>) is
%       structurally unaffected: ./ancestor::w:tbl[position()=1] resolves NEAREST.
%     * (RAISES) test_invalid_span_raises -- the 9 raises (6 span_dimensions L/tee +
%       2 swallow + tr_above): mat2doc:InvalidSpanError / mat2doc:ValueError with
%       the four VERBATIM messages ("requested span not rectangular", "not enough
%       grid columns", "span is not rectangular", "no tr above topmost tr in
%       w:tbl"). Identifiers AND messages pinned.
%     * (HANDLE-IDENTITY) test_handle_identity_uniform_3x3 -- a UNIFORM 3x3 (byte-
%       identical rows) -> tr_idx_ per row == [0 1 2] (NOT [0 0 0]). THE anti-
%       regression guard against a future isequal-on-content rewrite of _tr_idx
%       (H5 handle identity). Loudly commented.
%     * (M1) test_m1_styles_and_document -- mat2doc.Document().save() ->
%       word/styles.xml == 349458 B / 02d71a68... (the 595-node CT_TcPr parse-path
%       neutrality guard) AND word/document.xml == 1548 B / 0e4dd503... SHA is L1.
%     * (H11) test_ct_tcpr -- the CT_TcPr _tag_seq successor order: a SCRAMBLED-order
%       build lands [tcW, gridSpan, vMerge, vAlign] and the full serialized <w:tcPr>
%       is byte-identical to the frozen oracle string.
%
%   Provenance (all Gate-3 frozen 2026-07-31):
%     * Audit    : validation\mat2doc\audit_P6-3a_ct_tc.md
%     * Validate : validation\mat2doc\validate_P6-3a_ct_tc.md
%     * Scenarios: validation\mat2doc\scenarios\s0063_p6_3a_merge_matrix.{py,m}
%                  (the merge-matrix twin; resolve_corner replayed VERBATIM below),
%                  s0064_p6_3a_probe.{py,m} (the CT_Tc/CT_TcPr/CT_Row-unstub probe;
%                  runProbe() below replays its body VERBATIM).
%     * Frozen refs (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE at Gate-3):
%         references\s0063\ (10 src_*.xml + 10 merged_*.xml + manifest.json) copied
%           verbatim into tests\oxml\data\s0063\ WITH a co-located `.gitattributes`
%           `* binary` pin (frozen-byte fixtures must not be line-ending mangled on
%           the master checkout -- the Gate-4 byte-fixture lesson).
%         references\s0064\ (probe.json + 8 input parts + manifest.json) copied into
%           tests\oxml\data\s0064\ WITH the same `* binary` pin (the input snippet
%           parts are parsed byte-fixtures; the probe.json oracle is the Equivalence
%           replay target).
%         references\s0001\parts\word\{styles,document}.xml -- the M1 byte
%           references (SHA of what Document().save() emits); NOT copied.
%
%   Coverage taxonomy
%   -----------------
%   * Nominal      -- CT_Tc.new grid_span/vMerge/width get+set; CT_TcPr accessors;
%                     grid geometry (offset/left/right/top/bottom); CT_Row un-stub;
%                     the 10 documented merges.
%   * Edge         -- bare <w:vMerge/> continuation (m02); H4 falsy-zero width skip
%                     (m08); grid_span==1 emits NO <w:gridSpan> (H4 strict >1);
%                     vMerge/vAlign/width set-None removal; _tr_below last-row->None
%                     (no wrap); tc_at_grid_offset(1) over a gridSpan-2 first cell
%                     -> ValueError; nested inner-merge (m10); the 9 error raises.
%   * Equivalence  -- test_equivalence_full_battery_vs_frozen_oracle replays the
%                     ENTIRE s0064 probe (runProbe, the .m twin's body verbatim) and
%                     round-trips it through jsondecode(jsonencode(...)) to normalize
%                     container shapes, then isequaln-compares each section to the
%                     frozen probe.json oracle (Gate-3 probe_diff was exit 0).
%   * Regression   -- hard-coded expected serialized-XML strings (ASCII ==
%                     byte-identical L1) + the 10 merge SHA-256 pins + the M1
%                     styles/document part SHAs + the frozen s0064 h11_serialized.
%   * Upstream     -- the s0063 source tables are REAL python-docx add_table output
%                     (full 18-decl Word nsmap); the frozen merged fixtures ARE
%                     lxml's expected serialization of the merged trees. Upstream
%                     provenance of the merge semantics: tests/oxml/test_table.py,
%                     tests/test_files/snippets/tbl-cells.txt, features/
%                     tbl-merge-cells.feature (audit_P6-3a §12).
%
%   Byte-level (L1) note: every merge comparison is BOTH a raw uint8 byte-equality
%   AND a SHA-256 equality vs the frozen oracle -- the ladder demanded L1 and Gate-3
%   delivered 10/10 byte-identical with ZERO new D-numbers, so every pin here is L1.
%   The only looser-than-byte assertion is the Equivalence section isequaln (values,
%   not bytes) and its non-triviality floor -- commented at its site. The
%   mat2doc:InvalidSpanError / mat2doc:ValueError identifier segments equal the
%   Python class names; the message texts are verbatim (the signed exception-model
%   mapping, design.md section 2 -- non-byte, non-output, NOT a D-number).
%
%   Determinism: no network, no absolute paths. The worktree root and the co-located
%   s0063/s0064 fixtures resolve relative to this file via
%   fileparts(mfilename('fullpath')); the M1 save goes to a tempname .docx deleted
%   via onCleanup; every fixture read is binary ('r','n'). The +mat2doc package
%   resolves via the MANDATORY PathFixture(worktree-root) in TestClassSetup
%   (WP9-F4 lesson).

    properties (Constant)
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        % --- registered class names (P6-3a rows, +oxml\registry.m 309/315/316) ---
        CT_TC     = 'mat2doc.oxml.table.CT_Tc'
        CT_TCPR   = 'mat2doc.oxml.table.CT_TcPr'
        CT_ROW    = 'mat2doc.oxml.table.CT_Row'
        CT_DECNUM = 'mat2doc.oxml.shared.CT_DecimalNumber'

        % --- frozen s0001 M1 byte references (the CT_TcPr 595-node styles.xml
        %     parse-path neutrality guard) ---
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
            % Idiom copied verbatim from tests\oxml\Test_p6_2_table_props.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\oxml
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. registry rows (the 3 P6-3a flips)                             %
        % =============================================================== %

        function test_registry_resolves_three_rows(testCase)
            % Nominal / Regression (registry.m 309/315/316): REGISTRATION is what
            % flips the parse class of a table-cell subtree. w:gridSpan closes the
            % P4-6 deferral (CT_TcPr.grid_span dep); w:tc/w:tcPr are the merge engine.
            pairs = { ...
                "w:gridSpan", testCase.CT_DECNUM; ...
                "w:tc",       testCase.CT_TC; ...
                "w:tcPr",     testCase.CT_TCPR };
            for i = 1:size(pairs, 1)
                tag = pairs{i, 1}; cls = pairs{i, 2};
                r = mat2doc.oxml.registry(mat2doc.oxml.qn(tag));
                testCase.verifyEqual(char(r), cls, ...
                    sprintf('registry must resolve %s -> %s (P6-3a row)', tag, cls));
                testCase.verifyEqual(class(mat2doc.oxml.OxmlElement(tag)), cls, ...
                    sprintf('OxmlElement(%s) must be a %s', tag, cls));
            end
        end

        % =============================================================== %
        % 2. ★ THE MERGE BYTE-MATRIX -- the headline L1 pins (10/10)        %
        % =============================================================== %

        function test_merge_byte_matrix(testCase)
            % ★ (MATRIX) Regression + Upstream (byte-identical L1). For each of the
            % 10 frozen s0063 cases: read the SOURCE bytes (verify size+SHA intact),
            % parse_xml, navigate to the two merge-corner cells by the frozen coord
            % path, CT_Tc.merge, serialize the whole <w:tbl>, and assert the result
            % is BYTE-IDENTICAL (raw uint8 AND SHA-256) to the frozen MERGED oracle.
            % This IS the permanent table-merge byte oracle (Gate-3 froze 10/10).
            tbl = mergeFixtureTable();
            here = fileparts(mfilename('fullpath'));
            for i = 1:size(tbl, 1)
                name   = tbl{i, 1};
                mergeA = tbl{i, 2};   mergeB = tbl{i, 3};
                srcSz  = tbl{i, 4};   srcSha = tbl{i, 5};
                mrgSz  = tbl{i, 6};   mrgSha = tbl{i, 7};

                srcBytes = readBytes(fullfile(here, 'data', 's0063', 'parts', ['src_' name '.xml']));
                testCase.verifyEqual(numel(srcBytes), srcSz, ...
                    sprintf('%s SOURCE fixture size intact', name));
                testCase.verifyEqual(sha256hex(srcBytes), srcSha, ...
                    sprintf('%s SOURCE fixture SHA-256 intact', name));

                root = mat2doc.oxml.parse_xml(srcBytes);
                a = resolve_corner(root, mergeA);
                b = resolve_corner(root, mergeB);
                top_tc = a.merge(b);
                % merge() returns the top-left cell of the new span -- a CT_Tc.
                testCase.verifyClass(top_tc, 'mat2doc.oxml.table.CT_Tc', ...
                    sprintf('%s merge() returns a CT_Tc top-left cell', name));

                outBytes = mat2doc.oxml.serialize_part_xml(root);
                % raw byte-equality (L1) vs the frozen MERGED oracle
                frozen = readBytes(fullfile(here, 'data', 's0063', 'parts', ['merged_' name '.xml']));
                testCase.verifyEqual(numel(outBytes), mrgSz, ...
                    sprintf('%s MERGED size == frozen oracle (%d B)', name, mrgSz));
                testCase.verifyEqual(uint8(outBytes(:)'), uint8(frozen(:)'), ...
                    sprintf('%s CT_Tc.merge -> serialize must be BYTE-IDENTICAL to the frozen merged oracle', name));
                testCase.verifyEqual(sha256hex(outBytes), mrgSha, ...
                    sprintf('%s MERGED SHA-256 == frozen manifest (L1)', name));
            end
        end

        function test_merge_m02_bare_vmerge_guard(testCase)
            % ★ (BARE vMerge) Edge -- THE SHARPEST NEW-D GUARD. The m02 vertical
            % merge (0,0)+(1,0) MUST emit the continuation cell as a BARE <w:vMerge/>
            % (NO @w:val) and the top cell as <w:vMerge w:val="restart"/>; the
            % literal "continue" must be ABSENT from the output. The port NEVER
            % writes w:val="continue": CT_TcPr.vMerge_val="continue" -> CT_VMerge's
            % OptionalAttribute setter sees value==default("continue") and DELETES
            % @w:val (D-delta-1), producing the bare element. A regression here would
            % write <w:vMerge w:val="continue"/> and corrupt every vertical merge.
            out = doMerge('m02_vert_2x1', "0,0", "1,0");
            s = string(native2unicode(out, "UTF-8"));
            testCase.verifyTrue(contains(s, '<w:vMerge w:val="restart"/>'), ...
                'm02 top cell serializes <w:vMerge w:val="restart"/>');
            testCase.verifyTrue(contains(s, '<w:vMerge/>'), ...
                'm02 continuation cell serializes a BARE <w:vMerge/> (no @w:val)');
            testCase.verifyFalse(contains(s, 'continue'), ...
                'm02 must NEVER contain the literal "continue" (D-delta-1 deletes @w:val)');
            % exactly two w:vMerge elements (one restart top + one bare continuation)
            testCase.verifyEqual(count(s, '<w:vMerge'), 2, ...
                'm02 has exactly two w:vMerge elements (restart top + bare continuation)');
        end

        function test_merge_width_sum_and_h4_skip(testCase)
            % ★ (H4 SKIP) Edge/Regression. m07: two 1440-twip cells merge -> summed
            % <w:tcW w:type="dxa" w:w="2880"/> (tcW BEFORE gridSpan, H11). m08: the
            % next cell's width is 0 EMU (FALSY in Python; Length(0) is falsy), so
            % _add_width_of SKIPS the add and w STAYS 1440 (NOT 2880). This is the
            % H4 falsy-zero replication -- a naive `if width is not None` port would
            % wrongly sum to 1440 and emit 1440 here too, but would sum a REAL 0-plus
            % elsewhere; the guard is that a 0-width contributes NOTHING.
            s7 = string(native2unicode(doMerge('m07_width_sum', "0,0", "0,1"), "UTF-8"));
            testCase.verifyTrue(contains(s7, '<w:tcW w:type="dxa" w:w="2880"/>'), ...
                'm07 sums 1440+1440 -> <w:tcW w:w="2880"/> (H4 both-truthy add)');
            % tcW BEFORE gridSpan (H11 _tag_seq successor slice)
            iW = strfind(char(s7), '<w:tcW');  iG = strfind(char(s7), '<w:gridSpan');
            testCase.verifyTrue(~isempty(iW) && ~isempty(iG) && iW(1) < iG(1), ...
                'm07 tcW serializes BEFORE gridSpan (H11)');

            s8 = string(native2unicode(doMerge('m08_zerow_skip', "0,0", "0,1"), "UTF-8"));
            testCase.verifyTrue(contains(s8, '<w:tcW w:type="dxa" w:w="1440"/>'), ...
                'm08 H4 falsy-zero: 0-EMU next-cell width SKIPPED -> w STAYS 1440');
            testCase.verifyFalse(contains(s8, 'w:w="2880"'), ...
                'm08 must NOT sum to 2880 (a 0-width contributes nothing)');
        end

        function test_merge_m09_sdt_passenger(testCase)
            % ★ (SDT) Edge. m09: the merged cells carry multi-paragraph content plus
            % an <w:sdt> passenger; the merge consolidates all block content into the
            % top-left cell (iter_block_items snapshot moves each once -- H9) and
            % restores the required trailing <w:p/> in the emptied cell. Byte-identity
            % to the frozen oracle already proves it; this guard asserts the sdt
            % survived the content-move (moved once, not dropped/doubled).
            s = string(native2unicode(doMerge('m09_content_sdt', "0,0", "0,1"), "UTF-8"));
            % exact start-tag '<w:sdt>' (not the '<w:sdtContent>' child) -- exactly
            % one passenger, moved once into the top-left cell (not dropped/doubled).
            testCase.verifyEqual(count(s, '<w:sdt>'), 1, ...
                'm09 the sdt passenger is moved exactly once (not dropped, not doubled)');
            testCase.verifyTrue(contains(s, '<w:gridSpan w:val="2"/>'), ...
                'm09 merged cell is gridSpan=2');
        end

        function test_merge_m10_nested(testCase)
            % ★ (NESTED) Edge. m10: a 2x2 outer table with a 2x2 INNER table inside
            % outer cell (1,1); the merge targets the INNER cells (0,0)+(0,1). The
            % inner merge emits <w:gridSpan w:val="2"/>, and the OUTER host is
            % structurally unaffected (exactly one outer + one inner <w:tbl>). This
            % proves ./ancestor::w:tbl[position()=1] resolves to the NEAREST ancestor
            % (the inner tbl), so the inner _tr_idx/_tr_below walk stays inside the
            % nested table. A reverse-axis regression would reach the OUTER tbl.
            out = doMerge('m10_nested', "1,1;0,0", "1,1;0,1");
            s = string(native2unicode(out, "UTF-8"));
            % exactly two w:tbl start-tags: outer '<w:tbl ' (attrs) + inner '<w:tbl>'
            testCase.verifyEqual(count(s, '<w:tbl>') + count(s, '<w:tbl '), 2, ...
                'm10 exactly one outer + one inner <w:tbl> (outer host structure intact)');
            testCase.verifyTrue(contains(s, '<w:gridSpan w:val="2"/>'), ...
                'm10 inner merge emits gridSpan=2 (nearest-ancestor tbl resolution)');
        end

        function test_merge_chain_cross_links(testCase)
            % Regression (independent corroboration of the merge chain). The frozen
            % manifest encodes three SHA cross-links proving the engine is path-
            % independent + idempotent at the byte level:
            %   m01.merged == m05.src  (a pre-existing gridSpan=2 IS an m01 result)
            %   m02.merged == m06.src  (a pre-existing vspan IS an m02 result)
            %   m04.merged == m05.merged (two routes to the same gridSpan=3 -> same bytes)
            t = mergeFixtureTable();
            % columns: 1=name 2=merge_a 3=merge_b 4=src_size 5=src_sha 6=merged_size 7=merged_sha
            cell_ = @(nm, col) t{strcmp(t(:,1), nm), col};
            testCase.verifyEqual(cell_('m01_horiz_1x2', 7), cell_('m05_into_hspan', 5), ...
                'm01.merged SHA == m05.src SHA (the gridSpan=2 input IS an m01 result)');
            testCase.verifyEqual(cell_('m02_vert_2x1', 7), cell_('m06_into_vspan', 5), ...
                'm02.merged SHA == m06.src SHA (the vspan input IS an m02 result)');
            testCase.verifyEqual(cell_('m04_rowspan_g3', 7), cell_('m05_into_hspan', 7), ...
                'm04.merged SHA == m05.merged SHA (two routes -> identical gridSpan=3 bytes)');
        end

        % =============================================================== %
        % 3. ★ The 9 invalid-span raises (verbatim ids + messages)          %
        % =============================================================== %

        function test_invalid_span_raises(testCase)
            % ★ (RAISES) Edge -- the 9 raises: 6 span_dimensions (inverted-L + tee),
            % 2 swallow, 1 tr_above. Pin BOTH the mat2doc: identifier AND the verbatim
            % message (the four distinct texts). Replays the frozen s0064 raise pairs.
            here = fileparts(mfilename('fullpath'));
            P = @(n) mat2doc.oxml.parse_xml(readBytes(fullfile(here, 'data', 's0064', 'parts', n)));

            % -- 6 span_dimensions raises (all "requested span not rectangular") --
            spanPairs = { ...
                'snip1.xml', [0 0], [1 0]; ...   % inverted-L (left-equal/right-differ)
                'snip2.xml', [0 0], [0 1]; ...   % inverted-L (top-equal/bottom-differ)
                'snip4.xml', [0 0], [1 0]; ...   % combo span
                'snip5.xml', [0 1], [1 0]; ...   % middle-row span
                'snip6.xml', [0 0], [1 1]; ...   % middle-col span
                'snip2.xml', [0 1], [0 2] };     % vspan-restart vs uniform (tee)
            for i = 1:size(spanPairs, 1)
                rows = P(spanPairs{i, 1}).xpath("./w:tr");
                arc = spanPairs{i, 2};  brc = spanPairs{i, 3};
                a = rows(arc(1) + 1).tc_at_grid_offset(arc(2));   % H1: row +1; c RAW 0-based
                b = rows(brc(1) + 1).tc_at_grid_offset(brc(2));
                fn = @() a.span_dimensions_(b);
                testCase.verifyError(fn, 'mat2doc:InvalidSpanError', ...
                    sprintf('span_dimensions %s (%d,%d)+(%d,%d) raises InvalidSpanError', ...
                        spanPairs{i, 1}, arc(1), arc(2), brc(1), brc(2)));
                testCase.verifyEqual(errmsg(fn), "requested span not rectangular", ...
                    'span_dimensions raise message verbatim');
            end

            % -- 2 swallow raises (the TWO distinct non-rectangular texts) --
            % swallow_a: 1x3, grow_to_(4,1) exhausts the row -> "not enough grid columns"
            rowsA = P('swallow_a.xml').xpath("./w:tr");
            fnA = @() rowsA(1).tc_at_grid_offset(0).grow_to_(4, 1);
            testCase.verifyError(fnA, 'mat2doc:InvalidSpanError', ...
                'swallow grow_to_(4,1) raises InvalidSpanError');
            testCase.verifyEqual(errmsg(fnA), "not enough grid columns", ...
                'swallow "not enough grid columns" verbatim');
            % swallow_b: grow_to_(2,1) over a gridSpan-2 next cell -> "span is not rectangular"
            rowsB = P('swallow_b.xml').xpath("./w:tr");
            fnB = @() rowsB(1).tc_at_grid_offset(0).grow_to_(2, 1);
            testCase.verifyError(fnB, 'mat2doc:InvalidSpanError', ...
                'swallow grow_to_(2,1) raises InvalidSpanError');
            testCase.verifyEqual(errmsg(fnB), "span is not rectangular", ...
                'swallow "span is not rectangular" verbatim');

            % -- tr_above on the topmost row -> ValueError, verbatim --
            rows0 = P('snip0.xml').xpath("./w:tr");
            fnT = @() rows0(1).tc_at_grid_offset(0).tr_above_();
            testCase.verifyError(fnT, 'mat2doc:ValueError', ...
                'tr_above_ on the topmost row raises mat2doc:ValueError');
            testCase.verifyEqual(errmsg(fnT), "no tr above topmost tr in w:tbl", ...
                'tr_above_ ValueError message verbatim');
        end

        % =============================================================== %
        % 4. ★ Handle-identity (H5) -- the anti-isequal guard               %
        % =============================================================== %

        function test_handle_identity_uniform_3x3(testCase)
            % ★ (HANDLE-IDENTITY) THE anti-regression guard against a future
            % isequal-on-content rewrite of _tr_idx. snip0.xml is a UNIFORM 3x3 grid
            % (every <w:tr> byte-identical). tr_idx_ uses HANDLE identity
            % (find(trLst == myTr, 1)), so the col-0 cell of each row returns its
            % TRUE 0-based row index [0 1 2]. An isequal(content)-based match would
            % return the FIRST matching row for all three -> [0 0 0] and silently
            % corrupt every merge walk while passing non-uniform fixtures.
            here = fileparts(mfilename('fullpath'));
            t0 = mat2doc.oxml.parse_xml(readBytes(fullfile(here, 'data', 's0064', 'parts', 'snip0.xml')));
            rows0 = t0.xpath("./w:tr");
            testCase.verifyEqual(numel(rows0), 3, 'snip0 is a 3-row uniform grid');
            idx = [ rows0(1).tc_at_grid_offset(0).tr_idx_(), ...
                    rows0(2).tc_at_grid_offset(0).tr_idx_(), ...
                    rows0(3).tc_at_grid_offset(0).tr_idx_() ];
            testCase.verifyEqual(idx, [0 1 2], ...
                'H5: tr_idx_ on a UNIFORM 3x3 must be [0 1 2] (handle identity), NOT [0 0 0]');
        end

        % =============================================================== %
        % 5. CT_Tc geometry + get/set delegators                           %
        % =============================================================== %

        function test_ct_tc_getset(testCase)
            % Nominal + Edge (s0064 tc_getset): CT_Tc.new() defaults + the three
            % tcPr delegators (grid_span/vMerge/width). Bare tc has NO tcPr -> the
            % None/1 defaults; setters route through get_or_add_tcPr.
            tc = mat2doc.oxml.table.CT_Tc.new();
            testCase.verifyEqual(tc.grid_span, 1, 'default grid_span 1 (no tcPr)');
            testCase.verifyTrue(isequal(tc.vMerge, []), 'default vMerge [] (None)');
            testCase.verifyTrue(isequal(tc.width, []),  'default width [] (None)');

            % grid_span set 3 -> emits <w:gridSpan w:val="3"/>; set 1 -> H4 strict >1
            % removes it entirely (no <w:gridSpan> element).
            tc = mat2doc.oxml.table.CT_Tc.new(); tc.grid_span = 3;
            testCase.verifyEqual(tc.grid_span, 3, 'grid_span read-back 3');
            testCase.verifyTrue(contains(ser(tc), '<w:gridSpan w:val="3"/>'), ...
                'grid_span=3 emits <w:gridSpan w:val="3"/>');
            tc = mat2doc.oxml.table.CT_Tc.new(); tc.grid_span = 1;
            testCase.verifyFalse(contains(ser(tc), '<w:gridSpan'), ...
                'H4 strict: grid_span=1 emits NO <w:gridSpan> element');

            % vMerge restart -> w:val="restart"; continue -> BARE; None -> removed
            tc = mat2doc.oxml.table.CT_Tc.new(); tc.vMerge = "restart";
            testCase.verifyEqual(string(tc.vMerge), "restart", 'vMerge restart read-back');
            testCase.verifyTrue(contains(ser(tc), '<w:vMerge w:val="restart"/>'), ...
                'vMerge restart -> <w:vMerge w:val="restart"/>');
            tc = mat2doc.oxml.table.CT_Tc.new(); tc.vMerge = "continue";
            testCase.verifyEqual(string(tc.vMerge), "continue", 'vMerge continue read-back');
            testCase.verifyTrue(contains(ser(tc), '<w:vMerge/>'), ...
                'vMerge continue -> BARE <w:vMerge/> (D-delta-1)');
            tc = mat2doc.oxml.table.CT_Tc.new(); tc.vMerge = "restart"; tc.vMerge = [];
            testCase.verifyFalse(contains(ser(tc), '<w:vMerge'), ...
                'vMerge None removes the <w:vMerge> element');

            % width Twips(1440) -> 914400 EMU read-back
            tc = mat2doc.oxml.table.CT_Tc.new(); tc.width = mat2doc.shared.Twips(1440);
            testCase.verifyEqual(double(tc.width), 914400, 'width Twips(1440) -> 914400 EMU');
        end

        function test_ct_tc_geometry(testCase)
            % Nominal + Edge (s0064 geometry / tr_below): the 0-based grid-geometry
            % read tier over a merged grid. grid_offset accumulates preceding-sibling
            % grid_span (RAW 0-based); left==grid_offset; right==grid_offset+grid_span;
            % top/bottom recurse over a vertical span; bottom is the EXCLUSIVE lower
            % bound. _tr_below on the last row -> None (no wrap).
            here = fileparts(mfilename('fullpath'));
            P = @(n) mat2doc.oxml.parse_xml(readBytes(fullfile(here, 'data', 's0064', 'parts', n)));

            % snip1 row0: a gridSpan=2 first cell then a span-1 cell.
            r = P('snip1.xml').xpath("./w:tr");
            c0 = r(1).tc_lst(1);  c1 = r(1).tc_lst(2);
            testCase.verifyEqual(c0.grid_span, 2, 'snip1 (0,0) grid_span 2');
            testCase.verifyEqual(c0.grid_offset, 0, 'snip1 (0,0) grid_offset 0');
            testCase.verifyEqual(c0.left, 0, 'snip1 (0,0) left 0 (== grid_offset)');
            testCase.verifyEqual(c0.right, 2, 'snip1 (0,0) right 2 (offset+span)');
            testCase.verifyEqual(c0.top, 0, 'snip1 (0,0) top 0');
            testCase.verifyEqual(c0.bottom, 1, 'snip1 (0,0) bottom 1 (EXCLUSIVE, tr_idx+1)');
            testCase.verifyEqual(c1.grid_offset, 2, 'snip1 (0,1) grid_offset 2 (after the gridSpan=2)');
            testCase.verifyEqual(c1.right, 3, 'snip1 (0,1) right 3');

            % snip2 col-1 is a vertical span: row0 restart, row1 continuation. top/
            % bottom RECURSE across the span (continuation.top -> above.top == 0;
            % restart.bottom -> below.bottom == 2).
            r = P('snip2.xml').xpath("./w:tr");
            top = r(1).tc_lst(2);   cont = r(2).tc_lst(2);
            testCase.verifyEqual(string(top.vMerge), "restart", 'snip2 (0,1) vMerge restart');
            testCase.verifyEqual(top.top, 0, 'snip2 restart top 0');
            testCase.verifyEqual(top.bottom, 2, 'snip2 restart bottom 2 (recurses to continuation)');
            testCase.verifyEqual(string(cont.vMerge), "continue", 'snip2 (1,1) vMerge continue');
            testCase.verifyEqual(cont.top, 0, 'snip2 continuation top 0 (recurses to restart)');
            testCase.verifyEqual(cont.bottom, 2, 'snip2 continuation bottom 2');

            % _tr_below: snip0 uniform 3x3; last-row cell -> None (no wrap);
            % row-0 cell -> the next row (its tr_idx == 1).
            rows0 = P('snip0.xml').xpath("./w:tr");
            last = rows0(3).tc_at_grid_offset(0);
            testCase.verifyTrue(isequal(last.tr_below_(), []), ...
                '_tr_below on the LAST row -> [] (None, no wrap)');
            below = rows0(1).tc_at_grid_offset(0).tr_below_();
            testCase.verifyFalse(isequal(below, []), 'row-0 cell _tr_below is not None');
            testCase.verifyEqual(below.tr_idx, 1, 'row-0 cell _tr_below is the row at tr_idx 1');
        end

        % =============================================================== %
        % 6. CT_TcPr accessors + H11 _tag_seq order                        %
        % =============================================================== %

        function test_ct_tcpr(testCase)
            % Nominal + Edge + Regression (s0064 tcpr): the four accessors
            % (grid_span/vMerge_val/width/vAlign_val) get+set+None, and the H11
            % _tag_seq successor order via a SCRAMBLED-order build.
            tp = mat2doc.oxml.OxmlElement("w:tcPr");
            testCase.verifyEqual(class(tp), testCase.CT_TCPR, 'OxmlElement(w:tcPr) is CT_TcPr');
            testCase.verifyEqual(tp.grid_span, 1, 'default grid_span 1');
            testCase.verifyTrue(isequal(tp.vMerge_val, []), 'default vMerge_val [] (None)');
            testCase.verifyTrue(isequal(tp.width, []), 'default width [] (None)');
            testCase.verifyTrue(isequal(tp.vAlign_val, []), 'default vAlign_val [] (None)');

            % vAlign parse "center" -> CENTER member; set None removes
            center = parse_tcpr('<w:vAlign w:val="center"/>').vAlign_val;
            testCase.verifyEqual(string(center), "CENTER", 'parse <w:vAlign val="center"> -> CENTER');
            tp = mat2doc.oxml.OxmlElement("w:tcPr"); tp.vAlign_val = center;
            testCase.verifyEqual(string(tp.vAlign_val), "CENTER", 'vAlign_val set/read CENTER');
            tp.vAlign_val = [];
            testCase.verifyTrue(isequal(tp.vAlign_val, []), 'vAlign_val set None removes w:vAlign');

            % grid_span H4 strict; vMerge_val bare; width None-guards
            tp = mat2doc.oxml.OxmlElement("w:tcPr"); tp.grid_span = 2;
            testCase.verifyTrue(contains(ser(tp), '<w:gridSpan w:val="2"/>'), 'grid_span 2 emitted');
            tp = mat2doc.oxml.OxmlElement("w:tcPr"); tp.grid_span = 1;
            testCase.verifyFalse(contains(ser(tp), '<w:gridSpan'), 'grid_span 1 emits nothing (H4)');
            tp = mat2doc.oxml.OxmlElement("w:tcPr"); tp.vMerge_val = "continue";
            testCase.verifyTrue(contains(ser(tp), '<w:vMerge/>'), 'vMerge_val continue -> bare');
            tp = mat2doc.oxml.OxmlElement("w:tcPr"); tp.width = mat2doc.shared.Twips(1440);
            testCase.verifyEqual(double(tp.width), 914400, 'width Twips(1440) -> 914400 EMU');

            % ★ H11: set in SCRAMBLED order (vMerge, width, vAlign, gridSpan) -> the
            % children land in _tag_seq order [tcW, gridSpan, vMerge, vAlign], and the
            % full serialized <w:tcPr> is byte-identical to the frozen s0064 oracle.
            tp = mat2doc.oxml.OxmlElement("w:tcPr");
            tp.vMerge_val = "restart";
            tp.width = mat2doc.shared.Twips(1440);
            tp.vAlign_val = center;
            tp.grid_span = 2;
            testCase.verifyEqual(childLocalnames(tp), ["tcW" "gridSpan" "vMerge" "vAlign"], ...
                'H11: scrambled-order build lands in _tag_seq order [tcW, gridSpan, vMerge, vAlign]');
            expected = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>" + newline + ...
                "<w:tcPr xmlns:w=""" + testCase.W + """>" + ...
                "<w:tcW w:type=""dxa"" w:w=""1440""/><w:gridSpan w:val=""2""/>" + ...
                "<w:vMerge w:val=""restart""/><w:vAlign w:val=""center""/></w:tcPr>";
            testCase.verifyEqual(ser(tp), expected, ...
                'H11: full serialized <w:tcPr> byte-identical to the frozen s0064 oracle (L1)');
        end

        % =============================================================== %
        % 7. CT_Row un-stub (new_tc_ / tc_at_grid_offset now live)          %
        % =============================================================== %

        function test_ct_row_unstub(testCase)
            % (CT_Row un-stub, s0064 row_unstub): new_tc_ -> CT_Tc.new();
            % tc_at_grid_offset walks tc_lst summing grid_span (now live);
            % add_tc -> real CT_Tc. tc_at_grid_offset(1) over a gridSpan-2 first cell
            % steps PAST offset 1 (inside the span) and raises ValueError.
            here = fileparts(mfilename('fullpath'));
            row = mat2doc.oxml.OxmlElement("w:tr");
            testCase.verifyClass(row.new_tc_(), 'mat2doc.oxml.table.CT_Tc', ...
                'CT_Row.new_tc_ returns a CT_Tc (un-stubbed at P6-3a)');
            row2 = mat2doc.oxml.OxmlElement("w:tr");
            testCase.verifyClass(row2.add_tc(), 'mat2doc.oxml.table.CT_Tc', ...
                'CT_Row.add_tc adds+returns a CT_Tc');

            % snip1 row0: gridSpan=2 first cell (covers grid 0..1), then span-1 at 2.
            hrow = mat2doc.oxml.parse_xml(readBytes(fullfile(here, 'data', 's0064', 'parts', 'snip1.xml'))).xpath("./w:tr");
            hrow = hrow(1);
            testCase.verifyClass(hrow.tc_at_grid_offset(0), 'mat2doc.oxml.table.CT_Tc', ...
                'tc_at_grid_offset(0) -> the gridSpan=2 first cell');
            testCase.verifyClass(hrow.tc_at_grid_offset(2), 'mat2doc.oxml.table.CT_Tc', ...
                'tc_at_grid_offset(2) -> the span-1 second cell (walk past gridSpan=2)');
            % offset 1 is INSIDE the first cell's span -> no tc starts there
            fn = @() hrow.tc_at_grid_offset(1);
            testCase.verifyError(fn, 'mat2doc:ValueError', ...
                'tc_at_grid_offset(1) inside a gridSpan-2 cell raises mat2doc:ValueError');
            testCase.verifyEqual(errmsg(fn), "no `tc` element at grid_offset=1", ...
                'tc_at_grid_offset ValueError message verbatim');
        end

        % =============================================================== %
        % 8. ★ M1 styles.xml + document.xml byte-pins (neutrality guard)     %
        % =============================================================== %

        function test_m1_styles_and_document(testCase)
            % ★ (M1) Regression (byte-neutrality, L1): mat2doc.Document().save()
            % emits word/styles.xml at EXACTLY 349458 B (frozen s0001 SHA) even
            % though its 595 <w:tcPr> nodes now transit CT_TcPr on the M1 parse path
            % -- registering w:tcPr is byte-neutral (descriptors-only, no parse-time
            % behavior; the CT_TblPr/P4-6 precedent). word/document.xml stays 1548 B
            % (w:tc/w:gridSpan have 0 occurrences in default.docx). SHA-256 == L1.
            [styBytes, docBytes] = emitTwoParts('styles.xml', 'document.xml');

            testCase.verifyEqual(numel(styBytes), testCase.STYLES_SIZE_M1, ...
                sprintf('word/styles.xml must be exactly %d B (595 w:tcPr nodes transit CT_TcPr)', ...
                    testCase.STYLES_SIZE_M1));
            testCase.verifyEqual(sha256hex(styBytes), testCase.STYLES_SHA_M1, ...
                'word/styles.xml SHA-256 == frozen s0001 oracle (CT_TcPr parse-path neutral, L1)');

            testCase.verifyEqual(numel(docBytes), testCase.DOC_SIZE_M1, ...
                sprintf('word/document.xml must be exactly %d B', testCase.DOC_SIZE_M1));
            testCase.verifyEqual(sha256hex(docBytes), testCase.DOC_SHA_M1, ...
                'word/document.xml SHA-256 == frozen s0001 oracle (byte-identical L1)');
        end

        % =============================================================== %
        % 9. fixture-drift guard: manifest == the hard-coded pin table       %
        % =============================================================== %

        function test_s0063_manifest_matches_pinned(testCase)
            % Regression (fixture-drift guard): the shipped data\s0063\manifest.json
            % SHAs/sizes/coords must equal the hard-coded pin table in this class. A
            % silent re-freeze would flip the manifest but not the constant.
            here = fileparts(mfilename('fullpath'));
            raw = readBytes(fullfile(here, 'data', 's0063', 'manifest.json'));
            man = jsondecode(native2unicode(raw, 'UTF-8'));
            tbl = mergeFixtureTable();
            testCase.verifyEqual(numel(man.fixtures), size(tbl, 1), '10 manifest fixtures');
            for i = 1:numel(man.fixtures)
                fx = man.fixtures(i);
                row = tbl(strcmp(tbl(:, 1), fx.name), :);
                testCase.verifyEqual(size(row, 1), 1, sprintf('manifest fixture %s pinned', fx.name));
                testCase.verifyEqual(string(fx.merge_a), string(row{1, 2}), sprintf('%s merge_a', fx.name));
                testCase.verifyEqual(string(fx.merge_b), string(row{1, 3}), sprintf('%s merge_b', fx.name));
                testCase.verifyEqual(fx.src_size, row{1, 4}, sprintf('%s src_size', fx.name));
                testCase.verifyEqual(string(fx.src_sha256), row{1, 5}, sprintf('%s src_sha', fx.name));
                testCase.verifyEqual(fx.merged_size, row{1, 6}, sprintf('%s merged_size', fx.name));
                testCase.verifyEqual(string(fx.merged_sha256), row{1, 7}, sprintf('%s merged_sha', fx.name));
            end
        end

        % =============================================================== %
        % 10. EQUIVALENCE -- the full s0064 probe vs the frozen oracle        %
        % =============================================================== %

        function test_equivalence_full_battery_vs_frozen_oracle(testCase)
            % Equivalence: replay the ENTIRE s0064 probe (runProbe -- the .m twin's
            % body VERBATIM: geometry / tr_below / handle_id / tc_getset / tcpr /
            % row_unstub / raises) and compare EACH section to the frozen python-docx
            % 1.2.0 oracle copied into data\s0064\probe.json. Gate-3 probe_diff was
            % exit 0, so every section must be value-identical.
            %
            % Both sides are normalized through jsondecode(jsonencode(...)) so that
            % struct-array vs cell-of-struct container shapes collapse the SAME way
            % (the .m twin writes the port with jsonencode; the oracle came from the
            % identical JSON shape) -- the only looser-than-byte comparison in this
            % class (values, not bytes), justified because probe_diff already proved
            % byte/value equivalence at Gate-3.
            here = fileparts(mfilename('fullpath'));
            port   = runProbe(fullfile(here, 'data', 's0064'));
            portN  = jsondecode(jsonencode(port));               % normalize shapes
            oracleN = jsondecode(native2unicode( ...
                readBytes(fullfile(here, 'data', 's0064', 'probe.json')), 'UTF-8'));

            % Non-triviality floor (guards a silent-empty replay).
            testCase.verifyEqual(sort(fieldnames(portN)), sort(fieldnames(oracleN)), ...
                'the replayed probe and the frozen oracle expose the same top-level sections');
            testCase.verifyGreaterThanOrEqual(numel(fieldnames(oracleN)), 7, ...
                'the oracle must expose all 7 probe sections');

            % Section-by-section isequaln (localizes any regression to a section).
            secs = fieldnames(oracleN);
            for i = 1:numel(secs)
                s = secs{i};
                testCase.verifyTrue(isfield(portN, s), sprintf('port is missing section %s', s));
                testCase.verifyTrue(isequaln(portN.(s), oracleN.(s)), ...
                    sprintf('section "%s" must be value-identical to the frozen s0064 oracle', s));
            end
        end

    end
end

% ===================== file-local helpers ============================== %

function e = parse(xml)
    e = mat2doc.oxml.parse_xml(uint8(unicode2native(char(xml), "UTF-8")));
end

function s = ser(e)
    % Decode serialize_part_xml bytes to a string (ASCII content -> string-equality
    % is a byte-identical L1 assertion; non-ASCII round-trips via UTF-8).
    s = string(native2unicode(mat2doc.oxml.serialize_part_xml(e), "UTF-8"));
end

function names = childLocalnames(e)
    kids = e.xpath("./*");
    names = strings(1, numel(kids));
    for k = 1:numel(kids)
        names(k) = string(kids(k).local_part);
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

function e = parse_tcpr(inner)
    % Parse a <w:tcPr> with the given inner XML (mirrors the s0064 twin helper).
    xml = "<w:tcPr xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""" + ...
          ">" + inner + "</w:tcPr>";
    e = parse(xml);
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

% ---- s0063 merge-matrix pin table + corner navigation ----

function tbl = mergeFixtureTable()
    % Hard-coded pin table for the 10 frozen s0063 merge cases (manifest.json):
    % {name, merge_a, merge_b, src_size, src_sha256, merged_size, merged_sha256}.
    % SHAs are of the python-docx serialize_part_xml bytes == the MATLAB
    % parse->merge->serialize output (byte-identical, Gate-3 10/10 L1).
    SRC = "8ba252936bd65bf96640aa8f06d0f608b253bdf14cc38c80a527a3a3bf1b23a4"; % src of m01-m04
    tbl = { ...
        'm01_horiz_1x2',  "0,0",     "0,1",     2379, SRC, ...
            2340, "425401a74699931664bcf64d9bfba399b129f23d580dbefe9c464e386951bf1c"; ...
        'm02_vert_2x1',   "0,0",     "1,0",     2379, SRC, ...
            2423, "2a78bdb88109eaba7ea9be25e8a84f5a9338f8bce9facf4cee0d19bea903fd57"; ...
        'm03_block_2x2',  "0,0",     "1,1",     2379, SRC, ...
            2345, "598e97c14adba84fbf6eb6ac9184b1bac57949161b185c6b28fdfa614673c137"; ...
        'm04_rowspan_g3', "0,0",     "0,2",     2379, SRC, ...
            2278, "4018ccaf6106724c52bc36669a558e289e3beb01549066bea40d26371d3eb5c3"; ...
        'm05_into_hspan', "0,0",     "0,2",     2340, ...
            "425401a74699931664bcf64d9bfba399b129f23d580dbefe9c464e386951bf1c", ...
            2278, "4018ccaf6106724c52bc36669a558e289e3beb01549066bea40d26371d3eb5c3"; ...
        'm06_into_vspan', "0,0",     "1,1",     2423, ...
            "2a78bdb88109eaba7ea9be25e8a84f5a9338f8bce9facf4cee0d19bea903fd57", ...
            2345, "cc7ad8db4b931d2ec56eb5eb973825aace018fc9b3b46b380e59e61c3cbd4f99"; ...
        'm07_width_sum',  "0,0",     "0,1",     1658, ...
            "9b93a44567154712e93abc3b620924011ca78d58fd6bfcc01125b6b296803856", ...
            1619, "8bc2d3bd440ccd80cea57a8e30438baffd21f0add05e2187c64a1bc557de6bd8"; ...
        'm08_zerow_skip', "0,0",     "0,1",     1655, ...
            "7e45fe20b111fc2c972ca51d15e87b9ac0195887b7fb4bb880c756ff7b4130ac", ...
            1619, "f71f556afef3ab2908eff94242bcf37a4a08cbdb3b9ff12e1d5e45834fa81b51"; ...
        'm09_content_sdt',"0,0",     "0,1",     2506, ...
            "87b7701f969817b4e7c8c7e69314491f3ecf245a559b14de180608acef90beb0", ...
            2467, "653fe9e900fa603eb7380db399025292e2ec4238ebc2a18337ef45a5fb7a86f6"; ...
        'm10_nested',     "1,1;0,0", "1,1;0,1", 2542, ...
            "11965bc4fb5175c50b19871cf3e42a95db59d54ed4b28b2cdc5cac794d34b3be", ...
            2503, "58686bdc68e7d359b2d8c7b29ae95784dcff97c290d5ed30ebe390c5dcca0c38" };
end

function out = doMerge(name, mergeA, mergeB)
    % Read the frozen s0063 SOURCE, parse, navigate to the two corners, merge, and
    % return the serialized <w:tbl> bytes (uint8). Used by the loud content guards.
    here = fileparts(mfilename('fullpath'));
    src = readBytes(fullfile(here, 'data', 's0063', 'parts', ['src_' char(name) '.xml']));
    root = mat2doc.oxml.parse_xml(src);
    a = resolve_corner(root, mergeA);
    b = resolve_corner(root, mergeB);
    a.merge(b);
    out = mat2doc.oxml.serialize_part_xml(root);
end

function tc = resolve_corner(tbl, spec)
    % Walk the descent path "r,c" / "or,oc;ir,ic" to the corner CT_Tc (VERBATIM
    % from s0063_p6_3a_merge_matrix.m). rows(r+1).tc_at_grid_offset(c) at each level
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
            node = tc.tbl_lst(1);                % descend into nested w:tbl (generic)
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

% ---- s0064 probe replay (the .m twin body, VERBATIM) ----

function probe = runProbe(refdir)
    % Replay the s0064 probe sequence over the co-located frozen input parts and
    % return the nested struct of tagged canonical values (all leaves -> strings via
    % rv). Embedded here so the Equivalence leg is self-contained (the validation-
    % folder scenario is NOT on the toolbox path and must not be a dependency).
    % Mirrors validation\mat2doc\scenarios\s0064_p6_3a_probe.m lines 14-131 VERBATIM.
    refdir = char(refdir);
    P = @(n) mat2doc.oxml.parse_xml(readBytes(fullfile(refdir, 'parts', n)));

    probe = struct();

    % ---- geometry over the span snippets ----
    geo = struct();
    for key = ["snip1" "snip2" "snip4"]
        geo.(key) = cell_geometry(P(char(key) + ".xml"));
    end
    probe.geometry = geo;

    % ---- tr_below ----
    t0 = P("snip0.xml");
    rows0 = t0.xpath("./w:tr");
    last = rows0(3).tc_at_grid_offset(0);
    row0 = rows0(1).tc_at_grid_offset(0);
    below = row0.tr_below_();
    tb = struct();
    tb.last_row_is_none = rv(isequal(last.tr_below_(), []));
    if isequal(below, [])
        tb.row0_below_idx = rv([]);
    else
        tb.row0_below_idx = rv(below.tr_idx);
    end
    probe.tr_below = tb;

    % ---- handle-identity re-prove ----
    hid = cell(1, 3);
    for r = 1:3
        hid{r} = rv(rows0(r).tc_at_grid_offset(0).tr_idx_());
    end
    probe.handle_id = struct("tr_idx", {hid});

    % ---- CT_Tc.new() get/set ----
    g = struct();
    tc = mat2doc.oxml.table.CT_Tc.new(); g.default_grid_span = rv(tc.grid_span);
    tc = mat2doc.oxml.table.CT_Tc.new(); g.default_vMerge = rv(tc.vMerge);
    tc = mat2doc.oxml.table.CT_Tc.new(); g.default_width = rv(tc.width);
    tc = mat2doc.oxml.table.CT_Tc.new(); tc.grid_span = 3;
    g.set3_grid_span = rv(tc.grid_span);
    g.set3_has_gridSpan = rv(contains(ser(tc), "<w:gridSpan"));
    tc = mat2doc.oxml.table.CT_Tc.new(); tc.grid_span = 1;
    g.set1_has_gridSpan = rv(contains(ser(tc), "<w:gridSpan"));   % H4 strict >1
    tc = mat2doc.oxml.table.CT_Tc.new(); tc.vMerge = "restart";
    g.restart_read = rv(tc.vMerge);
    g.restart_bytes = rv(contains(ser(tc), '<w:vMerge w:val="restart"/>'));
    tc = mat2doc.oxml.table.CT_Tc.new(); tc.vMerge = "continue";
    g.continue_read = rv(tc.vMerge);
    g.continue_is_bare = rv(contains(ser(tc), "<w:vMerge/>"));     % bare (D-delta-1)
    tc = mat2doc.oxml.table.CT_Tc.new(); tc.vMerge = "restart"; tc.vMerge = [];
    g.none_has_vMerge = rv(contains(ser(tc), "<w:vMerge"));
    tc = mat2doc.oxml.table.CT_Tc.new(); tc.width = mat2doc.shared.Twips(1440);
    g.width_emu = rv(tc.width);
    probe.tc_getset = g;

    % ---- CT_TcPr direct + H11 child order ----
    p = struct();
    tp = mat2doc.oxml.OxmlElement("w:tcPr");
    p.default_grid_span = rv(tp.grid_span);
    p.default_vMerge_val = rv(tp.vMerge_val);
    p.default_width = rv(tp.width);
    p.default_vAlign_val = rv(tp.vAlign_val);
    center = parse_tcpr('<w:vAlign w:val="center"/>').vAlign_val;
    p.vAlign_parse_center = rv(center);
    % H11: set in SCRAMBLED order
    tp = mat2doc.oxml.OxmlElement("w:tcPr");
    tp.vMerge_val = "restart";
    tp.width = mat2doc.shared.Twips(1440);
    tp.vAlign_val = center;
    tp.grid_span = 2;
    kids = tp.xpath("./*");
    order = cell(1, numel(kids));
    for k = 1:numel(kids); order{k} = char(kids(k).local_part); end
    p.h11_child_order = order;
    p.h11_serialized = char(native2unicode(mat2doc.oxml.serialize_part_xml(tp), "UTF-8"));
    probe.tcpr = p;

    % ---- CT_Row un-stub ----
    r = struct();
    row = mat2doc.oxml.OxmlElement("w:tr");
    r.new_tc_class = simple_name(row.new_tc_());
    hrow = P("snip1.xml").xpath("./w:tr"); hrow = hrow(1);
    r.at0_class = simple_name(hrow.tc_at_grid_offset(0));
    r.at2_class = simple_name(hrow.tc_at_grid_offset(2));
    r.at1 = cap(@() hrow.tc_at_grid_offset(1));
    row2 = mat2doc.oxml.OxmlElement("w:tr");
    r.add_tc_class = simple_name(row2.add_tc());
    probe.row_unstub = r;

    % ---- raises: 6 span_dimensions + 2 swallow + 1 tr_above ----
    RAISE_PAIRS = { ...
        "snip1", [0 0], [1 0]; ...
        "snip2", [0 0], [0 1]; ...
        "snip4", [0 0], [1 0]; ...
        "snip5", [0 1], [1 0]; ...
        "snip6", [0 0], [1 1]; ...
        "snip2", [0 1], [0 2]};
    raises = cell(1, 9);
    for i = 1:size(RAISE_PAIRS, 1)
        key = RAISE_PAIRS{i,1}; a_rc = RAISE_PAIRS{i,2}; b_rc = RAISE_PAIRS{i,3};
        rows = P(char(key) + ".xml").xpath("./w:tr");
        a = rows(a_rc(1)+1).tc_at_grid_offset(a_rc(2));
        b = rows(b_rc(1)+1).tc_at_grid_offset(b_rc(2));
        rec = cap(@() a.span_dimensions_(b));
        rec.case = sprintf("%s_%d%d_%d%d", key, a_rc(1),a_rc(2),b_rc(1),b_rc(2));
        raises{i} = rec;
    end
    rowsA = P("swallow_a.xml").xpath("./w:tr");
    rec = cap(@() rowsA(1).tc_at_grid_offset(0).grow_to_(4,1));
    rec.case = "swallow_notenough"; raises{7} = rec;
    rowsB = P("swallow_b.xml").xpath("./w:tr");
    rec = cap(@() rowsB(1).tc_at_grid_offset(0).grow_to_(2,1));
    rec.case = "swallow_notrect"; raises{8} = rec;
    rec = cap(@() rows0(1).tc_at_grid_offset(0).tr_above_());
    rec.case = "tr_above_topmost"; raises{9} = rec;
    probe.raises = raises;
end

function rows = cell_geometry(tbl)
    trs = tbl.xpath("./w:tr");
    rows = cell(1, numel(trs));
    for ri = 1:numel(trs)
        tcs = trs(ri).tc_lst;
        cells = cell(1, numel(tcs));
        for ci = 1:numel(tcs)
            tc = tcs(ci);
            cells{ci} = struct( ...
                "grid_offset", rv(tc.grid_offset), "left", rv(tc.left), ...
                "right", rv(tc.right), "top", rv(tc.top), "bottom", rv(tc.bottom), ...
                "grid_span", rv(tc.grid_span), "vMerge", rv(tc.vMerge));
        end
        rows{ri} = cells;
    end
end

function n = simple_name(x)
    parts = split(string(class(x)), ".");
    n = parts(end);
end

function rec = cap(fn)
    try
        fn();
        rec = struct("raised", "False", "kind", "None", "msg", "None");
    catch me
        kind = extractAfter(string(me.identifier), "mat2doc:");
        if strlength(kind) == 0; kind = string(me.identifier); end
        rec = struct("raised", "True", "kind", kind, "msg", string(me.message));
    end
end

function s = rv(x)
    % Uniform canonicalizer (mirrors the s0064 rv): None->"None",
    % bool->"True"/"False", enum member NAME, Length/int->EMU/int decimal.
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

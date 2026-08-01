classdef Test_p6_4a_table_api < matlab.unittest.TestCase
% TEST_P6_4A_TABLE_API  Gate-4 permanent unit tests for Mat2Doc P6-4a [N] -- the
%   Table / _Rows / _Columns / _Row / _Column proxy API + add_table un-stub.
%   src/docx/table.py::{Table,_Columns,_Column,_Rows,_Row} -> +mat2doc\+table\
%   {Table, Columns_, Column_, Rows_, Row_}; and the add_table / iter_inner_content
%   un-stubs in src/docx/blkcntnr.py -> +mat2doc\BlockItemContainer.m,
%   src/docx/document.py -> +mat2doc\+document\Document.m, and the w:tbl branch of
%   src/docx/section.py::Section.iter_inner_content -> +mat2doc\+section\Section.m.
%
%   THIS IS THE FIRST END-TO-END TABLE the toolbox can create: doc.add_table(r,c)
%   reaches CT_Tbl.new_tbl (byte-proven P6-3b), inserts it into the body, and wraps
%   it in a live Table proxy. The public authoring path is exercised here as a
%   FULL-PACKAGE byte oracle -- every one of the 17 parts a real Word .docx carries
%   is asserted byte-identical to python-docx 1.2.0. The proxy read-tier
%   (rows/columns/getitem_/width/height/alignment/...) is frozen value-identical to
%   the frozen probe. This class is the permanent regression net for both.
%
%   THE HIGHEST-VALUE PERMANENT PINS IN THIS WP (a regression goes RED loudly):
%     * ★ (ADD_TABLE FULL-PACKAGE) test_add_table_full_package_byte_pins -- THE FIRST
%       END-TO-END TABLE. For each of the 5 frozen add_table scenarios the MATLAB
%       Document is built through the IDENTICAL public API, saved, unzipped, and
%       compared to the frozen python-docx package: word/document.xml raw uint8
%       byte-identical to the shipped frozen part AND all 17 parts size+SHA-256 ==
%       the frozen manifest. Covers the headline add_table(2,3) (document.xml 2256 B
%       / a1eda043..., styles.xml UNCHANGED 349458 B / 02d71a68...), the with-STYLE
%       add_table(2,2,"My Table Style") (document.xml gains <w:tblStyle
%       w:val="MyTableStyle"/> 2131 B / 6b568902..., styles.xml GAINS the style
%       349567 B / 593e27a4...), the explicit-width _body.add_table(2,3,Inches(2))
%       (2247 B / 1e6e66a1...), two-tables (2565 B / fd4ef864...), and table+paragraph
%       order (2300 B / f6a8d83e...). If a single byte of the table-authoring path
%       drifts, this goes RED. (frozen s0067..s0071, Gate-3 17/17 each.)
%     * ★ (M1) test_m1_byte_pins -- add_table un-stub NEUTRALITY: a bare
%       mat2doc.Document().save() STILL emits word/styles.xml at EXACTLY 349458 B /
%       02d71a68... AND word/document.xml at 1548 B / 0e4dd503... (un-stubbing
%       add_table/iter_inner_content touches no default-save byte). L1.
%     * ★ (A2 ALIGNMENT) test_table_alignment -- Table.alignment is the ratified A2
%       cross-enum: the getter returns a WD_PARAGRAPH_ALIGNMENT member VERBATIM
%       (WD_ALIGN_PARAGRAPH is an alias of that canonical enum, NOT the runtime
%       class), so it is compared by NAME ("CENTER") and .value (1), NEVER by
%       cross-class == (which diverges from Python's int-subclass duck-equality --
%       the documented consequence). set writes <w:jc w:val="center"/>; None removes
%       it. NOT a D-number (bytes/name/value identical).
%     * (COLLECTION SURFACE) test_columns / test_rows -- _Columns/_Column and
%       _Rows/_Row ported as EXPLICIT getitem_/to_array/len_ (the Sections/TabStops
%       VERIFY-COLLECTION precedent; the shared 1-based () base is a future WP).
%       0-based getitem_ + negative wrap; the VERBATIM out-of-range identifiers and
%       messages ("column index [3]/[-9] is out of range" with the ORIGINAL idx;
%       "list index out of range"); width/height/height_rule get/set + None;
%       _index (0-based DATA); grid_cols_before/after; the CPython slice port.
%     * (BOUNDARY) test_p6_4b_stubs -- the 8 P6-4b-owned members (add_column,
%       add_row, cell, column_cells, row_cells, _cells, _Row.cells, _Column.cells)
%       each raise the IDENTIFIER mat2doc:notYetPorted (need _Cell, P6-4b) -- never
%       a silent no-op.
%     * (ITER_INNER_CONTENT) test_iter_inner_content / test_section_iter_inner_content
%       -- a paragraph/table/paragraph document yields [Paragraph, Table, Paragraph]
%       in document order; a table in a section yields a Table (the C2 debt
%       discharged: the w:tbl branch no longer raises notYetPorted).
%
%   Provenance (Gate-1..3, all 2026-07-31):
%     * Audit    : validation\mat2doc\audit_P6-4a_table_api.md (Porter Gate-1
%                  self-probe 52/52 + mso-auditor Gate-2 APPROVE -- 12/12 add_table
%                  byte scenarios all full-package L1, incl. the s10 style-write and
%                  s7 property-sweep; ZERO new D-numbers).
%     * Validate : Gate-3 PASS -- 79/81 targeted regression (the 2 EXPECTED
%                  registry-flip stale-pins re-pinned at Gate-4); the add_table
%                  full-package oracles frozen at references\s0067..s0071 (17/17
%                  each) and the Table-API probe references\s0072\probe.json; M1
%                  17/17 preserved; ZERO new D-numbers.
%     * Scenarios: validation\mat2doc\scenarios\s0067..s0071_p6_4a_*_gscenario.m
%                  (the add_table byte twins -- their build bodies replayed VERBATIM
%                  by the build helpers below); s0072_p6_4a_table_api_probe.m (the
%                  full proxy-surface probe -- runProbe() below replays its body
%                  VERBATIM for the Equivalence leg).
%     * Frozen refs (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE at Gate-3):
%         references\s0067..s0071\ (manifest.json + word/document.xml [+ s0068
%           word/styles.xml]) copied verbatim into tests\table\data\s00XX\ WITH a
%           co-located `.gitattributes` `* binary` pin (frozen-byte fixtures must not
%           be line-ending mangled on the master checkout -- the Gate-4 byte-fixture
%           lesson). The manifest SHAs ARE the python-docx package part SHAs.
%         references\s0072\probe.json copied into tests\table\data\
%           s0072_probe_oracle.json (value JSON; jsondecode is line-ending agnostic).
%         references\s0001\parts\word\{styles,document}.xml -- the M1 byte references
%           (SHA of what Document().save() emits); NOT copied.
%
%   Coverage taxonomy
%   -----------------
%   * Nominal      -- add_table returns a Table; alignment/autofit/style/
%                     table_direction get; rows/columns (Rows_/Columns_, lazy-cached);
%                     column_count_/tblPr_; getitem_ int; to_array; len_; width/height
%                     get; grid_cols_*; index_.
%   * Edge         -- H3 None ([]) round-trips (alignment/table_direction/width/
%                     height/height_rule); NEGATIVE getitem_ wrap; the CPython slice;
%                     the mat2doc:IndexError error paths (VERBATIM identifier + the
%                     ORIGINAL-idx message); the 8 P6-4b notYetPorted stubs; single-
%                     row/column tables (add_table(1,1)); non-ASCII paragraph text
%                     through iter_inner_content order.
%   * Equivalence  -- test_equivalence_probe_vs_frozen_oracle replays the ENTIRE s0072
%                     probe (table_props / columns / rows / iter_inner_content) and
%                     section-compares every value to the frozen python-docx oracle
%                     (Gate-3 probe_diff was exit 0).
%   * Regression   -- hard-coded expected property values (EMU/name) + the 5 add_table
%                     full-package (document.xml raw uint8 + 17-part SHA) pins + the
%                     M1 styles/document part SHAs + the tblPr serialized-XML strings.
%   * Upstream     -- the frozen document.xml/styles.xml ARE python-docx add_table
%                     output; the "column index [%d] is out of range" / "list index
%                     out of range" messages and the CPython slice.indices semantics
%                     ARE the python-docx table.py contract.
%
%   Byte-level (L1) note: every add_table document.xml comparison is BOTH a raw uint8
%   byte-equality AND a SHA-256 equality vs the frozen oracle, and every one of the
%   17 package parts is SHA-256-pinned to the frozen manifest -- the ladder demanded
%   L1 and Gate-3 delivered 17/17 with ZERO new D-numbers, so every byte pin here is
%   L1. The tblPr serialized-XML string checks are ASCII (string-equality ==
%   byte-identical). The Equivalence section's jsondecode(jsonencode(...)) value
%   comparison and its non-triviality floor are the only looser-than-byte checks and
%   are commented at their sites. The mat2doc:notYetPorted / mat2doc:IndexError
%   identifiers equal the Python exception-class mapping (design.md section 2 --
%   non-byte, non-output, NOT a D-number).
%
%   Determinism: no network, no absolute paths -- the worktree root and the
%   co-located s0067..s0072 fixtures resolve relative to this file via
%   fileparts(mfilename('fullpath')); saves go to tempname .docx / tempname dirs
%   deleted via onCleanup; every fixture read is binary ('r','n'). The +mat2doc
%   package resolves via the MANDATORY PathFixture(worktree-root) in TestClassSetup
%   (WP9-F4 lesson).

    properties (Constant)
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        % --- proxy classes under test ---
        TABLE       = 'mat2doc.table.Table'
        COLUMNS_    = 'mat2doc.table.Columns_'
        COLUMN_     = 'mat2doc.table.Column_'
        ROWS_       = 'mat2doc.table.Rows_'
        ROW_        = 'mat2doc.table.Row_'
        PARAGRAPH   = 'mat2doc.text.Paragraph'
        TABLE_STYLE = 'mat2doc.styles.TableStyle_'

        % --- frozen s0001 M1 byte references (add_table un-stub neutrality guard) ---
        STYLES_SIZE_M1 = 349458
        STYLES_SHA_M1  = "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"
        DOC_SIZE_M1    = 1548
        DOC_SHA_M1     = "0e4dd503bc095219cd24e8312b372c588ca56bb834aacac53eafb7bf47836327"

        % --- headline s0067 document.xml SHA (public add_table(2,3)) ---
        S0067_DOC_SHA  = "a1eda0439dbf02fb02c109ca218c5c8b0fbe7136a9e54022cd0f28a4ea1820bf"

        % --- regression property values (frozen s0072 probe, EMU) ---
        WIDTH_PER_COL_IN6 = 1828800   % Inches(6)//3 == 5486400//3
        WIDTH_IN1P5       = 1371600   % Inches(1.5)
        HEIGHT_PT20       = 254000    % Pt(20)
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied from tests\section\Test_p5_3a_sections_api.m. here is
            % tests\table; the worktree root is two levels up.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\table
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. ★ add_table FULL-PACKAGE byte pins (s0067..s0071)             %
        % =============================================================== %

        function test_add_table_full_package_byte_pins(testCase)
            % ★ (ADD_TABLE FULL-PACKAGE) Regression + Upstream (byte-identical L1) --
            % THE FIRST END-TO-END TABLE and the headline permanent pin. For each of
            % the 5 frozen scenarios: build a mat2doc.Document through the IDENTICAL
            % public API, save(), unzip, and assert
            %   (a) word/document.xml is raw-uint8 byte-identical to the shipped
            %       frozen part (the part the table actually writes), AND
            %   (b) all 17 package parts are size+SHA-256 == the frozen python-docx
            %       manifest (genuine full-package 17/17).
            % s0067 (the public add_table(2,3)) is checked FIRST and its document.xml
            % SHA is additionally hard-pinned to the headline constant. Gate-3 froze
            % 17/17 byte-identical per scenario with ZERO new D-numbers.
            here = fileparts(mfilename('fullpath'));
            scenarios = { ...
                's0067', @buildS0067; ...   % public add_table(2,3)
                's0068', @buildS0068; ...   % with-STYLE add_table(2,2,"My Table Style")
                's0069', @buildS0069; ...   % explicit width _body.add_table(2,3,Inches(2))
                's0070', @buildS0070; ...   % two tables
                's0071', @buildS0071 };     % table + paragraph order
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
                    sprintf('%s: add_table word/document.xml must be BYTE-IDENTICAL to the frozen python-docx part', name));

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

            % headline s0067 document.xml SHA hard-pin (loudest single guard)
            d0 = buildS0067();
            doc0 = saveAndUnzipParts(d0); doc0 = doc0('word/document.xml');
            testCase.verifyEqual(sha256hex(doc0), testCase.S0067_DOC_SHA, ...
                'headline: add_table(2,3) document.xml SHA-256 == frozen s0067 oracle a1eda043... (L1)');
        end

        % =============================================================== %
        % 2. ★ M1 byte-pins (add_table un-stub neutrality guard)          %
        % =============================================================== %

        function test_m1_byte_pins(testCase)
            % ★ (M1) Regression (byte-neutrality, L1): un-stubbing add_table /
            % iter_inner_content touches NO default-save byte -- a bare
            % mat2doc.Document().save() STILL emits word/styles.xml at EXACTLY 349458
            % B / 02d71a68... AND word/document.xml at 1548 B / 0e4dd503... (frozen
            % s0001 SHAs). A single save() emits both parts. SHA == L1.
            parts = saveAndUnzipParts(mat2doc.Document());

            sty = parts('word/styles.xml');
            testCase.verifyEqual(numel(sty), testCase.STYLES_SIZE_M1, ...
                sprintf('bare save word/styles.xml must be exactly %d B', testCase.STYLES_SIZE_M1));
            testCase.verifyEqual(sha256hex(sty), testCase.STYLES_SHA_M1, ...
                'M1 word/styles.xml SHA-256 unchanged (add_table un-stub neutrality, L1)');

            doc = parts('word/document.xml');
            testCase.verifyEqual(numel(doc), testCase.DOC_SIZE_M1, ...
                sprintf('bare save word/document.xml must be exactly %d B', testCase.DOC_SIZE_M1));
            testCase.verifyEqual(sha256hex(doc), testCase.DOC_SHA_M1, ...
                'M1 word/document.xml SHA-256 unchanged (add_table un-stub neutrality, L1)');
        end

        % =============================================================== %
        % 3. Table -- alignment (★ A2 cross-enum), autofit, direction      %
        % =============================================================== %

        function test_table_alignment(testCase)
            % ★ (A2 ALIGNMENT) Nominal + Edge (table.py 57-69): the getter returns a
            % WD_PARAGRAPH_ALIGNMENT member VERBATIM (A2 -- python-docx's cast is a
            % runtime no-op). Compare by NAME and .value, NEVER cross-class == (which
            % diverges from Python's int-subclass duck-equality). set writes
            % <w:jc w:val="center"/>; None ([]) removes it.
            WTA = @(n) mat2doc.enum.table.WD_TABLE_ALIGNMENT.(n);

            t = mat2doc.Document().add_table(2, 3);
            testCase.verifyTrue(isequal(t.alignment, []), 'default alignment -> [] (None)');

            % --- set CENTER: assert BY NAME + .value, never cross-class == ---
            t.alignment = WTA("CENTER");
            testCase.verifyEqual(string(t.alignment), "CENTER", ...
                'A2: alignment compared by NAME ("CENTER"), not cross-class ==');
            testCase.verifyEqual(double(t.alignment.value), 1, ...
                'A2: alignment .value == 1 (WD_PARAGRAPH_ALIGNMENT.CENTER)');
            testCase.verifyTrue(contains(ser(t.tblPr_()), '<w:jc w:val="center"/>'), ...
                'set CENTER writes <w:jc w:val="center"/> into tblPr (byte)');

            % --- None removes the w:jc (H3) ---
            t.alignment = [];
            testCase.verifyTrue(isequal(t.alignment, []), 'alignment = [] -> [] (None)');
            testCase.verifyFalse(contains(ser(t.tblPr_()), "<w:jc"), ...
                'alignment = [] removes the <w:jc> element');
        end

        function test_table_autofit(testCase)
            % Nominal + Edge (table.py 71-83): default autofit True (no fixed layout);
            % set false writes <w:tblLayout w:type="fixed"/>; set true -> "autofit".
            t = mat2doc.Document().add_table(2, 3);
            testCase.verifyTrue(t.autofit, 'default autofit True');

            t.autofit = false;
            testCase.verifyFalse(t.autofit, 'autofit = false round-trip');
            testCase.verifyTrue(contains(ser(t.tblPr_()), '<w:tblLayout w:type="fixed"/>'), ...
                'autofit = false writes <w:tblLayout w:type="fixed"/> (byte)');

            t.autofit = true;
            testCase.verifyTrue(t.autofit, 'autofit = true round-trip');
        end

        function test_table_direction(testCase)
            % Nominal + Edge (table.py 150-161, WD_TABLE_DIRECTION -- a REAL distinct
            % enum, NO A2): default []; set RTL -> logical true + BARE <w:bidiVisual/>
            % (the True default deletes @w:val, D-delta-1); set LTR -> logical false +
            % <w:bidiVisual w:val="0"/>; set [] -> [] + element removed. The getter
            % returns a logical (CT_OnOff.val) verbatim -- Python's cast is a no-op.
            WTD = @(n) mat2doc.enum.table.WD_TABLE_DIRECTION.(n);

            t = mat2doc.Document().add_table(2, 3);
            testCase.verifyTrue(isequal(t.table_direction, []), 'default table_direction -> [] (None)');

            t.table_direction = WTD("RTL");
            testCase.verifyEqual(t.table_direction, true, 'RTL -> table_direction logical true');
            sRTL = ser(t.tblPr_());
            testCase.verifyTrue(contains(sRTL, "<w:bidiVisual/>"), ...
                'RTL -> BARE <w:bidiVisual/> (True default deletes @w:val)');
            testCase.verifyFalse(contains(sRTL, "w:bidiVisual w:val"), 'RTL bidiVisual carries no @w:val');

            t.table_direction = WTD("LTR");
            testCase.verifyEqual(t.table_direction, false, 'LTR -> table_direction logical false');
            testCase.verifyTrue(contains(ser(t.tblPr_()), '<w:bidiVisual w:val="0"/>'), ...
                'LTR -> <w:bidiVisual w:val="0"/> (byte)');

            t.table_direction = [];
            testCase.verifyTrue(isequal(t.table_direction, []), 'table_direction = [] -> [] (None)');
            testCase.verifyFalse(contains(ser(t.tblPr_()), "<w:bidiVisual"), ...
                'table_direction = [] removes the <w:bidiVisual> element');
        end

        % =============================================================== %
        % 4. Table -- style (get TableStyle_ / set tblStyle_val byte)       %
        % =============================================================== %

        function test_table_style(testCase)
            % Nominal + Edge (table.py 119-138): default style -> the "Normal Table"
            % TableStyle_ (no <w:tblStyle> directly applied); after add_style +
            % set "My Table Style" the tblPr gains <w:tblStyle w:val="MyTableStyle"/>
            % (the style id strips spaces) and style.name == "My Table Style"; None
            % ([]) removes the directly-applied style.
            d = mat2doc.Document();
            t = d.add_table(2, 2);

            st = t.style;
            testCase.verifyClass(st, testCase.TABLE_STYLE, 'default style -> a TableStyle_');
            testCase.verifyEqual(string(st.name), "Normal Table", 'default style name "Normal Table"');
            testCase.verifyFalse(contains(ser(t.tblPr_()), "<w:tblStyle"), ...
                'default table applies NO <w:tblStyle> directly');

            d.styles.add_style("My Table Style", mat2doc.enum.style.WD_STYLE_TYPE.TABLE);
            t.style = "My Table Style";
            testCase.verifyEqual(string(t.style.name), "My Table Style", ...
                'set style -> style.name "My Table Style"');
            % the style id strips spaces -> <w:tblStyle w:val="MyTableStyle"/> (byte);
            % Table exposes no tblStyle_val accessor, so the serialized tblPr IS the pin.
            testCase.verifyTrue(contains(ser(t.tblPr_()), '<w:tblStyle w:val="MyTableStyle"/>'), ...
                'set style writes <w:tblStyle w:val="MyTableStyle"/> (id "MyTableStyle", spaces stripped) (byte)');

            t.style = [];
            testCase.verifyFalse(contains(ser(t.tblPr_()), "<w:tblStyle"), ...
                'style = [] removes the directly-applied <w:tblStyle>');
        end

        % =============================================================== %
        % 5. Table -- rows/columns/column_count + table terminus + cache   %
        % =============================================================== %

        function test_table_rows_columns_count(testCase)
            % Nominal (table.py 99-117, 140-148, 182-189): rows -> Rows_ (len 2),
            % columns -> Columns_ (len 3), column_count_ == 3; table returns self
            % (identity terminus); rows/columns are @lazyproperty cached (same handle
            % on repeat). tblPr_ returns a CT_TblPr.
            t = mat2doc.Document().add_table(2, 3);

            testCase.verifyClass(t.rows, testCase.ROWS_, 'rows -> a Rows_');
            testCase.verifyClass(t.columns, testCase.COLUMNS_, 'columns -> a Columns_');
            testCase.verifyEqual(t.rows.len_(), 2, 'rows.len_ == 2');
            testCase.verifyEqual(t.columns.len_(), 3, 'columns.len_ == 3');
            testCase.verifyEqual(t.column_count_(), 3, 'column_count_ == 3');

            testCase.verifyTrue(t.table == t, 'table returns self (identity terminus)');
            testCase.verifyTrue(t.rows == t.rows, 'rows is @lazyproperty cached (same handle)');
            testCase.verifyTrue(t.columns == t.columns, 'columns is @lazyproperty cached (same handle)');
            testCase.verifyClass(t.tblPr_(), 'mat2doc.oxml.table.CT_TblPr', 'tblPr_ -> a CT_TblPr');
        end

        function test_table_add_row_column_stub(testCase)
            % Edge / Error path (table.py 37-55): the mutators add_row / add_column
            % raise mat2doc:notYetPorted (owner P6-4b) -- never a silent no-op.
            t = mat2doc.Document().add_table(2, 3);
            testCase.verifyError(@() t.add_row(), 'mat2doc:notYetPorted', ...
                'Table.add_row -> mat2doc:notYetPorted (P6-4b)');
            testCase.verifyError(@() t.add_column(mat2doc.shared.Inches(1)), 'mat2doc:notYetPorted', ...
                'Table.add_column -> mat2doc:notYetPorted (P6-4b)');
        end

        % =============================================================== %
        % 6. _Columns / _Column -- getitem_/iter/len_/width/index/errors    %
        % =============================================================== %

        function test_columns(testCase)
            % Nominal + Edge (table.py 314-384): a fresh 2x3 Inches(6) table.
            %   len_ 3; to_array 3 Column_ with 0-based index_ [0 1 2] and per-col
            %   width Inches(6)//3 == 1828800; getitem_(0)/(-1) (negative wrap);
            %   width set Inches(1.5)==1371600 / set [] -> [] (H3); table() terminus;
            %   OUT-OF-RANGE -> mat2doc:IndexError with the VERBATIM ORIGINAL-idx
            %   message.
            IN = @(v) mat2doc.shared.Inches(v);
            t = mat2doc.Document().add_table(2, 3);
            cols = t.columns;

            testCase.verifyEqual(cols.len_(), 3, 'columns.len_ == 3');
            arr = cols.to_array();
            testCase.verifyEqual(numel(arr), 3, 'to_array 3 Column_');
            testCase.verifyClass(arr(1), testCase.COLUMN_, 'to_array entries are Column_');
            testCase.verifyEqual(arrayfun(@(c) c.index_(), arr), [0 1 2], ...
                'Column_.index_ is 0-based grid DATA [0 1 2]');
            testCase.verifyEqual(arrayfun(@(c) double(c.width), arr), ...
                testCase.WIDTH_PER_COL_IN6 * [1 1 1], ...
                'per-column width == Inches(6)//3 == 1828800 EMU');

            % 0-based getitem_ + negative wrap
            testCase.verifyEqual(double(cols.getitem_(0).width), testCase.WIDTH_PER_COL_IN6, ...
                'getitem_(0).width == 1828800');
            testCase.verifyEqual(cols.getitem_(-1).index_(), 2, ...
                'getitem_(-1) is the LAST column (negative wrap) -> index_ 2');

            % table() terminus (property-as-method)
            testCase.verifyTrue(cols.getitem_(0).table == t, 'Column_.table() is the owning Table');
            testCase.verifyTrue(cols.table == t, 'Columns_.table() is the owning Table');

            % width set / None (H3) -- fresh proxies read the shared element
            c1 = cols.getitem_(1); c1.width = IN(1.5);
            testCase.verifyEqual(double(cols.getitem_(1).width), testCase.WIDTH_IN1P5, ...
                'set width Inches(1.5) -> 1371600 EMU readback');
            c1b = cols.getitem_(1); c1b.width = [];
            testCase.verifyTrue(isequal(cols.getitem_(1).width, []), 'set width [] -> [] (None, H3)');

            % out-of-range -> mat2doc:IndexError, VERBATIM ORIGINAL-idx message
            MEp = captureError(@() cols.getitem_(3));
            testCase.verifyEqual(string(MEp.identifier), "mat2doc:IndexError", ...
                'columns.getitem_(3) -> mat2doc:IndexError');
            testCase.verifyEqual(string(MEp.message), "column index [3] is out of range", ...
                'IndexError message formats the ORIGINAL idx [3]');
            MEn = captureError(@() cols.getitem_(-9));
            testCase.verifyEqual(string(MEn.identifier), "mat2doc:IndexError", ...
                'columns.getitem_(-9) -> mat2doc:IndexError');
            testCase.verifyEqual(string(MEn.message), "column index [-9] is out of range", ...
                'IndexError message formats the ORIGINAL idx [-9] (NOT the wrapped value)');
        end

        % =============================================================== %
        % 7. _Rows / _Row -- getitem_/iter/len_/height/rule/grid/index/slice %
        % =============================================================== %

        function test_rows(testCase)
            % Nominal + Edge (table.py 387-537): a fresh 2x3 Inches(6) table.
            %   len_ 2; to_array 2 Row_ with 0-based index_ [0 1] and default height
            %   [] (None); getitem_(0)/(-1) (negative wrap); grid_cols_before/after 0;
            %   height set Pt(20)==254000 / set [] -> [] (H3); height_rule EXACTLY
            %   (WD_ROW_HEIGHT_RULE) / default []; the CPython slice rows[1:2] -> the
            %   single last row; OUT-OF-RANGE -> mat2doc:IndexError "list index out of
            %   range".
            PT  = @(v) mat2doc.shared.Pt(v);
            WRH = @(n) mat2doc.enum.table.WD_ROW_HEIGHT_RULE.(n);
            t = mat2doc.Document().add_table(2, 3);
            rows = t.rows;

            testCase.verifyEqual(rows.len_(), 2, 'rows.len_ == 2');
            arr = rows.to_array();
            testCase.verifyEqual(numel(arr), 2, 'to_array 2 Row_');
            testCase.verifyClass(arr(1), testCase.ROW_, 'to_array entries are Row_');
            testCase.verifyEqual(arrayfun(@(r) r.index_(), arr), [0 1], ...
                'Row_.index_ is 0-based row DATA [0 1]');
            testCase.verifyTrue(all(arrayfun(@(r) isequal(r.height, []), arr)), ...
                'default row height is [] (None) on both rows');

            r0 = rows.getitem_(0);
            testCase.verifyEqual(r0.grid_cols_before, 0, 'grid_cols_before 0');
            testCase.verifyEqual(r0.grid_cols_after, 0, 'grid_cols_after 0');
            testCase.verifyTrue(isequal(r0.height_rule, []), 'default height_rule [] (None)');
            testCase.verifyEqual(rows.getitem_(-1).index_(), 1, ...
                'getitem_(-1) is the LAST row (negative wrap) -> index_ 1');
            testCase.verifyTrue(rows.getitem_(0).table == t, 'Row_.table() is the owning Table');
            testCase.verifyTrue(rows.table == t, 'Rows_.table() is the owning Table');

            % height / height_rule set + None (H3) -- fresh proxies, shared element
            ra = rows.getitem_(0); ra.height = PT(20);
            testCase.verifyEqual(double(rows.getitem_(0).height), testCase.HEIGHT_PT20, ...
                'set height Pt(20) -> 254000 EMU readback');
            rb = rows.getitem_(0); rb.height_rule = WRH("EXACTLY");
            testCase.verifyEqual(string(rows.getitem_(0).height_rule), "EXACTLY", ...
                'set height_rule EXACTLY (WD_ROW_HEIGHT_RULE) readback by name');
            rc = rows.getitem_(0); rc.height = [];
            testCase.verifyTrue(isequal(rows.getitem_(0).height, []), 'set height [] -> [] (None, H3)');

            % CPython slice rows[1:2] -> the single last row (index_ 1)
            sl = rows.getitem_(struct('start', 1, 'stop', 2, 'step', []));
            testCase.verifyEqual(numel(sl), 1, 'slice rows[1:2] has one row');
            testCase.verifyEqual(sl(1).index_(), 1, 'slice rows[1:2] is the row at index_ 1');

            % out-of-range -> mat2doc:IndexError "list index out of range"
            MEp = captureError(@() rows.getitem_(2));
            testCase.verifyEqual(string(MEp.identifier), "mat2doc:IndexError", ...
                'rows.getitem_(2) -> mat2doc:IndexError');
            testCase.verifyEqual(string(MEp.message), "list index out of range", ...
                'Rows_ IndexError verbatim message');
            MEn = captureError(@() rows.getitem_(-3));
            testCase.verifyEqual(string(MEn.identifier), "mat2doc:IndexError", ...
                'rows.getitem_(-3) -> mat2doc:IndexError');
        end

        % =============================================================== %
        % 8. iter_inner_content -- Paragraph|Table interleaved order        %
        % =============================================================== %

        function test_iter_inner_content(testCase)
            % Nominal + Edge (blkcntnr.py 74-79): a document with a paragraph, a
            % table, then a paragraph yields [Paragraph, Table, Paragraph] in DOCUMENT
            % order (a heterogeneous 1xN cell). Non-ASCII paragraph text survives.
            d = mat2doc.Document();
            d.add_paragraph("above-caf" + string(char(233)));   % "above-café" (U+00E9)
            d.add_table(2, 3);
            d.add_paragraph("below table");
            items = d.iter_inner_content();
            testCase.verifyClass(items, 'cell', 'iter_inner_content -> a cell array');
            testCase.verifyEqual(numel(items), 3, 'three inner blocks');
            testCase.verifyClass(items{1}, testCase.PARAGRAPH, 'block 1 -> Paragraph');
            testCase.verifyClass(items{2}, testCase.TABLE,     'block 2 -> Table');
            testCase.verifyClass(items{3}, testCase.PARAGRAPH, 'block 3 -> Paragraph');
            testCase.verifyEqual(string(items{1}.text), "above-caf" + string(char(233)), ...
                'Paragraph text survives non-ASCII (é)');
        end

        function test_section_iter_inner_content(testCase)
            % C2 (section.py 157-163): a table in a section yields a Table -- the
            % w:tbl branch no longer raises notYetPorted (the C2 debt is discharged).
            d = mat2doc.Document();
            d.add_table(2, 3);
            items = d.sections.getitem_(0).iter_inner_content();
            testCase.verifyEqual(numel(items), 1, 'the section has one inner block (the table)');
            testCase.verifyClass(items{1}, testCase.TABLE, ...
                'Section.iter_inner_content wraps the w:tbl as a Table (C2 discharged)');
        end

        % =============================================================== %
        % 9. P6-4b boundary -- the 8 notYetPorted stubs                     %
        % =============================================================== %

        function test_p6_4b_stubs(testCase)
            % Edge / Error path (P6-4b boundary): the 8 members that need _Cell (not
            % ported this WP) each raise the IDENTIFIER mat2doc:notYetPorted. Covers
            % Table.{add_column, add_row, cell, column_cells, row_cells, _cells} and
            % _Row.cells / _Column.cells.
            IN = @(v) mat2doc.shared.Inches(v);
            t = mat2doc.Document().add_table(2, 3);
            r0 = t.rows.getitem_(0);
            c0 = t.columns.getitem_(0);
            calls = { ...
                'Table.add_column',  @() t.add_column(IN(1)); ...
                'Table.add_row',     @() t.add_row(); ...
                'Table.cell',        @() t.cell(0, 0); ...
                'Table.column_cells',@() t.column_cells(0); ...
                'Table.row_cells',   @() t.row_cells(0); ...
                'Table._cells',      @() t.cells_(); ...
                '_Row.cells',        @() r0.cells(); ...
                '_Column.cells',     @() c0.cells() };
            testCase.verifyEqual(size(calls, 1), 8, 'the P6-4b stub battery covers exactly 8 members');
            for k = 1:size(calls, 1)
                ME = captureError(calls{k, 2});
                testCase.verifyEqual(string(ME.identifier), "mat2doc:notYetPorted", ...
                    sprintf('stub %s must raise mat2doc:notYetPorted (P6-4b)', calls{k, 1}));
            end
        end

        % =============================================================== %
        % 10. EQUIVALENCE -- full s0072 probe vs the frozen oracle          %
        % =============================================================== %

        function test_equivalence_probe_vs_frozen_oracle(testCase)
            % Equivalence: replay the ENTIRE s0072 probe (runProbe -- the .m twin's
            % body VERBATIM: table_props / columns / rows / iter_inner_content) and
            % compare EACH section to the frozen python-docx 1.2.0 oracle copied into
            % data\s0072_probe_oracle.json. Gate-3 probe_diff was exit 0, so every
            % section must be value-identical.
            %
            % Both sides are normalized through jsondecode(jsonencode(...)) so struct
            % vs cell-of-string container shapes collapse the SAME way (the .m twin
            % writes the port with jsonencode; the oracle came from the identical JSON
            % shape) -- the only looser-than-byte comparison in this class (values,
            % not bytes), justified because probe_diff already proved value
            % equivalence at Gate-3.
            here = fileparts(mfilename('fullpath'));
            port    = runProbe();
            portN   = jsondecode(jsonencode(port));               % normalize shapes
            oracleN = jsondecode(native2unicode( ...
                readBytes(fullfile(here, 'data', 's0072_probe_oracle.json')), 'UTF-8'));

            % Non-triviality floor (guards a silent-empty replay).
            testCase.verifyEqual(sort(fieldnames(portN)), sort(fieldnames(oracleN)), ...
                'the replayed probe and the frozen oracle expose the same top-level sections');
            testCase.verifyGreaterThanOrEqual(numel(fieldnames(oracleN)), 4, ...
                'the oracle must expose all 4 probe sections');

            secs = fieldnames(oracleN);
            for i = 1:numel(secs)
                s = secs{i};
                testCase.verifyTrue(isfield(portN, s), sprintf('port is missing section %s', s));
                testCase.verifyTrue(isequaln(portN.(s), oracleN.(s)), ...
                    sprintf('section "%s" must be value-identical to the frozen s0072 oracle', s));
            end
        end

    end
end

% ===================== add_table build twins (VERBATIM) ================ %

function d = buildS0067()
    % s0067: public add_table(2, 3) (width = _block_width). VERBATIM s0067 twin.
    d = mat2doc.Document();
    d.add_table(2, 3);
end

function d = buildS0068()
    % s0068: with-STYLE add_table(2, 2, "My Table Style"). VERBATIM s0068 twin.
    d = mat2doc.Document();
    d.styles.add_style("My Table Style", mat2doc.enum.style.WD_STYLE_TYPE.TABLE);
    d.add_table(2, 2, "My Table Style");
end

function d = buildS0069()
    % s0069: explicit width _body.add_table(2, 3, Inches(2)). VERBATIM s0069 twin.
    d = mat2doc.Document();
    d.body_().add_table(2, 3, mat2doc.shared.Inches(2));
end

function d = buildS0070()
    % s0070: two tables add_table(2,3); add_table(1,1). VERBATIM s0070 twin.
    d = mat2doc.Document();
    d.add_table(2, 3);
    d.add_table(1, 1);
end

function d = buildS0071()
    % s0071: table then paragraph. VERBATIM s0071 twin.
    d = mat2doc.Document();
    d.add_table(2, 3);
    d.add_paragraph("below table");
end

% ===================== file-local helpers ============================== %

function s = ser(e)
    % Decode serialize_part_xml bytes to a string (ASCII content -> string-equality
    % is a byte-identical L1 assertion; non-ASCII round-trips via UTF-8).
    s = string(native2unicode(mat2doc.oxml.serialize_part_xml(e), "UTF-8"));
end

function parts = saveAndUnzipParts(d)
    % Document.save() to a temp .docx, unzip once, return a containers.Map of every
    % package part: relative POSIX path (e.g. 'word/document.xml',
    % '[Content_Types].xml') -> raw uint8 bytes. Both temp artifacts cleaned on exit.
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

function b = readBytes(p)
    f = fopen(p, 'r', 'n');            % binary read (no CRLF translation)
    assert(f >= 0, 'could not open for read: %s', p);
    b = fread(f, Inf, '*uint8')';
    fclose(f);
end

function deleteIfExists(p)
    if isfile(p), delete(p); end
end

function rmdirIfExists(p)
    if isfolder(p), rmdir(p, 's'); end
end

function h = sha256hex(bytes)
    % Lowercase hex SHA-256 of a uint8 vector (base MATLAB / Java MessageDigest;
    % matches the python hashlib manifest SHAs).
    md  = java.security.MessageDigest.getInstance('SHA-256');
    dig = md.digest(typecast(uint8(bytes(:)'), 'int8'));
    b   = typecast(dig, 'uint8');
    h   = lower(string(sprintf('%02x', b)));
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

% ---- s0072 probe replay (the .m twin body, VERBATIM) ------------------ %

function P = runProbe()
    % Replay the s0072 Table-API probe (the .m twin's P-struct body, VERBATIM
    % tags/inputs/order) and return the nested struct of tagged canonical values.
    % Embedded so the Equivalence leg is self-contained (the validation-folder
    % scenario is NOT on the toolbox path and must not be a dependency). Mirrors
    % validation\mat2doc\scenarios\s0072_p6_4a_table_api_probe.m lines 27-118.
    IN  = @(v) mat2doc.shared.Inches(v);
    PT  = @(v) mat2doc.shared.Pt(v);
    WTA = @(n) mat2doc.enum.table.WD_TABLE_ALIGNMENT.(n);
    WTD = @(n) mat2doc.enum.table.WD_TABLE_DIRECTION.(n);
    WRH = @(n) mat2doc.enum.table.WD_ROW_HEIGHT_RULE.(n);
    STY = @(n) mat2doc.enum.style.WD_STYLE_TYPE.(n);

    P = struct();

    % ======================= table_props ==============================
    d = mat2doc.Document();
    t = d.add_table(2, 3);
    def = struct();
    def.alignment = rv(t.alignment);
    def.autofit = rv(t.autofit);
    def.table_direction = rv(t.table_direction);
    def.column_count = rv(t.column_count_());
    def.rows_len = rv(t.rows.len_());
    def.columns_len = rv(t.columns.len_());
    def.style_name = style_name(t);
    def.table_is_self = rv(t.table == t);
    def.rows_cached = rv(t.rows == t.rows);
    def.columns_cached = rv(t.columns == t.columns);
    tp = struct();
    tp.default = def;

    d.styles.add_style("My Table Style", STY("TABLE"));
    t.alignment = WTA("CENTER");
    t.autofit = false;
    t.table_direction = WTD("RTL");
    t.style = "My Table Style";
    aset = struct();
    aset.alignment = rv(t.alignment);
    aset.autofit = rv(t.autofit);
    aset.table_direction = rv(t.table_direction);
    aset.style_name = style_name(t);
    tp.after_set = aset;

    t.alignment = [];
    t.table_direction = [];
    anone = struct();
    anone.alignment = rv(t.alignment);
    anone.table_direction = rv(t.table_direction);
    tp.after_none = anone;
    P.table_props = tp;

    % ============================ columns ==============================
    d2 = mat2doc.Document();
    t2 = d2.add_table(2, 3);
    cols = t2.columns;
    col = struct();
    col.len = rv(cols.len_());
    col.iter_widths = map_rv(cols.to_array(), @(c) rv(c.width));
    col.iter_indices = map_rv(cols.to_array(), @(c) rv(c.index_()));
    col.getitem0_width = rv(cols.getitem_(0).width);
    col.getitem_neg1_index = rv(cols.getitem_(-1).index_());
    col.col0_table_is_t = rv(cols.getitem_(0).table == t2);
    col.cols_table_is_t = rv(cols.table == t2);
    col.err_oor_pos = errmsg(@() cols.getitem_(3));
    col.err_oor_neg = errmsg(@() cols.getitem_(-9));
    c1 = cols.getitem_(1); c1.width = IN(1.5);
    col.set_width_readback = rv(cols.getitem_(1).width);
    c1b = cols.getitem_(1); c1b.width = [];
    col.set_none_readback = rv(cols.getitem_(1).width);
    P.columns = col;

    % ============================ rows =================================
    d3 = mat2doc.Document();
    t3 = d3.add_table(2, 3);
    rows = t3.rows;
    row = struct();
    row.len = rv(rows.len_());
    row.iter_heights = map_rv(rows.to_array(), @(r) rv(r.height));
    row.iter_indices = map_rv(rows.to_array(), @(r) rv(r.index_()));
    row.getitem0_grid_cols_after = rv(rows.getitem_(0).grid_cols_after);
    row.getitem0_grid_cols_before = rv(rows.getitem_(0).grid_cols_before);
    row.getitem0_height_rule = rv(rows.getitem_(0).height_rule);
    row.getitem_neg1_index = rv(rows.getitem_(-1).index_());
    row.row0_table_is_t = rv(rows.getitem_(0).table == t3);
    row.rows_table_is_t = rv(rows.table == t3);
    row.slice_1_2_indices = map_rv(rows.getitem_(mkslice(1, 2, [])), @(r) rv(r.index_()));
    row.err_oor_pos = errmsg(@() rows.getitem_(2));
    row.err_oor_neg = errmsg(@() rows.getitem_(-3));
    r0 = rows.getitem_(0); r0.height = PT(20);
    r0h = rows.getitem_(0); r0h.height_rule = WRH("EXACTLY");
    row.set_height_readback = rv(rows.getitem_(0).height);
    row.set_height_rule_readback = rv(rows.getitem_(0).height_rule);
    r0n = rows.getitem_(0); r0n.height = [];
    row.set_height_none_readback = rv(rows.getitem_(0).height);
    P.rows = row;

    % ====================== iter_inner_content =========================
    d4 = mat2doc.Document();
    d4.add_table(2, 3);
    d4.add_paragraph("iic-below");
    iic = struct();
    iic.doc_types = map_short(d4.iter_inner_content());
    iic.section_types = map_short(d4.sections.getitem_(0).iter_inner_content());
    P.iter_inner_content = iic;
end

function c = map_rv(arr, fn)
    c = cell(1, numel(arr));
    for j = 1:numel(arr)
        c{j} = fn(arr(j));
    end
end

function c = map_short(items)
    c = cell(1, numel(items));
    for j = 1:numel(items)
        c{j} = shortcls(items{j});
    end
end

function s = mkslice(a, b, cc)
    s = struct();
    s.start = a; s.stop = b; s.step = cc;
end

function s = style_name(t)
    st = t.style;
    if isequal(st, [])
        s = "None";
    else
        s = string(st.name);
    end
end

function s = shortcls(x)
    parts = split(string(class(x)), ".");
    s = parts(end);
end

function s = rv(x)
    % Uniform accessor repr mirroring the s0072 rv(): None->"None", enum->member NAME
    % (A2 invisible), bool->"True"/"False", Length/int->EMU decimal.
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

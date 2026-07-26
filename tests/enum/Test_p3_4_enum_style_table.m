classdef Test_p3_4_enum_style_table < matlab.unittest.TestCase
% TEST_P3_4_ENUM_STYLE_TABLE  Gate-4 permanent unit tests for Mat2Doc P3-4
%   (the FINAL P3 WP: 7 concrete WD_* content enums under
%   +mat2doc.enum.{+style,+table,+shape} plus the 4 module-level class aliases).
%
%   WHAT THIS FREEZES (the P3-4 guarantees P4+ consumers lean on):
%
%   1. THE 153-MEMBER SET (the crux regression pin). For ALL 7 enums, every
%      member's (name, ms_api_value, xml_value, docstring length) is pinned in
%      DECLARATION ORDER against the frozen python-docx 1.2.0 oracle
%      (data\s0019_probe.json, copied verbatim from the Gate-3 reference
%      validation\mat2doc\references\s0019\probe.json). The load-bearing member
%      set is WD_BUILTIN_STYLE's 132 members (they feed the P4-7 styles tier): a
%      wrong / missing / extra member, a reordering, a changed value or xml
%      token, or a changed docstring length all go RED. Declaration order is
%      behavioural (from_xml is a first-match scan), so order is verified, not
%      just the set. numel guard + index-ordered loop.
%      (test_member_set_matches_oracle, parameterized over the 7 enums.)
%      The 2 period-less docstrings ride here (INDEX_HEADING doclen 13,
%      WD_TABLE_ALIGNMENT.LEFT doclen 12) and are additionally pinned explicitly
%      in test_periodless_docstrings.
%
%   2. THE 4 BaseXmlEnum ROUND-TRIPS. WD_STYLE_TYPE / WD_CELL_VERTICAL_ALIGNMENT
%      / WD_ROW_HEIGHT_RULE / WD_TABLE_ALIGNMENT: every member round-trips
%      from_xml(xml)->name, to_xml(member)->xml, AND to_xml(int)->xml (the
%      cls(value) int path, e.g. to_xml(1)->"paragraph"). Plus the two error
%      paths per enum: from_xml("bogus") and to_xml(99) verify the verbatim
%      mat2doc:ValueError message AND identifier (D-005), not merely that they
%      throw. (test_basexmlenum_roundtrip / test_basexmlenum_error_paths,
%      parameterized over the 4.)
%
%   3. WD_INLINE_SHAPE_TYPE (plain enum.Enum -> plain value classdef). 5 members
%      (CHART 12, LINKED_PICTURE 4, PICTURE 3, SMART_ART 15, NOT_IMPLEMENTED -6);
%      value stored int32; NOT_IMPLEMENTED negative; NOT a Base(Xml)Enum; no
%      from_xml/to_xml; no internal alias; string(member)==member name.
%
%   4. THE 2 BaseEnum (no XML). WD_BUILTIN_STYLE + WD_TABLE_DIRECTION:
%      int(member)==ms_api_value via double(member.value), and
%      str(member)==member.str_()=="NAME (value)" including the negatives
%      (NORMAL "NORMAL (-1)", BOOK_TITLE "BOOK_TITLE (-265)", RTL "RTL (1)").
%
%   5. THE 4 MODULE ALIASES. WD_STYLE (132 Constant re-exports, no statics),
%      WD_ALIGN_VERTICAL (4 Constants + forwards from_xml/to_xml), WD_ROW_HEIGHT
%      (3 Constants + forwards), WD_INLINE_SHAPE (5 Constants, no statics). Each
%      re-exports the IDENTICAL canonical members (identity ==, isa canonical);
%      the 2 BaseXmlEnum aliases forward the statics, the 2 non-XML aliases have
%      none (faithful to the base kind).
%
%   Provenance (Gates 1-3, all 2026-07-26):
%     * Audit    : validation\mat2doc\audit_P3-4_enum_style_table.md (Porter
%                  Gate-1 132-member WD_BUILTIN_STYLE dump + Gate-2 mso-auditor
%                  Opus APPROVE, 132/132 + 153-line member diff EXACT, 51/51
%                  runtime probes, 0 defects, ZERO new D-numbers).
%     * Validate : validation\mat2doc\validate_P3-4_enum_style_table.md (Gate-3
%                  PASS, 598/598 probe facts byte-identical via s0019 probe_diff
%                  exit 0, 0 new D-numbers, regression 511/511).
%     * Equivalence fixture (copied here so the suite is self-contained):
%       data\s0019_probe.json == the frozen python-docx oracle
%       validation\mat2doc\references\s0019\probe.json. Value/ASCII-token based
%       (no serialized bytes) -> no '* binary' .gitattributes needed
%       (s0016/s0017/s0018 precedent). Attributed scenario:
%       validation\mat2doc\scenarios\s0019_p3_4_enum_style_table.{py,m}.
%
%   Coverage taxonomy
%   -----------------
%   * Nominal     -- from_xml("character")->CHARACTER; to_xml(int)->token; the
%     happy 153-member set; the BaseEnum int/str facts.
%   * Edge        -- negative int32 members (NOT_IMPLEMENTED -6, all 132
%     WD_BUILTIN_STYLE values); period-less docstrings (doclen shorter than the
%     period-terminated form); single-alias identity; non-ASCII "cafe"
%     (e=U+00E9 via char(233)) unmapped-token error path; error paths verify the
%     mat2doc:ValueError IDENTIFIER + verbatim message, not merely that they throw.
%   * Equivalence -- every pinned value is read from the frozen s0019 oracle JSON
%     (the Gate-3 reference), so the class replays the validator's frozen output.
%   * Regression  -- hard-coded member counts (132 / 5 / 4 / 3 / 3 / 2 / 4), the
%     WD_INLINE_SHAPE_TYPE plain-enum shape, the alias Constant census, and the
%     verbatim ValueError forms.
%
%   Deviations exercised: 0 new D-numbers (Gate-3). Emitted errors ride the
%   standing D-005 adopt-only mat2doc:ValueError identifier convention. No
%   package bytes -- nothing serialized, so no L-ladder leg.
%
%   Determinism: no network, no absolute paths (the oracle resolves relative to
%   this class via fileparts(mfilename('fullpath'))), no file writes. The
%   +mat2doc package resolves via the MANDATORY PathFixture(worktree-root) added
%   in TestClassSetup (WP9-F4 lesson): R2024b runtests cd's into the test folder,
%   so a COLD run cannot resolve the +mat2doc package without it.

    properties
        % Frozen python-docx 1.2.0 oracle (jsondecode of data\s0019_probe.json).
        Oracle
    end

    properties (TestParameter)
        % The 7 concrete content enums, short name -> fully-qualified class name.
        % Drives the 153-member-set pin (one parameterized case per enum, so a
        % break names the offending enum). Field NAMES are the oracle members.*
        % keys; field VALUES are the +mat2doc class names.
        memberEnum = struct( ...
            'WD_BUILTIN_STYLE',           "mat2doc.enum.style.WD_BUILTIN_STYLE", ...
            'WD_STYLE_TYPE',              "mat2doc.enum.style.WD_STYLE_TYPE", ...
            'WD_CELL_VERTICAL_ALIGNMENT', "mat2doc.enum.table.WD_CELL_VERTICAL_ALIGNMENT", ...
            'WD_ROW_HEIGHT_RULE',         "mat2doc.enum.table.WD_ROW_HEIGHT_RULE", ...
            'WD_TABLE_ALIGNMENT',         "mat2doc.enum.table.WD_TABLE_ALIGNMENT", ...
            'WD_TABLE_DIRECTION',         "mat2doc.enum.table.WD_TABLE_DIRECTION", ...
            'WD_INLINE_SHAPE_TYPE',       "mat2doc.enum.shape.WD_INLINE_SHAPE_TYPE");

        % The 4 BaseXmlEnum enums (from_xml/to_xml surface). Each value carries
        % the fully-qualified class name, the oracle short name (= members/roundtrip
        % key), and the roundtrip.err_* key abbreviation.
        xmlEnum = struct( ...
            'WD_STYLE_TYPE', struct( ...
                'fqcn',  "mat2doc.enum.style.WD_STYLE_TYPE", ...
                'short', "WD_STYLE_TYPE", 'abbr', "wst"), ...
            'WD_CELL_VERTICAL_ALIGNMENT', struct( ...
                'fqcn',  "mat2doc.enum.table.WD_CELL_VERTICAL_ALIGNMENT", ...
                'short', "WD_CELL_VERTICAL_ALIGNMENT", 'abbr', "wcva"), ...
            'WD_ROW_HEIGHT_RULE', struct( ...
                'fqcn',  "mat2doc.enum.table.WD_ROW_HEIGHT_RULE", ...
                'short', "WD_ROW_HEIGHT_RULE", 'abbr', "wrhr"), ...
            'WD_TABLE_ALIGNMENT', struct( ...
                'fqcn',  "mat2doc.enum.table.WD_TABLE_ALIGNMENT", ...
                'short', "WD_TABLE_ALIGNMENT", 'abbr', "wta"));
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): idiom copied from the proven
            % tests\enum\Test_p3_3_enum_content.m so a COLD run resolves +mat2doc.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\enum
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end

        function loadOracle(testCase)
            % Load the frozen s0019 oracle (self-contained equivalence reference).
            here = fileparts(mfilename('fullpath'));
            f = fullfile(here, 'data', 's0019_probe.json');
            testCase.Oracle = jsondecode(fileread(f));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. THE 153-MEMBER SET -- the crux regression pin (Bar 1)         %
        % =============================================================== %

        function test_member_set_matches_oracle(testCase, memberEnum)
            % For the given enum, pin EVERY member's (name, ms_api_value,
            % xml_value, doclen) in declaration order against the frozen oracle.
            % Fields are gated on isfield: WD_INLINE_SHAPE_TYPE (plain enum)
            % carries neither xml nor doclen; the 2 BaseEnum carry doclen but no
            % xml; the 4 BaseXmlEnum carry both. So the fixture and the port agree
            % on which properties exist. (No member in this WP carries a None/""
            % xml_value, so the <missing> leg is defensive, not exercised.)
            fullName  = memberEnum;
            parts     = split(fullName, ".");
            shortName = char(parts(end));
            expected  = testCase.oracleMembers(shortName);
            members   = enumeration(fullName);

            testCase.verifyEqual(numel(members), numel(expected), ...
                sprintf('%s must have exactly %d members (declaration order set)', ...
                    shortName, numel(expected)));

            n = min(numel(members), numel(expected));
            for k = 1:n
                s = expected{k};
                m = members(k);
                ctx = sprintf('%s member #%d', shortName, k);

                testCase.verifyEqual(string(m), string(s.name), ...
                    [ctx ': name must match the oracle in declaration order']);
                testCase.verifyEqual(double(m.value), double(s.value), ...
                    [ctx ' (' char(string(s.name)) '): ms_api_value must match']);

                if isfield(s, 'xml')
                    if isempty(s.xml)
                        % JSON null -> Python None -> <missing> sentinel.
                        testCase.verifyTrue(ismissing(m.xml_value), ...
                            [ctx ' (' char(string(s.name)) '): xml_value must be <missing> (None)']);
                    else
                        testCase.verifyEqual(m.xml_value, string(s.xml), ...
                            [ctx ' (' char(string(s.name)) '): xml_value token must match']);
                    end
                end

                if isfield(s, 'doclen')
                    testCase.verifyEqual(strlength(m.doc), double(s.doclen), ...
                        [ctx ' (' char(string(s.name)) '): stripped docstring length must match']);
                end
            end
        end

        function test_periodless_docstrings(testCase)
            % Explicit pin for the 2 period-less docstrings (also caught by the
            % Bar-1 doclen loop, restated here as a load-bearing regression):
            % INDEX_HEADING = "Index Heading" (13, no trailing period) and
            % WD_TABLE_ALIGNMENT.LEFT = "Left-aligned" (12). Each is one char
            % shorter than the period-terminated sibling form.
            ihLen  = testCase.oracleDoclen('WD_BUILTIN_STYLE',  'INDEX_HEADING');
            leftLen = testCase.oracleDoclen('WD_TABLE_ALIGNMENT', 'LEFT');
            testCase.verifyEqual(double(ihLen), 13, ...
                'oracle fixture drift: INDEX_HEADING doclen must be 13');
            testCase.verifyEqual(double(leftLen), 12, ...
                'oracle fixture drift: WD_TABLE_ALIGNMENT.LEFT doclen must be 12');
            testCase.verifyEqual(strlength(mat2doc.enum.style.WD_BUILTIN_STYLE.INDEX_HEADING.doc), ...
                13, 'INDEX_HEADING docstring must be period-less ("Index Heading", len 13)');
            testCase.verifyEqual(strlength(mat2doc.enum.table.WD_TABLE_ALIGNMENT.LEFT.doc), ...
                12, 'WD_TABLE_ALIGNMENT.LEFT docstring must be period-less ("Left-aligned", len 12)');
        end

        % =============================================================== %
        % 2. The 4 BaseXmlEnum round-trips + error paths (Bar 2)           %
        % =============================================================== %

        function test_basexmlenum_roundtrip(testCase, xmlEnum)
            % Every member of the BaseXmlEnum round-trips both directions plus the
            % int (cls(value)) path, pinned against the frozen roundtrip oracle.
            fqcn    = xmlEnum.fqcn;
            rtlist  = testCase.oracleRoundtrip(char(xmlEnum.short));
            members = enumeration(fqcn);
            names   = arrayfun(@string, members);

            testCase.verifyEqual(numel(rtlist), numel(members), ...
                sprintf('%s roundtrip fixture must cover every member', char(xmlEnum.short)));

            for k = 1:numel(rtlist)
                e    = rtlist{k};
                name = string(e.name);
                m    = members(names == name);
                testCase.assertEqual(numel(m), 1, ...
                    ['roundtrip name not found as a member: ' char(name)]);

                % from_xml(xml) -> member name
                got = feval(char(fqcn + ".from_xml"), string(e.xml));
                testCase.verifyEqual(string(got), string(e.from_xml), ...
                    sprintf('%s.from_xml("%s") must return %s', ...
                        char(xmlEnum.short), char(string(e.xml)), char(string(e.from_xml))));

                % to_xml(member) -> xml token
                testCase.verifyEqual(feval(char(fqcn + ".to_xml"), m), string(e.to_xml_member), ...
                    sprintf('%s.to_xml(%s member) must return "%s"', ...
                        char(xmlEnum.short), char(name), char(string(e.to_xml_member))));

                % to_xml(int) -> xml token (the cls(value) int path)
                testCase.verifyEqual(feval(char(fqcn + ".to_xml"), double(m.value)), ...
                    string(e.to_xml_int), ...
                    sprintf('%s.to_xml(%d int) must return "%s"', ...
                        char(xmlEnum.short), double(m.value), char(string(e.to_xml_int))));
            end
        end

        function test_basexmlenum_error_paths(testCase, xmlEnum)
            % Both docx error forms, byte-verbatim, with the D-005 identifier:
            %   from_xml("bogus") -> "<ENUM> has no XML mapping for 'bogus'"
            %   to_xml(99)        -> "99 is not a valid <ENUM>"  (unquoted int)
            fqcn    = xmlEnum.fqcn;
            errFrom = testCase.Oracle.roundtrip.(char("err_" + xmlEnum.abbr + "_from_bogus"));
            errTo   = testCase.Oracle.roundtrip.(char("err_" + xmlEnum.abbr + "_to_99"));

            [idF, msgF] = raiseAndCapture(@() feval(char(fqcn + ".from_xml"), "bogus"));
            testCase.verifyEqual(idF, char(string(errFrom.id)), ...
                sprintf('%s.from_xml("bogus") identifier must be %s', ...
                    char(xmlEnum.short), char(string(errFrom.id))));
            testCase.verifyEqual(string(msgF), string(errFrom.msg), ...
                sprintf('%s.from_xml("bogus") message must be byte-verbatim', char(xmlEnum.short)));

            [idT, msgT] = raiseAndCapture(@() feval(char(fqcn + ".to_xml"), 99));
            testCase.verifyEqual(idT, char(string(errTo.id)), ...
                sprintf('%s.to_xml(99) identifier must be %s', ...
                    char(xmlEnum.short), char(string(errTo.id))));
            testCase.verifyEqual(string(msgT), string(errTo.msg), ...
                sprintf('%s.to_xml(99) message must be byte-verbatim (unquoted int)', char(xmlEnum.short)));
        end

        % =============================================================== %
        % 3. WD_INLINE_SHAPE_TYPE (plain enum.Enum) -- Bar 3               %
        % =============================================================== %

        function test_wd_inline_shape_type(testCase)
            % Plain value classdef: 5 members, negative int32 NOT_IMPLEMENTED,
            % NOT a Base(Xml)Enum, no from_xml/to_xml, no internal alias.
            sh = testCase.Oracle.shape;
            members = enumeration("mat2doc.enum.shape.WD_INLINE_SHAPE_TYPE");
            testCase.verifyEqual(numel(members), double(sh.member_count), ...
                'WD_INLINE_SHAPE_TYPE must have exactly 5 members (no internal alias)');

            testCase.verifyEqual(double(mat2doc.enum.shape.WD_INLINE_SHAPE_TYPE.CHART.value), ...
                double(sh.chart_value), 'CHART value must be 12');
            testCase.verifyEqual(double(mat2doc.enum.shape.WD_INLINE_SHAPE_TYPE.LINKED_PICTURE.value), ...
                double(sh.linked_picture_value), 'LINKED_PICTURE value must be 4');
            testCase.verifyEqual(double(mat2doc.enum.shape.WD_INLINE_SHAPE_TYPE.PICTURE.value), ...
                double(sh.picture_value), 'PICTURE value must be 3');
            testCase.verifyEqual(double(mat2doc.enum.shape.WD_INLINE_SHAPE_TYPE.SMART_ART.value), ...
                double(sh.smart_art_value), 'SMART_ART value must be 15');
            testCase.verifyEqual(double(mat2doc.enum.shape.WD_INLINE_SHAPE_TYPE.NOT_IMPLEMENTED.value), ...
                double(sh.not_implemented_value), 'NOT_IMPLEMENTED value must be -6 (negative)');

            % Stored as int32 (negative-safe); member name via string().
            testCase.verifyEqual(class(mat2doc.enum.shape.WD_INLINE_SHAPE_TYPE.NOT_IMPLEMENTED.value), ...
                'int32', 'value property must be stored int32');
            testCase.verifyEqual(string(mat2doc.enum.shape.WD_INLINE_SHAPE_TYPE.PICTURE), ...
                string(sh.picture_name), 'string(PICTURE) must be the member name "PICTURE"');

            % NOT a Base(Xml)Enum; no from_xml/to_xml surface.
            testCase.verifyFalse(isa(mat2doc.enum.shape.WD_INLINE_SHAPE_TYPE.PICTURE, ...
                'mat2doc.enum.base.BaseXmlEnum'), 'WD_INLINE_SHAPE_TYPE must NOT be a BaseXmlEnum');
            testCase.verifyFalse(isa(mat2doc.enum.shape.WD_INLINE_SHAPE_TYPE.PICTURE, ...
                'mat2doc.enum.base.BaseEnum'), 'WD_INLINE_SHAPE_TYPE must NOT be a BaseEnum');
            ml = methods('mat2doc.enum.shape.WD_INLINE_SHAPE_TYPE');
            testCase.verifyFalse(any(strcmp(ml, 'from_xml')), ...
                'plain enum WD_INLINE_SHAPE_TYPE must have NO from_xml');
            testCase.verifyFalse(any(strcmp(ml, 'to_xml')), ...
                'plain enum WD_INLINE_SHAPE_TYPE must have NO to_xml');
        end

        % =============================================================== %
        % 4. The 2 BaseEnum (no XML): int/str facts -- Bar 4              %
        % =============================================================== %

        function test_wd_builtin_style_baseenum(testCase)
            % int(member)==ms_api_value and str(member)=="NAME (value)" over the
            % load-bearing negatives (all 132 values negative).
            be = testCase.Oracle.baseenum;
            testCase.verifyEqual(numel(enumeration("mat2doc.enum.style.WD_BUILTIN_STYLE")), ...
                double(be.wbs_count), 'WD_BUILTIN_STYLE must have exactly 132 members');

            testCase.verifyEqual(double(mat2doc.enum.style.WD_BUILTIN_STYLE.NORMAL.value), ...
                double(be.wbs_normal_int), 'NORMAL int must be -1');
            testCase.verifyEqual(mat2doc.enum.style.WD_BUILTIN_STYLE.NORMAL.str_(), ...
                string(be.wbs_normal_str), 'str(NORMAL) must be "NORMAL (-1)"');
            testCase.verifyEqual(double(mat2doc.enum.style.WD_BUILTIN_STYLE.BODY_TEXT.value), ...
                double(be.wbs_body_text_int), 'BODY_TEXT int must be -67');
            testCase.verifyEqual(mat2doc.enum.style.WD_BUILTIN_STYLE.BODY_TEXT.str_(), ...
                string(be.wbs_body_text_str), 'str(BODY_TEXT) must be "BODY_TEXT (-67)"');
            testCase.verifyEqual(double(mat2doc.enum.style.WD_BUILTIN_STYLE.BOOK_TITLE.value), ...
                double(be.wbs_book_title_int), 'BOOK_TITLE int must be -265 (most-negative)');
            testCase.verifyEqual(mat2doc.enum.style.WD_BUILTIN_STYLE.BOOK_TITLE.str_(), ...
                string(be.wbs_book_title_str), 'str(BOOK_TITLE) must be "BOOK_TITLE (-265)"');
            testCase.verifyEqual(double(mat2doc.enum.style.WD_BUILTIN_STYLE.INDEX_HEADING.value), ...
                double(be.wbs_index_heading_int), 'INDEX_HEADING int must be -34');
            testCase.verifyEqual(mat2doc.enum.style.WD_BUILTIN_STYLE.INDEX_HEADING.str_(), ...
                string(be.wbs_index_heading_str), 'str(INDEX_HEADING) must be "INDEX_HEADING (-34)"');
            testCase.verifyEqual(double(mat2doc.enum.style.WD_BUILTIN_STYLE.HEADING_1.value), ...
                double(be.wbs_heading_1_int), 'HEADING_1 int must be -2');
            testCase.verifyEqual(double(mat2doc.enum.style.WD_BUILTIN_STYLE.HEADING_9.value), ...
                double(be.wbs_heading_9_int), 'HEADING_9 int must be -10');
            testCase.verifyEqual(double(mat2doc.enum.style.WD_BUILTIN_STYLE.TOC_9.value), ...
                double(be.wbs_toc_9_int), 'TOC_9 int must be -28');
        end

        function test_wd_table_direction_baseenum(testCase)
            % The second BaseEnum: LTR 0 / RTL 1, str "LTR (0)" / "RTL (1)".
            be = testCase.Oracle.baseenum;
            testCase.verifyEqual(numel(enumeration("mat2doc.enum.table.WD_TABLE_DIRECTION")), ...
                double(be.wtd_count), 'WD_TABLE_DIRECTION must have exactly 2 members');
            testCase.verifyEqual(double(mat2doc.enum.table.WD_TABLE_DIRECTION.LTR.value), ...
                double(be.wtd_ltr_int), 'LTR int must be 0');
            testCase.verifyEqual(mat2doc.enum.table.WD_TABLE_DIRECTION.LTR.str_(), ...
                string(be.wtd_ltr_str), 'str(LTR) must be "LTR (0)"');
            testCase.verifyEqual(double(mat2doc.enum.table.WD_TABLE_DIRECTION.RTL.value), ...
                double(be.wtd_rtl_int), 'RTL int must be 1');
            testCase.verifyEqual(mat2doc.enum.table.WD_TABLE_DIRECTION.RTL.str_(), ...
                string(be.wtd_rtl_str), 'str(RTL) must be "RTL (1)"');
        end

        % =============================================================== %
        % 5. The 4 module aliases -- Bar 5 (identity + forwarding)         %
        % =============================================================== %

        function test_alias_wd_style(testCase)
            % WD_STYLE re-exports the 132 canonical members as Constants (== and
            % isa behave as one enumeration) and carries NO statics (BaseEnum).
            al = testCase.Oracle.aliases;
            testCase.verifyTrue(mat2doc.enum.style.WD_STYLE.BODY_TEXT == ...
                mat2doc.enum.style.WD_BUILTIN_STYLE.BODY_TEXT, ...
                'WD_STYLE.BODY_TEXT must be identical to the canonical member');
            testCase.verifyEqual(logical(al.ws_body_text_eq_canonical), true, ...
                'oracle fixture drift: ws_body_text_eq_canonical must be true');
            testCase.verifyTrue(isa(mat2doc.enum.style.WD_STYLE.TOC_9, ...
                'mat2doc.enum.style.WD_BUILTIN_STYLE'), ...
                'WD_STYLE members must be WD_BUILTIN_STYLE instances');
            testCase.verifyEqual(string(mat2doc.enum.style.WD_STYLE.BODY_TEXT), ...
                string(al.ws_body_text_name), 'WD_STYLE.BODY_TEXT name must be "BODY_TEXT"');
            testCase.verifyEqual(double(mat2doc.enum.style.WD_STYLE.BODY_TEXT.value), ...
                double(al.ws_body_text_value), 'WD_STYLE.BODY_TEXT.value must be -67');
            testCase.verifyEqual(double(mat2doc.enum.style.WD_STYLE.NORMAL.value), ...
                double(al.ws_normal_value), 'WD_STYLE.NORMAL.value must be -1');
            testCase.verifyEqual(string(mat2doc.enum.style.WD_STYLE.TOC_9), ...
                string(al.ws_toc_9_name), 'WD_STYLE.TOC_9 name must be "TOC_9"');

            % 132 Constant re-exports; no from_xml/to_xml (BaseEnum has none).
            mc = ?mat2doc.enum.style.WD_STYLE;
            testCase.verifyEqual(numel(mc.PropertyList), double(al.ws_count), ...
                'WD_STYLE must re-export exactly 132 Constant members');
            ml = methods('mat2doc.enum.style.WD_STYLE');
            testCase.verifyFalse(any(strcmp(ml, 'from_xml')), ...
                'WD_STYLE alias must have NO from_xml (BaseEnum canonical)');
            testCase.verifyFalse(any(strcmp(ml, 'to_xml')), ...
                'WD_STYLE alias must have NO to_xml (BaseEnum canonical)');
        end

        function test_alias_wd_align_vertical(testCase)
            % BaseXmlEnum alias: identity + forwards from_xml/to_xml.
            al = testCase.Oracle.aliases;
            testCase.verifyTrue(mat2doc.enum.table.WD_ALIGN_VERTICAL.BOTTOM == ...
                mat2doc.enum.table.WD_CELL_VERTICAL_ALIGNMENT.BOTTOM, ...
                'WD_ALIGN_VERTICAL.BOTTOM must be identical to the canonical member');
            testCase.verifyEqual(string(mat2doc.enum.table.WD_ALIGN_VERTICAL.BOTTOM), ...
                string(al.wav_bottom_name), 'WD_ALIGN_VERTICAL.BOTTOM name must be "BOTTOM"');
            testCase.verifyEqual(string(mat2doc.enum.table.WD_ALIGN_VERTICAL.from_xml("top")), ...
                string(al.wav_from_top), 'WD_ALIGN_VERTICAL.from_xml("top") must forward -> TOP');
            testCase.verifyEqual(string(mat2doc.enum.table.WD_ALIGN_VERTICAL.from_xml("both")), ...
                string(al.wav_from_both), 'WD_ALIGN_VERTICAL.from_xml("both") must forward -> BOTH');
            testCase.verifyEqual(mat2doc.enum.table.WD_ALIGN_VERTICAL.to_xml( ...
                mat2doc.enum.table.WD_ALIGN_VERTICAL.BOTTOM), ...
                string(al.wav_to_bottom), 'WD_ALIGN_VERTICAL.to_xml(BOTTOM) must forward -> "bottom"');
        end

        function test_alias_wd_row_height(testCase)
            % BaseXmlEnum alias: identity + forwards from_xml/to_xml.
            al = testCase.Oracle.aliases;
            testCase.verifyTrue(mat2doc.enum.table.WD_ROW_HEIGHT.AT_LEAST == ...
                mat2doc.enum.table.WD_ROW_HEIGHT_RULE.AT_LEAST, ...
                'WD_ROW_HEIGHT.AT_LEAST must be identical to the canonical member');
            testCase.verifyEqual(string(mat2doc.enum.table.WD_ROW_HEIGHT.AT_LEAST), ...
                string(al.wrh_atleast_name), 'WD_ROW_HEIGHT.AT_LEAST name must be "AT_LEAST"');
            testCase.verifyEqual(string(mat2doc.enum.table.WD_ROW_HEIGHT.from_xml("auto")), ...
                string(al.wrh_from_auto), 'WD_ROW_HEIGHT.from_xml("auto") must forward -> AUTO');
            testCase.verifyEqual(mat2doc.enum.table.WD_ROW_HEIGHT.to_xml( ...
                mat2doc.enum.table.WD_ROW_HEIGHT.EXACTLY), ...
                string(al.wrh_to_exactly), 'WD_ROW_HEIGHT.to_xml(EXACTLY) must forward -> "exact"');
        end

        function test_alias_wd_inline_shape(testCase)
            % Plain-enum alias: identity; NO statics (canonical has none).
            al = testCase.Oracle.aliases;
            testCase.verifyTrue(mat2doc.enum.shape.WD_INLINE_SHAPE.PICTURE == ...
                mat2doc.enum.shape.WD_INLINE_SHAPE_TYPE.PICTURE, ...
                'WD_INLINE_SHAPE.PICTURE must be identical to the canonical member');
            testCase.verifyEqual(string(mat2doc.enum.shape.WD_INLINE_SHAPE.PICTURE), ...
                string(al.wis_picture_name), 'WD_INLINE_SHAPE.PICTURE name must be "PICTURE"');
            testCase.verifyEqual(double(mat2doc.enum.shape.WD_INLINE_SHAPE.CHART.value), ...
                double(al.wis_chart_value), 'WD_INLINE_SHAPE.CHART.value must be 12');
            testCase.verifyEqual(double(mat2doc.enum.shape.WD_INLINE_SHAPE.NOT_IMPLEMENTED.value), ...
                double(al.wis_not_implemented_value), 'WD_INLINE_SHAPE.NOT_IMPLEMENTED.value must be -6');
            ml = methods('mat2doc.enum.shape.WD_INLINE_SHAPE');
            testCase.verifyFalse(any(strcmp(ml, 'from_xml')), ...
                'WD_INLINE_SHAPE alias must have NO from_xml (plain-enum canonical)');
            testCase.verifyFalse(any(strcmp(ml, 'to_xml')), ...
                'WD_INLINE_SHAPE alias must have NO to_xml (plain-enum canonical)');
        end

        % =============================================================== %
        % 6. Non-ASCII edge / error path                                  %
        % =============================================================== %

        function test_from_xml_nonascii_bogus_raises(testCase)
            % Non-ASCII edge: query "cafe" with e = U+00E9 built via char(233) so
            % the expectation is independent of THIS file's source encoding. The
            % unmapped token raises mat2doc:ValueError with the query rendered
            % byte-identically inside the literal single quotes (UTF-8 pipeline).
            query    = "caf" + string(char(233));
            expected = "WD_STYLE_TYPE has no XML mapping for '" + query + "'";
            [id, msg] = raiseAndCapture(@() mat2doc.enum.style.WD_STYLE_TYPE.from_xml(query));
            testCase.verifyEqual(id, 'mat2doc:ValueError', ...
                'non-ASCII unmapped query must raise mat2doc:ValueError');
            testCase.verifyEqual(string(msg), expected, ...
                'the non-ASCII query must render byte-identically in the message');
        end

    end

    methods (Access = private)
        function mlist = oracleMembers(testCase, shortName)
            % The oracle member list for `shortName` as a cell of scalar structs.
            % jsondecode yields a struct ARRAY when every element has identical
            % fields (num2cell -> cell of structs) or a CELL of structs when
            % heterogeneous; normalize both. All members within a single P3-4
            % enum are homogeneous, so the struct-array leg is the live one.
            raw = testCase.Oracle.members.(char(shortName));
            if iscell(raw)
                mlist = raw;
            else
                mlist = num2cell(raw);
            end
        end

        function rlist = oracleRoundtrip(testCase, shortName)
            % The oracle roundtrip list for `shortName` as a cell of structs.
            raw = testCase.Oracle.roundtrip.(char(shortName));
            if iscell(raw)
                rlist = raw;
            else
                rlist = num2cell(raw);
            end
        end

        function n = oracleDoclen(testCase, shortName, memberName)
            % doclen of the named member from the oracle member list (for the
            % explicit period-less-docstring pin).
            mlist = testCase.oracleMembers(shortName);
            n = NaN;
            for k = 1:numel(mlist)
                if string(mlist{k}.name) == string(memberName)
                    n = mlist{k}.doclen;
                    return
                end
            end
        end
    end
end

% ===================== file-local helpers ============================== %

function [id, msg] = raiseAndCapture(fn)
    % Run fn() expecting it to raise; return (identifier, message) as char, or
    % ('','') if it did not raise. (Idiom from tests\enum\Test_p3_3_enum_content.m.)
    id = ''; msg = '';
    try
        fn();
    catch ME
        id = ME.identifier;
        msg = ME.message;
    end
end

classdef Test_p3_3_enum_content < matlab.unittest.TestCase
% TEST_P3_3_ENUM_CONTENT  Gate-4 permanent unit tests for Mat2Doc P3-3
%   (the 12 concrete WD_*/MSO_* content enums under
%   +mat2doc.enum.{+text,+section,+dml}, plus the WD_BREAK_TYPE.TEXT_WRAPPING
%   member-alias and the 7 module-level class aliases).
%
%   WHAT THIS FREEZES (the P3-3 guarantees every P4/P5/P6 consumer leans on):
%
%   1. THE 108-MEMBER SET (the crux regression pin). For ALL 12 enums, every
%      member's (name, ms_api_value, xml_value, docstring length) is pinned in
%      DECLARATION ORDER against the frozen python-docx 1.2.0 oracle
%      (data\s0018_probe.json, copied verbatim from the Gate-3 reference
%      validation\mat2doc\references\s0018\probe.json). A wrong / missing / extra
%      member, a reordering, a changed value or xml token, or a changed docstring
%      length all go RED. Declaration order is behavioural: from_xml is a
%      first-match scan, so order is verified, not just the set.
%      (test_member_set_matches_oracle, parameterized over the 12 enums.)
%
%   2. None-VALUED members (H3 tri-state). WD_COLOR_INDEX.INHERITED and
%      WD_UNDERLINE.INHERITED carry xml_value == <missing> (Python None). Under
%      the P3-1 docx no-None-guard, from_xml([]) (Python None) returns INHERITED
%      on both; to_xml(INHERITED) -- by member AND by its int -1 -- raises
%      mat2doc:ValueError "<Enum>.INHERITED has no XML representation" verbatim.
%
%   3. THE MSO_THEME_COLOR_INDEX docx DELTA (must never regress to the pptx
%      shape). 17 members, full-word xml tokens (accent1/dark1/hyperlink/...),
%      NOT_THEME_COLOR.xml_value == "UNMAPPED" (a truthy placeholder, NOT ""),
%      and NO MIXED member. to_xml(NOT_THEME_COLOR)/to_xml(0) -> "UNMAPPED";
%      from_xml("UNMAPPED") -> NOT_THEME_COLOR; from_xml("") raises (empty string
%      is a real string, NOT the None sentinel). None of the Mat2Ppt
%      abbreviations may leak in.
%
%   4. THE WD_BREAK_TYPE member-alias. TEXT_WRAPPING resolves to LINE_CLEAR_ALL
%      (name "LINE_CLEAR_ALL", value 11), == LINE_CLEAR_ALL (true), == LINE
%      (false; LINE == 6); the canonical iteration is exactly 10 members (the
%      alias excluded), in declaration order. Plain value classdef (no XML).
%
%   5. THE 7 MODULE ALIASES. WD_ALIGN_PARAGRAPH / WD_BREAK / WD_COLOR /
%      WD_HEADER_FOOTER / WD_ORIENT / WD_SECTION / MSO_THEME_COLOR each re-export
%      the IDENTICAL canonical members (identity ==) and forward from_xml/to_xml
%      (WD_BREAK carries the TEXT_WRAPPING member-alias and has no statics, its
%      canonical WD_BREAK_TYPE being a plain enum).
%
%   6. MSO_COLOR_TYPE (BaseEnum, no XML). int(member) == ms_api_value
%      (RGB 1 / THEME 2 / AUTO 101); str(member) == "NAME (value)".
%
%   7. from_xml/to_xml round-trips per BaseXmlEnum, incl. the badint form
%      to_xml(999) -> "999 is not a valid <Enum>" (no repr quotes for an int).
%
%   Provenance (Gates 1-3, all 2026-07-26):
%     * Audit    : validation\mat2doc\audit_P3-3_enum_content.md (Porter Gate-1
%                  108-member dump + Gate-2 mso-auditor Opus APPROVE, 108/108
%                  member lines byte-identical, 0 defects).
%     * Validate : validation\mat2doc\validate_P3-3_enum_content.md (Gate-3 PASS,
%                  498/498 probe facts byte-identical via s0018 probe_diff exit 0,
%                  0 new D-numbers, regression 474/474).
%     * Equivalence fixture (copied here so the suite is self-contained):
%       data\s0018_probe.json == the frozen python-docx oracle
%       validation\mat2doc\references\s0018\probe.json. Value/ASCII-token based
%       (no serialized bytes) -> no '* binary' .gitattributes needed
%       (s0016/s0017 precedent). Attributed scenario:
%       validation\mat2doc\scenarios\s0018_p3_3_enum_content.{py,m}.
%
%   Coverage taxonomy
%   -----------------
%   * Nominal     -- from_xml("center")->CENTER; to_xml(int)->token; MSO_COLOR_TYPE
%     int/str; the happy member set.
%   * Edge        -- None-valued members (from_xml([])->INHERITED, to_xml raises);
%     empty-string != None (from_xml("") raises); the truthy "UNMAPPED"
%     placeholder; single-member alias identity; non-ASCII "cafe" (e=U+00E9 via
%     char(233)) error path; error paths verify the mat2doc:ValueError IDENTIFIER
%     + verbatim message (not merely that they throw).
%   * Equivalence -- every pinned value is read from the frozen s0018 oracle JSON
%     (the Gate-3 reference), so the class replays the validator's frozen output.
%   * Regression  -- hard-coded member counts, the no-MIXED guard, the docx
%     full-word tokens, the WD_BREAK_TYPE alias, and the verbatim ValueError forms.
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
        % Frozen python-docx 1.2.0 oracle (jsondecode of data\s0018_probe.json).
        Oracle
    end

    properties (TestParameter)
        % The 12 concrete content enums, short name -> fully-qualified class name.
        % Drives the 108-member-set pin (one parameterized case per enum, so a
        % break names the offending enum). Field NAMES are the oracle members.*
        % keys; field VALUES are the +mat2doc class names.
        memberEnum = struct( ...
            'WD_PARAGRAPH_ALIGNMENT', "mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT", ...
            'WD_BREAK_TYPE',          "mat2doc.enum.text.WD_BREAK_TYPE", ...
            'WD_COLOR_INDEX',         "mat2doc.enum.text.WD_COLOR_INDEX", ...
            'WD_LINE_SPACING',        "mat2doc.enum.text.WD_LINE_SPACING", ...
            'WD_TAB_ALIGNMENT',       "mat2doc.enum.text.WD_TAB_ALIGNMENT", ...
            'WD_TAB_LEADER',          "mat2doc.enum.text.WD_TAB_LEADER", ...
            'WD_UNDERLINE',           "mat2doc.enum.text.WD_UNDERLINE", ...
            'WD_HEADER_FOOTER_INDEX', "mat2doc.enum.section.WD_HEADER_FOOTER_INDEX", ...
            'WD_ORIENTATION',         "mat2doc.enum.section.WD_ORIENTATION", ...
            'WD_SECTION_START',       "mat2doc.enum.section.WD_SECTION_START", ...
            'MSO_COLOR_TYPE',         "mat2doc.enum.dml.MSO_COLOR_TYPE", ...
            'MSO_THEME_COLOR_INDEX',  "mat2doc.enum.dml.MSO_THEME_COLOR_INDEX");
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): idiom copied from the proven
            % tests\enum\Test_p3_1_enum_base.m so a COLD run resolves +mat2doc.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\enum
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end

        function loadOracle(testCase)
            % Load the frozen s0018 oracle (self-contained equivalence reference).
            here = fileparts(mfilename('fullpath'));
            f = fullfile(here, 'data', 's0018_probe.json');
            testCase.Oracle = jsondecode(fileread(f));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. THE 108-MEMBER SET -- the crux regression pin (Bar 1)         %
        % =============================================================== %

        function test_member_set_matches_oracle(testCase, memberEnum)
            % For the given enum, pin EVERY member's (name, ms_api_value,
            % xml_value, doclen) in declaration order against the frozen oracle.
            % xml_value: JSON null (empty in the decoded struct) => <missing>;
            % doclen/xml fields are absent for enums that do not carry them
            % (WD_BREAK_TYPE: no doc, no xml; MSO_COLOR_TYPE: doc but no xml), so
            % each field is gated on isfield -- the fixture and the port agree on
            % which properties exist.
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

        % =============================================================== %
        % 2. None-VALUED members (H3 tri-state) -- Bar 2                   %
        % =============================================================== %

        function test_wd_color_index_none_and_roundtrip(testCase)
            % from_xml([]) (Python None) -> INHERITED (docx no-None-guard); a
            % mapped token round-trips; to_xml(INHERITED) raises by member AND by
            % its int -1, with the verbatim message.
            rt = testCase.Oracle.roundtrip;
            WCI = "mat2doc.enum.text.WD_COLOR_INDEX";

            testCase.verifyEqual(string(mat2doc.enum.text.WD_COLOR_INDEX.from_xml([])), ...
                string(rt.wci_from_none), 'from_xml([]) must return INHERITED (docx None member)');
            testCase.verifyEqual(string(mat2doc.enum.text.WD_COLOR_INDEX.from_xml("red")), ...
                string(rt.wci_from_red), 'from_xml("red") must return RED');
            testCase.verifyEqual(mat2doc.enum.text.WD_COLOR_INDEX.to_xml(6), ...
                string(rt.wci_to_6), 'to_xml(6) must return "red"');

            [id1, msg1] = raiseAndCapture(@() mat2doc.enum.text.WD_COLOR_INDEX.to_xml( ...
                mat2doc.enum.text.WD_COLOR_INDEX.INHERITED));
            testCase.verifyEqual(id1, 'mat2doc:ValueError');
            testCase.verifyEqual(string(msg1), string(rt.wci_to_inherited_msg), ...
                'to_xml(INHERITED member) must raise the verbatim "has no XML representation"');
            [id2, msg2] = raiseAndCapture(@() mat2doc.enum.text.WD_COLOR_INDEX.to_xml(-1));
            testCase.verifyEqual(id2, 'mat2doc:ValueError');
            testCase.verifyEqual(string(msg2), string(rt.wci_to_neg1_msg), ...
                'to_xml(-1) must raise the same verbatim message (int reaches INHERITED)');
            testCase.assertTrue(strlength(WCI) > 0);   % keep WCI referenced (doc anchor)
        end

        function test_wd_underline_none_and_roundtrip(testCase)
            % Same None machinery for the second None-bearing enum.
            rt = testCase.Oracle.roundtrip;
            testCase.verifyEqual(string(mat2doc.enum.text.WD_UNDERLINE.from_xml([])), ...
                string(rt.wu_from_none), 'from_xml([]) must return INHERITED');
            testCase.verifyEqual(string(mat2doc.enum.text.WD_UNDERLINE.from_xml("single")), ...
                string(rt.wu_from_single), 'from_xml("single") must return SINGLE');
            testCase.verifyEqual(mat2doc.enum.text.WD_UNDERLINE.to_xml(1), ...
                string(rt.wu_to_1), 'to_xml(1) must return "single"');
            [id, msg] = raiseAndCapture(@() mat2doc.enum.text.WD_UNDERLINE.to_xml( ...
                mat2doc.enum.text.WD_UNDERLINE.INHERITED));
            testCase.verifyEqual(id, 'mat2doc:ValueError');
            testCase.verifyEqual(string(msg), string(rt.wu_to_inherited_msg), ...
                'to_xml(INHERITED) must raise the verbatim WD_UNDERLINE message');
        end

        function test_inherited_members_xml_value_missing(testCase)
            % Property-level H3 pin: both INHERITED members store <missing>, not
            % a real string. Distinguishes None from "" at the store.
            testCase.verifyTrue(ismissing(mat2doc.enum.text.WD_COLOR_INDEX.INHERITED.xml_value), ...
                'WD_COLOR_INDEX.INHERITED.xml_value must be <missing> (Python None)');
            testCase.verifyTrue(ismissing(mat2doc.enum.text.WD_UNDERLINE.INHERITED.xml_value), ...
                'WD_UNDERLINE.INHERITED.xml_value must be <missing> (Python None)');
        end

        % =============================================================== %
        % 3. MSO_THEME_COLOR_INDEX docx DELTA -- Bar 3                     %
        % =============================================================== %

        function test_mso_theme_color_index_docx_shape(testCase)
            % 17 members, NO MIXED, NOT_THEME_COLOR carries "UNMAPPED" (truthy,
            % not ""). Pins the docx shape so a future edit cannot regress it to
            % the Mat2Ppt (abbreviated-token, MIXED-bearing) shape.
            members = enumeration("mat2doc.enum.dml.MSO_THEME_COLOR_INDEX");
            names   = arrayfun(@string, members);
            testCase.verifyEqual(numel(members), 17, ...
                'MSO_THEME_COLOR_INDEX must have exactly 17 members (docx delta)');
            testCase.verifyFalse(any(names == "MIXED"), ...
                'MSO_THEME_COLOR_INDEX must NOT have a MIXED member (docx has none; pptx does)');
            testCase.verifyEqual(mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.NOT_THEME_COLOR.xml_value, ...
                "UNMAPPED", 'NOT_THEME_COLOR.xml_value must be the truthy placeholder "UNMAPPED", not ""');
        end

        function test_mso_theme_color_index_unmapped_roundtrip(testCase)
            % "UNMAPPED" is truthy, so to_xml returns it (not guarded);
            % from_xml("UNMAPPED") resolves the first such member NOT_THEME_COLOR.
            rt = testCase.Oracle.roundtrip;
            testCase.verifyEqual(mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.to_xml( ...
                mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.NOT_THEME_COLOR), ...
                string(rt.mtci_to_not_theme), 'to_xml(NOT_THEME_COLOR) must return "UNMAPPED"');
            testCase.verifyEqual(mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.to_xml(0), ...
                string(rt.mtci_to_0), 'to_xml(0) must return "UNMAPPED"');
            testCase.verifyEqual(string(mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.from_xml("UNMAPPED")), ...
                string(rt.mtci_from_unmapped), 'from_xml("UNMAPPED") must return NOT_THEME_COLOR (first match)');
        end

        function test_mso_theme_color_index_fullword_roundtrip(testCase)
            % Full-word docx tokens round-trip (accent1/dark1/hyperlink), proving
            % the pptx abbreviations did not leak in.
            rt = testCase.Oracle.roundtrip;
            testCase.verifyEqual(mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.to_xml(5), ...
                string(rt.mtci_to_5), 'to_xml(5) must return "accent1"');
            testCase.verifyEqual(mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.to_xml(11), ...
                string(rt.mtci_to_11), 'to_xml(11) must return "hyperlink"');
            testCase.verifyEqual(string(mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.from_xml("accent1")), ...
                string(rt.mtci_from_accent1), 'from_xml("accent1") must return ACCENT_1');
            testCase.verifyEqual(string(mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.from_xml("dark1")), ...
                string(rt.mtci_from_dark1), 'from_xml("dark1") must return DARK_1');
        end

        function test_mso_theme_color_index_from_empty_raises(testCase)
            % empty string != None: from_xml("") finds no member (NOT_THEME_COLOR
            % carries "UNMAPPED", not "") and raises the verbatim mapping message.
            rt = testCase.Oracle.roundtrip;
            [id, msg] = raiseAndCapture(@() mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.from_xml(""));
            testCase.verifyEqual(id, 'mat2doc:ValueError', ...
                'from_xml("") must raise mat2doc:ValueError (empty string is a real string, not None)');
            testCase.verifyEqual(string(msg), string(rt.mtci_from_empty_msg), ...
                'from_xml("") message must be byte-verbatim ("... has no XML mapping for \'\'")');
        end

        % =============================================================== %
        % 4. WD_BREAK_TYPE member-alias -- Bar 4                           %
        % =============================================================== %

        function test_wd_break_type_text_wrapping_alias(testCase)
            % TEXT_WRAPPING resolves to LINE_CLEAR_ALL (name "LINE_CLEAR_ALL",
            % value 11), == LINE_CLEAR_ALL true, == LINE false (LINE == 6).
            ba = testCase.Oracle.break_alias;
            tw = mat2doc.enum.text.WD_BREAK_TYPE.TEXT_WRAPPING;
            testCase.verifyEqual(string(tw), string(ba.tw_name), ...
                'TEXT_WRAPPING.name must be "LINE_CLEAR_ALL" (Python same-value alias)');
            testCase.verifyEqual(double(tw.value), double(ba.tw_value), ...
                'TEXT_WRAPPING.value must be 11');
            testCase.verifyTrue(tw == mat2doc.enum.text.WD_BREAK_TYPE.LINE_CLEAR_ALL, ...
                'TEXT_WRAPPING == LINE_CLEAR_ALL must be true (same member object)');
            testCase.verifyFalse(tw == mat2doc.enum.text.WD_BREAK_TYPE.LINE, ...
                'TEXT_WRAPPING == LINE must be false');
            testCase.verifyEqual(double(mat2doc.enum.text.WD_BREAK_TYPE.LINE.value), ...
                double(ba.line_value), 'LINE.value must be 6');
        end

        function test_wd_break_type_iteration_ten_members(testCase)
            % Canonical iteration is exactly 10 members (the alias excluded), in
            % declaration order -- pinned against the oracle member_names list.
            ba = testCase.Oracle.break_alias;
            members = enumeration("mat2doc.enum.text.WD_BREAK_TYPE");
            names   = arrayfun(@string, members);
            testCase.verifyEqual(numel(members), double(ba.member_count), ...
                'WD_BREAK_TYPE must iterate exactly 10 canonical members (alias excluded)');
            testCase.verifyEqual(names(:), string(ba.member_names(:)), ...
                'WD_BREAK_TYPE member names must match the oracle in declaration order');
        end

        % =============================================================== %
        % 5. The 7 module aliases -- Bar 5 (identity + forwarding)         %
        % =============================================================== %

        function test_alias_wd_align_paragraph(testCase)
            ma = testCase.Oracle.mod_alias;
            testCase.verifyTrue(mat2doc.enum.text.WD_ALIGN_PARAGRAPH.CENTER == ...
                mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER, ...
                'WD_ALIGN_PARAGRAPH.CENTER must be identical to the canonical member');
            testCase.verifyEqual(string(mat2doc.enum.text.WD_ALIGN_PARAGRAPH.from_xml("both")), ...
                string(ma.wap_from_both), 'WD_ALIGN_PARAGRAPH.from_xml("both") must forward -> JUSTIFY');
            testCase.verifyEqual(mat2doc.enum.text.WD_ALIGN_PARAGRAPH.to_xml( ...
                mat2doc.enum.text.WD_ALIGN_PARAGRAPH.CENTER), ...
                string(ma.wap_to_center), 'WD_ALIGN_PARAGRAPH.to_xml(CENTER) must forward -> "center"');
        end

        function test_alias_wd_break(testCase)
            % WD_BREAK has NO statics (canonical WD_BREAK_TYPE is a plain enum);
            % it re-exports members incl. the TEXT_WRAPPING member-alias.
            ma = testCase.Oracle.mod_alias;
            testCase.verifyTrue(mat2doc.enum.text.WD_BREAK.PAGE == ...
                mat2doc.enum.text.WD_BREAK_TYPE.PAGE, ...
                'WD_BREAK.PAGE must be identical to WD_BREAK_TYPE.PAGE');
            testCase.verifyEqual(double(mat2doc.enum.text.WD_BREAK.PAGE.value), ...
                double(ma.wb_page_value), 'WD_BREAK.PAGE.value must be 7');
            testCase.verifyEqual(string(mat2doc.enum.text.WD_BREAK.TEXT_WRAPPING), ...
                string(ma.wb_text_wrapping_name), 'WD_BREAK.TEXT_WRAPPING must be the LINE_CLEAR_ALL alias');
        end

        function test_alias_wd_color(testCase)
            ma = testCase.Oracle.mod_alias;
            testCase.verifyTrue(mat2doc.enum.text.WD_COLOR.AUTO == ...
                mat2doc.enum.text.WD_COLOR_INDEX.AUTO, ...
                'WD_COLOR.AUTO must be identical to the canonical member');
            testCase.verifyEqual(string(mat2doc.enum.text.WD_COLOR.from_xml("red")), ...
                string(ma.wc_from_red), 'WD_COLOR.from_xml("red") must forward -> RED');
            testCase.verifyEqual(mat2doc.enum.text.WD_COLOR.to_xml( ...
                mat2doc.enum.text.WD_COLOR.AUTO), ...
                string(ma.wc_to_auto), 'WD_COLOR.to_xml(AUTO) must forward -> "default"');
        end

        function test_alias_wd_header_footer(testCase)
            ma = testCase.Oracle.mod_alias;
            testCase.verifyTrue(mat2doc.enum.section.WD_HEADER_FOOTER.PRIMARY == ...
                mat2doc.enum.section.WD_HEADER_FOOTER_INDEX.PRIMARY, ...
                'WD_HEADER_FOOTER.PRIMARY must be identical to the canonical member');
            testCase.verifyEqual(string(mat2doc.enum.section.WD_HEADER_FOOTER.from_xml("first")), ...
                string(ma.whf_from_first), 'WD_HEADER_FOOTER.from_xml("first") must forward -> FIRST_PAGE');
            testCase.verifyEqual(mat2doc.enum.section.WD_HEADER_FOOTER.to_xml(2), ...
                string(ma.whf_to_2), 'WD_HEADER_FOOTER.to_xml(2) must forward -> "first"');
        end

        function test_alias_wd_orient(testCase)
            ma = testCase.Oracle.mod_alias;
            testCase.verifyTrue(mat2doc.enum.section.WD_ORIENT.LANDSCAPE == ...
                mat2doc.enum.section.WD_ORIENTATION.LANDSCAPE, ...
                'WD_ORIENT.LANDSCAPE must be identical to the canonical member');
            testCase.verifyEqual(string(mat2doc.enum.section.WD_ORIENT.from_xml("portrait")), ...
                string(ma.wor_from_portrait), 'WD_ORIENT.from_xml("portrait") must forward -> PORTRAIT');
            testCase.verifyEqual(mat2doc.enum.section.WD_ORIENT.to_xml(1), ...
                string(ma.wor_to_1), 'WD_ORIENT.to_xml(1) must forward -> "landscape"');
        end

        function test_alias_wd_section(testCase)
            ma = testCase.Oracle.mod_alias;
            testCase.verifyTrue(mat2doc.enum.section.WD_SECTION.NEW_PAGE == ...
                mat2doc.enum.section.WD_SECTION_START.NEW_PAGE, ...
                'WD_SECTION.NEW_PAGE must be identical to the canonical member');
            testCase.verifyEqual(string(mat2doc.enum.section.WD_SECTION.from_xml("continuous")), ...
                string(ma.wsec_from_continuous), 'WD_SECTION.from_xml("continuous") must forward -> CONTINUOUS');
            testCase.verifyEqual(mat2doc.enum.section.WD_SECTION.to_xml(2), ...
                string(ma.wsec_to_2), 'WD_SECTION.to_xml(2) must forward -> "nextPage"');
        end

        function test_alias_mso_theme_color(testCase)
            ma = testCase.Oracle.mod_alias;
            testCase.verifyTrue(mat2doc.enum.dml.MSO_THEME_COLOR.ACCENT_1 == ...
                mat2doc.enum.dml.MSO_THEME_COLOR_INDEX.ACCENT_1, ...
                'MSO_THEME_COLOR.ACCENT_1 must be identical to the canonical member');
            testCase.verifyEqual(string(mat2doc.enum.dml.MSO_THEME_COLOR.from_xml("accent6")), ...
                string(ma.mtc_from_accent6), 'MSO_THEME_COLOR.from_xml("accent6") must forward -> ACCENT_6');
            testCase.verifyEqual(mat2doc.enum.dml.MSO_THEME_COLOR.to_xml( ...
                mat2doc.enum.dml.MSO_THEME_COLOR.ACCENT_6), ...
                string(ma.mtc_to_accent6), 'MSO_THEME_COLOR.to_xml(ACCENT_6) must forward -> "accent6"');
        end

        % =============================================================== %
        % 6. MSO_COLOR_TYPE (BaseEnum, no XML) -- Bar 6                    %
        % =============================================================== %

        function test_mso_color_type_int_str(testCase)
            % int(member) == ms_api_value and str(member) == "NAME (value)".
            mc = testCase.Oracle.mso_color;
            testCase.verifyEqual(double(mat2doc.enum.dml.MSO_COLOR_TYPE.RGB.value), ...
                double(mc.rgb_int), 'MSO_COLOR_TYPE.RGB int must be 1');
            testCase.verifyEqual(mat2doc.enum.dml.MSO_COLOR_TYPE.RGB.str_(), ...
                string(mc.rgb_str), 'str(RGB) must be "RGB (1)"');
            testCase.verifyEqual(double(mat2doc.enum.dml.MSO_COLOR_TYPE.THEME.value), ...
                double(mc.theme_int), 'MSO_COLOR_TYPE.THEME int must be 2');
            testCase.verifyEqual(mat2doc.enum.dml.MSO_COLOR_TYPE.THEME.str_(), ...
                string(mc.theme_str), 'str(THEME) must be "THEME (2)"');
            testCase.verifyEqual(double(mat2doc.enum.dml.MSO_COLOR_TYPE.AUTO.value), ...
                double(mc.auto_int), 'MSO_COLOR_TYPE.AUTO int must be 101');
            testCase.verifyEqual(mat2doc.enum.dml.MSO_COLOR_TYPE.AUTO.str_(), ...
                string(mc.auto_str), 'str(AUTO) must be "AUTO (101)"');
        end

        % =============================================================== %
        % 7. from_xml/to_xml round-trips + badint form -- Bar 2/7          %
        % =============================================================== %

        function test_wd_paragraph_alignment_roundtrip(testCase)
            rt = testCase.Oracle.roundtrip;
            testCase.verifyEqual(string(mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.from_xml("center")), ...
                string(rt.wpa_from_center), 'from_xml("center") -> CENTER');
            testCase.verifyEqual(string(mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.from_xml("both")), ...
                string(rt.wpa_from_both), 'from_xml("both") -> JUSTIFY');
            testCase.verifyEqual(mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.to_xml(3), ...
                string(rt.wpa_to_3), 'to_xml(3) -> "both"');
            [id, msg] = raiseAndCapture(@() mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.to_xml(999));
            testCase.verifyEqual(id, 'mat2doc:ValueError');
            testCase.verifyEqual(string(msg), string(rt.wpa_to_999_msg), ...
                'to_xml(999) must raise the unquoted "999 is not a valid WD_PARAGRAPH_ALIGNMENT"');
        end

        function test_wd_tab_alignment_roundtrip(testCase)
            rt = testCase.Oracle.roundtrip;
            testCase.verifyEqual(string(mat2doc.enum.text.WD_TAB_ALIGNMENT.from_xml("start")), ...
                string(rt.wta_from_start), 'from_xml("start") -> START');
            testCase.verifyEqual(mat2doc.enum.text.WD_TAB_ALIGNMENT.to_xml(104), ...
                string(rt.wta_to_104), 'to_xml(104) -> "start"');
            [id, msg] = raiseAndCapture(@() mat2doc.enum.text.WD_TAB_ALIGNMENT.to_xml(999));
            testCase.verifyEqual(id, 'mat2doc:ValueError');
            testCase.verifyEqual(string(msg), string(rt.wta_to_999_msg), ...
                'to_xml(999) must raise "999 is not a valid WD_TAB_ALIGNMENT"');
        end

        function test_wd_tab_leader_roundtrip(testCase)
            rt = testCase.Oracle.roundtrip;
            testCase.verifyEqual(string(mat2doc.enum.text.WD_TAB_LEADER.from_xml("middleDot")), ...
                string(rt.wtl_from_middledot), 'from_xml("middleDot") -> MIDDLE_DOT');
            testCase.verifyEqual(mat2doc.enum.text.WD_TAB_LEADER.to_xml(5), ...
                string(rt.wtl_to_5), 'to_xml(5) -> "middleDot"');
        end

        function test_wd_line_spacing_unmapped_and_badxml(testCase)
            % UNMAPPED is truthy (returned by to_xml); from_xml("UNMAPPED")
            % first-matches SINGLE; an unmapped token raises the verbatim message.
            rt = testCase.Oracle.roundtrip;
            testCase.verifyEqual(string(mat2doc.enum.text.WD_LINE_SPACING.from_xml("UNMAPPED")), ...
                string(rt.wls_from_unmapped), 'from_xml("UNMAPPED") -> SINGLE (first match)');
            testCase.verifyEqual(mat2doc.enum.text.WD_LINE_SPACING.to_xml(0), ...
                string(rt.wls_to_single), 'to_xml(0) -> "UNMAPPED"');
            testCase.verifyEqual(mat2doc.enum.text.WD_LINE_SPACING.to_xml(4), ...
                string(rt.wls_to_4), 'to_xml(4) -> "exact"');
            [id, msg] = raiseAndCapture(@() mat2doc.enum.text.WD_LINE_SPACING.from_xml("nosuch"));
            testCase.verifyEqual(id, 'mat2doc:ValueError');
            testCase.verifyEqual(string(msg), string(rt.wls_from_nosuch_msg), ...
                'from_xml("nosuch") must raise the verbatim WD_LINE_SPACING mapping message');
        end

        function test_wd_header_footer_index_roundtrip(testCase)
            rt = testCase.Oracle.roundtrip;
            testCase.verifyEqual(string(mat2doc.enum.section.WD_HEADER_FOOTER_INDEX.from_xml("default")), ...
                string(rt.whfi_from_default), 'from_xml("default") -> PRIMARY');
            testCase.verifyEqual(mat2doc.enum.section.WD_HEADER_FOOTER_INDEX.to_xml(1), ...
                string(rt.whfi_to_1), 'to_xml(1) -> "default"');
        end

        function test_wd_orientation_roundtrip(testCase)
            rt = testCase.Oracle.roundtrip;
            testCase.verifyEqual(string(mat2doc.enum.section.WD_ORIENTATION.from_xml("landscape")), ...
                string(rt.wo_from_landscape), 'from_xml("landscape") -> LANDSCAPE');
            testCase.verifyEqual(mat2doc.enum.section.WD_ORIENTATION.to_xml(1), ...
                string(rt.wo_to_1), 'to_xml(1) -> "landscape"');
        end

        function test_wd_section_start_roundtrip(testCase)
            rt = testCase.Oracle.roundtrip;
            testCase.verifyEqual(string(mat2doc.enum.section.WD_SECTION_START.from_xml("nextPage")), ...
                string(rt.wss_from_nextpage), 'from_xml("nextPage") -> NEW_PAGE');
            testCase.verifyEqual(mat2doc.enum.section.WD_SECTION_START.to_xml(2), ...
                string(rt.wss_to_2), 'to_xml(2) -> "nextPage"');
        end

        % =============================================================== %
        % 8. Non-ASCII edge / error path                                  %
        % =============================================================== %

        function test_from_xml_nonascii_bogus_raises(testCase)
            % Non-ASCII edge: query "cafe" with e = U+00E9 built via char(233) so
            % the expectation is independent of THIS file's source encoding. The
            % unmapped token raises mat2doc:ValueError with the query rendered
            % byte-identically inside the literal single quotes (UTF-8 pipeline).
            query    = "caf" + string(char(233));
            expected = "WD_PARAGRAPH_ALIGNMENT has no XML mapping for '" + query + "'";
            [id, msg] = raiseAndCapture(@() mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.from_xml(query));
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
            % fields (num2cell -> cell of structs) or a CELL of structs when a
            % null xml_value makes the elements heterogeneous; normalize both.
            raw = testCase.Oracle.members.(char(shortName));
            if iscell(raw)
                mlist = raw;
            else
                mlist = num2cell(raw);
            end
        end
    end
end

% ===================== file-local helpers ============================== %

function [id, msg] = raiseAndCapture(fn)
    % Run fn() expecting it to raise; return (identifier, message) as char, or
    % ('','') if it did not raise. (Idiom from tests\enum\Test_p3_1_enum_base.m.)
    id = ''; msg = '';
    try
        fn();
    catch ME
        id = ME.identifier;
        msg = ME.message;
    end
end

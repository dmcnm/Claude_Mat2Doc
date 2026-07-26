classdef Test_p3_1_enum_base < matlab.unittest.TestCase
% TEST_P3_1_ENUM_BASE  Gate-4 permanent unit tests for Mat2Doc P3-1
%   (the enum base machinery: mat2doc.enum.base.BaseEnum + BaseXmlEnum).
%
%   Freezes the P3-1 guarantees every concrete docx enum (P3-3..P3-6:
%   WD_TAB_ALIGNMENT, WD_COLOR_INDEX, WD_SECTION_START, ...) will lean on:
%     - mat2doc.enum.base.BaseEnum      -- plain MS-API enum: int(member) ==
%       double(member.value); str(member) == "NAME (value)".
%     - mat2doc.enum.base.BaseXmlEnum   -- adds xml_value + static from_xml_/
%       to_xml_ machinery, ported to python-docx 1.2.0 semantics (NOT pptx's).
%
%   THE LOAD-BEARING DOCX DELTA (why this must never regress to the pptx form):
%   docx `from_xml` has NO None/empty short-circuit (pptx added one). It is a
%   straight None-tolerant equality scan, so:
%     * from_xml(None)  -> the member whose xml_value IS None  (here INHERIT).
%       The pptx guard would instead RAISE here, which would break
%       WD_COLOR_INDEX.INHERITED downstream. This is the single most important
%       pin in the class (test_from_xml_none_returns_none_member).
%     * from_xml("")    -> the member whose xml_value is the empty string (BLANK).
%     * None and "" are DISTINCT (H3 tri-state): they route to distinct members
%       and BOTH are rejected by to_xml ("has no XML representation").
%
%   THE F1 PIN (Gate-2 auditor fix, must not regress): the string-input branch
%   of to_xml renders CPython's `%r` repr quotes -- to_xml("center") raises
%   'center' is not a valid s0016_VXml  WITH the single quotes
%   (test_to_xml_string_input_raises_with_repr_quotes).
%
%   REGRESSION GUARD: BaseXmlEnum has NO `validate` method in docx v1.2.0 (pptx
%   does). A future edit that re-adds the pptx validate() is caught by
%   test_basexmlenum_has_no_validate_method.
%
%   Provenance (Gate-1..3, all 2026-07-26):
%     * Audit    : validation\mat2doc\audit_P3-1_enum_base.md  (Porter Gate-1 +
%                  Opus Gate-2 adversarial APPROVE; F1 to_xml repr-quote fix
%                  applied inline; the 4 docx-vs-pptx deltas documented).
%     * Validate : validation\mat2doc\validate_P3-1_enum_base.md  (Gate-3 PASS,
%                  probe_diff 34/34 facts byte-identical, 0 new D-numbers,
%                  regression 414/414).
%     * Sample enums (Equivalence fixtures, copied verbatim into this folder so
%       the suite is self-contained): s0016_VXml.m / s0016_VXmlNoNone.m /
%       s0016_VPlain.m, provenance validation\mat2doc\scenarios\s0016_V*.m. The
%       concrete class NAMES are kept byte-identical to the Python twin
%       (s0016_enum_base_probes.py), so every ValueError message -- which
%       interpolates cls.__name__ / shortName_ -- is a genuine byte-for-byte
%       comparison against the frozen references\s0016\probe.json oracle.
%
%   Coverage taxonomy
%   -----------------
%   * Nominal   -- from_xml("center")->CENTER; to_xml(1)/(member)->"center";
%     BaseEnum int/str/name.
%   * Edge      -- from_xml(None)->None member (the delta); from_xml("")->empty
%     member; non-ASCII "café" error path; single None/empty member absent ->
%     raise; error paths verify the mat2doc:ValueError IDENTIFIER + verbatim
%     message (not merely that it throws); value-0 member (CONTINUOUS).
%   * Equivalence -- messages pinned byte-verbatim to the frozen s0016 oracle
%     (validate_P3-1 bars 1-5), class names preserved for exact interpolation.
%   * Regression -- hard-coded expected xml_value strings / str() forms / the
%     verbatim CPython-derived "is not a valid" messages; the no-validate guard.
%
%   Deviations exercised: 0 new D-numbers (Gate-3). Emitted errors ride the
%   standing D-005 adopt-only mat2doc:ValueError identifier convention (no new
%   ledger row). No package bytes -- nothing serialized, so no L-ladder leg.
%
%   EXPECTED BENIGN WARNING (not a failure): the three co-located sample enums
%   extend mat2doc.enum.base.Base(Xml)Enum, which is not on the path at
%   suite-CREATION time (before TestClassSetup's PathFixture runs). testsuite's
%   folder scan therefore emits a "superclass ... cannot be found ... was
%   excluded" warning for each and CORRECTLY declines to collect them as tests
%   (they are helper fixtures, not TestCase classes). At RUN time -- after the
%   PathFixture is active -- they resolve and every case passes. This mirrors the
%   loose co-located helper pattern of tests\shared\Test_p2_1_proxy_tier.m
%   (LazyHost_p2_1.m / PartStub_p2_1.m); those don't warn only because they do
%   not extend a +mat2doc class. The warning is cosmetic and adds 0 to the suite
%   count (439 = 414 baseline + 25 here).
%
%   Determinism: no network, no absolute paths, no file writes at all (pure
%   value/behaviour machinery). The +mat2doc package and the co-located sample
%   enums resolve via the MANDATORY PathFixture(worktree-root) added in
%   TestClassSetup (WP9-F4 lesson).

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied from the proven tests\shared\Test_p2_1_proxy_tier.m.
            % The co-located sample enums (s0016_V*.m) resolve because their
            % folder is the one runtests operates in.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\enum
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. from_xml -- the KEY DOCX DELTA (no None/empty short-circuit)  %
        %    (validate_P3-1 Bar 1)                                         %
        % =============================================================== %

        function test_from_xml_center_returns_member(testCase)
            % Nominal happy path: from_xml("center") -> CENTER (documented example).
            m = s0016_VXml.from_xml("center");
            testCase.verifyEqual(m, s0016_VXml.CENTER, ...
                'from_xml("center") must return the CENTER member');
            testCase.verifyEqual(string(m), "CENTER", ...
                'string(member) must be the member name CENTER');
            testCase.verifyEqual(double(m.value), 1, ...
                'CENTER.value must be its MS-API integer 1');
        end

        function test_from_xml_both_returns_member(testCase)
            % Second nominal mapping: from_xml("both") -> BOTH (xml_value "both").
            m = s0016_VXml.from_xml("both");
            testCase.verifyEqual(m, s0016_VXml.BOTH, ...
                'from_xml("both") must return the BOTH member');
        end

        function test_from_xml_none_returns_none_member(testCase)
            % THE LOAD-BEARING DOCX DELTA. from_xml(None) returns the member whose
            % xml_value IS None -- here INHERIT -- NOT a short-circuit raise (the
            % pptx guard would raise, breaking WD_COLOR_INDEX.INHERITED). Both
            % None spellings map to it: the [] None sentinel AND <missing>.
            % (validate_P3-1 Bar 1: from_xml(None) -> INHERIT.)
            mFromEmpty   = s0016_VXml.from_xml([]);             % [] None sentinel
            mFromMissing = s0016_VXml.from_xml(string(missing)); % <missing> None
            testCase.verifyEqual(mFromEmpty, s0016_VXml.INHERIT, ...
                'from_xml([]) must return the None-xml member INHERIT (docx delta -- NOT a raise)');
            testCase.verifyEqual(mFromMissing, s0016_VXml.INHERIT, ...
                'from_xml(<missing>) must also return INHERIT (None-tolerant scan)');
        end

        function test_from_xml_empty_returns_empty_member(testCase)
            % from_xml("") returns the member whose xml_value is the REAL empty
            % string -- here BLANK -- distinct from the None member.
            % (validate_P3-1 Bar 1: from_xml("") -> BLANK.)
            m = s0016_VXml.from_xml("");
            testCase.verifyEqual(m, s0016_VXml.BLANK, ...
                'from_xml("") must return the empty-string member BLANK');
        end

        function test_from_xml_none_and_empty_are_distinct(testCase)
            % H3 tri-state crux: from_xml(None) and from_xml("") route to DISTINCT
            % members (INHERIT != BLANK). If the None sentinel and "" ever
            % collapsed, these would alias. (validate_P3-1 Bar 4.)
            mNone  = s0016_VXml.from_xml([]);
            mEmpty = s0016_VXml.from_xml("");
            testCase.verifyNotEqual(mNone, mEmpty, ...
                'from_xml(None) and from_xml("") must be DISTINCT members (H3 tri-state)');
        end

        function test_from_xml_bogus_raises_valueerror_verbatim(testCase)
            % Error path: an unmapped token raises mat2doc:ValueError with the
            % byte-verbatim docx message (str-interpolation, NOT repr, of the
            % query inside literal single quotes). Verifies IDENTIFIER + message.
            % (validate_P3-1 Bar 1.)
            [id, msg] = raiseAndCapture(@() s0016_VXml.from_xml("bogus"));
            testCase.verifyEqual(id, 'mat2doc:ValueError', ...
                'unmapped from_xml must raise mat2doc:ValueError (D-005 convention)');
            testCase.verifyEqual(msg, 's0016_VXml has no XML mapping for ''bogus''', ...
                'from_xml message must be byte-verbatim to the python-docx oracle');
        end

        function test_from_xml_nonascii_raises_verbatim(testCase)
            % Non-ASCII edge: query "café" (é = U+00E9) built from char(233) so the
            % expectation is independent of THIS file's source encoding -- exactly
            % the validator's UTF-8-pipeline proof. Message renders byte-identical.
            % (validate_P3-1 Bar 1: from_xml("café").)
            query = "caf" + string(char(233));
            expected = "s0016_VXml has no XML mapping for '" + query + "'";
            [id, msg] = raiseAndCapture(@() s0016_VXml.from_xml(query));
            testCase.verifyEqual(id, 'mat2doc:ValueError', ...
                'non-ASCII unmapped query must still raise mat2doc:ValueError');
            testCase.verifyEqual(string(msg), expected, ...
                'the non-ASCII query must render byte-identically in the message (UTF-8 pipeline)');
        end

        function test_from_xml_none_without_none_member_raises(testCase)
            % Error path -- the raise-on-None branch: a class with NO None member
            % (s0016_VXmlNoNone) raises on from_xml(None); the query str is "None".
            % (validate_P3-1 Bar 1: no-None-member None.)
            [id, msg] = raiseAndCapture(@() s0016_VXmlNoNone.from_xml([]));
            testCase.verifyEqual(id, 'mat2doc:ValueError', ...
                'from_xml(None) with no None member must raise mat2doc:ValueError');
            testCase.verifyEqual(msg, 's0016_VXmlNoNone has no XML mapping for ''None''', ...
                'None query with no None member must render ''None'' verbatim');
        end

        function test_from_xml_empty_without_empty_member_raises(testCase)
            % Error path: a class with no empty-string member raises on
            % from_xml("") with the empty query interpolated between the quotes.
            % (validate_P3-1 Bar 1: no-empty-member "".)
            [id, msg] = raiseAndCapture(@() s0016_VXmlNoNone.from_xml(""));
            testCase.verifyEqual(id, 'mat2doc:ValueError', ...
                'from_xml("") with no empty member must raise mat2doc:ValueError');
            testCase.verifyEqual(msg, 's0016_VXmlNoNone has no XML mapping for ''''', ...
                'empty query with no empty member must render '''' (two quotes, nothing between)');
        end

        % =============================================================== %
        % 2. to_xml -- the F1 repr-quote fix + falsy guard                 %
        %    (validate_P3-1 Bar 2)                                         %
        % =============================================================== %

        function test_to_xml_int_returns_xml_value(testCase)
            % Nominal: cls(int) value-lookup then xml_value. to_xml(1)->"center",
            % to_xml(3)->"both". (validate_P3-1 Bar 2.)
            testCase.verifyEqual(s0016_VXml.to_xml(1), "center", ...
                'to_xml(1) must resolve CENTER and return its xml_value "center"');
            testCase.verifyEqual(s0016_VXml.to_xml(3), "both", ...
                'to_xml(3) must resolve BOTH and return its xml_value "both"');
        end

        function test_to_xml_member_returns_xml_value(testCase)
            % Nominal: cls(member) identity path then xml_value.
            % (validate_P3-1 Bar 2.)
            testCase.verifyEqual(s0016_VXml.to_xml(s0016_VXml.CENTER), "center", ...
                'to_xml(CENTER member) must return "center"');
        end

        function test_to_xml_none_xml_member_raises(testCase)
            % Falsy guard: a resolved member whose xml_value IS None raises "has no
            % XML representation" -- reachable both by MEMBER and by its INT (5).
            % (validate_P3-1 Bar 2.)
            [id1, msg1] = raiseAndCapture(@() s0016_VXml.to_xml(s0016_VXml.INHERIT));
            testCase.verifyEqual(id1, 'mat2doc:ValueError');
            testCase.verifyEqual(msg1, 's0016_VXml.INHERIT has no XML representation', ...
                'None-xml member (by member) must raise the verbatim representation message');
            [id2, msg2] = raiseAndCapture(@() s0016_VXml.to_xml(5));
            testCase.verifyEqual(id2, 'mat2doc:ValueError');
            testCase.verifyEqual(msg2, 's0016_VXml.INHERIT has no XML representation', ...
                'None-xml member (by int 5) must raise the same verbatim message');
        end

        function test_to_xml_empty_xml_member_raises(testCase)
            % Falsy guard also fires for the REAL empty string "" (Python `if not
            % xml_value` is falsy for both None and ""). BLANK by member and int 7.
            % (validate_P3-1 Bar 2 + Bar 4: "" is rejected by to_xml too.)
            [id1, msg1] = raiseAndCapture(@() s0016_VXml.to_xml(s0016_VXml.BLANK));
            testCase.verifyEqual(id1, 'mat2doc:ValueError');
            testCase.verifyEqual(msg1, 's0016_VXml.BLANK has no XML representation', ...
                'empty-xml member (by member) must raise the representation message');
            [id2, msg2] = raiseAndCapture(@() s0016_VXml.to_xml(7));
            testCase.verifyEqual(id2, 'mat2doc:ValueError');
            testCase.verifyEqual(msg2, 's0016_VXml.BLANK has no XML representation', ...
                'empty-xml member (by int 7) must raise the same verbatim message');
        end

        function test_to_xml_bad_int_raises(testCase)
            % Error path: an int with no matching member raises the CPython-stdlib
            % form "<v> is not a valid <Cls>"; an int renders UNQUOTED.
            % (validate_P3-1 Bar 2.)
            [id, msg] = raiseAndCapture(@() s0016_VXml.to_xml(999));
            testCase.verifyEqual(id, 'mat2doc:ValueError');
            testCase.verifyEqual(msg, '999 is not a valid s0016_VXml', ...
                'bad int must raise the unquoted "999 is not a valid ..." form');
        end

        function test_to_xml_none_raises(testCase)
            % Error path: None ([]) is not a valid member value; renders as the
            % bare word None (unquoted). (validate_P3-1 Bar 2.)
            [id, msg] = raiseAndCapture(@() s0016_VXml.to_xml([]));
            testCase.verifyEqual(id, 'mat2doc:ValueError');
            testCase.verifyEqual(msg, 'None is not a valid s0016_VXml', ...
                'None value must raise the unquoted "None is not a valid ..." form');
        end

        function test_to_xml_string_input_raises_with_repr_quotes(testCase)
            % THE F1 PIN (Gate-2 auditor fix, must never regress): a STRING passed
            % to to_xml is not a valid member value and CPython's EnumCls(value)
            % raises `%r is not a valid %s`; for a string %r yields SINGLE QUOTES.
            % to_xml("center") must therefore raise WITH the quotes -- distinct
            % from the unquoted int/None forms above. (validate_P3-1 Bar 2, F1.)
            [id, msg] = raiseAndCapture(@() s0016_VXml.to_xml("center"));
            testCase.verifyEqual(id, 'mat2doc:ValueError');
            testCase.verifyEqual(msg, '''center'' is not a valid s0016_VXml', ...
                'F1: a string value must render WITH CPython %r single quotes (''center'')');
        end

        % =============================================================== %
        % 3. BaseEnum (non-xml, MS-API-only path)                         %
        %    (validate_P3-1 Bar 3)                                         %
        % =============================================================== %

        function test_baseenum_int_value(testCase)
            % int(member) ports as double(member.value): NEW_PAGE->2, and the
            % value-0 member CONTINUOUS->0 exercises the non-truthy value path.
            % (validate_P3-1 Bar 3.)
            testCase.verifyEqual(double(s0016_VPlain.NEW_PAGE.value), 2, ...
                'double(NEW_PAGE.value) must be its MS-API int 2');
            testCase.verifyEqual(double(s0016_VPlain.CONTINUOUS.value), 0, ...
                'double(CONTINUOUS.value) must be 0 (non-truthy MS-API value)');
        end

        function test_baseenum_str_form(testCase)
            % str(member) == f"{name} ({value})": "NEW_PAGE (2)" / "CONTINUOUS (0)".
            % (validate_P3-1 Bar 3.)
            testCase.verifyEqual(s0016_VPlain.NEW_PAGE.str_(), "NEW_PAGE (2)", ...
                'str_(NEW_PAGE) must be "NEW_PAGE (2)"');
            testCase.verifyEqual(s0016_VPlain.CONTINUOUS.str_(), "CONTINUOUS (0)", ...
                'str_(CONTINUOUS) must be "CONTINUOUS (0)"');
        end

        function test_baseenum_name(testCase)
            % string(member) yields the member name (Python self.name).
            % (validate_P3-1 Bar 3.)
            testCase.verifyEqual(string(s0016_VPlain.NEW_PAGE), "NEW_PAGE", ...
                'string(member) must be the member name NEW_PAGE');
        end

        function test_basexmlenum_str_form(testCase)
            % BaseXmlEnum inherits the same str form: CENTER -> "CENTER (1)".
            % (validate_P3-1 Bar 5 / Bar 3 idiom carried to the xml subclass.)
            testCase.verifyEqual(s0016_VXml.CENTER.str_(), "CENTER (1)", ...
                'str_(CENTER) must be "CENTER (1)"');
        end

        % =============================================================== %
        % 4. H3 tri-state property inspection (None vs "" distinct)       %
        %    (validate_P3-1 Bar 4)                                         %
        % =============================================================== %

        function test_h3_none_member_xml_value_ismissing(testCase)
            % The None member stores <missing> (the Python-None sentinel):
            % ismissing -> true. (validate_P3-1 Bar 4.)
            testCase.verifyTrue(ismissing(s0016_VXml.INHERIT.xml_value), ...
                'INHERIT.xml_value must be <missing> (the Python None sentinel)');
        end

        function test_h3_empty_member_xml_value_is_real_empty_string(testCase)
            % The empty member stores the REAL empty string: NOT missing, strlen 0.
            % Distinguishes "" from None at the property level. (validate_P3-1 Bar 4.)
            xv = s0016_VXml.BLANK.xml_value;
            testCase.verifyFalse(ismissing(xv), ...
                'BLANK.xml_value must NOT be <missing> (it is a real "")');
            testCase.verifyEqual(strlength(xv), 0, ...
                'BLANK.xml_value must be the empty string (length 0)');
        end

        % =============================================================== %
        % 5. Regression guard -- NO validate() (docx v1.2.0, not pptx)     %
        % =============================================================== %

        function test_basexmlenum_has_no_validate_method(testCase)
            % REGRESSION GUARD: python-docx 1.2.0 BaseXmlEnum has NO `validate`
            % classmethod (python-pptx DOES). A future edit that copies the pptx
            % base verbatim would re-add it; this test goes RED if that happens.
            mc = ?mat2doc.enum.base.BaseXmlEnum;
            names = string({mc.MethodList.Name});
            testCase.verifyFalse(any(names == "validate"), ...
                'BaseXmlEnum must have NO validate() method (docx v1.2.0 delta vs pptx)');
            % Positive control: the shared machinery IS present (guards a typo'd
            % rename from silently passing the negative check above).
            testCase.verifyTrue(any(names == "from_xml_") && any(names == "to_xml_"), ...
                'the shared from_xml_/to_xml_ machinery must be present');
        end

        % =============================================================== %
        % 6. Idiom soundness (cls(int) lookup; string(member)->name)      %
        %    (validate_P3-1 Bar 5)                                         %
        % =============================================================== %

        function test_idiom_cls_int_lookup_first_declared(testCase)
            % cls(intval) lookup: the value-scan resolves the member with a
            % matching MS-API int (find(...,1) -> first-declared, matching Python
            % alias resolution). Proven through to_xml's resolve path: int 3 -> the
            % BOTH member's xml_value. (validate_P3-1 Bar 5.)
            testCase.verifyEqual(s0016_VXml.to_xml(3), "both", ...
                'cls(3) int-lookup must resolve BOTH (first member with value 3)');
        end

        function test_idiom_string_member_is_name(testCase)
            % string(member) -> member name for the xml enum too (Python self.name);
            % char/string deliberately NOT overridden so the name stays accessible.
            % (validate_P3-1 Bar 5.)
            testCase.verifyEqual(string(s0016_VXml.BOTH), "BOTH", ...
                'string(member) must yield the member name BOTH');
        end

    end
end

% ===================== file-local helpers ============================== %

function [id, msg] = raiseAndCapture(fn)
    % Run fn() expecting it to raise; return (identifier, message) as char, or
    % ('','') if it did not raise. (Idiom copied from Test_p2_1_proxy_tier.m.)
    id = ''; msg = '';
    try
        fn();
    catch ME
        id = ME.identifier;
        msg = ME.message;
    end
end

classdef Test_p1_7_coreprops < matlab.unittest.TestCase
% TEST_P1_7_COREPROPS  Gate-4 permanent unit tests for Mat2Doc P1-7 (coreprops).
%
%   Freezes the P1-7 core-properties tier ported in +mat2doc\ from python-docx
%   v1.2.0 src/docx/oxml/coreprops.py (CT_CoreProperties),
%   src/docx/opc/coreprops.py (CoreProperties), src/docx/opc/parts/coreprops.py
%   (CorePropertiesPart + default()), the shared helper +mat2doc\+shared\pyStr.m,
%   and the one registry.m main-map row (cp:coreProperties -> CT_CoreProperties,
%   docx/oxml/__init__.py:96).
%
%   Provenance (Gate-1..3, all 2026-07-26):
%     * Audit    : validation\mat2doc\audit_P1-7_coreprops.md
%                  (Porter Gate-1 + Opus Gate-2 adversarial APPROVE with the 3
%                  inline fixes F1 (lowercase-t W3CDTF), F2 (pyStr in
%                  set_element_text_), F3 (D-002 doc) applied).
%     * Validate : validation\mat2doc\validate_P1-7_coreprops.md (Gate-3 PASS,
%                  all 5 charter bars proven, 0 FAIL, 0 new D-numbers, regression
%                  329/329). Both byte-scenarios three-way byte-identical.
%     * Scenarios: validation\mat2doc\scenarios\
%                    s0009_coreprops_reserialize.{py,m}  (bar 1, 721 B)
%                    s0010_xsi_hoist_battery.{py,m}      (bar 3, 681 B)
%                    s0011_coreprops_probes.{py,m}       (bars 2/4/5 + pyStr)
%     * Frozen refs (python-docx 1.2.0 oracle, frozen ONCE):
%         references\s0009\parts\docProps\core.xml (721 B, sha d14be828...) ==
%           references\s0001\parts\docProps\core.xml (the M1 core.xml)
%         references\s0010\parts\docProps\core.xml (681 B, sha 3be041bb...)
%         references\s0011\probe.json (exact-match) + probe_info.json (divergences)
%       The two byte-parts are copied byte-for-byte into tests\coreprops\data\
%       (core721.bin, core681.bin -- SHA-verified == the frozen refs; co-located
%       `* binary` .gitattributes) so this suite is self-contained relative to
%       the worktree.
%
%   Coverage taxonomy
%   -----------------
%   * Regression / L1 byte (bar 1, the M1 pin): parse the frozen 721 B core.xml
%     -> CT_CoreProperties -> serialize_part_xml -> byte-identical + SHA-256
%     d14be828... . SHA-256 equality IS byte-identity (collision-resistant), an
%     L1 assertion. RED on ANY serializer/coreprops drift.
%   * Regression / L1 byte (bar 3, the xsi-hoist pin): fresh new() + the
%     created/modified/title/author/keywords/revision battery -> serialize ->
%     byte-identical + SHA-256 3be041bb..., 681 B, to the frozen s0010 ref. Pins
%     xmlns:xsi once-on-root + cp/dc/dcterms/xsi order + xsi:type on the 2 date
%     children (H8 hoist, D-serializer-nsdecl). Also re-proves H7 escaping (& < >)
%     and H2 non-ASCII (e-acute + emoji U+1F600) on GENERATED content.
%   * Equivalence (bar 2, s0011): the 36-vector W3CDTF wall-clock grammar (never
%     tzinfo -- the getters are naive by design, VERIFY-tz), asserted against the
%     frozen s0011 probe.json wall-clock outputs.
%   * Regression (bar 4, D-002): revision getter 11-exact + the 3 signed-D-002
%     divergences pinned as deliberate-red guards; setter positive-only else
%     mat2doc:ValueError with the verbatim message incl. the pyStr type-token.
%   * Regression (F2 pin, H14): set_element_text_ routes through pyStr -- 1/3 ->
%     "0.3333333333333333" (NOT "0.33333"); a logical -> "True"/"False" (NOT
%     "true"/"false"). Locks the F2 fix against regression.
%   * Equivalence (bar 5, s0011): CoreProperties 15-prop delegation (set ->
%     read-back) + write-through to the element; default() structure (excluding
%     the wall-clock `modified`, D-coreprops-time; default() is M1-unreachable).
%   * Equivalence (aux, s0011): pyStr 20-vector CPython str() battery; strftime
%     %S second truncation.
%
%   Deviations exercised (adopt-only, ZERO new D-numbers -- Gate-3): D-001 (own
%   OOXML parser/serializer, via parse_xml / serialize_part_xml), D-002 (the
%   ASCII-[0-9] int/date grammar -- the 3 revision divergences below are inside
%   the signed row's named scope), D-serializer-nsdecl (the xsi hoist, byte-proven
%   by the 681 B pin), D-coreprops-time (default()'s wall-clock `modified`,
%   excluded from the pinned structure). The 721 B + 681 B L1 results prove ZERO
%   output-visible divergence.
%
%   Determinism: no network, no absolute paths -- the byte fixtures resolve
%   relative to this file via fileparts(mfilename('fullpath')). No temp files are
%   written (all serialization is in-memory). Fixtures are read BINARY ('r','n').

    properties (Constant)
        % SHA-256 of the frozen references (== byte-identity). From the Gate-3
        % report and the co-located data\ fixtures.
        SHA_CORE721 = "d14be8284e406d14dc2a576b19f15e7f637b955f9e27a3d9ed871ec99606f7ba"
        SHA_CORE681 = "3be041bb7420f66ab139419fd48abed9a3f606f86003ae5ca8c6f3fc9747d802"

        % The 36 W3CDTF input vectors (s0011_coreprops_probes.m wvec, 1:1) and
        % their frozen wall-clock outputs (references\s0011\probe.json
        % w3cdtf_wallclock). "" is the JSON null sentinel here (getter returns []).
        % Index 33 (2003+newline) is built at runtime (a Constant cannot hold a
        % literal newline cleanly); see w3cdtfVectors().
        W3CDTF_EXPECT = [ ...
            "2003-01-01 00:00:00"; "2003-12-01 00:00:00"; "2003-12-31 00:00:00"; ...
            "2003-12-31 10:14:55"; "2003-12-31 10:14:55"; "2003-12-31 18:14:55"; ...
            "2003-12-31 04:44:55"; "2003-12-31 10:14:55"; "2003-12-31 10:44:55"; ...
            "2003-12-31 10:14:55"; "2003-12-31 10:14:55"; ""; ...
            ""; "2013-12-23 23:15:00"; ""; ""; ""; ""; ""; ""; ""; ...
            "0001-01-01 00:00:00"; "9999-01-01 00:00:00"; "0999-01-01 00:00:00"; ...
            ""; ""; "2003-01-01 00:00:00"; "2003-01-02 00:00:00"; ...
            "2003-01-02 03:04:05"; ""; ""; ""; ""; ""; ...
            "2003-12-31 10:14:55"; "2003-12-31 10:14:55"]

        % Revision getter EXACT subset (s0011 revExact / revision_get). "" and
        % "<ABSENT>"/"<EMPTYNONE>" are structural sentinels; see revisionOf().
        REV_EXACT_IN  = ["<ABSENT>", "<EMPTYNONE>", "1", " 42 ", "+7", "-5", ...
                         "abc", "2.5", "7\n", "0x1F", "007"]
        REV_EXACT_OUT = [0, 0, 1, 42, 7, 0, 0, 0, 7, 0, 7]

        % pyStr 20-vector battery (s0011 pvec) + frozen CPython str() outputs
        % (references\s0011\probe.json pystr).
        PYSTR_KIND = ["int" "int" "int" "int" "float" "float" "float" "float" ...
            "float" "float" "float" "float" "float" "float" "float" "float" ...
            "float" "float" "float" "float"]
        PYSTR_EXPECT = ["0" "7" "-5" "914400" "0.5" "2.0" "-0.0" "0.1" ...
            "0.3333333333333333" "1e+16" "1000000000000000.0" "0.0001" "1e-05" ...
            "inf" "-inf" "nan" "5e-324" "1.7976931348623157e+308" "123456.789" ...
            "2.5e+16"]
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied from tests\opc\Test_p1_6b_package_part.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\coreprops
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % Bar 1 -- core.xml 721 B re-serialize (the M1 pin)               %
        % =============================================================== %

        function test_core721_root_is_CT_CoreProperties(testCase)
            % Regression (bar 1): the registry routes the parsed <cp:coreProperties>
            % root to mat2doc.oxml.coreprops.CT_CoreProperties (docx/oxml/
            % __init__.py:96 registration). Leaf class name matches Python __name__.
            root = mat2doc.oxml.parse_xml(testCase.loadFixture('core721'));
            testCase.verifyEqual(class(root), ...
                'mat2doc.oxml.coreprops.CT_CoreProperties', ...
                'parsed cp:coreProperties root must be CT_CoreProperties (registry)');
        end

        function test_core721_reserialize_byte_identical(testCase)
            % Regression / L1 (bar 1, THE M1 pin): parse the frozen 721 B core.xml
            % and re-serialize through serialize_part_xml -> byte-identical + SHA-256
            % d14be828... . SHA equality is a byte-level (L1) assertion. RED on ANY
            % serializer/coreprops drift. (validate_P1-7 bar 1, three-way identical.)
            frozen = testCase.loadFixture('core721');
            root   = mat2doc.oxml.parse_xml(frozen);
            reser  = uint8(mat2doc.opc.oxml.serialize_part_xml(root));
            reser  = reser(:)';
            testCase.verifyEqual(numel(reser), 721, ...
                're-serialized core.xml must be exactly 721 B');
            testCase.verifyEqual(sha256hex(reser), testCase.SHA_CORE721, ...
                're-serialize must be byte-identical to the frozen 721 B oracle (L1)');
            verifyByteIdentical(testCase, reser, frozen, 'core.xml 721 B re-serialize');
        end

        % =============================================================== %
        % Bar 3 -- xsi-hoist byte identity (fresh build, 681 B)           %
        % =============================================================== %

        function test_xsi_hoist_681_byte_identical(testCase)
            % Regression / L1 (bar 3): a fresh new() + the created/modified/title/
            % author/keywords/revision battery serializes byte-identical + SHA-256
            % 3be041bb..., 681 B, to the frozen s0010 ref. Pins the whole generated
            % serializer path (H8 xsi hoist + H7 escaping + H2 non-ASCII + H14
            % revision). (validate_P1-7 bar 3.)
            blob = testCase.buildXsiBattery();
            want = testCase.loadFixture('core681');
            testCase.verifyEqual(numel(blob), 681, ...
                'xsi-hoist battery must be exactly 681 B');
            testCase.verifyEqual(sha256hex(blob), testCase.SHA_CORE681, ...
                'xsi-hoist battery must be byte-identical to the frozen s0010 ref (L1)');
            verifyByteIdentical(testCase, blob, want, 'xsi-hoist 681 B battery');
        end

        function test_xsi_hoist_decl_once_type_twice_order(testCase)
            % Regression (bar 3 structural, H8/D-serializer-nsdecl): xmlns:xsi
            % appears EXACTLY once (hoisted onto the root, trailing after
            % cp/dc/dcterms -- lxml order), and xsi:type="dcterms:W3CDTF" on EXACTLY
            % the two date children (created + modified).
            blob = testCase.buildXsiBattery();
            txt  = string(native2unicode(blob, 'UTF-8'));
            testCase.verifyEqual(count(txt, "xmlns:xsi"), 1, ...
                'xmlns:xsi must be declared exactly once (root hoist)');
            testCase.verifyEqual(count(txt, "xsi:type=""dcterms:W3CDTF"""), 2, ...
                'xsi:type must appear on exactly the 2 date children');
            rootTag = extractBefore(extractAfter(txt, "<cp:coreProperties"), ">");
            order = strings(0, 1);
            for d = ["xmlns:cp" "xmlns:dc" "xmlns:dcterms" "xmlns:dcmitype" "xmlns:xsi"]
                if contains(rootTag, d + "=")
                    order(end+1, 1) = d; %#ok<AGROW>
                end
            end
            testCase.verifyEqual(order, ["xmlns:cp"; "xmlns:dc"; "xmlns:dcterms"; ...
                "xmlns:xsi"], 'root nsdecl order must be cp, dc, dcterms, xsi');
        end

        % =============================================================== %
        % Bar 2 -- W3CDTF wall-clock grammar (36 vectors)                 %
        % =============================================================== %

        function test_w3cdtf_36_vector_wallclock(testCase)
            % Equivalence (bar 2, s0011): each of the 36 W3CDTF inputs read through
            % created_datetime as a WALL-CLOCK value (never tzinfo -- VERIFY-tz)
            % equals the frozen s0011 wall-clock output, or [] where Python is None.
            vecs = testCase.w3cdtfVectors();
            exp  = testCase.W3CDTF_EXPECT;
            for i = 1:numel(vecs)
                e = mat2doc.oxml.coreprops.CT_CoreProperties.new();
                c = e.get_or_add_created();
                c.text = vecs(i);
                got = e.created_datetime;
                if exp(i) == ""
                    testCase.verifyEmpty(got, sprintf( ...
                        'vector %d (%s) must be [] (Python None)', i, vecs(i)));
                else
                    testCase.verifyEqual(string(got, "yyyy-MM-dd HH:mm:ss"), ...
                        exp(i), sprintf('vector %d (%s) wall-clock mismatch', ...
                        i, vecs(i)));
                end
            end
        end

        function test_w3cdtf_template_created_modified_roundtrip(testCase)
            % Regression (bar 2, the template stamp): setting created/modified to
            % the template's 2013-12-23T23:15:00 datetime and reading back yields
            % the same whole-second wall-clock. The round-trip the M1 core.xml
            % relies on.
            e = mat2doc.oxml.coreprops.CT_CoreProperties.new();
            e.created_datetime  = datetime(2013, 12, 23, 23, 15, 0);
            e.modified_datetime = datetime(2013, 12, 23, 23, 15, 0);
            testCase.verifyEqual(string(e.created_datetime, "yyyy-MM-dd HH:mm:ss"), ...
                "2013-12-23 23:15:00", 'created round-trip');
            testCase.verifyEqual(string(e.modified_datetime, "yyyy-MM-dd HH:mm:ss"), ...
                "2013-12-23 23:15:00", 'modified round-trip');
        end

        function test_w3cdtf_inverted_offset_sign(testCase)
            % Regression (bar 2, offset_dt_ INVERTED sign, coreprops.py 211-225):
            % a '+'/'-' numeric offset shifts with sign_factor = -1 if '+' else 1,
            % so -08:00 ADDS 8 h (10:14:55 -> 18:14:55) and +05:30 SUBTRACTS 5:30
            % (10:14:55 -> 04:44:55).
            e = mat2doc.oxml.coreprops.CT_CoreProperties.new();
            e.get_or_add_created().text = "2003-12-31T10:14:55-08:00";
            testCase.verifyEqual(string(e.created_datetime, "yyyy-MM-dd HH:mm:ss"), ...
                "2003-12-31 18:14:55", '-08:00 offset must add 8 h (inverted sign)');
            e2 = mat2doc.oxml.coreprops.CT_CoreProperties.new();
            e2.get_or_add_created().text = "2003-12-31T10:14:55+05:30";
            testCase.verifyEqual(string(e2.created_datetime, "yyyy-MM-dd HH:mm:ss"), ...
                "2003-12-31 04:44:55", '+05:30 offset must subtract 5:30 (inverted sign)');
        end

        function test_w3cdtf_lowercase_t_parses_F1(testCase)
            % Regression (bar 2, fix F1): CPython _strptime compiles with
            % re.IGNORECASE, so the "T" separator matches lowercase "t". The [Tt]
            % template must parse "2003-12-31t10:14:55" (pre-F1 it returned []).
            e = mat2doc.oxml.coreprops.CT_CoreProperties.new();
            e.get_or_add_created().text = "2003-12-31t10:14:55";
            testCase.verifyEqual(string(e.created_datetime, "yyyy-MM-dd HH:mm:ss"), ...
                "2003-12-31 10:14:55", 'lowercase-t W3CDTF must parse (F1)');
        end

        function test_w3cdtf_garbage_returns_empty(testCase)
            % Edge (bar 2, error path): an unparseable string yields [] (the
            % ValueError is caught in datetime_of_element_ and mapped to None).
            e = mat2doc.oxml.coreprops.CT_CoreProperties.new();
            e.get_or_add_created().text = "garbage";
            testCase.verifyEmpty(e.created_datetime, 'garbage W3CDTF must be []');
            e.get_or_add_created().text = "";
            testCase.verifyEmpty(e.created_datetime, 'empty W3CDTF must be []');
        end

        function test_w3cdtf_getter_naive_tzinfo(testCase)
            % Regression (VERIFY-tz, API-value): created_datetime is a NAIVE datetime
            % (TimeZone = '') by design -- Python returns tz-aware UTC, but the XML
            % bytes are unaffected (proven byte-identical by bars 1/3). Comparisons
            % are wall-clock, never tzinfo.
            e = mat2doc.oxml.coreprops.CT_CoreProperties.new();
            e.get_or_add_created().text = "2013-12-23T23:15:00Z";
            testCase.verifyEqual(char(e.created_datetime.TimeZone), '', ...
                'getter must return a naive datetime (TimeZone='''') -- VERIFY-tz');
        end

        function test_strftime_second_truncation(testCase)
            % Equivalence (aux, s0011 strftime_trunc): MATLAB 'ss' truncates a
            % fractional second exactly like Python strftime %S (55.9 -> "...55Z").
            got = string(datetime(2003, 12, 31, 10, 14, 55.9), ...
                "yyyy-MM-dd'T'HH:mm:ss") + "Z";
            testCase.verifyEqual(got, "2003-12-31T10:14:55Z", ...
                'fractional second must truncate like strftime %S');
        end

        % =============================================================== %
        % Bar 4 -- revision D-002 grammar                                 %
        % =============================================================== %

        function test_revision_getter_11_exact(testCase)
            % Regression (bar 4, s0011 revision_get): the 11 exact getter vectors
            % -- absent->0, <cp:revision/> (str(None))->0, "1"->1, " 42 "->42,
            % "+7"->7, "-5"->0 (neg), "abc"->0, "2.5"->0, "7\n"->7, "0x1F"->0,
            % "007"->7.
            for i = 1:numel(testCase.REV_EXACT_IN)
                in = replace(testCase.REV_EXACT_IN(i), "\n", newline);
                testCase.verifyEqual(testCase.revisionOf(in), ...
                    testCase.REV_EXACT_OUT(i), sprintf( ...
                    'revision getter vector %d (%s)', i, testCase.REV_EXACT_IN(i)));
            end
        end

        function test_revision_getter_D002_divergences(testCase)
            % Regression (bar 4, SIGNED D-002 scope): the 3 known divergent getter
            % vectors are pinned as deliberate-red guards -- any future change to
            % the ASCII-[0-9] grammar surfaces here. These fall INSIDE the signed
            % D-002 row's named WP9 scope (probe_info.json), NOT new deviations:
            %   "1_0"                  -> 0     (Python int() 10; underscore group)
            %   "٣" (Arabic 3)    -> 0     (Python int() 3; Unicode digit)
            %   "99999999999999999999" -> 1e20  (Python exact big-int; nearest double)
            testCase.verifyEqual(testCase.revisionOf("1_0"), 0, ...
                'D-002: "1_0" -> 0 (Python 10) -- underscore not accepted');
            arabic3 = string(char(1635));   % U+0663 ARABIC-INDIC DIGIT THREE
            testCase.verifyEqual(testCase.revisionOf(arabic3), 0, ...
                'D-002: Arabic digit -> 0 (Python 3) -- Unicode digit not accepted');
            testCase.verifyEqual(testCase.revisionOf("99999999999999999999"), 1e20, ...
                'D-002: 20-digit int -> 1e20 nearest double (Python exact big-int)');
        end

        function test_revision_setter_positive_ok(testCase)
            % Regression (bar 4, setter happy path): a positive int writes str(int)
            % via pyStr into <cp:revision>.
            e = mat2doc.oxml.coreprops.CT_CoreProperties.new();
            e.revision_number = 7;
            testCase.verifyEqual(e.revision.text, "7", 'revision=7 -> text "7"');
            e2 = mat2doc.oxml.coreprops.CT_CoreProperties.new();
            e2.revision_number = 1;
            testCase.verifyEqual(e2.revision.text, "1", 'revision=1 -> text "1"');
        end

        function test_revision_setter_nonpositive_raises_valueerror(testCase)
            % Edge / error path (bar 4, setter guard): 0, -1, 2.5, NaN, "5" each
            % raise mat2doc:ValueError with the VERBATIM message incl. the pyStr
            % type-token ('2.5' not '2', 'nan', '5' for a string input). Asserts
            % the IDENTIFIER (mat2doc:ValueError = mat2ppt:<PyExceptionName> family),
            % not merely that it throws. (s0011 revision_set, byte-identical to
            % Python's ValueError.)
            cases = { 0,   "revision property requires positive int, got '0'"; ...
                     -1,   "revision property requires positive int, got '-1'"; ...
                     2.5,  "revision property requires positive int, got '2.5'"; ...
                     NaN,  "revision property requires positive int, got 'nan'"; ...
                     "5",  "revision property requires positive int, got '5'" };
            for i = 1:size(cases, 1)
                val = cases{i, 1};
                wantMsg = cases{i, 2};
                e = mat2doc.oxml.coreprops.CT_CoreProperties.new();
                caught = [];
                try
                    e.revision_number = val;
                catch ME
                    caught = ME;
                end
                testCase.assertNotEmpty(caught, sprintf( ...
                    'setter case %d must raise', i));
                testCase.verifyEqual(caught.identifier, 'mat2doc:ValueError', ...
                    sprintf('setter case %d identifier', i));
                testCase.verifyEqual(caught.message, char(wantMsg), sprintf( ...
                    'setter case %d message must be byte-identical (pyStr token)', i));
            end
        end

        % =============================================================== %
        % F2 pin -- pyStr in set_element_text_ (H14 regression)           %
        % =============================================================== %

        function test_setter_pyStr_full_precision_F2(testCase)
            % Regression (fix F2, H14): a non-str value through a CoreProperties
            % string setter routes through pyStr, so cp.comments = 1/3 writes the
            % FULL-precision "0.3333333333333333" -- NOT the raw string(1/3)
            % "0.33333". Locks F2 so it cannot regress. (audit probe T03.)
            e  = mat2doc.oxml.coreprops.CT_CoreProperties.new();
            cp = mat2doc.opc.CoreProperties(e);
            cp.comments = 1/3;
            testCase.verifyEqual(cp.comments, "0.3333333333333333", ...
                'cp.comments = 1/3 must write full-precision pyStr (F2)');
        end

        function test_setter_pyStr_logical_F2(testCase)
            % Regression (fix F2, H14): a logical through a string setter writes the
            % CPython bool str "True"/"False" (via pyStr), NOT MATLAB "true"/"false".
            e  = mat2doc.oxml.coreprops.CT_CoreProperties.new();
            cp = mat2doc.opc.CoreProperties(e);
            cp.comments = true;
            testCase.verifyEqual(cp.comments, "True", 'logical true -> "True" (F2)');
            cp.comments = false;
            testCase.verifyEqual(cp.comments, "False", 'logical false -> "False" (F2)');
        end

        function test_pyStr_20_vector(testCase)
            % Equivalence (aux, s0011 pystr): the 20-vector CPython str() battery --
            % int + float shortest-repr (2.0->"2.0", 1/3->"0.3333333333333333",
            % 1e16->"1e+16", inf/-inf/nan, 5e-324) -- against the frozen outputs.
            vals = {0, 7, -5, 914400, 0.5, 2.0, -0.0, 0.1, 1/3, 1e16, 1e15, ...
                1e-4, 1e-5, Inf, -Inf, NaN, 5e-324, ...
                1.7976931348623157e308, 123456.789, 2.5e16};
            for i = 1:numel(vals)
                testCase.verifyEqual( ...
                    mat2doc.shared.pyStr(vals{i}, testCase.PYSTR_KIND(i)), ...
                    testCase.PYSTR_EXPECT(i), sprintf('pyStr vector %d', i));
            end
        end

        % =============================================================== %
        % Bar 5 -- CoreProperties API wrapper + default()                 %
        % =============================================================== %

        function test_api_15_prop_delegation(testCase)
            % Equivalence (bar 5, s0011 api_delegation): all 15 CoreProperties
            % accessors set through the wrapper and read back agree with the frozen
            % s0011 values (datetimes wall-clock, revision int 7).
            e  = mat2doc.oxml.coreprops.CT_CoreProperties.new();
            cp = mat2doc.opc.CoreProperties(e);
            cp.author = "Ada";           cp.category = "Report";
            cp.comments = "a comment";   cp.content_status = "Final";
            cp.created = datetime(2013, 1, 2, 3, 4, 5);
            cp.identifier = "id-0001";   cp.keywords = "alpha beta";
            cp.language = "en-US";       cp.last_modified_by = "LMB";
            cp.last_printed = datetime(2014, 5, 6, 7, 8, 9);
            cp.modified = datetime(2015, 9, 10, 11, 12, 13);
            cp.revision = 7;             cp.subject = "the subject";
            cp.title = "The Title";      cp.version = "v1.2";

            testCase.verifyEqual(cp.author, "Ada");
            testCase.verifyEqual(cp.category, "Report");
            testCase.verifyEqual(cp.comments, "a comment");
            testCase.verifyEqual(cp.content_status, "Final");
            testCase.verifyEqual(string(cp.created, "yyyy-MM-dd HH:mm:ss"), ...
                "2013-01-02 03:04:05");
            testCase.verifyEqual(cp.identifier, "id-0001");
            testCase.verifyEqual(cp.keywords, "alpha beta");
            testCase.verifyEqual(cp.language, "en-US");
            testCase.verifyEqual(cp.last_modified_by, "LMB");
            testCase.verifyEqual(string(cp.last_printed, "yyyy-MM-dd HH:mm:ss"), ...
                "2014-05-06 07:08:09");
            testCase.verifyEqual(string(cp.modified, "yyyy-MM-dd HH:mm:ss"), ...
                "2015-09-10 11:12:13");
            testCase.verifyEqual(cp.revision, 7);
            testCase.verifyEqual(cp.subject, "the subject");
            testCase.verifyEqual(cp.title, "The Title");
            testCase.verifyEqual(cp.version, "v1.2");
        end

        function test_api_writes_through(testCase)
            % Equivalence (bar 5, s0011 api_writes_through): the wrapper writes
            % through to the underlying CT_CoreProperties element accessors.
            e  = mat2doc.oxml.coreprops.CT_CoreProperties.new();
            cp = mat2doc.opc.CoreProperties(e);
            cp.author  = "Ada";
            cp.title   = "The Title";
            cp.created = datetime(2013, 1, 2, 3, 4, 5);
            cp.revision = 7;
            testCase.verifyEqual(e.author_text, "Ada", 'write-through author_text');
            testCase.verifyEqual(e.title_text, "The Title", 'write-through title_text');
            testCase.verifyEqual(string(e.created_datetime, "yyyy-MM-dd HH:mm:ss"), ...
                "2013-01-02 03:04:05", 'write-through created_datetime');
            testCase.verifyEqual(e.revision_number, 7, 'write-through revision_number');
        end

        function test_default_structure(testCase)
            % Equivalence (bar 5, s0011 default_structure): CorePropertiesPart.
            % default() builds the docx-faithful default -- part_class, partname,
            % content_type, title "Word Document", last_modified_by "python-docx",
            % revision 1, created [], last_printed []. The wall-clock `modified` is
            % EXCLUDED (D-coreprops-time, signed). default() is M1-UNREACHABLE
            % (default.docx HAS a core.xml, and PartFactory maps
            % OPC_CORE_PROPERTIES -> base XmlPart at M1) -- ported for P2 wiring.
            dp = mat2doc.opc.parts.CorePropertiesPart.default([]);
            testCase.verifyEqual(class(dp), ...
                'mat2doc.opc.parts.CorePropertiesPart', 'default() part class');
            testCase.verifyEqual(string(dp.partname), "/docProps/core.xml", ...
                'default() partname');
            testCase.verifyEqual(string(dp.content_type), ...
                "application/vnd.openxmlformats-package.core-properties+xml", ...
                'default() content_type');
            cp = dp.core_properties;
            testCase.verifyEqual(cp.title, "Word Document", 'default() title');
            testCase.verifyEqual(cp.last_modified_by, "python-docx", ...
                'default() last_modified_by');
            testCase.verifyEqual(cp.revision, 1, 'default() revision');
            testCase.verifyEmpty(cp.created, 'default() created is [] (not set)');
            testCase.verifyEmpty(cp.last_printed, 'default() last_printed is []');
            % NOTE: cp.modified is a wall-clock stamp (D-coreprops-time) -- NOT
            % pinned here; it is a live-clock value excluded from the frozen probe.
        end

    end

    % ===================== instance helpers ============================ %
    methods (Access = private)

        function vecs = w3cdtfVectors(testCase) %#ok<MANU>
            % The 36 W3CDTF inputs, 1:1 with s0011_coreprops_probes.m wvec.
            % Index 33 carries a trailing newline (built here to keep it literal).
            vecs = [ ...
                "2003", "2003-12", "2003-12-31", "2003-12-31T10:14:55", ...
                "2003-12-31T10:14:55Z", "2003-12-31T10:14:55-08:00", ...
                "2003-12-31T10:14:55+05:30", "2003-12-31T10:14:55+00:00", ...
                "2003-12-31T10:14:55-00:30", "2003-12-31T10:14:55.123", ...
                "2003-12-31T10:14:55+0800", "2003-12-31T10:14:55 08:00", ...
                "2003-12-31T10:14:55ABCDEF", "2013-12-23T23:15:00Z", ...
                "2003-13-01", "2003-02-31", "2003-00-01", "2003-12-31T10:14:61", ...
                "2003-12-31T24:14:55", "2003-12-31T10:61:55", ...
                "0000", "0001", "9999", "0999", "203", "20031", ...
                "2003-1", "2003-1-2", "2003-1-2T3:4:5", "garbage", "", ...
                " 2003", "2003 ", "2003" + newline, ...
                "2003-12-31t10:14:55", "2003-12-31T10:14:55z"];
        end

        function blob = buildXsiBattery(testCase) %#ok<MANU>
            % The s0010 xsi-hoist battery, 1:1 with s0010_xsi_hoist_battery.m:
            % fresh new() + title (H7 & < > + H2 e-acute + U+1F600 emoji) + author
            % + keywords (quotes) + created/modified/lastPrinted + revision 7,
            % serialized to the 681 B expected bytes.
            emoji = string(char([55357 56832]));   % U+1F600
            e = mat2doc.oxml.coreprops.CT_CoreProperties.new();
            e.title_text    = "Report & <T" + string(char(233)) + emoji + ">";
            e.author_text   = "Ada";
            e.keywords_text  = "k1 ""quoted"" k2";
            e.created_datetime     = datetime(2013, 12, 23, 23, 15, 0);
            e.modified_datetime    = datetime(2020, 1, 2, 3, 4, 5);
            e.lastPrinted_datetime = datetime(2021, 6, 7, 8, 9, 10);
            e.revision_number = 7;
            blob = uint8(mat2doc.opc.oxml.serialize_part_xml(e));
            blob = blob(:)';
        end

        function n = revisionOf(testCase, v) %#ok<MANU>
            % Mirror of s0011 revisionOf: build a CT_CoreProperties, optionally add
            % <cp:revision> (absent for "<ABSENT>"), optionally set its text (empty
            % <cp:revision/> for "<EMPTYNONE>"), read revision_number.
            e = mat2doc.oxml.coreprops.CT_CoreProperties.new();
            if v ~= "<ABSENT>"
                r = e.get_or_add_revision();
                if v ~= "<EMPTYNONE>"
                    r.text = v;
                end
            end
            n = e.revision_number;
        end

        function b = loadFixture(testCase, base) %#ok<MANU>
            here = fileparts(mfilename('fullpath'));   % tests\coreprops
            p = fullfile(here, 'data', [base '.bin']);
            b = readBytes(p);
        end
    end
end

% ===================== file-local helpers ============================== %

function b = readBytes(p)
    f = fopen(p, 'r', 'n');            % binary read (no CRLF translation)
    assert(f >= 0, 'could not open for read: %s', p);
    b = fread(f, Inf, '*uint8')';
    fclose(f);
end

function verifyByteIdentical(testCase, got, want, label)
    % Byte-level (L1) assertion. On mismatch report sizes and first diff offset.
    got = uint8(got(:)');  want = uint8(want(:)');
    if ~isequal(got, want)
        n = min(numel(got), numel(want));
        d = find(got(1:n) ~= want(1:n), 1);
        if isempty(d), d = n + 1; end
        gv = 0; wv = 0;
        if d <= numel(got), gv = double(got(d)); end
        if d <= numel(want), wv = double(want(d)); end
        diag = sprintf(['%s: bytes differ (got %d B, want %d B); first diff @%d ' ...
            '(got 0x%02X, want 0x%02X)'], char(label), numel(got), numel(want), ...
            d, gv, wv);
    else
        diag = char(label);
    end
    testCase.verifyEqual(got, want, diag);
end

function h = sha256hex(bytes)
    % Lowercase hex SHA-256 of a uint8 vector (base MATLAB / Java MessageDigest).
    md  = java.security.MessageDigest.getInstance('SHA-256');
    dig = md.digest(typecast(uint8(bytes(:)'), 'int8'));
    b   = typecast(dig, 'uint8');
    h   = lower(string(sprintf('%02x', b)));
end

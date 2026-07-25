classdef Test_p1_1_shared < matlab.unittest.TestCase
% TEST_P1_1_SHARED  Gate-4 permanent unit tests for Mat2Doc P1-1.
%
%   The FIRST Mat2Doc test class -- it establishes the tests\ scaffold and
%   conventions the rest of the Mat2Doc suite follows.
%
%   Surface under test (ported from python-docx v1.2.0 src/docx/shared.py):
%     - the Length family: Length base + the Emu/Inches/Cm/Mm/Pt/Twips
%       convenience constructors, their EMU constants, the conversion
%       properties (.emu/.inches/.cm/.mm/.pt/.twips), and the twips
%       half-to-even rounding with the mandated negative-band +0.0 fix.
%     - RGBColor: str_() (UPPER hex), repr_() (docx lowercase form),
%       from_string round-trip, and the TypeError/ValueError error split.
%
%   Coverage taxonomy:
%     * Nominal      -- documented conversions and constructor round trips.
%     * Edge         -- twips half-to-even ties, the -1..-317 signed-zero
%                       band and its -318 edge, empty/error paths, string
%                       parse with whitespace/sign/underscore.
%     * Regression   -- hard-coded expected values / hex strings / bit
%                       patterns (num2hex) baked into the assertions.
%     * Equivalence  -- test_replay_all_88_records replays the entire frozen
%                       Gate-3 probe battery (tests\shared\data\p1_1\probe.json,
%                       the Validator's frozen reference copied in verbatim so
%                       the suite is self-contained) and asserts the port
%                       reproduces every record bit-exact (num2hex canon,
%                       signed-zero aware -- matching the Gate-3 comparator in
%                       harness\mat2doc\validate_p1_1.m).
%     * Error paths  -- verify the error IDENTIFIER (mat2doc:TypeError /
%                       mat2doc:ValueError) plus byte-exact message where the
%                       message is cross-faithful (not the D-003 port-authored
%                       multiplier-string messages).
%
%   Adopted deviations exercised (validation\summary\
%   decision_2026-07-25_mat2doc_deviation_preadoption.md):
%     D-002  string-multiplier Twips(str) -> port raises (raw CPython succeeds)
%     D-003  Inches/Pt/Cm/Mm(str) message is port-authored; class faithful
%     D-004  Emu(nan)/Emu(inf) raise a MATLAB-native identifier (raising
%            decision faithful; identifier not observed by any ported caller)
%     D-STYPE-1  an integer-valued double is accepted as the Python int
%
%   Determinism: no network, no absolute paths -- the frozen fixture is
%   resolved relative to this file via fileparts(mfilename('fullpath')).
%
%   Ported surface: python-docx v1.2.0 src/docx/shared.py lines 25-149.

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into
            % the test folder, so without the worktree root on the path a COLD
            % run cannot resolve the +mat2doc package. Copy of the proven
            % Mat2Ppt idiom (tests\shared\Test_ElementProxy.m).
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\shared
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =================================================== EMU constants ====
        function test_emu_constants(testCase)
            % Regression: the five exact class constants (shared.py 26-30).
            % Accessed via the class name (Constant properties need no instance;
            % Length has no no-arg constructor path).
            testCase.verifyEqual(mat2doc.shared.Length.EMUS_PER_INCH_, 914400);
            testCase.verifyEqual(mat2doc.shared.Length.EMUS_PER_CM_,   360000);
            testCase.verifyEqual(mat2doc.shared.Length.EMUS_PER_MM_,    36000);
            testCase.verifyEqual(mat2doc.shared.Length.EMUS_PER_PT_,    12700);
            testCase.verifyEqual(mat2doc.shared.Length.EMUS_PER_TWIP_,     635);
        end

        % =============================================== multiplier ctors =====
        function test_multiplier_constructors_to_emu(testCase)
            % Nominal: each convenience constructor's unit value -> 914400 EMU
            % (one inch), plus the raw Emu passthrough (shared.py 88-124).
            testCase.verifyEqual(double(mat2doc.shared.Inches(1).emu), 914400);
            testCase.verifyEqual(double(mat2doc.shared.Cm(2.54).emu),  914400);
            testCase.verifyEqual(double(mat2doc.shared.Mm(25.4).emu),  914400);
            testCase.verifyEqual(double(mat2doc.shared.Pt(72).emu),    914400);
            testCase.verifyEqual(double(mat2doc.shared.Twips(1440).emu), 914400);
            testCase.verifyEqual(double(mat2doc.shared.Emu(914400).emu), 914400);
        end

        function test_constructor_truncates_toward_zero(testCase)
            % Edge: Python int() truncates toward zero (H6 -> fix()), NOT round.
            testCase.verifyEqual(double(mat2doc.shared.Emu(2.5)),   2);
            testCase.verifyEqual(double(mat2doc.shared.Emu(-2.5)), -2);
            % Inches(0.5) = int(457200.0) exact; Inches(-0.5) = -457200.
            testCase.verifyEqual(double(mat2doc.shared.Inches(0.5)),   457200);
            testCase.verifyEqual(double(mat2doc.shared.Inches(-0.5)), -457200);
        end

        % ================================================= conversions =========
        function test_conversions_one_inch(testCase)
            % Nominal: Emu(914400) round-trips to 1 inch / 2.54 cm / 25.4 mm /
            % 72 pt / 1440 twips (shared.py 33-84). Exact-integer conversions
            % pinned exactly; the fractional cm/mm pinned bit-exact by num2hex
            % (regression) since 914400/360000 is not representable finitely.
            L = mat2doc.shared.Emu(914400);
            testCase.verifyEqual(double(L.emu), 914400);
            testCase.verifyEqual(L.inches, 1);          % exact
            testCase.verifyEqual(L.pt,      72);        % exact
            testCase.verifyEqual(L.twips,   1440);      % exact (integer)
            % Bit-exact regression for the non-terminating ratios:
            testCase.verifyEqual(num2hex(L.cm), '400451eb851eb852');   % 2.54
            testCase.verifyEqual(num2hex(L.mm), '4039666666666666');   % 25.4
        end

        function test_emu_getter_returns_self(testCase)
            % .emu returns the Length instance itself (Python: return self).
            L = mat2doc.shared.Emu(12345);
            testCase.verifyEqual(double(L.emu), 12345);
            testCase.verifyTrue(isa(L.emu, 'mat2doc.shared.Length'));
        end

        % ============================ twips: half-to-even + signed-zero band ===
        function test_twips_negative_band_is_positive_zero(testCase)
            % *** The Gate-2 fix ***: Emu(e).twips for e in -1..-317 rounds to
            % ZERO, and it must be +0.0 (num2hex '0000000000000000'), never
            % -0.0. Python int() never yields -0; MATLAB fix()/round() of a
            % value in (-1,0) would leak -0.0 without the normalization.
            for e = [-1, -50, -158, -317]
                t = mat2doc.shared.Emu(e).twips;
                testCase.verifyEqual(t, 0, ...
                    sprintf('Emu(%d).twips must be 0', e));
                % Sign assertion: +0 -> 1/0 == +Inf; -0 -> 1/0 == -Inf.
                testCase.verifyEqual(1/t, Inf, ...
                    sprintf('Emu(%d).twips must be +0.0, not -0.0', e));
                % Direct bit-pattern check (belt and braces).
                testCase.verifyEqual(num2hex(t), '0000000000000000');
            end
        end

        function test_twips_positive_band_is_zero(testCase)
            % Edge (mirror side): Emu(1..317).twips == +0.
            for e = [1, 100, 317]
                testCase.verifyEqual(mat2doc.shared.Emu(e).twips, 0);
            end
        end

        function test_twips_band_edge_neg318(testCase)
            % Edge boundary: -318 EMU is the first value that rounds to -1
            % (|-318/635| = 0.5008 > 0.5).
            testCase.verifyEqual(mat2doc.shared.Emu(-318).twips, -1);
            testCase.verifyEqual(mat2doc.shared.Emu(318).twips,   1);
        end

        function test_twips_rounding_edges(testCase)
            % Edge: half-to-even boundary crossings both directions
            % (regression values decoded from the frozen probe canon).
            testCase.verifyEqual(mat2doc.shared.Emu(635).twips,   1);
            testCase.verifyEqual(mat2doc.shared.Emu(-635).twips, -1);
            testCase.verifyEqual(mat2doc.shared.Emu(952).twips,   1);   % 1.499->1
            testCase.verifyEqual(mat2doc.shared.Emu(953).twips,   2);   % 1.500->2
            testCase.verifyEqual(mat2doc.shared.Emu(-952).twips, -1);
            testCase.verifyEqual(mat2doc.shared.Emu(-953).twips, -2);
            testCase.verifyEqual(mat2doc.shared.Emu(20002).twips,  31);
            testCase.verifyEqual(mat2doc.shared.Emu(20003).twips,  32);
            testCase.verifyEqual(mat2doc.shared.Emu(-20003).twips, -32);
        end

        % ==================================== string-argument construction =====
        function test_emu_string_parse(testCase)
            % Edge (LIVE path): the docx oxml ST_ length lexer builds
            % Emu(str_value). int(str) PARSES base-10 (never truncates):
            % surrounding whitespace, a single sign, and underscores strictly
            % between digits are all accepted.
            testCase.verifyEqual(double(mat2doc.shared.Emu("914400")), 914400);
            testCase.verifyEqual(double(mat2doc.shared.Emu(" +914_400 ")), 914400);
            testCase.verifyEqual(double(mat2doc.shared.Emu("-0")), 0);   % int('-0')==0
        end

        % ================================================= RGBColor: format ====
        function test_rgb_str_uppercase(testCase)
            % Nominal: __str__ = "%02X%02X%02X" -- fixed-width UPPER hex.
            testCase.verifyEqual(mat2doc.shared.RGBColor(255, 0, 0).str_(),   "FF0000");
            testCase.verifyEqual(mat2doc.shared.RGBColor(60, 47, 128).str_(), "3C2F80");
            testCase.verifyEqual(mat2doc.shared.RGBColor(0, 0, 0).str_(),     "000000");
            testCase.verifyEqual(mat2doc.shared.RGBColor(255, 255, 255).str_(), "FFFFFF");
        end

        function test_rgb_repr_lowercase(testCase)
            % Nominal: docx __repr__ = "RGBColor(0x%02x, 0x%02x, 0x%02x)".
            testCase.verifyEqual(mat2doc.shared.RGBColor(60, 47, 128).repr_(), ...
                "RGBColor(0x3c, 0x2f, 0x80)");
            testCase.verifyEqual(mat2doc.shared.RGBColor(255, 0, 0).repr_(), ...
                "RGBColor(0xff, 0x00, 0x00)");
        end

        function test_rgb_from_string_roundtrip(testCase)
            % Nominal + edge: from_string slices the 6-hex string into 3 bytes,
            % is case-insensitive, and its final slice [4:] is greedy so a
            % 5-char string maps the last char into the low nibble of blue.
            testCase.verifyEqual( ...
                mat2doc.shared.RGBColor.from_string("FF0000").str_(), "FF0000");
            testCase.verifyEqual( ...
                mat2doc.shared.RGBColor.from_string("3c2f80").str_(), "3C2F80");
            testCase.verifyEqual( ...
                mat2doc.shared.RGBColor.from_string("3C2F80").repr_(), ...
                "RGBColor(0x3c, 0x2f, 0x80)");
            % "3C2F8" -> r=3C g=2F b=int("8",16)=08  (Python slice [4:] == "8").
            testCase.verifyEqual( ...
                mat2doc.shared.RGBColor.from_string("3C2F8").str_(), "3C2F08");
        end

        function test_rgb_value_equality(testCase)
            % H2/H5: tuple-subclass value semantics -- equal by triplet;
            % comparison against a non-RGBColor is false.
            a = mat2doc.shared.RGBColor(255, 0, 0);
            b = mat2doc.shared.RGBColor.from_string("FF0000");
            testCase.verifyTrue(a == b);
            testCase.verifyFalse(mat2doc.shared.RGBColor(1, 2, 3) == ...
                                 mat2doc.shared.RGBColor(1, 2, 4));
            testCase.verifyTrue(mat2doc.shared.RGBColor(1, 2, 3) ~= ...
                                mat2doc.shared.RGBColor(1, 2, 4));
            % Non-RGBColor comparands -> false / true (tuple == guard).
            testCase.verifyFalse(a == "FF0000");
            testCase.verifyFalse(a == 16711680);
            testCase.verifyTrue(a ~= []);
        end

        % ================================================ RGBColor: errors =====
        function test_rgb_valueerror_out_of_range(testCase)
            % Error path: an in-type but out-of-range component -> ValueError,
            % message byte-exact (cross-faithful, shared.py 130-134).
            msg = 'RGBColor() takes three integer values 0-255';
            checkRaises(testCase, @() mat2doc.shared.RGBColor(256, 0, 0), ...
                'mat2doc:ValueError', msg);
            checkRaises(testCase, @() mat2doc.shared.RGBColor(-1, 0, 0), ...
                'mat2doc:ValueError', msg);
            checkRaises(testCase, @() mat2doc.shared.RGBColor(0, 300, 0), ...
                'mat2doc:ValueError', msg);
            checkRaises(testCase, @() mat2doc.shared.RGBColor(0, 0, 256), ...
                'mat2doc:ValueError', msg);
        end

        function test_rgb_typeerror_non_int(testCase)
            % Error path: a non-int component (non-integral float or a string)
            % -> TypeError (docx splits TypeError vs ValueError; pptx did not).
            msg = 'RGBColor() takes three integer values 0-255';
            checkRaises(testCase, @() mat2doc.shared.RGBColor(2.5, 0, 0), ...
                'mat2doc:TypeError', msg);
            checkRaises(testCase, @() mat2doc.shared.RGBColor("ff", 0, 0), ...
                'mat2doc:TypeError', msg);
        end

        function test_from_string_errors(testCase)
            % Error path: from_string hex-parse failures.
            checkRaises(testCase, ...
                @() mat2doc.shared.RGBColor.from_string("GGGGGG"), ...
                'mat2doc:ValueError', ...
                "invalid literal for int() with base 16: 'GG'");
            % 4-char: slice [4:] is empty -> int('',16) ValueError.
            checkRaises(testCase, ...
                @() mat2doc.shared.RGBColor.from_string("FF00"), ...
                'mat2doc:ValueError', ...
                "invalid literal for int() with base 16: ''");
            % 7-char "3C2F80A": b = int("80A",16) = 2058 > 255 -> range msg.
            checkRaises(testCase, ...
                @() mat2doc.shared.RGBColor.from_string("3C2F80A"), ...
                'mat2doc:ValueError', ...
                "RGBColor() takes three integer values 0-255");
        end

        function test_emu_parse_errors(testCase)
            % Error path: int(str) parse failures carry Python's exact message.
            checkRaises(testCase, @() mat2doc.shared.Emu("2.5"), ...
                'mat2doc:ValueError', ...
                "invalid literal for int() with base 10: '2.5'");
            checkRaises(testCase, @() mat2doc.shared.Emu(""), ...
                'mat2doc:ValueError', ...
                "invalid literal for int() with base 10: ''");
        end

        % ================================================ adopted deviations ===
        function test_dstype1_integer_valued_float_accepted(testCase)
            % D-STYPE-1: a MATLAB double cannot distinguish int from float, so
            % an integer-valued double is accepted as the Python int.
            testCase.verifyEqual( ...
                mat2doc.shared.RGBColor(255.0, 0.0, 0.0).str_(), "FF0000");
            testCase.verifyEqual(double(mat2doc.shared.Emu(914400.0)), 914400);
        end

        function test_d004_nonfinite_raises(testCase)
            % D-004: non-finite input RAISES (the raising decision is faithful;
            % the identifier is MATLAB-native and observed by no ported caller,
            % so it is deliberately NOT pinned here -- assert only that it
            % raises, per the D-004 ruling).
            checkThrows(testCase, @() mat2doc.shared.Emu(nan));
            checkThrows(testCase, @() mat2doc.shared.Emu(inf));
            checkThrows(testCase, @() mat2doc.shared.Emu(-inf));
        end

        function test_d002_string_multiplier_raises(testCase)
            % D-002: Inches/Pt/Cm/Mm/Twips evaluate int(str * K) (string
            % repetition) in Python; the port rejects every string. Class is
            % ValueError (D-003 -- message port-authored, so class only).
            checkThrows(testCase, @() mat2doc.shared.Twips("1"), 'mat2doc:ValueError');
            checkThrows(testCase, @() mat2doc.shared.Twips("123456"), 'mat2doc:ValueError');
            checkThrows(testCase, @() mat2doc.shared.Inches("1"), 'mat2doc:ValueError');
            checkThrows(testCase, @() mat2doc.shared.Pt("1"), 'mat2doc:ValueError');
            checkThrows(testCase, @() mat2doc.shared.Cm("1"), 'mat2doc:ValueError');
            checkThrows(testCase, @() mat2doc.shared.Mm("1"), 'mat2doc:ValueError');
        end

        % ============================================= FULL probe equivalence ==
        function test_replay_all_88_records(testCase)
            % Equivalence: replay the entire frozen Gate-3 battery and assert
            % the port reproduces every record's canon bit-exact. The battery
            % is rebuilt here identically to harness\mat2doc\validate_p1_1.m
            % (num2hex bit pattern for numerics -- signed-zero aware; raw
            % string for str; "true"/"false" for bool; class[|message] for
            % err; RAISES for dev-raises). Fixture is the frozen reference
            % copied into tests\shared\data\p1_1\probe.json (self-contained).
            here = fileparts(mfilename('fullpath'));
            refPath = fullfile(here, 'data', 'p1_1', 'probe.json');
            testCase.assertTrue(isfile(refPath), ...
                sprintf('frozen probe fixture missing: %s', refPath));
            ref = jsondecode(fileread(refPath));
            recs = ref.records;

            obs = buildObservedBattery();

            n = numel(recs);
            testCase.verifyEqual(n, 88, 'frozen battery must hold 88 records');
            for i = 1:n
                if iscell(recs); r = recs{i}; else; r = recs(i); end
                id = char(string(r.id));
                testCase.assertTrue(isKey(obs, id), ...
                    sprintf('record id not reproduced by port: %s', id));
                testCase.verifyEqual(obs(id), char(string(r.canon)), ...
                    sprintf('canon mismatch for %s (policy %s)', ...
                            id, char(string(r.policy))));
            end
        end

    end
end

% ============================= file-local helpers ==========================
% Local functions in a classdef file are visible to the class methods above.

function checkRaises(testCase, fn, expId, expMsg)
    % Assert fn() raises; verify the identifier and the byte-exact message.
    caught = false;
    try
        fn();
    catch ME
        caught = true;
        testCase.verifyEqual(ME.identifier, expId, ...
            sprintf('expected identifier %s, got %s', expId, ME.identifier));
        testCase.verifyEqual(ME.message, char(expMsg), ...
            'error message must be byte-exact');
    end
    testCase.verifyTrue(caught, 'expected an error but none was raised');
end

function checkThrows(testCase, fn, expId)
    % Assert fn() raises. If expId given, pin the identifier; else (D-004)
    % assert only that it raised.
    caught = false;
    try
        fn();
    catch ME
        caught = true;
        if nargin >= 3
            testCase.verifyEqual(ME.identifier, expId);
        end
    end
    testCase.verifyTrue(caught, 'expected an error but none was raised');
end

function obs = buildObservedBattery()
    % Rebuild the P1-1 canon battery through mat2doc.shared.*, keyed by id --
    % a faithful mirror of harness\mat2doc\validate_p1_1.m so the permanent
    % suite is a standalone re-derivation of Gate-3. containers.Map is a
    % handle, so putCanon assigns into the caller's map.
    Inches = @mat2doc.shared.Inches;
    Cm     = @mat2doc.shared.Cm;
    Emu    = @mat2doc.shared.Emu;
    Mm     = @mat2doc.shared.Mm;
    Pt     = @mat2doc.shared.Pt;
    Twips  = @mat2doc.shared.Twips;
    Length = @mat2doc.shared.Length;
    RGB    = @mat2doc.shared.RGBColor;
    nh     = @(v) num2hex(double(v));   % == Python struct.pack('>d',v).hex()

    obs = containers.Map('KeyType', 'char', 'ValueType', 'char');

    % ---- len_const (5 EMU constants, via the constructors) ----
    putCanon(obs, 'len_const/inch', nh(Inches(1).emu));
    putCanon(obs, 'len_const/cm',   nh(Cm(1).emu));
    putCanon(obs, 'len_const/mm',   nh(Mm(1).emu));
    putCanon(obs, 'len_const/pt',   nh(Pt(1).emu));
    putCanon(obs, 'len_const/twip', nh(Twips(1).emu));

    % ---- ctor ----
    putCanon(obs, 'ctor/inches_1',        nh(Inches(1).emu));
    putCanon(obs, 'ctor/inches_half',     nh(Inches(0.5).emu));
    putCanon(obs, 'ctor/inches_neg_half', nh(Inches(-0.5).emu));
    putCanon(obs, 'ctor/cm_254',          nh(Cm(2.54).emu));
    putCanon(obs, 'ctor/mm_254',          nh(Mm(25.4).emu));
    putCanon(obs, 'ctor/mm_2405',         nh(Mm(240.5).emu));
    putCanon(obs, 'ctor/pt_72',           nh(Pt(72).emu));
    putCanon(obs, 'ctor/pt_half',         nh(Pt(0.5).emu));
    putCanon(obs, 'ctor/pt_neg_half',     nh(Pt(-0.5).emu));
    putCanon(obs, 'ctor/twips_1440',      nh(Twips(1440).emu));
    putCanon(obs, 'ctor/twips_09',        nh(Twips(0.9).emu));
    putCanon(obs, 'ctor/emu_914400',      nh(Emu(914400).emu));
    putCanon(obs, 'ctor/emu_25',          nh(Emu(2.5).emu));
    putCanon(obs, 'ctor/emu_neg_25',      nh(Emu(-2.5).emu));
    putCanon(obs, 'ctor/length_neg_half_emu', nh(Length(-0.5).emu));
    putCanon(obs, 'ctor/emu_str',         nh(Emu("914400").emu));
    putCanon(obs, 'ctor/emu_str_ws_sign_us', nh(Emu(" +914_400 ").emu));

    % ---- conv ----
    for e = [914400, 1000000]
        L = Emu(e); pfx = "conv/" + string(e) + "_";
        putCanon(obs, pfx + "emu",    nh(L.emu));
        putCanon(obs, pfx + "inches", nh(L.inches));
        putCanon(obs, pfx + "cm",     nh(L.cm));
        putCanon(obs, pfx + "mm",     nh(L.mm));
        putCanon(obs, pfx + "pt",     nh(L.pt));
        putCanon(obs, pfx + "twips",  nh(L.twips));
    end

    % ---- twips_edge ----
    edges = [317, 318, -318, 635, -635, 952, 953, -952, -953, ...
             634999682, 634999683, 6349996825, 6350003175, ...
             20002, 20003, -20003];
    for e = edges
        if e >= 0; id = "twips_edge/e_" + string(e);
        else;      id = "twips_edge/e_neg_" + string(-e); end
        putCanon(obs, id, nh(Emu(e).twips));
    end
    bandCases = {-1, "twips_edge/band_neg_1"; -317, "twips_edge/band_neg_317"; ...
                  1, "twips_edge/band_pos_1";  317, "twips_edge/band_pos_317"};
    for i = 1:size(bandCases, 1)
        putCanon(obs, bandCases{i, 2}, nh(Emu(bandCases{i, 1}).twips));
    end

    % ---- rgb ----
    putCanon(obs, 'rgb/str_ff0000', string(RGB(255, 0, 0).str_()));
    putCanon(obs, 'rgb/str_3c2f80', string(RGB(60, 47, 128).str_()));
    putCanon(obs, 'rgb/str_000000', string(RGB(0, 0, 0).str_()));
    putCanon(obs, 'rgb/str_ffffff', string(RGB(255, 255, 255).str_()));
    putCanon(obs, 'rgb/repr_3c2f80', string(RGB(60, 47, 128).repr_()));
    putCanon(obs, 'rgb/repr_ff0000', string(RGB(255, 0, 0).repr_()));
    putCanon(obs, 'rgb/from_string_ff0000', string(mat2doc.shared.RGBColor.from_string("FF0000").str_()));
    putCanon(obs, 'rgb/from_string_lower',  string(mat2doc.shared.RGBColor.from_string("3c2f80").str_()));
    putCanon(obs, 'rgb/from_string_repr',   string(mat2doc.shared.RGBColor.from_string("3C2F80").repr_()));
    putCanon(obs, 'rgb/from_string_5char',  string(mat2doc.shared.RGBColor.from_string("3C2F8").str_()));
    putCanon(obs, 'rgb/eq_true',  boolTok(RGB(255, 0, 0) == mat2doc.shared.RGBColor.from_string("FF0000")));
    putCanon(obs, 'rgb/eq_false', boolTok(RGB(1, 2, 3) == RGB(1, 2, 4)));
    putCanon(obs, 'rgb/ne_true',  boolTok(RGB(1, 2, 3) ~= RGB(1, 2, 4)));

    % ---- err ----
    M = "err_class_msg_exact"; D = "err_class_only_d003";
    putCanon(obs, 'err/rgb_high',        errCanon(@() RGB(256, 0, 0), M));
    putCanon(obs, 'err/rgb_low_neg',     errCanon(@() RGB(-1, 0, 0), M));
    putCanon(obs, 'err/rgb_g_range',     errCanon(@() RGB(0, 300, 0), M));
    putCanon(obs, 'err/rgb_b_range',     errCanon(@() RGB(0, 0, 256), M));
    putCanon(obs, 'err/rgb_nonint_float', errCanon(@() RGB(2.5, 0, 0), M));
    putCanon(obs, 'err/rgb_nonint_str',  errCanon(@() RGB("ff", 0, 0), M));
    putCanon(obs, 'err/emu_parse_25',    errCanon(@() Emu("2.5"), M));
    putCanon(obs, 'err/emu_parse_empty', errCanon(@() Emu(""), M));
    putCanon(obs, 'err/from_string_badhex', errCanon(@() mat2doc.shared.RGBColor.from_string("GGGGGG"), M));
    putCanon(obs, 'err/from_string_4char',  errCanon(@() mat2doc.shared.RGBColor.from_string("FF00"), M));
    putCanon(obs, 'err/from_string_7char',  errCanon(@() mat2doc.shared.RGBColor.from_string("3C2F80A"), M));
    putCanon(obs, 'err/inches_str', errCanon(@() Inches("1"), D));
    putCanon(obs, 'err/pt_str',     errCanon(@() Pt("1"), D));
    putCanon(obs, 'err/cm_str',     errCanon(@() Cm("1"), D));
    putCanon(obs, 'err/mm_str',     errCanon(@() Mm("1"), D));

    % ---- dev ----
    putCanon(obs, 'dev/dstype1_rgb_float', string(RGB(255.0, 0.0, 0.0).str_()));
    putCanon(obs, 'dev/dstype1_emu_float', nh(Emu(914400.0).emu));
    putCanon(obs, 'dev/d004_emu_nan', raisesTok(@() Emu(nan)));
    putCanon(obs, 'dev/d004_emu_inf', raisesTok(@() Emu(inf)));
    putCanon(obs, 'dev/d002_twips_str_1',      raisesTok(@() Twips("1")));
    putCanon(obs, 'dev/d002_twips_str_123456', raisesTok(@() Twips("123456")));
end

function putCanon(map, id, canon)
    map(char(id)) = char(canon);
end

function s = boolTok(tf)
    if tf; s = "true"; else; s = "false"; end
end

function c = errCanon(thunk, policy)
    try
        thunk();
        c = "NO_RAISE";
    catch ME
        parts = split(string(ME.identifier), ":");
        cls = parts(end);                       % strip "mat2doc:" -> class token
        if policy == "err_class_msg_exact"
            c = cls + "|" + string(ME.message);
        else
            c = cls;                            % class only (D-003)
        end
    end
end

function c = raisesTok(thunk)
    try
        thunk();
        c = "NO_RAISE";
    catch
        c = "RAISES";
    end
end

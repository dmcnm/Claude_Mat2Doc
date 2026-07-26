classdef Test_p3_2_simpletypes < matlab.unittest.TestCase
% TEST_P3_2_SIMPLETYPES  Gate-4 permanent unit tests for Mat2Doc P3-2
%   (src/docx/oxml/simpletypes.py -> +mat2doc\+oxml\+simpletypes: the 35
%   Base*/Xsd*/ST_* attribute-VALUE validators/converters).
%
%   Simple-types emit NO package bytes -- they validate / convert attribute
%   VALUES. There is therefore no L0-L3 package ladder; every guarantee frozen
%   here is a returned VALUE or a raised IDENTIFIER+message, pinned against the
%   Gate-3 python-docx 1.2.0 oracle (validate_P3-2_simpletypes.md; oracle frozen
%   at references\s0017\probe.json, copied verbatim into
%   tests\oxml\data\s0017_probe_oracle.json so this suite is self-contained).
%
%   Provenance (Gate-1..3, all 2026-07-26):
%     * Audit    : validation\mat2doc\audit_P3-2_simpletypes.md (Porter Gate-1 +
%                  Opus Gate-2 adversarial APPROVE; the F1 XsdBoolean render fix
%                  applied inline).
%     * Validate : validation\mat2doc\validate_P3-2_simpletypes.md (Gate-3 PASS,
%                  175/175 probe facts across all 35 classes; 154 byte-identical,
%                  21 divergences EACH mapped to an already-adopted ruling; ZERO
%                  new D-numbers; regression 439/439).
%     * Scenario : validation\mat2doc\scenarios\s0017_simpletypes_probe.{py,m}
%                  (the 175-fact battery; the .m twin's probe sequence is replayed
%                  verbatim by test_equivalence_full_175_vs_frozen_oracle below).
%     * Frozen ref (self-contained): tests\oxml\data\s0017_probe_oracle.json --
%                  the python-docx 1.2.0 ORACLE (175 tagged canonical facts).
%
%   THE 21 ADOPTED DIVERGENCES (port != oracle) -- pinned to the PORT's actual
%   behaviour with the governing D-number, so a regression that "accidentally
%   fixes" one to the Python value goes RED and is reviewed, not silently merged:
%     * D-005  (MATLAB class token where CPython prints <class '...'>): B12, E13,
%              P06, P09, R04, R08, X04.
%     * D-STYPE-3 (rounded 2^63 / 2^64 long bound or value >2^53): H08, R06, R10, P27.
%     * D-STYPE-2 / D-004 (str2double vs CPython float() lexis): U12 (underscore
%              -> mustBeFinite), U13 (thousands-comma OVER-accept -> 12700000),
%              S07 (non-numeric -> mustBeFinite).
%     * D-002  (ASCII-only int() grammar): I06 (Arabic-Indic '5' rejected, not ->5).
%     * VERIFY-fromiso / P8-2 (narrow fromisoformat subset -> 1970 epoch on
%              schema-INVALID xsd:dateTime): D07, D11, D22, D23, D24. These pins
%              deliberately freeze the CURRENT port epoch behaviour per
%              validation\summary\decision_2026-07-26_st_datetime_underaccept.md
%              (defer-to-P8-2). A future P8-2 fix that widens the subset to full
%              CPython parity WILL flip these five -- that is an EXPECTED,
%              documented flip, not a surprise regression (see the dedicated test
%              test_datetime_epoch_underaccept_deferred_to_P8_2).
%     * dead-abstract-base: P01 (BaseSimpleType.from_xml unreachable -- from_xml/
%              to_xml are hosted on the branch bases per H10, not the abstract root).
%
%   THE F1 PIN (Gate-2 auditor fix, must never regress): XsdBoolean.to_xml([])
%   (the None analogue) raises mat2doc:TypeError with the verbatim template
%   "only True or False (and possibly None) may be assigned, got 'double'" --
%   the pre-F1 mat2doc:pyStr:unsupportedType crash is GONE
%   (test_xsdboolean_F1_typeerror_none_render). Only the value token differs from
%   Python's 'None' (D-005); the identifier + template are exact.
%
%   Coverage taxonomy
%   -----------------
%   * Nominal   -- every converter's documented happy path (all 6 measure units,
%     bool set, hex round-trip, canonical datetime round-trip, enum members).
%   * Edge      -- empty ("" -> ST_OnOff InvalidXmlError), non-ASCII (Arabic-Indic
%     digit built via char(1637), D-002 reject), single-element enum (ST_HexColorAuto
%     singleton), zero / dead range edges (2^63 / 2^64), error paths verifying the
%     mat2doc:<PyExc> IDENTIFIER (not merely that it throws).
%   * Equivalence -- test_equivalence_full_175_vs_frozen_oracle replays the whole
%     s0017 battery live and compares every fact to the frozen oracle (or, for the
%     21 adopted divergences, to the pinned port value), tying the suite to the
%     Gate-3 reference.
%   * Regression -- hard-coded expected VALUES / verbatim messages throughout.
%
%   Deviations exercised: adopt-only (D-002, D-004, D-005, D-STYPE-2, D-STYPE-3);
%   ZERO new D-numbers (Gate-3). Nothing serialized -> no L-ladder leg.
%
%   Determinism: no network, no absolute paths, no file WRITES (the one file READ
%   is the co-located frozen oracle, resolved via fileparts(mfilename('fullpath'))
%   and read in BINARY mode / decoded UTF-8, line-ending agnostic). The +mat2doc
%   package resolves via the MANDATORY PathFixture(worktree-root) added in
%   TestClassSetup (WP9-F4 lesson).

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied verbatim from tests\oxml\Test_p2_3_document_shell.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\oxml
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. Base / Xsd tier                                              %
        % =============================================================== %

        function test_baseinttype_nominal_and_passthru(testCase)
            % Nominal: BaseIntType round-trips int() <-> str; XsdString/Token/Id/
            % AnyUri are string pass-throughs. (s0017 P02/P03/P07/P08/P10-P15.)
            import mat2doc.oxml.simpletypes.*
            testCase.verifyEqual(double(BaseIntType.from_xml("42")), 42);
            testCase.verifyEqual(BaseIntType.to_xml(42), "42");
            testCase.verifyEqual(XsdString.from_xml("s"), "s");
            testCase.verifyEqual(XsdString.to_xml("s"), "s");
            testCase.verifyEqual(XsdToken.from_xml("  a  b  "), "  a  b  ");
            testCase.verifyEqual(XsdId.to_xml("_id1"), "_id1");
            testCase.verifyEqual(XsdAnyUri.to_xml("http://x/y"), "http://x/y");
        end

        function test_baseinttype_int_grammar(testCase)
            % D-002 grammar: int() accepts +/-, leading zeros, and (via CPython)
            % surrounding whitespace / underscores; rejects a decimal literal.
            % (s0017 I01-I05,I07.)
            import mat2doc.oxml.simpletypes.*
            testCase.verifyEqual(double(XsdInt.from_xml("+5")), 5);
            testCase.verifyEqual(double(XsdInt.from_xml("007")), 7);
            testCase.verifyEqual(double(XsdInt.from_xml(" 42 ")), 42);
            testCase.verifyEqual(double(XsdInt.from_xml("1_000")), 1000);
            testCase.verifyEqual(double(XsdInt.from_xml("-5")), -5);
            [id, msg] = raiseAndCapture(@() XsdInt.from_xml("2.5"));
            testCase.verifyEqual(id, 'mat2doc:ValueError');
            testCase.verifyEqual(msg, 'invalid literal for int() with base 10: ''2.5''', ...
                'int() grammar must reject a decimal literal, verbatim CPython message');
        end

        function test_baseinttype_nonascii_digit_rejected_D002(testCase)
            % D-002 (adopted): the int() grammar accepts ASCII [0-9] ONLY; CPython
            % int() would accept the Arabic-Indic digit U+0665 -> 5, the port
            % rejects it. The input is built via char(1637) so the expectation is
            % independent of THIS file's source encoding (P3-1 cafe precedent).
            % Pinned to the PORT (reject), NOT the oracle (accept). (s0017 I06.)
            import mat2doc.oxml.simpletypes.*
            digit = string(char(1637));                 % Arabic-Indic FIVE
            [id, msg] = raiseAndCapture(@() XsdInt.from_xml(digit));
            testCase.verifyEqual(id, 'mat2doc:ValueError', ...
                'D-002: non-ASCII digit must raise mat2doc:ValueError (safe under-accept)');
            testCase.verifyEqual(string(msg), ...
                "invalid literal for int() with base 10: '" + digit + "'", ...
                'the non-ASCII digit must render byte-identically in the message');
        end

        function test_baseinttype_nan_inf_reject(testCase)
            % D-002 grammar rejects non-finite / non-numeric lexis. "NaN"/"Inf" are
            % NOT valid xsd int literals -> int() ValueError (not a silent NaN).
            import mat2doc.oxml.simpletypes.*
            for tok = ["NaN" "Inf" "-Inf" "1e3"]
                [id, msg] = raiseAndCapture(@() XsdInt.from_xml(tok));
                testCase.verifyEqual(id, 'mat2doc:ValueError', ...
                    sprintf('int(%s) must raise ValueError, not accept a non-finite', tok));
                testCase.verifyEqual(string(msg), ...
                    "invalid literal for int() with base 10: '" + tok + "'");
            end
        end

        function test_xsd_int_ranges(testCase)
            % Range guards: XsdInt +/-2^31, XsdUnsignedInt 0..2^32-1. Off-by-one
            % over the top raises ValueError with the verbatim range message; the
            % negative floor and the exact bounds pass. (s0017 R01,R02,R09,P22,P25.)
            import mat2doc.oxml.simpletypes.*
            testCase.verifyEqual(XsdInt.to_xml(-2147483648), "-2147483648");
            testCase.verifyEqual(XsdUnsignedInt.to_xml(4294967295), "4294967295");
            [id, msg] = raiseAndCapture(@() XsdInt.to_xml(2147483648));
            testCase.verifyEqual(id, 'mat2doc:ValueError');
            testCase.verifyEqual(msg, ...
                'value must be in range -2147483648 to 2147483647 inclusive, got 2147483648');
            [id2, msg2] = raiseAndCapture(@() XsdUnsignedInt.to_xml(-1));
            testCase.verifyEqual(id2, 'mat2doc:ValueError');
            testCase.verifyEqual(msg2, ...
                'value must be in range 0 to 4294967295 inclusive, got -1');
        end

        function test_xsd_long_unsignedlong_D_STYPE3_dead_edge(testCase)
            % XsdLong/XsdUnsignedLong exact-zero + mid-range pass; the 2^63/2^64
            % bounds are a DEAD upper edge (>2^53 is not a representable double).
            % Pinned to the PORT's rounded bound/value (D-STYPE-3), NOT the exact
            % CPython digits -- unreachable (every concrete subclass sits << 2^53).
            % (s0017 R05,R07,P27,R10,H08,R06.)
            import mat2doc.oxml.simpletypes.*
            testCase.verifyEqual(XsdUnsignedLong.to_xml(0), "0");
            testCase.verifyEqual(XsdLong.to_xml(9007199254740992), "9007199254740992");
            % D-STYPE-3 value edge: 2^64-1 is not representable -> port returns 2^64.
            testCase.verifyEqual(double(XsdUnsignedLong.from_xml("18446744073709551615")), ...
                18446744073709551616, ...
                'D-STYPE-3: the 2^64-1 dead edge round-trips through the nearest double (2^64)');
            % D-STYPE-3 message bound: 2^64 renders as ...616 (not the exact ...615).
            [idU, msgU] = raiseAndCapture(@() XsdUnsignedLong.to_xml(-1));
            testCase.verifyEqual(idU, 'mat2doc:ValueError');
            testCase.verifyEqual(msgU, ...
                'value must be in range 0 to 18446744073709551616 inclusive, got -1', ...
                'D-STYPE-3: the unsigned-long upper bound renders as the rounded 2^64');
            % D-STYPE-3: 2^63 signed bound renders as ...808 (not the exact ...807).
            [idL, msgL] = raiseAndCapture(@() XsdLong.to_xml(-18446744073709551616));
            testCase.verifyEqual(idL, 'mat2doc:ValueError');
            testCase.verifyEqual(msgL, ...
                'value must be in range -9223372036854775808 to 9223372036854775808 inclusive, got -18446744073709551616');
        end

        function test_xsd_typeerror_int_type_token_D005(testCase)
            % D-005: passing a non-int to an int validator raises TypeError with
            % the MATLAB class token where CPython prints <class 'str'>/<class
            % 'float'>. Pinned to the PORT token. (s0017 R04,R08.)
            import mat2doc.oxml.simpletypes.*
            [id1, msg1] = raiseAndCapture(@() XsdInt.to_xml("5"));
            testCase.verifyEqual(id1, 'mat2doc:TypeError');
            testCase.verifyEqual(msg1, 'value must be <type ''int''>, got string', ...
                'D-005: str input -> port class token "string"');
            [id2, msg2] = raiseAndCapture(@() XsdInt.to_xml(2.5));
            testCase.verifyEqual(id2, 'mat2doc:TypeError');
            testCase.verifyEqual(msg2, 'value must be <type ''int''>, got double', ...
                'D-005: float input -> port class token "double"');
        end

        function test_xsdstring_typeerror_nonstring_D005(testCase)
            % D-005: a non-string to a string validator -> TypeError "value must be
            % a string, got double" (CPython prints <class 'int'>). BaseStringType
            % and XsdString share the guard. (s0017 P06/P09.)
            import mat2doc.oxml.simpletypes.*
            [id1, msg1] = raiseAndCapture(@() BaseStringType.to_xml(5));
            testCase.verifyEqual(id1, 'mat2doc:TypeError');
            testCase.verifyEqual(msg1, 'value must be a string, got double');
            [id2, msg2] = raiseAndCapture(@() XsdString.to_xml(5));
            testCase.verifyEqual(id2, 'mat2doc:TypeError');
            testCase.verifyEqual(msg2, 'value must be a string, got double');
        end

        % =============================================================== %
        % 2. XsdBoolean / ST_OnOff (incl. the F1 pin + InvalidXmlError)   %
        % =============================================================== %

        function test_xsdboolean_full_bool_set(testCase)
            % Nominal: '1'/'true' -> true, '0'/'false' -> false; to_xml maps back;
            % 1 (==True) accepted by validate. (s0017 B01-B04,B07-B09.)
            import mat2doc.oxml.simpletypes.*
            testCase.verifyTrue(XsdBoolean.from_xml("1"));
            testCase.verifyTrue(XsdBoolean.from_xml("true"));
            testCase.verifyFalse(XsdBoolean.from_xml("0"));
            testCase.verifyFalse(XsdBoolean.from_xml("false"));
            testCase.verifyEqual(XsdBoolean.to_xml(true), "1");
            testCase.verifyEqual(XsdBoolean.to_xml(false), "0");
            testCase.verifyEqual(XsdBoolean.to_xml(1), "1");
        end

        function test_xsdboolean_F1_typeerror_none_render(testCase)
            % ** THE F1 PIN (Gate-2 auditor fix, must never regress). **
            % XsdBoolean.to_xml([]) (the Python None analogue) raises
            % mat2doc:TypeError with the verbatim validate template -- the pre-F1
            % mat2doc:pyStr:unsupportedType CRASH is gone. The got-token is 'double'
            % (D-005), NOT Python's 'None'; pinned to the PORT. (s0017 B12.)
            import mat2doc.oxml.simpletypes.*
            [id, msg] = raiseAndCapture(@() XsdBoolean.to_xml([]));
            testCase.verifyEqual(id, 'mat2doc:TypeError', ...
                'F1: to_xml([]) must raise mat2doc:TypeError, NOT the old pyStr crash');
            testCase.verifyNotEqual(id, 'mat2doc:pyStr:unsupportedType', ...
                'F1 regression guard: the pre-fix pyStr crash identifier must never return');
            testCase.verifyEqual(msg, ...
                'only True or False (and possibly None) may be assigned, got ''double''', ...
                'F1: verbatim validate template with the D-005 port token (double)');
            % A non-0/1 numeric (2) and a string ('x') also raise the same template.
            [~, msg2] = raiseAndCapture(@() XsdBoolean.to_xml(2));
            testCase.verifyEqual(msg2, ...
                'only True or False (and possibly None) may be assigned, got ''2''');
            [~, msg3] = raiseAndCapture(@() XsdBoolean.to_xml("x"));
            testCase.verifyEqual(msg3, ...
                'only True or False (and possibly None) may be assigned, got ''x''');
        end

        function test_xsdboolean_invalidxmlerror_on_bad_token(testCase)
            % XsdBoolean.convert_from_xml raises docx.exceptions.InvalidXmlError
            % (NOT ValueError) for a token outside {1,0,true,false}. Case-sensitive:
            % 'on'/'TRUE' reject. (s0017 B05/B06.)
            import mat2doc.oxml.simpletypes.*
            [id, msg] = raiseAndCapture(@() XsdBoolean.from_xml("on"));
            testCase.verifyEqual(id, 'mat2doc:InvalidXmlError', ...
                'bad bool token must raise mat2doc:InvalidXmlError, not ValueError');
            testCase.verifyEqual(msg, ...
                'value must be one of ''1'', ''0'', ''true'' or ''false'', got ''on''');
            [id2, msg2] = raiseAndCapture(@() XsdBoolean.from_xml("TRUE"));
            testCase.verifyEqual(id2, 'mat2doc:InvalidXmlError');
            testCase.verifyEqual(msg2, ...
                'value must be one of ''1'', ''0'', ''true'' or ''false'', got ''TRUE''');
        end

        function test_st_onoff_full_set(testCase)
            % ST_OnOff widens the accepted XML set to add 'on'/'off'. Nominal:
            % all six tokens map; to_xml(true) -> "1". (s0017 O01-O06,O08.)
            import mat2doc.oxml.simpletypes.*
            expect = [true false true false true false];
            toks   = ["1" "0" "true" "false" "on" "off"];
            for i = 1:6
                testCase.verifyEqual(ST_OnOff.from_xml(toks(i)), expect(i), ...
                    sprintf('ST_OnOff.from_xml("%s")', toks(i)));
            end
            testCase.verifyEqual(ST_OnOff.to_xml(true), "1");
        end

        function test_st_onoff_invalidxmlerror_incl_empty(testCase)
            % Edge + error path: ST_OnOff.convert_from_xml raises InvalidXmlError
            % with the SIX-token message (note the two-line CPython literal split);
            % case-sensitive ('On' rejects); the EMPTY string "" rejects too.
            % (s0017 O07/O09/O10.)
            import mat2doc.oxml.simpletypes.*
            wants = 'value must be one of ''1'', ''0'', ''true'', ''false'', ''on'', or ''off'', got ''%s''';
            for pair = {"maybe","maybe"; "On","On"; "",""}'
                inTok = pair{1};
                [id, msg] = raiseAndCapture(@() ST_OnOff.from_xml(inTok));
                testCase.verifyEqual(id, 'mat2doc:InvalidXmlError', ...
                    sprintf('ST_OnOff.from_xml("%s") must raise InvalidXmlError', inTok));
                testCase.verifyEqual(string(msg), string(sprintf(wants, pair{2})));
            end
        end

        % =============================================================== %
        % 3. ST_ measures + banker's rounding                            %
        % =============================================================== %

        function test_st_universalmeasure_six_units(testCase)
            % Nominal: all SIX universal-measure units convert exactly to EMU.
            % (s0017 U01-U06,U14.)
            import mat2doc.oxml.simpletypes.*
            testCase.verifyEqual(double(ST_UniversalMeasure.from_xml("1in")),  914400);
            testCase.verifyEqual(double(ST_UniversalMeasure.from_xml("2.5cm")), 900000);
            testCase.verifyEqual(double(ST_UniversalMeasure.from_xml("12pt")),  152400);
            testCase.verifyEqual(double(ST_UniversalMeasure.from_xml("10mm")),  360000);
            testCase.verifyEqual(double(ST_UniversalMeasure.from_xml("1pc")),   152400);
            testCase.verifyEqual(double(ST_UniversalMeasure.from_xml("1pi")),   152400);
        end

        function test_st_universalmeasure_signed_fraction_exp(testCase)
            % Signed, leading-dot fraction, and 1e2 exponent forms. (s0017 U07-U10.)
            import mat2doc.oxml.simpletypes.*
            testCase.verifyEqual(double(ST_UniversalMeasure.from_xml("-1.5pt")), -19050);
            testCase.verifyEqual(double(ST_UniversalMeasure.from_xml(".5in")),   457200);
            testCase.verifyEqual(double(ST_UniversalMeasure.from_xml("1e2pt")),  1270000);
            testCase.verifyEqual(double(ST_UniversalMeasure.from_xml("0.5pt")),  6350);
        end

        function test_st_universalmeasure_unknown_unit_keyerror(testCase)
            % Error path: an unrecognized 2-char unit raises mat2doc:KeyError with
            % the CPython key repr 'xy' (quotes included). (s0017 U11.)
            import mat2doc.oxml.simpletypes.*
            [id, msg] = raiseAndCapture(@() ST_UniversalMeasure.from_xml("5xy"));
            testCase.verifyEqual(id, 'mat2doc:KeyError', ...
                'unknown unit must raise mat2doc:KeyError');
            testCase.verifyEqual(msg, '''xy''', ...
                'KeyError message is the CPython key repr with single quotes');
        end

        function test_st_universalmeasure_str2double_deviations(testCase)
            % D-STYPE-2 / D-004: str2double vs CPython float() lexis. Pinned to the
            % PORT. U12 underscore -> mustBeFinite (Python -> 127000); U13
            % thousands-comma is the one OVER-accept -> 12700000 (Python raises).
            % Both inputs are outside -?\d+(\.\d+)?(unit) and unreachable. (s0017 U12/U13.)
            import mat2doc.oxml.simpletypes.*
            [id, msg] = raiseAndCapture(@() ST_UniversalMeasure.from_xml("1_0pt"));
            testCase.verifyEqual(id, 'MATLAB:validators:mustBeFinite', ...
                'D-STYPE-2/D-004: underscore lexis -> str2double NaN -> mustBeFinite');
            testCase.verifyEqual(msg, 'Value must be finite.');
            testCase.verifyEqual(double(ST_UniversalMeasure.from_xml("1,000pt")), 12700000, ...
                'D-STYPE-2 OVER-accept: str2double("1,000")=1000 (Python float() would raise)');
        end

        function test_st_twips_and_hps_measures(testCase)
            % ST_TwipsMeasure / ST_SignedTwipsMeasure / ST_HpsMeasure: universal
            % strings AND bare numbers; to_xml re-derives the integer count.
            % (s0017 T01-T05,S01,S04-S06,H01-H07.)
            import mat2doc.oxml.simpletypes.*
            import mat2doc.shared.*
            testCase.verifyEqual(double(ST_TwipsMeasure.from_xml("1440")), 914400);
            testCase.verifyEqual(double(ST_TwipsMeasure.from_xml("1in")),  914400);
            testCase.verifyEqual(ST_TwipsMeasure.to_xml(Twips(1440)), "1440");
            testCase.verifyEqual(ST_TwipsMeasure.to_xml(Emu(635)), "1");
            testCase.verifyEqual(ST_TwipsMeasure.to_xml(Emu(300)), "0");
            testCase.verifyEqual(double(ST_SignedTwipsMeasure.from_xml("-360")), -228600);
            testCase.verifyEqual(double(ST_SignedTwipsMeasure.from_xml("1in")), 914400);
            testCase.verifyEqual(ST_SignedTwipsMeasure.to_xml(Twips(-360)), "-360");
            testCase.verifyEqual(double(ST_HpsMeasure.from_xml("24")), 152400);
            testCase.verifyEqual(double(ST_HpsMeasure.from_xml("12pt")), 152400);
            testCase.verifyEqual(double(ST_HpsMeasure.from_xml("25")), 158750);
            testCase.verifyEqual(ST_HpsMeasure.to_xml(Pt(12)), "24");
            testCase.verifyEqual(ST_HpsMeasure.to_xml(Pt(12.25)), "24");     % banker's -> 24
            testCase.verifyEqual(ST_HpsMeasure.to_xml(Emu(152400)), "24");
        end

        function test_st_signedtwips_float_parse_error_D004(testCase)
            % D-004: a non-numeric to str2double -> mustBeFinite (Python float()
            % raises "could not convert string to float"). Pinned to the PORT.
            % (s0017 S07.)
            import mat2doc.oxml.simpletypes.*
            [id, msg] = raiseAndCapture(@() ST_SignedTwipsMeasure.from_xml("abc"));
            testCase.verifyEqual(id, 'MATLAB:validators:mustBeFinite', ...
                'D-004: non-numeric float lexis -> port mustBeFinite (error-class divergence)');
            testCase.verifyEqual(msg, 'Value must be finite.');
        end

        function test_bankers_rounding_round_half_to_even(testCase)
            % ST_SignedTwipsMeasure exercises pyRound (round-half-to-even): 0.5->0,
            % 1.5->2(->1270 EMU), 2.5->2, -0.5->0, and non-half values round normal.
            % (s0017 PR01-PR09; the EMU values are the twips*635 products.)
            import mat2doc.oxml.simpletypes.*
            expect = [0 1270 1270 0 -1270 -1270 1905 7620 8890];
            inputs = ["0.5" "1.5" "2.5" "-0.5" "-1.5" "-2.5" "2.675" "12.5" "13.5"];
            for i = 1:numel(inputs)
                testCase.verifyEqual(double(ST_SignedTwipsMeasure.from_xml(inputs(i))), ...
                    expect(i), sprintf('banker''s rounding of %s', inputs(i)));
            end
        end

        function test_st_coordinate_range_and_parse(testCase)
            % ST_Coordinate / ST_PositiveCoordinate / ST_CoordinateUnqualified:
            % EMU int + universal string; the outer range guard; the int() lexis
            % error; the positive-coordinate floor. (s0017 C01-C12.)
            import mat2doc.oxml.simpletypes.*
            import mat2doc.shared.*
            testCase.verifyEqual(double(ST_Coordinate.from_xml("914400")), 914400);
            testCase.verifyEqual(double(ST_Coordinate.from_xml("-27273042329600")), -27273042329600);
            testCase.verifyEqual(double(ST_Coordinate.from_xml("2.5cm")), 900000);
            testCase.verifyEqual(ST_Coordinate.to_xml(Emu(914400)), "914400");
            testCase.verifyEqual(ST_CoordinateUnqualified.to_xml(-27273042329600), "-27273042329600");
            [id, msg] = raiseAndCapture(@() ST_Coordinate.to_xml(Emu(27273042316901)));
            testCase.verifyEqual(id, 'mat2doc:ValueError');
            testCase.verifyEqual(msg, ...
                'value must be in range -27273042329600 to 27273042316900 inclusive, got 27273042316901');
            [id2, msg2] = raiseAndCapture(@() ST_Coordinate.from_xml("abc"));
            testCase.verifyEqual(id2, 'mat2doc:ValueError');
            testCase.verifyEqual(msg2, 'invalid literal for int() with base 10: ''abc''');
            [id3, msg3] = raiseAndCapture(@() ST_PositiveCoordinate.to_xml(Emu(-1)));
            testCase.verifyEqual(id3, 'mat2doc:ValueError');
            testCase.verifyEqual(msg3, 'value must be in range 0 to 27273042316900 inclusive, got -1');
        end

        % =============================================================== %
        % 4. ST_ hex color + ST_ string-enum member sets                 %
        % =============================================================== %

        function test_st_hexcolor_roundtrip_and_auto(testCase)
            % ST_HexColor: from_xml("3C2F80") -> RGBColor; to_xml -> "3C2F80";
            % lowercase input normalises to UPPER; the "auto" literal passes
            % through both directions. (s0017 X01,X02,X03,X07.)
            import mat2doc.oxml.simpletypes.*
            rgb = ST_HexColor.from_xml("FF0000");
            testCase.verifyClass(rgb, 'mat2doc.shared.RGBColor');
            testCase.verifyEqual(rgb.str_(), "FF0000");
            testCase.verifyEqual(ST_HexColor.to_xml(ST_HexColor.from_xml("3C2F80")), "3C2F80");
            testCase.verifyEqual(ST_HexColor.to_xml(ST_HexColor.from_xml("3c2f80")), "3C2F80", ...
                'lowercase hex must normalise to UPPER on to_xml');
            testCase.verifyEqual(ST_HexColor.from_xml("auto"), "auto");
        end

        function test_st_hexcolor_invalid_and_typeerror(testCase)
            % Error paths: a non-hex pair -> int(base 16) ValueError; a non-RGBColor
            % to to_xml -> ValueError with the D-005 port class token ("string blue",
            % where CPython prints "<class 'str'> blue"). (s0017 X04,X05,X06.)
            import mat2doc.oxml.simpletypes.*
            [id, msg] = raiseAndCapture(@() ST_HexColor.from_xml("GGGGGG"));
            testCase.verifyEqual(id, 'mat2doc:ValueError');
            testCase.verifyEqual(msg, 'invalid literal for int() with base 16: ''GG''');
            [id2, msg2] = raiseAndCapture(@() ST_HexColor.from_xml("Auto"));  % case-sensitive
            testCase.verifyEqual(id2, 'mat2doc:ValueError');
            testCase.verifyEqual(msg2, 'invalid literal for int() with base 16: ''Au''');
            [id3, msg3] = raiseAndCapture(@() ST_HexColor.to_xml("blue"));
            testCase.verifyEqual(id3, 'mat2doc:ValueError');
            testCase.verifyEqual(msg3, 'rgb color value must be RGBColor object, got string blue', ...
                'D-005: non-RGBColor value renders the MATLAB class token "string"');
        end

        function test_st_hexcolorauto_singleton(testCase)
            % Single-element edge: ST_HexColorAuto is the ONE-member ('auto',)
            % enum; "auto" passes both directions, anything else -> ValueError with
            % the one-tuple repr "('auto',)". (s0017 A01,A02,A03.)
            import mat2doc.oxml.simpletypes.*
            testCase.verifyEqual(ST_HexColorAuto.to_xml("auto"), "auto");
            testCase.verifyEqual(ST_HexColorAuto.from_xml("auto"), "auto");
            [id, msg] = raiseAndCapture(@() ST_HexColorAuto.to_xml("x"));
            testCase.verifyEqual(id, 'mat2doc:ValueError');
            testCase.verifyEqual(msg, 'must be one of (''auto'',), got ''x''', ...
                'the singleton must render the one-element tuple repr with the trailing comma');
        end

        function test_st_string_enum_member_sets(testCase)
            % Regression: the EXACT accepted member set of each ST_ string enum is
            % frozen via its ValueError tuple-repr -- a missing OR extra member
            % moves the message and goes RED. Also pins one nominal member each.
            % (s0017 E01-E12,E14-E20.)
            import mat2doc.oxml.simpletypes.*
            cases = {
                @ST_BrClear,          "left",       "must one of ('none', 'left', 'right', 'all'), got 'x'"
                @ST_BrType,           "textWrapping","must one of ('page', 'column', 'textWrapping'), got 'x'"
                @ST_TblLayoutType,    "autofit",    "must one of ('fixed', 'autofit'), got 'x'"
                @ST_TblWidth,         "dxa",        "must one of ('auto', 'dxa', 'nil', 'pct'), got 'x'"
                @ST_Merge,            "restart",    "must one of ('continue', 'restart'), got 'x'"
                @ST_VerticalAlignRun, "subscript",  "must one of ('baseline', 'superscript', 'subscript'), got 'x'"
            };
            for k = 1:size(cases,1)
                cls = cases{k,1}; okMember = cases{k,2};
                wantMsg = "must be one of " + extractAfter(cases{k,3}, "must one of ");
                testCase.verifyEqual(cls().to_xml(okMember), okMember, ...
                    sprintf('%s.to_xml("%s") nominal member must pass through', func2str(cls), okMember));
                [id, msg] = raiseAndCapture(@() cls().to_xml("x"));
                testCase.verifyEqual(id, 'mat2doc:ValueError');
                testCase.verifyEqual(string(msg), wantMsg, ...
                    sprintf('%s member set frozen via its ValueError tuple-repr', func2str(cls)));
            end
            % A representative from_xml + a couple more nominal members.
            testCase.verifyEqual(ST_Merge.from_xml("restart"), "restart");
            testCase.verifyEqual(ST_BrClear.to_xml("none"), "none");
            testCase.verifyEqual(ST_TblWidth.to_xml("nil"), "nil");
        end

        function test_st_string_enum_typeerror_nonstring_D005(testCase)
            % D-005: a non-string to a string-enum -> TypeError "value must be a
            % string, got double" (CPython <class 'int'>). Pinned to the PORT token.
            % (s0017 E13.)
            import mat2doc.oxml.simpletypes.*
            [id, msg] = raiseAndCapture(@() ST_BrClear.to_xml(5));
            testCase.verifyEqual(id, 'mat2doc:TypeError');
            testCase.verifyEqual(msg, 'value must be a string, got double');
        end

        % =============================================================== %
        % 5. ST_DateTime                                                  %
        % =============================================================== %

        function test_datetime_canonical_z_roundtrip(testCase)
            % Nominal: a canonical Z form parses to a tz-aware UTC instant and
            % re-emits byte-identically; fractional seconds truncate on to_xml.
            % (s0017 D01,D02,D03,D12.)
            import mat2doc.oxml.simpletypes.*
            testCase.verifyEqual(canonDT(ST_DateTime.from_xml("2023-01-02T03:04:05Z")), ...
                "dt-aware|2023-01-02T03:04:05+00:00");
            testCase.verifyEqual(canonDT(ST_DateTime.from_xml("2023-01-02T03:04:05.123456Z")), ...
                "dt-aware|2023-01-02T03:04:05.123456+00:00");
            testCase.verifyEqual(ST_DateTime.to_xml(ST_DateTime.from_xml("2023-01-02T03:04:05Z")), ...
                "2023-01-02T03:04:05Z");
            testCase.verifyEqual(ST_DateTime.to_xml(ST_DateTime.from_xml("2023-01-02T03:04:05.999999Z")), ...
                "2023-01-02T03:04:05Z", 'fractional seconds truncate (whole-second strftime)');
        end

        function test_datetime_offset_to_utc_instant_VERIFYtz(testCase)
            % VERIFY-tz: a non-UTC offset is stored as the equivalent UTC INSTANT
            % (TimeZone 'UTC'), never the original offset. -08:00 -> 11:04:05Z;
            % +05:30 -> previous-day 21:34:05Z. The instant is compared, not the
            % zone label. (s0017 D04,D16.)
            import mat2doc.oxml.simpletypes.*
            testCase.verifyEqual(canonDT(ST_DateTime.from_xml("2023-01-02T03:04:05-08:00")), ...
                "dt-aware|2023-01-02T11:04:05+00:00");
            testCase.verifyEqual(canonDT(ST_DateTime.from_xml("2023-01-02T03:04:05+05:30")), ...
                "dt-aware|2023-01-01T21:34:05+00:00");
        end

        function test_datetime_naive_and_dateonly(testCase)
            % Naive (no Z / no offset) -> empty-TimeZone datetime; date-only ->
            % naive midnight; space separator + fractional seconds. (s0017 D05,D06,D15,D17.)
            import mat2doc.oxml.simpletypes.*
            testCase.verifyEqual(canonDT(ST_DateTime.from_xml("2023-01-02 03:04:05")), ...
                "dt-naive|2023-01-02T03:04:05");
            testCase.verifyEqual(canonDT(ST_DateTime.from_xml("2023-06-15")), ...
                "dt-naive|2023-06-15T00:00:00");
            testCase.verifyEqual(canonDT(ST_DateTime.from_xml("2023-01-02T03:04:05.5")), ...
                "dt-naive|2023-01-02T03:04:05.500000");
            testCase.verifyEqual(canonDT(ST_DateTime.from_xml("2023-01-02 03:04:05.123")), ...
                "dt-naive|2023-01-02T03:04:05.123000");
        end

        function test_datetime_to_xml_and_typeerror(testCase)
            % to_xml of an aware datetime; the epoch of a garbage parse re-emits
            % the 1970 Z; validate rejects a non-datetime with TypeError. (s0017 D13,D14,D18,D19.)
            import mat2doc.oxml.simpletypes.*
            testCase.verifyEqual( ...
                ST_DateTime.to_xml(datetime(2023,6,15,12,0,0,"TimeZone","UTC")), ...
                "2023-06-15T12:00:00Z");
            testCase.verifyEqual( ...
                ST_DateTime.to_xml(ST_DateTime.from_xml("not-a-date")), ...
                "1970-01-01T00:00:00Z", 'a garbage parse epochs, then re-emits 1970Z');
            [id, msg] = raiseAndCapture(@() ST_DateTime.to_xml("x"));
            testCase.verifyEqual(id, 'mat2doc:TypeError');
            testCase.verifyEqual(msg, 'only a datetime.datetime object may be assigned, got ''x''');
        end

        function test_datetime_epoch_on_invalid(testCase)
            % NEVER-RAISE parse: outright-garbage, impossible calendar dates, and
            % year 0 all fall to the 1970 UTC epoch (mutual with Python). Includes
            % D21 (secondless WITH Z) -- the true mutual-epoch case. (s0017 D08,D09,D10,D21.)
            import mat2doc.oxml.simpletypes.*
            for badInput = ["not-a-date" "2023-02-30T00:00:00Z" "0000-01-02T03:04:05Z" "2023-06-15T10:14Z"]
                testCase.verifyEqual(canonDT(ST_DateTime.from_xml(badInput)), ...
                    "dt-aware|1970-01-01T00:00:00+00:00", ...
                    sprintf('invalid xsd:dateTime "%s" must epoch, never raise', badInput));
            end
        end

        function test_datetime_epoch_underaccept_deferred_to_P8_2(testCase)
            % ** DELIBERATE UNDER-ACCEPT PIN (VERIFY-fromiso). **
            % These five inputs are lexically INVALID xsd:dateTime (seconds omitted,
            % non-zero-padded fields, basic no-separator form, ISO week date).
            % CPython datetime.fromisoformat PARSES them; the port's narrow subset
            % does NOT and reaches the 1970 epoch. This test freezes the CURRENT
            % PORT epoch behaviour per
            %   validation\summary\decision_2026-07-26_st_datetime_underaccept.md
            % (defer-to-P8-2 ruling; ST_DateTime is unreachable until comments land).
            % A future P8-2 fix that widens the subset to full CPython parity WILL
            % flip these to the parsed values -- that is the EXPECTED, documented
            % flip this pin exists to surface for review, NOT a surprise regression.
            import mat2doc.oxml.simpletypes.*
            underAccept = [ ...
                "2023-06-15T10:14"    % D07  seconds omitted, naive
                "2023-1-2T3:4:5Z"     % D11  non-zero-padded fields (via lenient strptime Z)
                "2023-06-15T10"       % D22  hour-only
                "20230615T101430"     % D23  basic (no-separator) form
                "2023-W24-4" ];       % D24  ISO week date
            for i = 1:numel(underAccept)
                testCase.verifyEqual(canonDT(ST_DateTime.from_xml(underAccept(i))), ...
                    "dt-aware|1970-01-01T00:00:00+00:00", ...
                    sprintf(['P8-2 deferred under-accept: "%s" currently epochs; ' ...
                    'a widened fromisoformat subset will deliberately flip this'], ...
                    underAccept(i)));
            end
        end

        % =============================================================== %
        % 6. Dead abstract-base path (P01) + raise-type audit             %
        % =============================================================== %

        function test_basesimpletype_abstract_deadpath_P01(testCase)
            % BaseSimpleType is the ABSTRACT root: from_xml/to_xml are hosted on the
            % branch bases (BaseIntType/BaseStringType) per the H10 no-late-binding
            % idiom, NOT the root. Calling them on the root raises MATLAB's
            % classHasNoPropertyOrMethod. UNREACHABLE (every concrete class descends
            % via a branch base or declares its own). Pinned to the PORT (raise);
            % Python would return int(42). (s0017 P01.)
            import mat2doc.oxml.simpletypes.*
            [id, ~] = raiseAndCapture(@() BaseSimpleType.from_xml("42"));
            testCase.verifyEqual(id, 'MATLAB:subscripting:classHasNoPropertyOrMethod', ...
                'the abstract root has no from_xml (H10); every real converter uses a branch base');
        end

        function test_raise_type_audit_representative(testCase)
            % Pins the RAISE-TYPE at representative sites (validate_P3-2 raise-site
            % audit): InvalidXmlError ONLY at the two docx.exceptions convert_from_xml
            % sites; ValueError for membership/range/hex/int-lexis; TypeError for the
            % validate guards; KeyError for the unknown measure unit.
            import mat2doc.oxml.simpletypes.*
            expect = { ...
                @() XsdBoolean.from_xml("on"),            'mat2doc:InvalidXmlError'
                @() ST_OnOff.from_xml("maybe"),           'mat2doc:InvalidXmlError'
                @() ST_BrClear.to_xml("x"),               'mat2doc:ValueError'
                @() ST_Coordinate.from_xml("abc"),        'mat2doc:ValueError'
                @() XsdInt.to_xml(2147483648),            'mat2doc:ValueError'
                @() ST_HexColor.from_xml("GGGGGG"),       'mat2doc:ValueError'
                @() XsdBoolean.to_xml(2),                 'mat2doc:TypeError'
                @() ST_DateTime.to_xml("x"),              'mat2doc:TypeError'
                @() XsdString.to_xml(5),                  'mat2doc:TypeError'
                @() ST_UniversalMeasure.from_xml("5xy"),  'mat2doc:KeyError' };
            for k = 1:size(expect,1)
                [id, ~] = raiseAndCapture(expect{k,1});
                testCase.verifyEqual(id, expect{k,2}, ...
                    sprintf('raise-type audit site #%d must be %s', k, expect{k,2}));
            end
        end

        % =============================================================== %
        % 7. EQUIVALENCE -- full 175-fact battery vs the frozen oracle    %
        % =============================================================== %

        function test_equivalence_full_175_vs_frozen_oracle(testCase)
            % Replays the ENTIRE s0017 175-fact battery (the .m twin's probe
            % sequence, embedded below) against the frozen python-docx 1.2.0 oracle
            % copied into tests\oxml\data\s0017_probe_oracle.json. For each of the
            % 175 tags: the 154 MATCHING facts are verified EQUAL to the oracle; the
            % 21 ADOPTED DIVERGENCES are verified equal to the pinned PORT value AND
            % NOT equal to the oracle (proving the deviation is real, governed, and
            % frozen). This single method is the class's Equivalence leg.
            oracle = loadOracle();
            port   = runProbes();
            div    = divergentPortValues();          % 21 tag -> pinned port string
            tags   = fieldnames(port);
            testCase.verifyEqual(numel(tags), 175, 'the battery must have exactly 175 facts');
            nMatch = 0; nDiv = 0;
            for i = 1:numel(tags)
                t = tags{i};
                got = string(port.(t));
                if isfield(div, t)
                    nDiv = nDiv + 1;
                    testCase.verifyEqual(got, string(div.(t)), ...
                        sprintf('%s: pinned adopted-divergence port value', t));
                    testCase.verifyNotEqual(got, string(oracle.(t)), ...
                        sprintf('%s: adopted divergence must still differ from the oracle', t));
                else
                    nMatch = nMatch + 1;
                    testCase.verifyEqual(got, string(oracle.(t)), ...
                        sprintf('%s: must be byte-identical to the frozen oracle', t));
                end
            end
            testCase.verifyEqual(nMatch, 154, 'exactly 154 facts must match the oracle');
            testCase.verifyEqual(nDiv, 21, 'exactly 21 facts must be adopted divergences');
        end

    end
end

% ===================== file-local helpers ============================== %

function [id, msg] = raiseAndCapture(fn)
    % Run fn() expecting a raise; return (identifier, message) as char, or ('','')
    % if it did not raise. (Idiom copied from Test_p3_1_enum_base.m.)
    id = ''; msg = '';
    try
        fn();
    catch ME
        id = ME.identifier;
        msg = ME.message;
    end
end

function s = canonDT(v)
    % Canonicalise a datetime EXACTLY as the s0017 probe does (dt-aware|<UTC
    % instant>+00:00 or dt-naive|<wall>), so datetime pins compare as strings and
    % sidestep timezone-object equality subtleties. Copied verbatim from
    % s0017_simpletypes_probe.m::canon (datetime branch).
    aware = ~isempty(v.TimeZone);
    if aware
        v.TimeZone = "UTC";
    end
    [Y, Mo, D] = ymd(v);
    [H, Mi, Sec] = hms(v);
    Sw = floor(Sec);
    us = round((Sec - Sw) * 1e6);
    if us >= 1e6
        Sw = Sw + 1; us = 0;
    end
    base = string(sprintf("%04d-%02d-%02dT%02d:%02d:%02d", Y, Mo, D, H, Mi, Sw));
    if us > 0
        base = base + string(sprintf(".%06d", us));
    end
    if aware
        s = "dt-aware|" + base + "+00:00";
    else
        s = "dt-naive|" + base;
    end
end

function o = loadOracle()
    % Read the co-located frozen oracle in BINARY mode (no CRLF translation) and
    % decode UTF-8 -> struct. jsondecode is whitespace/line-ending agnostic, so no
    % `* binary` .gitattributes pin is needed (value-based fixture, s0016 precedent).
    here = fileparts(mfilename('fullpath'));
    p = fullfile(here, 'data', 's0017_probe_oracle.json');
    fid = fopen(p, 'r', 'n');
    if fid < 0
        error('mat2doc:test:openFail', 'cannot open frozen oracle %s', p);
    end
    raw = fread(fid, Inf, '*uint8')';
    fclose(fid);
    o = jsondecode(native2unicode(raw, 'UTF-8'));
end

function d = divergentPortValues()
    % The 21 adopted divergences (port != oracle), each pinned to the PORT's actual
    % canonical output with the governing D-number. Sourced from the Gate-3
    % candidate capture. I06 is built via char(1637) for source-encoding independence.
    d = struct();
    % --- D-005 (MATLAB class token where CPython prints <class '...'>) ---
    d.B12 = "err|TypeError|only True or False (and possibly None) may be assigned, got 'double'";
    d.E13 = "err|TypeError|value must be a string, got double";
    d.P06 = "err|TypeError|value must be a string, got double";
    d.P09 = "err|TypeError|value must be a string, got double";
    d.R04 = "err|TypeError|value must be <type 'int'>, got string";
    d.R08 = "err|TypeError|value must be <type 'int'>, got double";
    d.X04 = "err|ValueError|rgb color value must be RGBColor object, got string blue";
    % --- D-STYPE-3 (rounded 2^63/2^64 bound or value >2^53) ---
    d.H08 = "err|ValueError|value must be in range 0 to 18446744073709551616 inclusive, got -1";
    d.R10 = "err|ValueError|value must be in range 0 to 18446744073709551616 inclusive, got -1";
    d.R06 = "err|ValueError|value must be in range -9223372036854775808 to 9223372036854775808 inclusive, got -18446744073709551616";
    d.P27 = "OK|int|18446744073709551616";
    % --- D-STYPE-2 / D-004 (str2double vs CPython float()) ---
    d.U12 = "err|MATLAB:validators:mustBeFinite|Value must be finite.";
    d.U13 = "OK|int|12700000";
    d.S07 = "err|MATLAB:validators:mustBeFinite|Value must be finite.";
    % --- D-002 (ASCII-only int() grammar) ---
    d.I06 = "err|ValueError|invalid literal for int() with base 10: '" + string(char(1637)) + "'";
    % --- VERIFY-fromiso / P8-2 (narrow subset -> 1970 epoch on schema-INVALID input) ---
    d.D07 = "OK|dt-aware|1970-01-01T00:00:00+00:00";
    d.D11 = "OK|dt-aware|1970-01-01T00:00:00+00:00";
    d.D22 = "OK|dt-aware|1970-01-01T00:00:00+00:00";
    d.D23 = "OK|dt-aware|1970-01-01T00:00:00+00:00";
    d.D24 = "OK|dt-aware|1970-01-01T00:00:00+00:00";
    % --- dead-abstract-base (H10) ---
    d.P01 = "err|MATLAB:subscripting:classHasNoPropertyOrMethod|The class mat2doc.oxml.simpletypes.BaseSimpleType has no Constant property or Static method named 'from_xml'.";
end

function p = runProbes()
    % Replay the s0017 probe sequence (the .m twin's body, verbatim tags/inputs/
    % order) and return a struct of tagged canonical strings. Embedded here so the
    % Equivalence leg is self-contained (the validation-folder scenario is NOT on
    % the toolbox path and must not be a dependency).
    import mat2doc.oxml.simpletypes.*
    import mat2doc.shared.*
    p = struct();

    % B  XsdBoolean
    p.B01 = P(@() XsdBoolean.from_xml("1"));
    p.B02 = P(@() XsdBoolean.from_xml("0"));
    p.B03 = P(@() XsdBoolean.from_xml("true"));
    p.B04 = P(@() XsdBoolean.from_xml("false"));
    p.B05 = P(@() XsdBoolean.from_xml("on"));
    p.B06 = P(@() XsdBoolean.from_xml("TRUE"));
    p.B07 = P(@() XsdBoolean.to_xml(true));
    p.B08 = P(@() XsdBoolean.to_xml(false));
    p.B09 = P(@() XsdBoolean.to_xml(1));
    p.B10 = P(@() XsdBoolean.to_xml(2));
    p.B11 = P(@() XsdBoolean.to_xml("x"));
    p.B12 = P(@() XsdBoolean.to_xml([]));

    % O  ST_OnOff
    toks = ["1", "0", "true", "false", "on", "off"];
    for i = 1:6
        p.(sprintf("O%02d", i)) = P(@() ST_OnOff.from_xml(toks(i)));
    end
    p.O07 = P(@() ST_OnOff.from_xml("maybe"));
    p.O08 = P(@() ST_OnOff.to_xml(true));
    p.O09 = P(@() ST_OnOff.from_xml("On"));
    p.O10 = P(@() ST_OnOff.from_xml(""));

    % E  string enums
    p.E01 = P(@() ST_BrClear.to_xml("left"));
    p.E02 = P(@() ST_BrClear.to_xml("x"));
    p.E03 = P(@() ST_BrType.to_xml("textWrapping"));
    p.E04 = P(@() ST_BrType.to_xml("x"));
    p.E05 = P(@() ST_TblLayoutType.to_xml("autofit"));
    p.E06 = P(@() ST_TblLayoutType.to_xml("x"));
    p.E07 = P(@() ST_TblWidth.to_xml("dxa"));
    p.E08 = P(@() ST_TblWidth.to_xml("x"));
    p.E09 = P(@() ST_Merge.to_xml("restart"));
    p.E10 = P(@() ST_Merge.to_xml("x"));
    p.E11 = P(@() ST_VerticalAlignRun.to_xml("subscript"));
    p.E12 = P(@() ST_VerticalAlignRun.to_xml("x"));
    p.E13 = P(@() ST_BrClear.to_xml(5));
    p.E14 = P(@() ST_Merge.from_xml("restart"));
    p.E15 = P(@() ST_BrClear.to_xml("none"));
    p.E16 = P(@() ST_BrType.to_xml("page"));
    p.E17 = P(@() ST_TblWidth.to_xml("nil"));
    p.E18 = P(@() ST_VerticalAlignRun.to_xml("baseline"));
    p.E19 = P(@() ST_Merge.to_xml("continue"));
    p.E20 = P(@() ST_TblLayoutType.to_xml("fixed"));

    % U  ST_UniversalMeasure
    um = ["1in", "2.5cm", "12pt", "10mm", "1pc", "1pi", "-1.5pt", ".5in", ...
          "1e2pt", "0.5pt"];
    for i = 1:numel(um)
        p.(sprintf("U%02d", i)) = P(@() ST_UniversalMeasure.from_xml(um(i)));
    end
    p.U11 = P(@() ST_UniversalMeasure.from_xml("5xy"));
    p.U12 = P(@() ST_UniversalMeasure.from_xml("1_0pt"));
    p.U13 = P(@() ST_UniversalMeasure.from_xml("1,000pt"));
    p.U14 = P(@() ST_UniversalMeasure.from_xml("1in"));

    % C  coordinates
    p.C01 = P(@() ST_Coordinate.from_xml("914400"));
    p.C02 = P(@() ST_Coordinate.from_xml("-27273042329600"));
    p.C03 = P(@() ST_Coordinate.from_xml("1in"));
    p.C04 = P(@() ST_Coordinate.from_xml("2.5cm"));
    p.C05 = P(@() ST_Coordinate.to_xml(Emu(914400)));
    p.C06 = P(@() ST_Coordinate.to_xml(Emu(27273042316901)));
    p.C07 = P(@() ST_Coordinate.from_xml("abc"));
    p.C08 = P(@() ST_Coordinate.from_xml("2.5"));
    p.C09 = P(@() ST_PositiveCoordinate.from_xml("914400"));
    p.C10 = P(@() ST_PositiveCoordinate.to_xml(Emu(-1)));
    p.C11 = P(@() ST_CoordinateUnqualified.to_xml(-27273042329600));
    p.C12 = P(@() ST_CoordinateUnqualified.to_xml(-27273042329599));

    % H  ST_HpsMeasure
    p.H01 = P(@() ST_HpsMeasure.from_xml("24"));
    p.H02 = P(@() ST_HpsMeasure.from_xml("12pt"));
    p.H03 = P(@() ST_HpsMeasure.from_xml("25"));
    p.H04 = P(@() ST_HpsMeasure.to_xml(Pt(12)));
    p.H05 = P(@() ST_HpsMeasure.to_xml(Pt(12.25)));
    p.H06 = P(@() ST_HpsMeasure.to_xml(Emu(152400)));
    p.H07 = P(@() ST_HpsMeasure.from_xml("1in"));
    p.H08 = P(@() ST_HpsMeasure.to_xml(-1));

    % T/S  twips
    p.T01 = P(@() ST_TwipsMeasure.from_xml("1440"));
    p.T02 = P(@() ST_TwipsMeasure.from_xml("1in"));
    p.T03 = P(@() ST_TwipsMeasure.to_xml(Twips(1440)));
    p.T04 = P(@() ST_TwipsMeasure.to_xml(Emu(635)));
    p.T05 = P(@() ST_TwipsMeasure.to_xml(Emu(300)));
    p.S01 = P(@() ST_SignedTwipsMeasure.from_xml("-360"));
    p.S02 = P(@() ST_SignedTwipsMeasure.from_xml("12.5"));
    p.S03 = P(@() ST_SignedTwipsMeasure.from_xml("13.5"));
    p.S04 = P(@() ST_SignedTwipsMeasure.from_xml("-12.5"));
    p.S05 = P(@() ST_SignedTwipsMeasure.from_xml("1in"));
    p.S06 = P(@() ST_SignedTwipsMeasure.to_xml(Twips(-360)));
    p.S07 = P(@() ST_SignedTwipsMeasure.from_xml("abc"));

    % X/A  hex color
    p.X01 = P(@() ST_HexColor.to_xml(ST_HexColor.from_xml("3C2F80")));
    p.X02 = P(@() ST_HexColor.from_xml("auto"));
    p.X03 = P(@() ST_HexColor.to_xml(ST_HexColor.from_xml("3c2f80")));
    p.X04 = P(@() ST_HexColor.to_xml("blue"));
    p.X05 = P(@() ST_HexColor.from_xml("GGGGGG"));
    p.X06 = P(@() ST_HexColor.from_xml("Auto"));
    p.X07 = P(@() ST_HexColor.from_xml("FF0000"));
    p.A01 = P(@() ST_HexColorAuto.to_xml("auto"));
    p.A02 = P(@() ST_HexColorAuto.to_xml("x"));
    p.A03 = P(@() ST_HexColorAuto.from_xml("auto"));

    % D  ST_DateTime
    p.D01 = P(@() ST_DateTime.from_xml("2023-01-02T03:04:05Z"));
    p.D02 = P(@() ST_DateTime.from_xml("2023-01-02T03:04:05.123456Z"));
    p.D03 = P(@() ST_DateTime.to_xml(ST_DateTime.from_xml("2023-01-02T03:04:05.999999Z")));
    p.D04 = P(@() ST_DateTime.from_xml("2023-01-02T03:04:05-08:00"));
    p.D05 = P(@() ST_DateTime.from_xml("2023-01-02 03:04:05"));
    p.D06 = P(@() ST_DateTime.from_xml("2023-06-15"));
    p.D07 = P(@() ST_DateTime.from_xml("2023-06-15T10:14"));
    p.D08 = P(@() ST_DateTime.from_xml("not-a-date"));
    p.D09 = P(@() ST_DateTime.from_xml("2023-02-30T00:00:00Z"));
    p.D10 = P(@() ST_DateTime.from_xml("0000-01-02T03:04:05Z"));
    p.D11 = P(@() ST_DateTime.from_xml("2023-1-2T3:4:5Z"));
    p.D12 = P(@() ST_DateTime.to_xml(ST_DateTime.from_xml("2023-01-02T03:04:05Z")));
    p.D13 = P(@() ST_DateTime.to_xml(datetime(2023, 6, 15, 12, 0, 0, "TimeZone", "UTC")));
    p.D14 = P(@() ST_DateTime.to_xml("x"));
    p.D15 = P(@() ST_DateTime.from_xml("2023-01-02T03:04:05.5"));
    p.D16 = P(@() ST_DateTime.from_xml("2023-01-02T03:04:05+05:30"));
    p.D17 = P(@() ST_DateTime.from_xml("2023-01-02 03:04:05.123"));
    p.D18 = P(@() ST_DateTime.to_xml(datetime(2023, 1, 2, 3, 4, 5.9, "TimeZone", "UTC")));
    p.D19 = P(@() ST_DateTime.to_xml(ST_DateTime.from_xml("not-a-date")));
    p.D20 = P(@() ST_DateTime.from_xml("2023-12-31T23:59:59Z"));
    p.D21 = P(@() ST_DateTime.from_xml("2023-06-15T10:14Z"));
    p.D22 = P(@() ST_DateTime.from_xml("2023-06-15T10"));
    p.D23 = P(@() ST_DateTime.from_xml("20230615T101430"));
    p.D24 = P(@() ST_DateTime.from_xml("2023-W24-4"));

    % R  Xsd ranges
    p.R01 = P(@() XsdInt.to_xml(2147483648));
    p.R02 = P(@() XsdInt.to_xml(-2147483648));
    p.R03 = P(@() XsdUnsignedInt.to_xml(-1));
    p.R04 = P(@() XsdInt.to_xml("5"));
    p.R05 = P(@() XsdLong.to_xml(9007199254740992));
    p.R06 = P(@() XsdLong.to_xml(-18446744073709551616));
    p.R07 = P(@() XsdUnsignedLong.to_xml(0));
    p.R08 = P(@() XsdInt.to_xml(2.5));
    p.R09 = P(@() XsdUnsignedInt.to_xml(4294967295));
    p.R10 = P(@() XsdUnsignedLong.to_xml(-1));

    % I  int() grammar
    p.I01 = P(@() XsdInt.from_xml("2.5"));
    p.I02 = P(@() XsdInt.from_xml("+5"));
    p.I03 = P(@() XsdInt.from_xml("007"));
    p.I04 = P(@() XsdInt.from_xml(" 42 "));
    p.I05 = P(@() XsdInt.from_xml("1_000"));
    p.I06 = P(@() XsdInt.from_xml(string(char(1637))));   % Arabic-Indic 5
    p.I07 = P(@() XsdInt.from_xml("-5"));

    % PR pyRound (banker's)
    prv = ["0.5", "1.5", "2.5", "-0.5", "-1.5", "-2.5", "2.675", "12.5", "13.5"];
    for i = 1:numel(prv)
        p.(sprintf("PR%02d", i)) = P(@() ST_SignedTwipsMeasure.from_xml(prv(i)));
    end

    % P  pass-through tier
    p.P01 = P(@() mat2doc.oxml.simpletypes.BaseSimpleType.from_xml("42"));
    p.P02 = P(@() mat2doc.oxml.simpletypes.BaseIntType.from_xml("42"));
    p.P03 = P(@() mat2doc.oxml.simpletypes.BaseIntType.to_xml(42));
    p.P04 = P(@() mat2doc.oxml.simpletypes.BaseStringType.from_xml("hello"));
    p.P05 = P(@() mat2doc.oxml.simpletypes.BaseStringType.to_xml("hello"));
    p.P06 = P(@() mat2doc.oxml.simpletypes.BaseStringType.to_xml(5));
    p.P07 = P(@() XsdString.from_xml("s"));
    p.P08 = P(@() XsdString.to_xml("s"));
    p.P09 = P(@() XsdString.to_xml(5));
    p.P10 = P(@() XsdToken.from_xml("  a  b  "));
    p.P11 = P(@() XsdToken.to_xml("tok"));
    p.P12 = P(@() XsdId.from_xml("_id1"));
    p.P13 = P(@() XsdId.to_xml("_id1"));
    p.P14 = P(@() XsdAnyUri.from_xml("http://x/y"));
    p.P15 = P(@() XsdAnyUri.to_xml("http://x/y"));
    p.P16 = P(@() ST_String.from_xml("free text"));
    p.P17 = P(@() ST_String.to_xml("free text"));
    p.P18 = P(@() ST_RelationshipId.from_xml("rId7"));
    p.P19 = P(@() ST_RelationshipId.to_xml("rId7"));
    p.P20 = P(@() ST_DecimalNumber.from_xml("-3"));
    p.P21 = P(@() ST_DecimalNumber.to_xml(-3));
    p.P22 = P(@() ST_DecimalNumber.to_xml(2147483648));
    p.P23 = P(@() ST_DrawingElementId.from_xml("42"));
    p.P24 = P(@() ST_DrawingElementId.to_xml(42));
    p.P25 = P(@() ST_DrawingElementId.to_xml(-1));
    p.P26 = P(@() XsdStringEnumeration.from_xml("x"));
    p.P27 = P(@() XsdUnsignedLong.from_xml("18446744073709551615"));
end

function s = P(fn)
    % Run fn(); return "OK|<canon>" or "err|<ExcName>|<message>". Verbatim from
    % s0017_simpletypes_probe.m::P.
    try
        v = fn();
        s = "OK|" + canon(v);
    catch e
        id = string(e.identifier);
        name = id;
        if startsWith(id, "mat2doc:")
            name = extractAfter(id, "mat2doc:");
        end
        s = "err|" + name + "|" + string(e.message);
    end
end

function s = canon(v)
    % Verbatim from s0017_simpletypes_probe.m::canon.
    if isa(v, "datetime")
        s = canonDT(v);
    elseif islogical(v)
        if v, s = "bool|True"; else, s = "bool|False"; end
    elseif isa(v, "mat2doc.shared.RGBColor")
        s = "rgb|" + v.str_();
    elseif isnumeric(v)
        dv = double(v);
        if mod(dv, 1) == 0
            s = "int|" + string(sprintf("%.0f", dv));
        else
            s = "float|" + string(num2str(dv, 17));
        end
    elseif isstring(v) || ischar(v)
        s = "str|" + string(v);
    else
        s = "obj|" + string(class(v));
    end
end

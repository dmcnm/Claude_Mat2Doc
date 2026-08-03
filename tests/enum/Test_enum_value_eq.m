classdef Test_enum_value_eq < matlab.unittest.TestCase
% TEST_ENUM_VALUE_EQ  Gate-4 permanent unit tests for the Mat2Doc
%   enum value-based ``==`` / ``~=`` semantics (the A2 section 2 resolution).
%
%   A new shared root ``mat2doc.enum.base.BaseIntEnum`` gives every int-enum
%   class (all 14 BaseXmlEnum + 3 BaseEnum subclasses) value-based equality,
%   restoring the python-docx int-subclass semantics (``BaseEnum(int, enum.Enum)``
%   / ``BaseXmlEnum(int, enum.Enum)``, base.py 15/33 ``int.__new__(cls, value)``).
%   ``BaseEnum``/``BaseXmlEnum`` now derive from ``BaseIntEnum`` and inherit its
%   ``eq``/``ne``; the two PLAIN ``enum.Enum`` ports (``WD_BREAK_TYPE``,
%   ``WD_INLINE_SHAPE_TYPE``) do NOT derive from it and keep MATLAB identity
%   ``==`` -- matching Python plain-enum identity semantics.
%
%   Provenance
%   ----------
%     * Gate-1 porter + Gate-2 auditor (Fable APPROVE, 0 divergent operand cells
%       vs a python-docx v1.2.0 REPL oracle):
%       validation\mat2doc\audit_enum_value_eq.md
%     * Changed code (worktree Mat2Doc-enumeq): NEW
%       +mat2doc\+enum\+base\BaseIntEnum.m (eq/ne + private valVec_);
%       one-line superclass edits to BaseEnum.m / BaseXmlEnum.m.
%     * Reference: python-docx v1.2.0 src/docx/enum/base.py.
%
%   This class LOCKS every row of the Gate-1/2 operand matrix plus the two
%   audit-flagged behaviors that Gate-4 was told to pin empirically rather than
%   assume (switch routing through eq; the residual plain-enum ==string delta).
%   All ``switch`` / plain-enum-==string assertions below were OBSERVED on
%   R2024b (scratchpad\probe_g4.m), not assumed -- so a future change to the
%   dispatch surface goes RED here.
%
%   Member values used (from the concrete classes, for reference):
%     WD_PARAGRAPH_ALIGNMENT (BaseXmlEnum): LEFT=0 CENTER=1 RIGHT=2 JUSTIFY=3
%     WD_TABLE_ALIGNMENT     (BaseXmlEnum): LEFT=0 CENTER=1 RIGHT=2
%     WD_TABLE_DIRECTION     (BaseEnum):    LTR=0 RTL=1
%     MSO_COLOR_TYPE         (BaseEnum):    RGB=1 THEME=2 AUTO=101
%     WD_BREAK_TYPE          (plain enum):  LINE=6 PAGE=7   (excluded from value-eq)
%
%   Coverage taxonomy
%   -----------------
%   * Nominal     -- same-class ==/~= true/false; the headline cross-class fix.
%   * Edge        -- cross-base; enum vs numeric/logical; enum vs string/char
%     (false) with the NAME idiom still working; enum vs [] (None) / missing;
%     empty-array broadcasting; the plain-enum exclusion + residual delta.
%   * Equivalence -- every asserted truth value matches the frozen python-docx
%     oracle cell in audit_enum_value_eq.md section 4 (cited per method).
%   * Regression  -- hard-coded expected logicals + element-wise array vectors;
%     the switch-routing behavior and the ==string residual pinned exactly.
%
%   Deviations exercised: 0 new D-numbers (byte-neutral; comparison operators
%   only, no serialization path touched).
%
%   Determinism: no network, no absolute paths, no file writes (except the
%   headline case which builds an IN-MEMORY Document -- nothing saved to disk).
%   The +mat2doc package resolves via the MANDATORY PathFixture(worktree-root)
%   in TestClassSetup (WP9-F4 lesson). Enum members are FULLY-QUALIFIED inline
%   (never `import mat2doc.enum...`) -- a bare import errors at suite-creation
%   parse time, before the PathFixture runs.

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied from the proven tests\enum\Test_p3_1_enum_base.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\enum
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. Same-class == / ~=  (audit section 4: same-class true/false)  %
        % =============================================================== %

        function test_same_class_eq(testCase)
            % Equal members -> true; unequal -> false (WD_PARAGRAPH_ALIGNMENT).
            testCase.verifyTrue( ...
                mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER == ...
                mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER, ...
                'same-class equal members must compare ==true');
            testCase.verifyFalse( ...
                mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER == ...
                mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.LEFT, ...
                'same-class distinct members must compare ==false');
        end

        function test_same_class_ne(testCase)
            % ~= is the inverse of == on the same class.
            testCase.verifyFalse( ...
                mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER ~= ...
                mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER, ...
                'equal members must compare ~=false');
            testCase.verifyTrue( ...
                mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER ~= ...
                mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.LEFT, ...
                'distinct members must compare ~=true');
        end

        % =============================================================== %
        % 2. Cross-class equal-by-value  --  THE HEADLINE FIX             %
        %    (audit section 4: PA.CENTER==TA.CENTER True; PA.JUSTIFY==TA.CENTER False) %
        % =============================================================== %

        function test_cross_class_equal_by_value_true(testCase)
            % THE FIX: two DIFFERENT int-enum classes sharing a value (1) compare
            % ==true -- python-docx int==int. Both directions (operand order).
            testCase.verifyTrue( ...
                mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER == ...
                mat2doc.enum.table.WD_TABLE_ALIGNMENT.CENTER, ...
                'cross-class equal-by-value (both value 1) must be true (THE HEADLINE FIX)');
            testCase.verifyTrue( ...
                mat2doc.enum.table.WD_TABLE_ALIGNMENT.CENTER == ...
                mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER, ...
                'cross-class == must be symmetric (RHS operand)');
        end

        function test_cross_class_non_shared_value_false(testCase)
            % A non-shared member (PA.JUSTIFY value 3, which WD_TABLE_ALIGNMENT
            % cannot represent) vs TA.CENTER value 1 -> false.
            testCase.verifyFalse( ...
                mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.JUSTIFY == ...
                mat2doc.enum.table.WD_TABLE_ALIGNMENT.CENTER, ...
                'cross-class members of different value (3 vs 1) must be false');
        end

        % =============================================================== %
        % 3. Cross-BASE (BaseXmlEnum subclass vs BaseEnum subclass)       %
        %    (audit section 4: TA.LEFT==TD.LTR True; MCT.RGB==... etc)     %
        % =============================================================== %

        function test_cross_base_equal_by_value_true(testCase)
            % A BaseXmlEnum member vs a BaseEnum member sharing a value both
            % dispatch to the single inherited BaseIntEnum.eq -> value compare.
            % WD_TABLE_ALIGNMENT.LEFT(0) [XmlEnum] == WD_TABLE_DIRECTION.LTR(0) [Enum].
            testCase.verifyTrue( ...
                mat2doc.enum.table.WD_TABLE_ALIGNMENT.LEFT == ...
                mat2doc.enum.table.WD_TABLE_DIRECTION.LTR, ...
                'cross-BASE (XmlEnum vs Enum) sharing value 0 must be true');
            % And an unshared cross-base pair is false: PA.CENTER(1) vs
            % WD_TABLE_DIRECTION.LTR(0).
            testCase.verifyFalse( ...
                mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER == ...
                mat2doc.enum.table.WD_TABLE_DIRECTION.LTR, ...
                'cross-BASE distinct value (1 vs 0) must be false');
        end

        function test_cross_base_baseenum_pair_true(testCase)
            % Two BaseEnum-derived classes sharing a value: MSO_COLOR_TYPE.RGB(1)
            % == WD_TABLE_DIRECTION.RTL(1) -> true (both derive from BaseIntEnum).
            testCase.verifyTrue( ...
                mat2doc.enum.dml.MSO_COLOR_TYPE.RGB == ...
                mat2doc.enum.table.WD_TABLE_DIRECTION.RTL, ...
                'BaseEnum vs BaseEnum sharing value 1 must be true');
        end

        % =============================================================== %
        % 4. enum vs numeric / logical  (Python bool is an int subclass)  %
        %    (audit section 4: ==1/==int32(1)/==2.0; ==true/==false)       %
        % =============================================================== %

        function test_enum_vs_numeric(testCase)
            % CENTER has value 1: ==1 true, ==int32(1) true, ==2.0 false, ==1.0 true.
            C = mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER;
            testCase.verifyTrue(C == 1,           'CENTER == 1 (double) must be true');
            testCase.verifyTrue(C == int32(1),    'CENTER == int32(1) must be true');
            testCase.verifyTrue(C == 1.0,         'CENTER == 1.0 must be true');
            testCase.verifyFalse(C == 2.0,        'CENTER == 2.0 must be false');
            testCase.verifyTrue(1 == C,           'RHS: 1 == CENTER must be true (operand order)');
        end

        function test_enum_vs_logical(testCase)
            % Python bool is an int subclass: member==true matches value 1,
            % member==false matches value 0. CENTER(1), LEFT(0).
            C = mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER;   % value 1
            L = mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.LEFT;     % value 0
            testCase.verifyTrue(C == true,   'CENTER(1) == true must be true (bool-is-int)');
            testCase.verifyFalse(C == false, 'CENTER(1) == false must be false');
            testCase.verifyTrue(L == false,  'LEFT(0) == false must be true');
            testCase.verifyFalse(L == true,  'LEFT(0) == true must be false');
        end

        % =============================================================== %
        % 5. enum vs string / char -> false; NAME idiom still works        %
        %    (audit section 4: =="CENTER" F, =='CENTER' F; name idiom True) %
        % =============================================================== %

        function test_enum_vs_string_is_false(testCase)
            % python-docx int != str -> False. THE FLIP from old Mat2Doc behavior
            % (member=="NAME" used to be true; now false).
            C = mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER;
            testCase.verifyFalse(C == "CENTER", ...
                'member == "CENTER" (string) must be false (int != str; A2 flip)');
            testCase.verifyFalse(C == 'CENTER', ...
                'member == ''CENTER'' (char) must be false');
            testCase.verifyTrue(C ~= "CENTER", ...
                'member ~= "CENTER" must be true (inverse)');
        end

        function test_name_idiom_still_works(testCase)
            % The correct NAME idiom under the new semantics: string(member).
            C = mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER;
            testCase.verifyTrue(string(C) == "CENTER", ...
                'string(member) == "CENTER" must remain true (name access idiom)');
            testCase.verifyEqual(string(C), "CENTER", ...
                'string(member) must be the member name');
        end

        % =============================================================== %
        % 6. enum vs [] (None) / missing -> false ; ~=[] -> true           %
        %    (audit section 4: ==[] scalar false; ~=[] true; ==missing F)  %
        % =============================================================== %

        function test_enum_vs_none_and_missing(testCase)
            % [] is the None sentinel -> scalar false (Python member==None -> False).
            C = mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER;
            r = (C == []);
            testCase.verifyFalse(r, 'member == [] (None) must be false');
            testCase.verifyEqual(size(r), [1 1], ...
                'member == [] must be a SCALAR false (None sentinel coerced to scalar)');
            testCase.verifyTrue(C ~= [], 'member ~= [] (None) must be true');
            testCase.verifyFalse(C == missing, 'member == missing must be false');
            testCase.verifyFalse(C == string(missing), 'member == string(missing) must be false');
        end

        % =============================================================== %
        % 7. ~= is the exact inverse of ==  (representative sample)        %
        % =============================================================== %

        function test_ne_is_exact_inverse(testCase)
            % For a representative spread of operand kinds, (a~=b) == ~(a==b).
            C  = mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER;
            samples = { ...
                mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER, ...   % same-class equal
                mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.LEFT, ...     % same-class unequal
                mat2doc.enum.table.WD_TABLE_ALIGNMENT.CENTER, ...      % cross-class equal
                1, int32(1), 2.0, true, false, ...                     % numeric/logical
                "CENTER", 'CENTER', ...                                % string/char
                missing };                                             % missing
            for k = 1:numel(samples)
                b = samples{k};
                testCase.verifyEqual(C ~= b, ~(C == b), ...
                    sprintf('~= must be the exact inverse of == for sample %d', k));
            end
        end

        % =============================================================== %
        % 8. Element-wise arrays (array==scalar; array==array)             %
        %    (audit section 4: [CENTER LEFT]==CENTER->[1 0]; ==[1 0]->[1 1]) %
        % =============================================================== %

        function test_array_eq_scalar(testCase)
            % array==scalar broadcasts element-wise.
            arr = [mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER, ...
                   mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.LEFT];
            testCase.verifyEqual( ...
                arr == mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER, ...
                [true false], ...
                '[CENTER LEFT] == CENTER must be [1 0] element-wise');
            % array == int broadcasts too (CENTER=1, LEFT=0).
            testCase.verifyEqual(arr == 1, [true false], ...
                '[CENTER LEFT] == 1 must be [1 0]');
        end

        function test_array_eq_array(testCase)
            % array==array compares element-wise (values 1 and 0).
            arr = [mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER, ...
                   mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.LEFT];
            testCase.verifyEqual(arr == [1 0], [true true], ...
                '[CENTER LEFT] == [1 0] must be [1 1]');
            testCase.verifyEqual(arr == [0 0], [false true], ...
                '[CENTER LEFT] == [0 0] must be [0 1]');
        end

        % =============================================================== %
        % 9. HEADLINE end-to-end through a real Table                     %
        %    (audit section 5: tbl.alignment == WD_TABLE_ALIGNMENT.CENTER) %
        % =============================================================== %

        function test_headline_table_alignment_eq(testCase)
            % Build a real in-memory Document, add a 2x3 table, set alignment via
            % WD_TABLE_ALIGNMENT.CENTER, read it back (faithfully returns
            % WD_PARAGRAPH_ALIGNMENT.CENTER per the A2 note), and assert
            % a == WD_TABLE_ALIGNMENT.CENTER -> TRUE. This is the whole point of
            % the WP. (Observed on R2024b: class=WD_PARAGRAPH_ALIGNMENT.)
            t = mat2doc.Document().add_table(2, 3);
            t.alignment = mat2doc.enum.table.WD_TABLE_ALIGNMENT.CENTER;
            a = t.alignment;
            testCase.verifyEqual(class(a), ...
                'mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT', ...
                'getter faithfully returns the paragraph-alignment member (A2)');
            testCase.verifyTrue(a == mat2doc.enum.table.WD_TABLE_ALIGNMENT.CENTER, ...
                'HEADLINE: t.alignment == WD_TABLE_ALIGNMENT.CENTER must be true (value-eq)');
            testCase.verifyFalse(a == mat2doc.enum.table.WD_TABLE_ALIGNMENT.LEFT, ...
                't.alignment == WD_TABLE_ALIGNMENT.LEFT must be false');
            testCase.verifyFalse(a ~= mat2doc.enum.table.WD_TABLE_ALIGNMENT.CENTER, ...
                't.alignment ~= WD_TABLE_ALIGNMENT.CENTER must be false (inverse)');
        end

        % =============================================================== %
        % 10. switch behavior over an int-enum member (audit finding #4)   %
        %     PINNED FROM OBSERVED R2024b BEHAVIOR (probe_g4.m), NOT assumed %
        % =============================================================== %

        function test_switch_routes_through_eq(testCase)
            % Because eq is overridden, `switch member` routes case-tests through
            % the new value-eq: a `case OtherEnumClass.MEMBER` matches CROSS-CLASS
            % by value. Corrects the Gate-1 record claim that switch is
            % identity-based (Gate-2 correction). Locked so a regression is caught.
            m = mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER;
            r = "nomatch";
            switch m
                case mat2doc.enum.table.WD_TABLE_ALIGNMENT.CENTER
                    r = "cross-class-CENTER";
                case mat2doc.enum.table.WD_TABLE_ALIGNMENT.LEFT
                    r = "cross-class-LEFT";
            end
            testCase.verifyEqual(r, "cross-class-CENTER", ...
                'switch over int-enum matches a cross-class value case (routes through eq)');
        end

        function test_switch_string_case_no_longer_matches(testCase)
            % A `case "NAME"` no longer matches an int-enum member (member=="NAME"
            % is now false), so string cases on int-enum members are DEAD.
            m = mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER;
            r = "fell-through";
            switch m
                case "CENTER"
                    r = "string-CENTER";
            end
            testCase.verifyEqual(r, "fell-through", ...
                'switch string case must NOT match an int-enum member (==NAME is false)');
        end

        % =============================================================== %
        % 11. Plain-enum exclusion + the known residual ==string delta     %
        %     (audit section 4 / Gate-2 RESIDUAL note)                     %
        % =============================================================== %

        function test_plain_enum_identity_retained(testCase)
            % WD_BREAK_TYPE is a plain enum.Enum port (NOT an int subclass, does
            % NOT derive from BaseIntEnum): identity == retained. Same member true,
            % different member false. (LINE=6, PAGE=7.)
            testCase.verifyTrue( ...
                mat2doc.enum.text.WD_BREAK_TYPE.LINE == ...
                mat2doc.enum.text.WD_BREAK_TYPE.LINE, ...
                'plain WD_BREAK_TYPE.LINE == LINE must be true (identity retained)');
            testCase.verifyFalse( ...
                mat2doc.enum.text.WD_BREAK_TYPE.LINE == ...
                mat2doc.enum.text.WD_BREAK_TYPE.PAGE, ...
                'plain WD_BREAK_TYPE.LINE == PAGE must be false');
        end

        function test_plain_enum_not_a_baseintenum(testCase)
            % Structural guard: the plain enums must NOT derive from BaseIntEnum
            % (that is what keeps them on identity ==). If a future edit rebased
            % them, value-eq would wrongly leak in -- caught here.
            testCase.verifyFalse( ...
                isa(mat2doc.enum.text.WD_BREAK_TYPE.LINE, ...
                    'mat2doc.enum.base.BaseIntEnum'), ...
                'WD_BREAK_TYPE must NOT be a BaseIntEnum (plain-enum exclusion)');
            testCase.verifyFalse( ...
                isa(mat2doc.enum.shape.WD_INLINE_SHAPE_TYPE.PICTURE, ...
                    'mat2doc.enum.base.BaseIntEnum'), ...
                'WD_INLINE_SHAPE_TYPE must NOT be a BaseIntEnum (plain-enum exclusion)');
        end

        function test_plain_enum_string_residual_divergence(testCase)
            % KNOWN RESIDUAL (pre-existing, NOT introduced by this WP): a plain
            % enum compared to its NAME string returns true in MATLAB (built-in
            % enumeration name-compare), whereas python-docx plain enum == "LINE"
            % is False. Pinned here as the CURRENT observed behavior so a future
            % change is caught; flagged for the design.md section 2 rewrite as the
            % only remaining enum-== divergence from Python. (Observed R2024b.)
            testCase.verifyTrue( ...
                mat2doc.enum.text.WD_BREAK_TYPE.LINE == "LINE", ...
                ['RESIDUAL DIVERGENCE (flagged for design.md): plain WD_BREAK_TYPE.LINE ' ...
                 '== "LINE" is TRUE in MATLAB (name-compare) vs False in python-docx. ' ...
                 'Pinned as current behavior -- NOT introduced by this WP.']);
        end

        % =============================================================== %
        % 12. Structural: int-enums ARE BaseIntEnum; isequal stays strict  %
        % =============================================================== %

        function test_int_enums_are_baseintenum(testCase)
            % Positive control for the exclusion guard above: the int-enum classes
            % DO derive from BaseIntEnum (both bases).
            testCase.verifyTrue( ...
                isa(mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER, ...
                    'mat2doc.enum.base.BaseIntEnum'), ...
                'WD_PARAGRAPH_ALIGNMENT (BaseXmlEnum) must be a BaseIntEnum');
            testCase.verifyTrue( ...
                isa(mat2doc.enum.table.WD_TABLE_DIRECTION.LTR, ...
                    'mat2doc.enum.base.BaseIntEnum'), ...
                'WD_TABLE_DIRECTION (BaseEnum) must be a BaseIntEnum');
        end

        function test_isequal_stays_class_strict(testCase)
            % isequal (hence verifyEqual) is NOT loosened by the eq override --
            % it compares class+properties, so cross-class members that are
            % ==-equal are still isequal-DISTINCT. This protects every other
            % test's verifyEqual from becoming value-loose.
            testCase.verifyFalse( ...
                isequal(mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER, ...
                        mat2doc.enum.table.WD_TABLE_ALIGNMENT.CENTER), ...
                'isequal must stay class-strict (cross-class value-equal members are NOT isequal)');
            testCase.verifyFalse( ...
                isequal(mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER, 1), ...
                'isequal(member, 1) must be false (class-strict)');
        end

    end
end

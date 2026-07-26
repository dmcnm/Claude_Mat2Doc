classdef Test_p2_2_xpath_hoist < matlab.unittest.TestCase
% TEST_P2_2_XPATH_HOIST  Gate-4 permanent unit tests for standing task #60
%   (hoist `xpath` from BaseOxmlElement onto XmlElement), folded into Mat2Doc
%   P2-2. THIS CLASS IS THE PERMANENT #60 CLOSURE.
%
%   Task #60 arose at P1-6b (validate_P1-6b Finding 1): `Part.drop_rel` on an
%   XmlPart, and `StoryPart.next_id`, both need `element.xpath("//@r:id")` /
%   `element.xpath("//@id")` on a PLAIN parsed root -- but at P1-3 the `xpath`
%   method lived only on `mat2doc.oxml.BaseOxmlElement`, so an UNREGISTERED root
%   (a parsed `w:document` before CT_Document registration at P2-3, whose runtime
%   class is the fallback `mat2doc.oxml.XmlElement`) had no `.xpath`. P2-2 HOISTS
%   the method to `mat2doc.oxml.XmlElement` (the lxml `_Element` analogue -- every
%   lxml element has `.xpath`), a pure relocation: BaseOxmlElement INHERITS it
%   verbatim. python-docx's `BaseOxmlElement.xpath` override (xmlchemy.py 687-692)
%   only injects the fixed `nsmap`; it never ADDED the capability.
%
%   Provenance (Gate-1..3, all 2026-07-26):
%     * Audit    : validation\mat2doc\audit_P2-2_storypart.md sec.4/sec.11
%                  (Porter Gate-1 + Opus Gate-2 [R] APPROVE -- "#60 hoist
%                  semantic-identity ... CONFIRMED PURE MOVE"; metaclass probe A4
%                  proves single definition; 42/42 MATLAB probes).
%     * Validate : validation\mat2doc\validate_P2-2_storypart.md Bar 2
%                  (#60 xpath hoist equivalence PASS -- DefiningClass on
%                  `mat2doc.oxml.XmlElement`, plain-element eval, Test_p1_3a/3b/3x
%                  GREEN inside the 383/383 full suite).
%     * Scenario : validation\mat2doc\scenarios\s0014_p2_2_storypart.{py,m}
%                  (probe `xpath_plain`, the identical parsed w:document blob).
%     * Frozen ref (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE):
%                  validation\mat2doc\references\s0014\probe.json ->
%                    "xpath_plain": { "id_values": ["7","12","notnum"],
%                                     "rid_values": ["rId9","rId9","rId3"] }
%                  Those frozen attribute-value vectors are embedded below as the
%                  ID_ORACLE / RID_ORACLE constants so this suite is self-contained.
%
%   Coverage taxonomy
%   -----------------
%   * Regression (metaclass, THE hoist pin) -- `xpath` is DEFINED exactly once,
%     on `mat2doc.oxml.XmlElement` (its DefiningClass), and `BaseOxmlElement`
%     resolves the SAME single definition (DefiningClass still XmlElement -- the
%     P1-3 copy was RELOCATED, not duplicated). RED if a future edit re-adds a
%     BaseOxmlElement copy or moves the definition back.
%   * Equivalence (frozen s0014 oracle) -- the parsed `w:document` root evaluates
%     `//@id` == ["7","12","notnum"] and `//@r:id` == ["rId9","rId9","rId3"]
%     byte-for-byte against the frozen python-docx/lxml oracle. (P2-2 NOTE, kept
%     for provenance: at P2-2 this root was the PLAIN unregistered
%     `mat2doc.oxml.XmlElement` -- the exact production condition of
%     StoryPart.next_id / XmlPart.rel_ref_count_. P2-3 registers CT_Document, so
%     the root is now a CT_Document; the hoisted `.xpath` is inherited unchanged
%     and the frozen vectors still hold. The plain-fallback #60 pin now lives in
%     test_registered_root_is_CT_Document via a genuinely-unregistered element.)
%   * Edge (H3 typed-empty return) -- `//@nomatch` on the plain root returns
%     `string.empty(1,0)` (a TYPED empty attribute-value array), NEVER [] (None):
%     callers use numel/~isempty, so the H3 contract must hold post-hoist.
%   * Regression (inheritance spot-check, P1-3 behaviour intact) -- a
%     BaseOxmlElement instance still evaluates `.xpath(".//w:p")` correctly WITH
%     the fixed-nsmap default injected (the whole point of the P1-3 override):
%     the `w:` prefix resolves without a `namespaces` argument. This pins that
%     the hoist did NOT break inheritance or the nsmap injection -- it does NOT
%     re-run the full P1-3x battery (Test_p1_3a/3b/3x own that).
%
%   Deviations exercised: none new. The `xpath` engine (evaluate_xpath) is the
%   Fable-audited P1-3x surface, RELOCATED unchanged; the hoist moved zero bytes
%   (validate_P2-2 Bar 2). No D-number.
%
%   Determinism: no network, no absolute paths -- the worktree root resolves
%   relative to this file via fileparts(mfilename('fullpath')); the parsed blob
%   is an in-memory literal.

    properties (Constant)
        % Frozen python-docx 1.2.0 / lxml 5.3.0 oracle (references\s0014\
        % probe.json "xpath_plain"): the attribute-value vectors a plain parsed
        % w:document root yields for //@id and //@r:id.
        ID_ORACLE  = ["7" "12" "notnum"]
        RID_ORACLE = ["rId9" "rId9" "rId3"]

        % The exact scenario blob from s0014 (parsed root is UNREGISTERED until
        % P2-3, so its runtime class is the fallback mat2doc.oxml.XmlElement --
        % the condition #60 exists to serve).
        W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        R_URI = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied from tests\oxml\Test_p1_3x_xpath.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\oxml
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. THE HOIST PIN (metaclass single-definition + inheritance)    %
        % =============================================================== %

        function test_xpath_defined_on_XmlElement(testCase)
            % Regression (#60 hoist, metaclass): `xpath` is defined ON
            % mat2doc.oxml.XmlElement -- its DefiningClass is XmlElement itself.
            % This is the relocation TARGET pin. (validate_P2-2 Bar 2.)
            dc = testCase.definingClassOf('mat2doc.oxml.XmlElement', 'xpath');
            testCase.verifyEqual(dc, "mat2doc.oxml.XmlElement", ...
                'xpath must be defined ON mat2doc.oxml.XmlElement (the #60 hoist target)');
        end

        function test_xpath_inherited_by_BaseOxmlElement_single_definition(testCase)
            % Regression (#60 hoist, THE pin): BaseOxmlElement resolves `xpath` to
            % the SAME single definition on XmlElement (DefiningClass still
            % mat2doc.oxml.XmlElement) -- the P1-3 copy was RELOCATED, not
            % duplicated. RED if a BaseOxmlElement copy is ever re-added or the
            % definition moves back. (audit_P2-2 sec.4 metaclass probe A4;
            % validate_P2-2 Bar 2.)
            dc = testCase.definingClassOf('mat2doc.oxml.BaseOxmlElement', 'xpath');
            testCase.verifyEqual(dc, "mat2doc.oxml.XmlElement", ...
                ['BaseOxmlElement.xpath must resolve to the single hoisted ' ...
                 'definition on XmlElement (inherited, not duplicated)']);
        end

        % =============================================================== %
        % 2. Registered root (CT_Document @ P2-3) evaluates the hoisted    %
        %    xpath; plain-fallback #60 pin preserved in the first method    %
        % =============================================================== %

        function test_registered_root_is_CT_Document(testCase)
            % P2-3 REGISTRATION FLIP (expected; this file's line ~130 comment,
            % below, pre-declared it): P2-3 registers w:document -> CT_Document, so
            % the parsed s0014 root now has runtime class EXACTLY
            % mat2doc.oxml.document.CT_Document -- NOT the old unregistered
            % XmlElement fallback. CT_Document IS-A mat2doc.oxml.XmlElement via
            % BaseOxmlElement, so it still carries the #60 hoisted `.xpath` (the
            % value tests below run on this same root and pass).
            % (validate_P2-3_document_shell.md JOB A, stale-pin #4.)
            root = testCase.plainRoot();
            testCase.verifyEqual(class(root), 'mat2doc.oxml.document.CT_Document', ...
                'parsed w:document root now resolves to the registered CT_Document (P2-3)');
            testCase.verifyInstanceOf(root, 'mat2doc.oxml.XmlElement', ...
                'CT_Document IS-A XmlElement, so the #60 hoisted xpath is still available');

            % #60 COVERAGE PRESERVED. The whole point of #60 is that the hoisted
            % `.xpath` reaches a PLAIN (unregistered) fallback element -- the
            % production condition of StoryPart.next_id / XmlPart.rel_ref_count_
            % on any not-yet-registered root. Now that w:document is registered,
            % that plain-element condition must be pinned with a GENUINELY
            % unregistered tag: a bare custom element still resolves to the plain
            % mat2doc.oxml.XmlElement fallback AND still evaluates the hoisted xpath.
            plain = testCase.unregisteredRoot();
            testCase.verifyEqual(class(plain), 'mat2doc.oxml.XmlElement', ...
                'a genuinely-unregistered root must still be the plain XmlElement fallback');
            ids = plain.xpath("//@id");
            testCase.verifyEqual(ids, ["3" "8"], ...
                'the #60 hoisted xpath must work on the plain XmlElement fallback');
        end

        function test_plain_root_xpath_id_values_match_oracle(testCase)
            % Equivalence (frozen s0014 oracle): `//@id` on the plain root ==
            % ["7","12","notnum"] -- the exact python-docx/lxml frozen vector.
            % Pins StoryPart.next_id's underlying xpath call reaching XmlElement
            % directly. (references\s0014\probe.json xpath_plain.id_values.)
            root = testCase.plainRoot();
            ids = root.xpath("//@id");
            testCase.verifyEqual(class(ids), 'string', ...
                '//@attr must return a string array');
            testCase.verifyEqual(ids, testCase.ID_ORACLE, ...
                '//@id must equal the frozen s0014 python-docx oracle');
        end

        function test_plain_root_xpath_rid_values_match_oracle(testCase)
            % Equivalence (frozen s0014 oracle): `//@r:id` on the plain root ==
            % ["rId9","rId9","rId3"] -- the r: prefix resolved via the injected
            % fixed nsmap, on a plain XmlElement. Pins XmlPart.rel_ref_count_'s
            % underlying call. (references\s0014\probe.json xpath_plain.rid_values.)
            root = testCase.plainRoot();
            rids = root.xpath("//@r:id");
            testCase.verifyEqual(rids, testCase.RID_ORACLE, ...
                '//@r:id must equal the frozen s0014 python-docx oracle (r: via fixed nsmap)');
        end

        function test_plain_root_xpath_no_match_typed_empty(testCase)
            % Edge (H3 typed-empty): `//@nomatch` returns string.empty(1,0), a
            % TYPED empty attribute-value array -- never [] (None). Callers use
            % numel/~isempty (e.g. next_id's `if isempty(used_ids)`), so the H3
            % contract must survive the hoist. (audit_P2-2 sec.4 probe A6.)
            root = testCase.plainRoot();
            empt = root.xpath("//@nomatch");
            testCase.verifyEqual(class(empt), 'string', ...
                'no-match //@attr must still be a string array (typed)');
            testCase.verifyEqual(numel(empt), 0, 'no-match //@attr must be empty');
            testCase.verifyEqual(empt, strings(1, 0), ...
                'no-match //@attr must be string.empty(1,0), NEVER [] (None)');
        end

        % =============================================================== %
        % 3. Inheritance spot-check -- P1-3 behaviour intact               %
        % =============================================================== %

        function test_baseoxmlelement_xpath_inherited_functional(testCase)
            % Regression (inheritance spot-check): a BaseOxmlElement instance
            % still evaluates `.xpath(".//w:p")` with the fixed-nsmap default
            % INJECTED -- the w: prefix resolves WITHOUT a namespaces argument
            % (the sole purpose of the P1-3 BaseOxmlElement.xpath override).
            % Proves the hoist neither broke inheritance nor the nsmap injection.
            % Does NOT duplicate the P1-3x battery (Test_p1_3a/3b/3x own that).
            root = mat2doc.oxml.BaseOxmlElement("w:document");
            child = mat2doc.oxml.BaseOxmlElement("w:p");
            root.append(child);
            got = root.xpath(".//w:p");   % w: resolved via injected fixed nsmap
            testCase.verifyEqual(numel(got), 1, ...
                'BaseOxmlElement.xpath(".//w:p") must find the one w:p child');
            testCase.verifyTrue(got(1) == child, ...
                'the found node must be the SAME handle (H5 element identity)');
        end

    end

    % ===================== instance helpers ============================ %
    methods (Access = private)

        function root = plainRoot(testCase)
            % Parse the frozen s0014 w:document blob into its root. NOTE (P2-3):
            % w:document is now registered, so this root is a CT_Document (an
            % XmlElement subclass) rather than the bare fallback it was at P2-2 --
            % the hoisted `.xpath` is inherited, so the frozen-oracle value tests
            % below are unchanged. Rebuilt per case (parsing is cheap; keeps the
            % cases independent).
            blob = "<w:document xmlns:w=""" + testCase.W_URI + """ " + ...
                "xmlns:r=""" + testCase.R_URI + """><w:body>" + ...
                "<w:p id=""7""/><w:p id=""12""/><w:p id=""notnum""/>" + ...
                "<w:x r:id=""rId9""/><w:y r:id=""rId9""/><w:z r:id=""rId3""/>" + ...
                "</w:body></w:document>";
            root = mat2doc.oxml.parse_xml(uint8(unicode2native(char(blob), "UTF-8")));
        end

        function root = unregisteredRoot(~)
            % A GENUINELY-unregistered root (a bare custom-namespace element that
            % no registry row claims -- unlike w:document, registered at P2-3).
            % Its runtime class is the plain mat2doc.oxml.XmlElement fallback, so
            % it is the post-P2-3 vehicle for the #60 "xpath on a plain element"
            % pin. Carries two unprefixed id attributes for the //@id probe.
            blob = "<z:doc xmlns:z=""urn:mat2doc:test"">" + ...
                "<z:a id=""3""/><z:b id=""8""/></z:doc>";
            root = mat2doc.oxml.parse_xml(uint8(unicode2native(char(blob), "UTF-8")));
        end

        function dc = definingClassOf(~, clsName, methodName)
            % The DefiningClass name (string) of methodName as resolved on the
            % metaclass of clsName -- the mechanical "where is this method
            % defined" probe. Returns "" if the method is absent.
            mc = meta.class.fromName(clsName);
            ml = mc.MethodList;
            i = find(strcmp({ml.Name}, methodName), 1);
            if isempty(i)
                dc = "";
            else
                dc = string(ml(i).DefiningClass.Name);
            end
        end
    end
end

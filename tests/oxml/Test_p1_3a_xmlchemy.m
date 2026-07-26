classdef Test_p1_3a_xmlchemy < matlab.unittest.TestCase
% TEST_P1_3A_XMLCHEMY  Gate-4 permanent unit tests for Mat2Doc P1-3a
%   (xmlchemy engine pt1: BaseOxmlElement tree-ops + attribute descriptors +
%   evaluate_xpath / BaseOxmlElement.xpath).
%
%   Surface under test (ported from python-docx v1.2.0 src/docx/oxml/xmlchemy.py,
%   realised in +mat2doc\+oxml\BaseOxmlElement.m + +oxml\evaluate_xpath.m):
%     * the mini-XPath engine reached via BaseOxmlElement.xpath / evaluate_xpath;
%     * the OptionalAttribute / RequiredAttribute descriptor engine
%       (getAttrTyped / setAttrTyped / getAttrRequired / setAttrRequired);
%     * the tree-ops first_child_found_in / insert_element_before / remove_all.
%
%   This class ENCODES THE FROZEN GATE-3 53-CASE BATTERY verbatim -- the
%   Validator's oracle copied byte-for-byte into tests\oxml\data\p1_3a\ so the
%   suite is self-contained. Provenance:
%     validation\mat2doc\references\p1_3a\ (manifest mso-p1_3a-oracle/1, frozen
%       2026-07-26; oracle lxml 5.3.0 / libxml2 2.13.9 under python-docx 1.2.0);
%     Gate-3 report validation\mat2doc\reports\p1_3a_validation.md (PASS 53/53).
%   The oracle values are lxml's own results, so value/byte-identical here ==
%   identical to python-docx: these are equivalence AND regression pins in one.
%
%   Case map (1 MATLAB test per oracle case, 53 total):
%     * XPath battery, 38  -- test_xpath_case parameterised over the 38 ids in
%         xpath\cases.json, dispatched on the frozen `expect` (nodes / strings /
%         raise) against xpath\oracle.json. Covers positional 1-based no-shift,
%         //, absolute, element/attr predicates, union->doc-ordered group (H11),
%         ancestor::, typed-empty no-match, //@r:id, @xml:space, and
%         mixed-content text() document ordering.
%     * Attribute descriptors, 10 -- O1..O7 (OptionalAttribute) + R1..R3
%         (RequiredAttribute), exercised through the test-only simple type
%         p13test.TST (no dependency on P3 +simpletypes). Serialize+byte-check
%         where the op emits (O2/O7/R2 vs the frozen .oracle.xml).
%     * Tree-ops, 5 -- T1/T2 insert_element_before (successor placement +
%         append-fallback, serialize+byte), T3 first_child_found_in (arg-order +
%         absent), T4 remove_all (serialize+byte), T5 xpath() public-nsmap
%         default injection.
%
%   *** FIX-1 self/parent name-test pins (F01-F05) ***
%   F01-F05 (inside the xpath battery) pin the Gate-2 engine fix that makes
%   self::NAME / parent::NAME apply the NODE-NAME TEST (libxml2/lxml semantics):
%     F01 self::w:tbl on a w:p            -> [] (name mismatch)
%     F02 self::w:p   on a w:p            -> [p]
%     F03 parent::w:body on a body-child -> [body]
%     F04 ./parent::w:r/parent::w:hyperlink from a w:t inside a hyperlink -> [hyperlink]
%     F05 same expr from a PLAIN (non-hyperlink) run's w:t              -> []
%   Any regression that drops the name test (matching self/parent unconditionally)
%   turns F01/F05 red.
%
%   *** VERIFY-2 value-pins (R01-R09) -- P1-3x MERGED, now SUPPORTED ***
%   R01-R09 (inside the xpath battery) are docx call-site XPath patterns that
%   were OUTSIDE the P1-3a evaluator subset (bare union `./w:p | ./w:tbl`,
%   not(self::), preceding/following-sibling::, preceding::, position()=1,
%   [last()], a predicate attribute-subpath, //@r:id[2]). At P1-3a they raised
%   mat2doc:XPathError ("never silently mis-evaluated", design.md section XPath).
%   The engine-extension WP P1-3x (xpath-engine-extension,
%   decision_2026-07-25_mat2doc_xpath_engine_extension.md) has now LANDED and
%   IMPLEMENTS every one of these forms, so R01-R09 now RETURN node-sets /
%   strings instead of raising. The `case "raise"` dispatch has accordingly been
%   FLIPPED from verifyError to a value/type assertion, per the frozen pin-flip
%   recipe validation\...\p1_3x\flip\flip_types.json (manifest mso-p1_3x-flip/1)
%   against the `would_be` node-set frozen in xpath\oracle.json:
%     R01,R02,R04,R06,R07 -> non-empty NODES == would_be (tag+path value-compare);
%     R03,R05,R08         -> typed-empty NODES (isa XmlElement && isempty);
%     R09 (//@r:id[2])    -> typed-empty STRING (isstring && isempty). R09's
%       would_be labels {kind:nodes,value:[]} only because the oracle generator's
%       encode() classifies every empty lxml list as "nodes"; the engine's
%       STRING typed-empty is the CORRECT H3 static type (attr terminal), so R09
%       asserts emptiness, NOT the would_be "nodes" kind (flip_types.json note_R09).
%   These are now permanent SUPPORTED-behaviour equivalence pins; the extended
%   surface is exercised in full by tests\oxml\Test_p1_3x_xpath.m.
%
%   VERIFY-3 (from the manifest): the battery runs over the PARSED w:document
%   tree; the registry is EMPTY at P1-3a so the parsed tree is homogeneous
%   XmlElement -- the heterogeneous / Sealed-method risk is not yet exercisable
%   and must be re-verified at P1-3b once CT_* classes register.
%
%   Determinism: no network, no absolute paths -- every fixture is resolved
%   relative to this file via fileparts(mfilename('fullpath')). The test-only
%   simple type lives at tests\oxml\+p13test\TST.m and is reached only because
%   TestClassSetup puts tests\oxml on the path.

    properties (TestParameter)
        % The 38 frozen XPath case ids (xpath\cases.json). One MATLAB test each;
        % the method dispatches on the frozen `expect` field.
        xpathId = { ...
            'X01','X02','X03','X04','X05','X06','X07','X08','X09','X10', ...
            'X11','X12','X13','X14','X15','X16', ...
            'S01','S02','S03','S04','S05','S06','S07','S08', ...
            'F01','F02','F03','F04','F05', ...
            'R01','R02','R03','R04','R05','R06','R07','R08','R09'};
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into
            % the test folder, so without the worktree root on the path a COLD
            % run cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % SECOND fixture puts tests\oxml on the path so the test-only
            % simple-type package p13test.TST resolves (the descriptor engine
            % feval's "p13test.TST.from_xml"/"...to_xml" via its dotted-name
            % currency). Both are auto-restored after the class.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\oxml
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
            testCase.applyFixture(PathFixture(here));
        end
    end

    methods (Test)

        % =============================================== XPATH battery (38) ===
        function test_xpath_case(testCase, xpathId)
            % Equivalence + regression: parse document.in.xml, resolve the case
            % context, evaluate the case expr, and assert the result equals the
            % frozen lxml oracle -- node-set (tag + 0-based document path),
            % string-set (ordered), or a mat2doc:XPathError raise (VERIFY-2).
            ns    = mat2doc.oxml.nsmap();
            root  = parseDoc();
            [ctxLoc, expr, expect] = caseSpec(xpathId);

            % resolve context (must NOT raise for any case, incl. the R-raises)
            if ctxLoc == ""
                ctx = root;
            else
                ctxArr = mat2doc.oxml.evaluate_xpath(root, ctxLoc, ns);
                testCase.assertNotEmpty(ctxArr, ...
                    sprintf('%s: context "%s" resolved empty', xpathId, ctxLoc));
                ctx = ctxArr(1);
            end

            switch expect
                case "raise"
                    % P1-3x MERGED (see class header): the former outside-subset
                    % forms R01-R09 are now SUPPORTED and return node-sets/strings
                    % rather than raising. Flipped to a value/type assertion per
                    % flip_types.json (manifest mso-p1_3x-flip/1) against the frozen
                    % `would_be` in xpath\oracle.json. `expect` is still "raise" in
                    % cases.json (the frozen P1-3a spec); the flip lives HERE so the
                    % frozen oracle stays untouched.
                    got = mat2doc.oxml.evaluate_xpath(ctx, expr, ns);
                    switch flipKind(xpathId)
                        case "string-empty"
                            % R09 //@r:id[2] -> typed-empty STRING (attr terminal).
                            % Assert emptiness, NOT the would_be 'nodes' kind
                            % (flip_types.json note_R09: encode() artifact).
                            testCase.verifyTrue(isstring(got), sprintf( ...
                                '%s: %s must return a string array (STRING typed-empty)', ...
                                xpathId, expr));
                            testCase.verifyTrue(isempty(got), sprintf( ...
                                '%s: %s must be the empty string-set', xpathId, expr));
                        case "nodes-empty"
                            % R03/R05/R08 -> typed-empty NODES.
                            testCase.verifyClass(got, 'mat2doc.oxml.XmlElement', ...
                                sprintf('%s: typed-empty node result must be XmlElement', xpathId));
                            testCase.verifyTrue(isempty(got), sprintf( ...
                                '%s: %s must be the empty node-set', xpathId, expr));
                        case "nodes-nonempty"
                            % R01/R02/R04/R06/R07 -> would_be node-set (tag+path).
                            testCase.verifyClass(got, 'mat2doc.oxml.XmlElement', ...
                                sprintf('%s: node result must be XmlElement', xpathId));
                            [expTags, expPaths] = expectedNodes(oracleEntry(xpathId).would_be);
                            testCase.verifyEqual(numel(got), numel(expTags), ...
                                sprintf('%s: node count vs would_be', xpathId));
                            for i = 1:numel(expTags)
                                testCase.verifyEqual(string(got(i).tag), expTags(i), ...
                                    sprintf('%s: node %d tag vs would_be', xpathId, i));
                                testCase.verifyEqual(path0(got(i)), expPaths{i}, ...
                                    sprintf('%s: node %d document path vs would_be', xpathId, i));
                            end
                        otherwise
                            testCase.assertFail(sprintf( ...
                                '%s: no P1-3x flip kind for a "raise" case', xpathId));
                    end
                case "nodes"
                    got = mat2doc.oxml.evaluate_xpath(ctx, expr, ns);
                    testCase.verifyClass(got, 'mat2doc.oxml.XmlElement', ...
                        sprintf('%s: node result must be XmlElement', xpathId));
                    [expTags, expPaths] = expectedNodes(oracleEntry(xpathId));
                    testCase.verifyEqual(numel(got), numel(expTags), ...
                        sprintf('%s: node count', xpathId));
                    for i = 1:numel(expTags)
                        testCase.verifyEqual(string(got(i).tag), expTags(i), ...
                            sprintf('%s: node %d tag', xpathId, i));
                        testCase.verifyEqual(path0(got(i)), expPaths{i}, ...
                            sprintf('%s: node %d document path', xpathId, i));
                    end
                case "strings"
                    got = mat2doc.oxml.evaluate_xpath(ctx, expr, ns);
                    testCase.verifyTrue(isstring(got), ...
                        sprintf('%s: string result must be a string array', xpathId));
                    expStr = expectedStrings(oracleEntry(xpathId));
                    testCase.verifyEqual(got, expStr, ...
                        sprintf('%s: string-set (ordered)', xpathId));
                otherwise
                    testCase.assertFail(sprintf('%s: unknown expect "%s"', xpathId, expect));
            end
        end

        % ====================================== Attribute descriptors (10) ====
        % OptionalAttribute (O*) / RequiredAttribute (R*) via the test-only
        % simple type p13test.TST. Property-value expectations are hard-coded
        % (regression) and cross-checked against attr\oracle.json; serialized
        % bytes are compared to the frozen .oracle.xml (equivalence, byte-exact).

        function test_O1_get_absent_default(testCase)
            % OptionalAttribute getter, attribute absent -> returns the default,
            % attribute NOT created. oracle O1: getter_return "DEF", attrib null.
            o = newFoo();
            testCase.verifyEqual(o.getAttrTyped("w:val", tt(), "DEF"), "DEF");
            testCase.verifyTrue(isequal(o.get(clark("w:val")), []), ...
                'O1: absent attribute must read [] (None)');
        end

        function test_O2_set_stored(testCase)
            % Set "Vx" (!= default) -> stored under the Clark key; getter now
            % round-trips through from_xml ("F:Vx"); serialized bytes byte-exact.
            % oracle O2.
            o = newFoo();
            o.setAttrTyped("w:val", tt(), "Vx", "DEF");
            testCase.verifyEqual(string(o.get(clark("w:val"))), "Vx", 'O2: stored value');
            testCase.verifyEqual(o.getAttrTyped("w:val", tt(), "DEF"), "F:Vx", ...
                'O2: getter round-trips via from_xml');
            verifyBytes(testCase, serialize(o), oracleXml('attr','O2_set_stored'), 'O2');
        end

        function test_O3_from_xml_roundtrip(testCase)
            % Getter on a set attribute invokes from_xml. oracle O3: "F:Vx".
            o = newFoo(); o.setAttrTyped("w:val", tt(), "Vx", "DEF");
            testCase.verifyEqual(o.getAttrTyped("w:val", tt(), "DEF"), "F:Vx");
        end

        function test_O4_setNone_removes(testCase)
            % D-delta-1: assigning [] (None) with a NON-None default REMOVES the
            % attribute (docx line 203 explicit `value is None`). oracle O4:
            % attrib null, getter_return "DEF".
            o = newFoo(); o.setAttrTyped("w:val", tt(), "Vx", "DEF");
            o.setAttrTyped("w:val", tt(), [], "DEF");
            testCase.verifyTrue(isequal(o.get(clark("w:val")), []), 'O4: attribute removed');
            testCase.verifyEqual(o.getAttrTyped("w:val", tt(), "DEF"), "DEF");
        end

        function test_O5_setDefault_removes(testCase)
            % Assigning value == default REMOVES the attribute. oracle O5: null.
            o = newFoo(); o.setAttrTyped("w:val", tt(), "Vx", "DEF");
            o.setAttrTyped("w:val", tt(), "DEF", "DEF");
            testCase.verifyTrue(isequal(o.get(clark("w:val")), []), 'O5: attribute removed');
        end

        function test_O6_toxmlNone_removes(testCase)
            % D-delta-2: to_xml returning None (sentinel "__TOXML_NONE__") REMOVES
            % the attribute after conversion (docx lines 208-211). oracle O6: null.
            o = newFoo(); o.setAttrTyped("w:val", tt(), "Vx", "DEF");
            o.setAttrTyped("w:val", tt(), "__TOXML_NONE__", "DEF");
            testCase.verifyTrue(isequal(o.get(clark("w:val")), []), 'O6: attribute removed');
        end

        function test_O7_set_emptystring(testCase)
            % H3 tri-state: "" is a real string (isequal("",[]) is false), so it
            % is STORED, not removed -> w:val="" serializes byte-exact. oracle O7.
            o = newFoo(); o.setAttrTyped("w:val", tt(), "", "DEF");
            testCase.verifyEqual(string(o.get(clark("w:val"))), "", 'O7: empty string stored');
            verifyBytes(testCase, serialize(o), oracleXml('attr','O7_set_emptystring'), 'O7');
        end

        function test_R1_get_absent_raises(testCase)
            % RequiredAttribute getter, attribute absent -> mat2doc:InvalidXmlError
            % with byte-exact message (obj.tag is the Clark name). oracle R1.
            o = newFoo();
            err = grabError(@() o.getAttrRequired("w:req", tt()));
            testCase.assertNotEmpty(err, 'R1: must raise');
            testCase.verifyEqual(err.identifier, 'mat2doc:InvalidXmlError');
            testCase.verifyEqual(string(err.message), ...
                "required 'w:req' attribute not present on element " + ...
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}foo", ...
                'R1: byte-exact message');
        end

        function test_R2_set_stored(testCase)
            % RequiredAttribute setter stores "Q"; getter round-trips ("F:Q");
            % serialized bytes byte-exact. oracle R2.
            o = newFoo(); o.setAttrRequired("w:req", tt(), "Q");
            testCase.verifyEqual(string(o.get(clark("w:req"))), "Q", 'R2: stored value');
            testCase.verifyEqual(o.getAttrRequired("w:req", tt()), "F:Q", ...
                'R2: getter round-trips via from_xml');
            verifyBytes(testCase, serialize(o), oracleXml('attr','R2_set_stored'), 'R2');
        end

        function test_R3_toxmlNone_raises(testCase)
            % D-delta-3: to_xml returning None on a REQUIRED attribute raises
            % mat2doc:ValueError with byte-exact message (never removes). oracle R3.
            o = newFoo();
            err = grabError(@() o.setAttrRequired("w:req", tt(), "__TOXML_NONE__"));
            testCase.assertNotEmpty(err, 'R3: must raise');
            testCase.verifyEqual(err.identifier, 'mat2doc:ValueError');
            testCase.verifyEqual(string(err.message), ...
                "cannot assign __TOXML_NONE__ to this required attribute", ...
                'R3: byte-exact message');
        end

        % ================================================= Tree-ops (5) =======
        function test_T1_insert_before(testCase)
            % insert_element_before places elm immediately before the first
            % successor found among tagnames (a w:pPr before an existing w:r).
            % Serialize byte-exact. oracle T1.
            p = newP();
            p.append(mat2doc.oxml.OxmlElement("w:r"));
            p.insert_element_before(mat2doc.oxml.OxmlElement("w:pPr"), "w:r");
            verifyBytes(testCase, serialize(p), oracleXml('treeops','T1_insert_before'), 'T1');
        end

        function test_T2_insert_append_fallback(testCase)
            % insert_element_before append-fallback: no successor present (no
            % w:pPr) -> elm APPENDED at the end. Serialize byte-exact. oracle T2.
            p = newP();
            p.append(mat2doc.oxml.OxmlElement("w:r"));
            p.insert_element_before(mat2doc.oxml.OxmlElement("w:br"), "w:pPr");
            verifyBytes(testCase, serialize(p), ...
                oracleXml('treeops','T2_insert_append_fallback'), 'T2');
        end

        function test_T3_first_child_found_in(testCase)
            % first_child_found_in searches in ARGUMENT order (first NAME that
            % matches wins, NOT document order); absent -> [] (None). oracle T3.
            wns = wuri();
            p = newP();
            p.append(mat2doc.oxml.OxmlElement("w:r"));     % doc order: w:r first
            p.append(mat2doc.oxml.OxmlElement("w:pPr"));   %            w:pPr second
            got1 = p.first_child_found_in("w:pPr", "w:r");
            testCase.verifyEqual(string(got1.tag), "{" + wns + "}pPr", ...
                'T3: arg-order pPr-first wins despite later doc position');
            got2 = p.first_child_found_in("w:r", "w:pPr");
            testCase.verifyEqual(string(got2.tag), "{" + wns + "}r", ...
                'T3: arg-order r-first wins');
            testCase.verifyTrue(isequal(p.first_child_found_in("w:zzz"), []), ...
                'T3: absent -> [] (None)');
        end

        function test_T4_remove_all(testCase)
            % remove_all strips every child whose tag is listed (all w:r),
            % leaving the rest (w:pPr). Serialize byte-exact. oracle T4.
            p = newP();
            p.append(mat2doc.oxml.OxmlElement("w:r"));
            p.append(mat2doc.oxml.OxmlElement("w:pPr"));
            p.append(mat2doc.oxml.OxmlElement("w:r"));
            p.append(mat2doc.oxml.OxmlElement("w:r"));
            p.remove_all("w:r");
            verifyBytes(testCase, serialize(p), oracleXml('treeops','T4_remove_all'), 'T4');
        end

        function test_T5_xpath_method_default_nsmap(testCase)
            % BaseOxmlElement.xpath with NO namespaces arg injects the docx
            % PUBLIC nsmap default (mat2doc.oxml.nsmap), so a prefixed query
            % resolves. oracle T5: count 1, tag w:r.
            p = newP();
            p.append(mat2doc.oxml.OxmlElement("w:r"));
            found = p.xpath("w:r");                          % default nsmap
            testCase.verifyEqual(numel(found), 1, 'T5: match count');
            testCase.verifyEqual(string(found(1).tag), "{" + wuri() + "}r", 'T5: matched tag');
        end
    end
end

% ============================= file-local helpers ==========================

function u = wuri()
    u = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
end

function s = tt()
    % Dotted type token -> resolveTypeCls_ uses it VERBATIM -> feval on the
    % test-only p13test.TST simple-type double.
    s = "p13test.TST";
end

function o = newFoo()
    % Fresh loose <w:foo> carrying the w prefix binding (matches the oracle).
    o = mat2doc.oxml.BaseOxmlElement("w:foo", struct('w', wuri()));
end

function p = newP()
    p = mat2doc.oxml.BaseOxmlElement("w:p", struct('w', wuri()));
end

function k = clark(name)
    k = mat2doc.oxml.qn(name);
end

function b = serialize(elm)
    b = uint8(mat2doc.oxml.serialize_part_xml(elm));
    b = b(:)';
end

function root = parseDoc()
    root = mat2doc.oxml.parse_xml(readFixture(dataPath("document.in.xml")));
end

function p = dataPath(varargin)
    here = fileparts(mfilename('fullpath'));   % tests\oxml
    p = fullfile(here, 'data', 'p1_3a', varargin{:});
end

function b = readFixture(p)
    f = fopen(p, 'r');
    if f < 0; b = uint8([]); return; end
    b = fread(f, Inf, '*uint8')';
    fclose(f);
end

function b = oracleXml(sub, base)
    b = readFixture(dataPath(char(sub), char(base) + ".oracle.xml"));
end

function [ctxLoc, expr, expect] = caseSpec(cid)
    cases = jsondecode(fileread(dataPath('xpath','cases.json')));
    ids = string({cases.id});
    k = find(ids == string(cid), 1);
    if isempty(k)
        error('Test:missingCase', 'case %s not in cases.json', cid);
    end
    ctxLoc = string(cases(k).ctx);
    expr   = string(cases(k).expr);
    expect = string(cases(k).expect);
end

function fk = flipKind(cid)
    % P1-3x pin-flip recipe: per-pin MEASURED return type of the extended engine
    % (validation\...\p1_3x\flip\flip_types.json, manifest mso-p1_3x-flip/1).
    %   R01,R02,R04,R06,R07 -> "nodes-nonempty" (would_be node-set)
    %   R03,R05,R08         -> "nodes-empty"     (typed-empty XmlElement)
    %   R09 //@r:id[2]      -> "string-empty"    (STRING typed-empty; attr terminal)
    switch string(cid)
        case {"R01","R02","R04","R06","R07"}; fk = "nodes-nonempty";
        case {"R03","R05","R08"};             fk = "nodes-empty";
        case "R09";                           fk = "string-empty";
        otherwise;                            fk = "";
    end
end

function e = oracleEntry(cid)
    oracle = jsondecode(fileread(dataPath('xpath','oracle.json')));
    e = oracle.(char(cid));
end

function [tags, paths] = expectedNodes(entry)
    % Normalise oracle.<cid>.value (list of {tag, path}) into a string array of
    % Clark tags and a cell of 0-based document-path row vectors.
    tags = string.empty(1, 0);
    paths = {};
    val = entry.value;
    if isempty(val); return; end
    if iscell(val)
        for i = 1:numel(val)
            tags(end+1) = string(val{i}.tag);          %#ok<AGROW>
            paths{end+1} = double(val{i}.path(:))';    %#ok<AGROW>
        end
    else
        for i = 1:numel(val)
            tags(end+1) = string(val(i).tag);          %#ok<AGROW>
            paths{end+1} = double(val(i).path(:))';    %#ok<AGROW>
        end
    end
end

function s = expectedStrings(entry)
    % Normalise oracle.<cid>.value (list of strings) into a (1,N) string row.
    val = entry.value;
    if isempty(val)
        s = strings(1, 0);
    elseif iscell(val)
        s = string(val(:))';
    else
        s = string(val(:))';
    end
end

function p = path0(el)
    % 0-based element-child indices from the document root down to el
    % (matches the oracle path encoding).
    p = [];
    cur = el;
    while true
        par = cur.getparent();
        if isequal(par, []); break; end
        arr = par.to_array();
        idx = find(arr == cur, 1) - 1;   % 0-based
        p = [idx, p]; %#ok<AGROW>
        cur = par;
    end
end

function err = grabError(fn)
    err = [];
    try
        fn();
    catch e
        err = e;
    end
end

function verifyBytes(testCase, got, want, label)
    % Byte-level equivalence (L1). On mismatch, report sizes + first diff offset.
    got = uint8(got(:)');
    want = uint8(want(:)');
    testCase.assertNotEmpty(want, sprintf('%s: oracle fixture missing/empty', label));
    if ~isequal(got, want)
        n = min(numel(got), numel(want));
        d = find(got(1:n) ~= want(1:n), 1);
        if isempty(d); d = n + 1; end
        diag = sprintf(['%s: bytes differ (got %d B, want %d B); first diff @%d ' ...
            '(got 0x%02X, want 0x%02X)'], char(label), numel(got), numel(want), ...
            d, byteAt(got, d), byteAt(want, d));
    else
        diag = char(label);
    end
    testCase.verifyEqual(got, want, diag);
end

function v = byteAt(b, i)
    if i >= 1 && i <= numel(b); v = double(b(i)); else; v = 0; end
end

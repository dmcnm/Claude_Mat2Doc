classdef Test_p1_3x_xpath < matlab.unittest.TestCase
% TEST_P1_3X_XPATH  Gate-4 permanent unit tests for Mat2Doc P1-3x
%   (xpath-engine-extension: the docx call-site XPath patterns the pptx-derived
%   P1-3a evaluator did not cover, now realised in +mat2doc\+oxml\evaluate_xpath).
%
%   Surface under test (extension forms, see evaluate_xpath.m "P1-3x EXTENSION"):
%     * bare top-level unions            ./w:p | ./w:tbl  (doc-order + identity dedup)
%     * not(path) / union-subpath preds  ./*[not(self::w:sectPr)]
%     * following-sibling:: / preceding-sibling:: axes (name-test + positional [1])
%     * preceding:: axis                 ./preceding::w:sectPr[1] (ancestor-exclusion)
%     * position()=n / (...)[last()]     w:p[position()=1]   (./w:p)[last()]
%     * group predicate                  (./w:comment[@w:id='5'])[1]
%     * predicate attribute sub-path     w:style[w:name/@w:val="Heading 1"]
%     * positional predicate on a        //@r:id[2]  ([] -- one r:id per element;
%       terminal @attr step               STRING typed-empty, not a node-set)
%     * compound reverse-axis positional ./ancestor::w:tr[position()=1]/preceding-sibling::w:tr[1]
%     * compound pagebreak forms         (./w:r)[last()]/w:lastRenderedPageBreak[not(...)]
%
%   This class ENCODES THE FROZEN GATE-3 P1-3x ORACLE verbatim -- the Validator's
%   value oracle copied byte-for-byte into tests\oxml\data\p1_3x\ so the suite is
%   self-contained. Provenance:
%     validation\mat2doc\references\p1_3x\ (manifest mso-p1_3x-references/1, frozen
%       2026-07-25; oracle lxml 5.3.0 / libxml2 2.13.9 under python-docx 1.2.0);
%     Gate-3 report validation\mat2doc\reports\p1_3x_validation.md (PASS 61/61).
%   The oracle values ARE lxml's own results, so value-identical here == identical
%   to python-docx: these are equivalence AND regression pins in one.
%
%   Case map (one MATLAB test per oracle case + 2 authored union pins):
%     * Call-site battery, 42 -- test_callsite_case parameterised over the 42 ids
%         in callsites\battery.json (frozen Gate-1, independently re-verified
%         against live lxml with 0 diffs). Fixtures DOC/TBL/PB1/PB2/CMTS/STYLES/
%         NUM/SECT/DRAW/RUN are the inline fixture strings inside battery.json.
%         Dispatched on the frozen `kind` (nodes / strings) against `value`.
%         EXCEPTION S1 (//@r:id[2]): battery.json labels it {kind:nodes,value:[]}
%         but the terminal @attr step returns a STRING typed-empty (the same
%         oracle-generator encode() artifact documented in the P1-3x flip
%         manifest note_R09) -- S1 asserts isstring && isempty, NOT a node kind.
%     * Real-parts replay, 10 -- test_realpart_case parameterised over RP1..RP10
%         in realparts\oracle.json, evaluated over the REAL word/document.xml
%         (python-docx Document() + injected w:lastRenderedPageBreak), a byte
%         fixture in realparts\document.xml. Exercises the extension patterns on
%         a real multi-row table + pagebreak tree.
%     * Union invariants, 2 -- authored regression pins (derived from lxml
%         node-set set-semantics, NOT frozen-oracle rows) that isolate the two
%         union guarantees the task calls out explicitly:
%           test_union_interleaved_docorder: `./w:tbl | ./w:p` returns DOCUMENT
%             order [p,p,tbl,p,p], NOT operand order [tbl,p,p,p,p] (H11 doc-order).
%           test_union_overlapping_dedup: `./w:p | ./w:p` returns the 4 distinct
%             paragraphs ONCE each, not 8 (H11 identity dedup -- the union trap).
%
%   Determinism: no network, no absolute paths -- every fixture is resolved
%   relative to this file via fileparts(mfilename('fullpath')).

    properties (TestParameter)
        % The 42 frozen call-site battery ids (callsites\battery.json).
        callsiteId = { ...
            'U1','U2','U3','U4','U5','U6','U7','U8','U9', ...
            'N1','N2','N3','N4','N5','N6', ...
            'A1','A2','A3','A4','A4b', ...
            'P1','P2','P3','P4','P5','P6','P7','P8','P9', ...
            'E1','E2','E3', ...
            'S1','S2','S3','S4', ...
            'G1','G2','G3','G4','G5','G6'};
        % The 10 real-parts replay ids (realparts\oracle.json).
        realpartId = {'RP1','RP2','RP3','RP4','RP5','RP6','RP7','RP8','RP9','RP10'};
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\oxml
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % ===================================== call-site battery (42) =========
        function test_callsite_case(testCase, callsiteId)
            % Equivalence + regression: parse the case fixture, resolve ctx,
            % evaluate expr, assert the result equals the frozen lxml oracle.
            ns  = mat2doc.oxml.nsmap();
            bat = battery();
            c   = batteryCase(bat, callsiteId);
            root = mat2doc.oxml.parse_xml(string(bat.fixtures.(char(c.fixture))));
            ctx  = resolveCtx(testCase, root, string(c.ctx), ns, callsiteId);

            got = mat2doc.oxml.evaluate_xpath(ctx, string(c.expr), ns);

            if string(callsiteId) == "S1"
                % S1 //@r:id[2] -- terminal @attr step -> STRING typed-empty.
                % battery.json labels {kind:nodes,value:[]} only as the oracle
                % encode() artifact (see class header / flip manifest note_R09).
                testCase.verifyTrue(isstring(got), ...
                    'S1: //@r:id[2] must return a string array (STRING typed-empty)');
                testCase.verifyTrue(isempty(got), ...
                    'S1: //@r:id[2] must be the empty string-set');
                return
            end

            switch string(c.kind)
                case "nodes"
                    testCase.verifyClass(got, 'mat2doc.oxml.XmlElement', ...
                        sprintf('%s: node result must be XmlElement', callsiteId));
                    assertNodes(testCase, got, c.value, callsiteId);
                case "strings"
                    testCase.verifyTrue(isstring(got), ...
                        sprintf('%s: string result must be a string array', callsiteId));
                    testCase.verifyEqual(got, expectedStrings(c.value), ...
                        sprintf('%s: string-set (ordered)', callsiteId));
                otherwise
                    testCase.assertFail(sprintf('%s: unknown kind "%s"', callsiteId, c.kind));
            end
        end

        % ======================================== real-parts replay (10) ======
        function test_realpart_case(testCase, realpartId)
            % Extension patterns replayed over the REAL word/document.xml byte
            % fixture; assert against realparts\oracle.json (all node-set cases,
            % RP7/RP8 empty).
            ns  = mat2doc.oxml.nsmap();
            root = mat2doc.oxml.parse_xml(readFixture(dataPath('realparts','document.xml')));
            c   = realpartCase(realpartId);
            ctx = resolveCtx(testCase, root, string(c.ctx), ns, realpartId);

            got = mat2doc.oxml.evaluate_xpath(ctx, string(c.expr), ns);
            testCase.verifyClass(got, 'mat2doc.oxml.XmlElement', ...
                sprintf('%s: node result must be XmlElement', realpartId));
            assertNodes(testCase, got, c.value, realpartId);
        end

        % ======================================== union invariants (2) ========
        function test_union_interleaved_docorder(testCase)
            % H11 doc-order: a union returns DOCUMENT order regardless of operand
            % order. Operands reversed (tbl before p) but result is still the
            % interleaved [p(0,0),p(0,1),tbl(0,2),p(0,3),p(0,4)] -- NOT
            % [tbl,p,p,p,p]. Authored pin over the frozen DOC fixture.
            ns  = mat2doc.oxml.nsmap();
            bat = battery();
            root = mat2doc.oxml.parse_xml(string(bat.fixtures.DOC));
            body = mat2doc.oxml.evaluate_xpath(root, "w:body", ns);
            got  = mat2doc.oxml.evaluate_xpath(body(1), "./w:tbl | ./w:p", ns);
            testCase.verifyClass(got, 'mat2doc.oxml.XmlElement', 'interleave: XmlElement');
            expTags  = ["p","p","tbl","p","p"];
            expPaths = {[0 0],[0 1],[0 2],[0 3],[0 4]};
            testCase.verifyEqual(numel(got), numel(expTags), 'interleave: node count');
            for i = 1:numel(expTags)
                testCase.verifyEqual(localName(got(i)), expTags(i), ...
                    sprintf('interleave: node %d tag (doc-order not operand-order)', i));
                testCase.verifyEqual(path0(got(i)), expPaths{i}, ...
                    sprintf('interleave: node %d document path', i));
            end
        end

        function test_union_overlapping_dedup(testCase)
            % H11 identity dedup (the union trap -- node-sets are sets): overlapping
            % operands `./w:p | ./w:p` return each matching paragraph ONCE (4), not
            % twice (8). Authored pin over the frozen DOC fixture.
            ns  = mat2doc.oxml.nsmap();
            bat = battery();
            root = mat2doc.oxml.parse_xml(string(bat.fixtures.DOC));
            body = mat2doc.oxml.evaluate_xpath(root, "w:body", ns);
            got  = mat2doc.oxml.evaluate_xpath(body(1), "./w:p | ./w:p", ns);
            testCase.verifyClass(got, 'mat2doc.oxml.XmlElement', 'dedup: XmlElement');
            expPaths = {[0 0],[0 1],[0 3],[0 4]};    % the 4 w:p (0,2 is w:tbl)
            testCase.verifyEqual(numel(got), 4, ...
                'dedup: overlapping union must yield 4 distinct paragraphs, not 8');
            for i = 1:numel(expPaths)
                testCase.verifyEqual(localName(got(i)), "p", sprintf('dedup: node %d tag', i));
                testCase.verifyEqual(path0(got(i)), expPaths{i}, ...
                    sprintf('dedup: node %d document path', i));
            end
        end
    end
end

% ============================= file-local helpers ==========================

function b = battery()
    % Frozen call-site battery (self-contained copy). jsondecode once per test;
    % the file is small.
    persistent cached
    if isempty(cached)
        cached = jsondecode(fileread(dataPath('callsites','battery.json')));
    end
    b = cached;
end

function c = batteryCase(bat, cid)
    ids = string({bat.cases.id});
    k = find(ids == string(cid), 1);
    if isempty(k)
        error('Test:missingCase', 'battery case %s not found', cid);
    end
    c = bat.cases(k);
end

function c = realpartCase(cid)
    oracle = jsondecode(fileread(dataPath('realparts','oracle.json')));
    ids = string({oracle.cases.id});
    k = find(ids == string(cid), 1);
    if isempty(k)
        error('Test:missingCase', 'realpart case %s not found', cid);
    end
    c = oracle.cases(k);
end

function ctx = resolveCtx(testCase, root, ctxLoc, ns, cid)
    if ctxLoc == ""
        ctx = root;
    else
        arr = mat2doc.oxml.evaluate_xpath(root, ctxLoc, ns);
        testCase.assertNotEmpty(arr, ...
            sprintf('%s: context "%s" resolved empty', cid, ctxLoc));
        ctx = arr(1);
    end
end

function assertNodes(testCase, got, val, cid)
    % Value-compare a node-set result against the oracle `value` (list of
    % {tag, path}) by Clark tag + 0-based document path. Empty oracle -> count 0.
    [expTags, expPaths] = expectedNodes(val);
    testCase.verifyEqual(numel(got), numel(expTags), ...
        sprintf('%s: node count', cid));
    for i = 1:min(numel(got), numel(expTags))
        testCase.verifyEqual(string(got(i).tag), expTags(i), ...
            sprintf('%s: node %d tag', cid, i));
        testCase.verifyEqual(path0(got(i)), expPaths{i}, ...
            sprintf('%s: node %d document path', cid, i));
    end
end

function [tags, paths] = expectedNodes(val)
    % Normalise an oracle node list (struct array / cell / [] ) into a string
    % array of Clark tags and a cell of 0-based document-path row vectors.
    tags = string.empty(1, 0);
    paths = {};
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

function s = expectedStrings(val)
    % Normalise an oracle string list into a (1,N) string row.
    if isempty(val)
        s = strings(1, 0);
    elseif iscell(val)
        s = string(val(:))';
    else
        s = string(val(:))';
    end
end

function p = dataPath(varargin)
    here = fileparts(mfilename('fullpath'));   % tests\oxml
    p = fullfile(here, 'data', 'p1_3x', varargin{:});
end

function b = readFixture(p)
    % Return the raw UTF-8 bytes (the on-disk part-blob currency parse_xml
    % accepts) -- NOT char-decoded, so parse_xml's native2unicode UTF-8 path
    % handles any non-ASCII byte faithfully.
    f = fopen(p, 'r');
    if f < 0; b = uint8([]); return; end
    b = fread(f, Inf, '*uint8')';
    fclose(f);
end

function nm = localName(el)
    % Local (unprefixed) tag name from an XmlElement's Clark tag {uri}local.
    t = string(el.tag);
    nm = extractAfter(t, "}");
    if ismissing(nm) || nm == ""
        nm = t;
    end
end

function p = path0(el)
    % 0-based element-child indices from the document root down to el (matches
    % the oracle path encoding).
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

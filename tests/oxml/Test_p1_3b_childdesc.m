classdef Test_p1_3b_childdesc < matlab.unittest.TestCase
% TEST_P1_3B_CHILDDESC  Gate-4 permanent unit tests for Mat2Doc P1-3b
%   (xmlchemy engine pt2: the CHILD-element descriptor engine -- the 11 generic
%   methods on BaseOxmlElement that every future CT_* child-descriptor
%   delegating member calls: getChild, getRequiredChild, getChildList, newChild,
%   insertChildInSequence, addChild, getOrAddChild, removeChild,
%   firstChildFoundIn, removeChildren, getOrChangeToChild).
%
%   Surface ported from python-docx v1.2.0 src/docx/oxml/xmlchemy.py, realised
%   in +mat2doc\+oxml\BaseOxmlElement.m (child-descriptor block, lines 430-620).
%
%   *** NATURE: engine-probe FREEZE, not a package ladder ***
%   The child-descriptor engine emits NO standalone XML part -- its accessors are
%   consumed by CT_* classes that do not exist yet (P4+). So this class ENCODES
%   THE FROZEN GATE-3 46-CHECK ENGINE BATTERY verbatim: H11 per-step serialized
%   w:pPr-fragment byte-equivalence (L1) + identity/materialization/tri-state
%   value probes. The engine is driven by supplying the REAL CT_PPr._tag_seq
%   (parfmt.py 64-119) as the successor slices -- exactly as the Gate-3 driver
%   did (there is no CT_PPr class yet); the slice supplier is the test-only
%   tests\oxml\+p13btest\pprTagSeq.m.
%
%   Provenance (self-contained; oracle copied byte-for-byte into
%   tests\oxml\data\p1_3b\):
%     validation\mat2doc\references\p1_3b\ (manifest mso-p1_3b-oracle/1, frozen
%       2026-07-26; oracle lxml 5.3.0 / libxml2 2.13.9 under python-docx 1.2.0;
%       serialize etree.tostring(elm, encoding='UTF-8', standalone=True));
%     Gate-3 report validation\mat2doc\reports\p1_3b_validation.md (PASS 46/46,
%       0 new D-numbers);
%     port driver harness\mat2doc\validate_p1_3b.m (mirrored here).
%   The oracle bytes are lxml's own results, so byte-identical here == identical
%   to python-docx: these are equivalence AND regression pins in one.
%
%   *** FOLD-FORWARD (recorded, from the manifest) ***
%   The PACKAGE-LEVEL L1 scenario for the child-descriptor engine folds forward
%   to the FIRST CT_* WP (CT_PPr out-of-order get_or_add_* build -> a real w:pPr
%   part inside a document.xml). THIS WP's Gate-3/Gate-4 is the engine-probe
%   freeze (H11 per-step serialized-fragment byte-match + identity/materialization
%   /tri-state value probes), NOT a pkgcompare L0-L3 part ladder. No deviation is
%   exercised: every serialized fragment is L1 byte-identical (no L2, no
%   PASS-DEVIATION) -- the fragments are built then MOVED into pPr, and the
%   serializer's move-time redundant-decl suppression keeps them byte-exact.
%
%   *** VERIFY-3 (re-verify at the first CT_* WP) ***
%   The registry is STILL EMPTY at P1-3b (no CT_* registered), so every parsed /
%   built tree here is homogeneous XmlElement -- the heterogeneous-array /
%   Sealed-method risk is NOT yet exercisable and must be re-verified at the
%   first CT_* registration (carried forward from P1-3a's VERIFY-3).
%
%   *** CHOICE / ZeroOrOneChoice: SYNTHETIC parity (READ THIS) ***
%   Choice / ZeroOrOneChoice are DEAD CODE in docx v1.2.0 (0 production uses), so
%   NO real-element byte-oracle exists or can exist. The test_choice_* points are
%   engine-CONTRACT parity on a synthetic loose host (getOrChangeToChild /
%   firstChildFoundIn / removeChildren) checked against choice\oracle.json VALUES
%   only -- they are NOT byte pins. They guard the engine members the first docx
%   choice group (if one ever ships) would delegate to.
%
%   Case map (mirrors the 46 Gate-3 checks; count reported at run time):
%     * H11 ordering, 36 -- test_h11_serialized_L1 (18) byte-identical to
%         h11\sNN.oracle.xml + test_h11_localnames (18) == h11\oracle.json.
%         Plus the explicit s10==s11 get-existing-no-op sha pin.
%     * identity/materialization/tri-state, 8 -- H5 same-handle (x2), H9
%         materialized 1x0 + snapshot-safe remove-during-iter, H3 getChild
%         absent->[] / list absent->1x0 / getRequiredChild absent-> byte-exact
%         mat2doc:InvalidXmlError, materialize_tabs serialized L1.
%     * D-delta-4, 1 -- public add_x and _add_x both route through addChild.
%     * choice synthetic parity, 5 -- add-before-successor, repeat-same-handle,
%         change-to-member, firstChildFoundIn arg-order, removeChildren.
%
%   Determinism: no network, no absolute paths -- every fixture is resolved
%   relative to this file via fileparts(mfilename('fullpath')). The test-only
%   _tag_seq supplier lives at tests\oxml\+p13btest\pprTagSeq.m and is reached
%   only because TestClassSetup puts tests\oxml on the path.

    properties (TestParameter)
        % The 18 H11 step ids (h11\cases.json). One serialized-L1 test and one
        % localnames test per step; assertions read the pre-computed cumulative
        % run cached in TestClassSetup (the 18-step sequence is stateful, so it
        % is replayed ONCE and each step's snapshot frozen).
        stepId = { ...
            's01','s02','s03','s04','s05','s06','s07','s08','s09', ...
            's10','s11','s12','s13','s14','s15','s16','s17','s18'};
    end

    properties
        % Cumulative H11 run captured once in TestClassSetup:
        % struct array (1x18) with fields id, bytes (1xN uint8), names (string).
        H11
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % SECOND fixture puts tests\oxml on the path so the test-only
            % _tag_seq supplier package p13btest.pprTagSeq resolves. Both are
            % auto-restored after the class.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\oxml
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
            testCase.applyFixture(PathFixture(here));
        end

        function replayH11(testCase)
            % Replay the 18-step out-of-order build ONCE (the sequence is
            % cumulative) and freeze each step's serialized fragment + localnames.
            % Runs AFTER addWorktreeToPath (both TestClassSetup; declaration order
            % is preserved), so mat2doc.* and p13btest.* are resolvable.
            [~, S] = p13btest.pprTagSeq();
            cases = jsondecode(fileread(dataPath('h11','cases.json')));
            pPr = newPPr();
            run = struct('id', {}, 'bytes', {}, 'names', {});
            for i = 1:numel(cases)
                [kind, arg] = opParts(cases(i).op);
                switch kind
                    case "get_or_add"
                        pPr.getOrAddChild("w:" + arg, S.(char(arg)));
                    case "remove"
                        pPr.removeChild("w:" + arg);
                    case "append_foreign"
                        pPr.append(mat2doc.oxml.OxmlElement("w:" + arg));
                    otherwise
                        testCase.assertFail(sprintf('unknown op "%s"', kind));
                end
                run(i).id    = string(cases(i).id); %#ok<AGROW>
                run(i).bytes = serialize(pPr);      %#ok<AGROW>
                run(i).names = localnames(pPr);     %#ok<AGROW>
            end
            testCase.H11 = run;
        end
    end

    methods (Test)

        % =============================================== H11 ordering (36) ====
        function test_h11_serialized_L1(testCase, stepId)
            % Equivalence + regression: after this step the loose w:pPr fragment
            % serializes BYTE-IDENTICAL to the frozen lxml oracle h11\sNN.oracle.xml
            % (L1; the ladder demanded byte-identical -- no D-number granted L2).
            got  = stepBytes(testCase, stepId);
            want = readFixture(dataPath('h11', stepId + ".oracle.xml"));
            verifyBytes(testCase, got, want, stepId + " serialized-fragment L1");
        end

        function test_h11_localnames(testCase, stepId)
            % The child localname transcript after this step matches the frozen
            % oracle (h11\oracle.json .<sid>.localnames) -- pins ordering/insert
            % slot independently of serialization.
            got  = stepNames(testCase, stepId);
            want = string(h11Oracle(stepId).localnames);
            testCase.verifyEqual(got, want, ...
                sprintf('%s: localnames transcript', stepId));
        end

        function test_h11_s10_eq_s11_get_existing_noop(testCase)
            % H5 get-existing no-op sha pin: s11 (2nd get_or_add tabs, already
            % present) adds NOTHING -> byte-identical to s10 (same sha). The getter
            % path returns the live child, mutates nothing. Guards against a
            % getOrAdd that re-creates/re-inserts a present child.
            b10 = stepBytes(testCase, 's10');
            b11 = stepBytes(testCase, 's11');
            testCase.verifyEqual(sha256hex(b11), sha256hex(b10), ...
                's10==s11: get-existing get_or_add must be a byte no-op (same sha)');
            % cross-check against the frozen oracle sha (self-contained regression)
            testCase.verifyEqual(sha256hex(b11), ...
                string(h11Oracle('s11').serialized_sha256), ...
                's11: sha matches frozen oracle');
        end

        % ============================ identity / materialization / tri-state ==
        function test_H5_get_or_add_same_handle(testCase)
            % H5: get_or_add_x on a PRESENT child returns the SAME handle every
            % call (getChild -> find returns the live node, not a copy). oracle
            % H5_get_or_add_same_handle = true.
            [~, S] = p13btest.pprTagSeq();
            p = newPPr();
            a = p.getOrAddChild("w:tabs", S.tabs);
            b = p.getOrAddChild("w:tabs", S.tabs);
            testCase.verifyTrue(a == b, 'H5: repeat get_or_add returns same handle');
            testCase.verifyTrue(idOracle().H5_get_or_add_same_handle, ...
                'H5: oracle pins same-handle true');
        end

        function test_H5_add_x_returns_live_intree(testCase)
            % H5: addChild returns the LIVE in-tree node (find() locates the very
            % handle it returned). oracle H5_add_x_returns_live_intree = true.
            tabs = newLoose("w:tabs");
            added = tabs.addChild("w:tab", strings(1,0));
            found = tabs.find(mat2doc.oxml.qn("w:tab"));
            testCase.verifyTrue(found == added, ...
                'H5: add_x return handle is the live in-tree child');
        end

        function test_H9_empty_list_materialized_1x0(testCase)
            % H9: getChildList when NONE present -> a materialized typed EMPTY
            % array (1x0 XmlElement), NOT [] (None). Python x_lst returns a list;
            % "none" is an empty LIST, distinct from the ZeroOrOne None. oracle
            % H9_empty_lst_len = 0.
            tabs = newLoose("w:tabs");
            lst0 = tabs.getChildList("w:tab");
            testCase.verifyClass(lst0, 'mat2doc.oxml.XmlElement', ...
                'H9: empty list is typed XmlElement, not []');
            testCase.verifyEqual(size(lst0), [1 0], 'H9: empty list is 1x0');
            testCase.verifyEqual(double(idOracle().H9_empty_lst_len), 0);
        end

        function test_H9_snapshot_safe_remove_during_iter(testCase)
            % H9: getChildList returns a MATERIALIZED snapshot -- iterating it and
            % removing each element does not skip/alias. Three tabs -> snapshot
            % len 3 stable -> after removing all, list len 0. oracle
            % H9_three_lst_len 3 / H9_snapshot_len_stable 3 / after 0.
            tabs = newLoose("w:tabs");
            tabs.addChild("w:tab", strings(1,0));
            tabs.addChild("w:tab", strings(1,0));
            tabs.addChild("w:tab", strings(1,0));
            len3 = numel(tabs.getChildList("w:tab"));
            snap = tabs.getChildList("w:tab");
            snapLen = numel(snap);
            for k = 1:numel(snap)
                tabs.remove(snap(k));      % mutate during iteration over snapshot
            end
            lenAfter = numel(tabs.getChildList("w:tab"));
            o = idOracle();
            testCase.verifyEqual(len3, double(o.H9_three_lst_len), 'H9: three present');
            testCase.verifyEqual(snapLen, double(o.H9_snapshot_len_stable), ...
                'H9: snapshot length stable');
            testCase.verifyEqual(lenAfter, double(o.H9_after_remove_during_iter_len), ...
                'H9: all removed via snapshot iteration');
        end

        function test_H3_getChild_absent_none(testCase)
            % H3 tri-state: getChild when absent -> [] (None), a double, isequal
            % to []. Distinct from the empty LIST of getChildList. oracle
            % H3_getchild_absent_is_none = true.
            p = newPPr();
            g = p.getChild("w:ind");
            testCase.verifyTrue(isequal(g, []), 'H3: getChild absent -> [] (None)');
            testCase.verifyTrue(idOracle().H3_getchild_absent_is_none);
        end

        function test_H3_list_absent_1x0_distinct(testCase)
            % H3/H9: getChildList when absent -> typed 1x0 (empty list), DISTINCT
            % from getChild's [] (None). Guards the tri-state boundary. oracle
            % H3_list_absent_len = 0.
            p = newPPr();
            gl = p.getChildList("w:ind");
            testCase.verifyClass(gl, 'mat2doc.oxml.XmlElement', ...
                'H3: absent list is typed, not []');
            testCase.verifyEqual(size(gl), [1 0], 'H3: absent list is 1x0');
            testCase.verifyEqual(double(idOracle().H3_list_absent_len), 0);
        end

        function test_H3_getRequiredChild_absent_raises(testCase)
            % H3: getRequiredChild when absent -> mat2doc:InvalidXmlError with a
            % BYTE-EXACT message (the literal RST double-backticks reproduced
            % verbatim from xmlchemy OneAndOnlyOne). error-IDENTIFIER checked, not
            % just that it throws. oracle H3_required_absent.message.
            num = newLoose("w:num");
            err = grabError(@() num.getRequiredChild("w:abstractNumId"));
            testCase.assertNotEmpty(err, 'H3: getRequiredChild absent must raise');
            testCase.verifyEqual(err.identifier, 'mat2doc:InvalidXmlError', ...
                'H3: error identifier');
            % Regression: hard-coded expected message (== idOracle.H3_required_absent.message)
            testCase.verifyEqual(string(err.message), ...
                "required ``<w:abstractNumId>`` child element not present", ...
                'H3: byte-exact InvalidXmlError message');
            testCase.verifyEqual(string(err.message), ...
                string(idOracle().H3_required_absent.message), ...
                'H3: message matches frozen oracle');
        end

        function test_identity_materialize_tabs_L1(testCase)
            % A single get_or_add tabs on an empty pPr serializes BYTE-IDENTICAL to
            % the frozen oracle identity\materialize_tabs.oracle.xml (L1). Confirms
            % the materialization path (create + insert-in-sequence + serialize) is
            % byte-exact in isolation.
            [~, S] = p13btest.pprTagSeq();
            p = newPPr();
            p.getOrAddChild("w:tabs", S.tabs);
            verifyBytes(testCase, serialize(p), ...
                readFixture(dataPath('identity','materialize_tabs.oracle.xml')), ...
                'materialize_tabs serialized L1');
        end

        % ======================================================= D-delta-4 ====
        function test_D_delta_4_engine_neutral(testCase)
            % D-delta-4 (engine-neutral, NO D-number): the public add_x adder and
            % the private _add_x both route through the SAME addChild primitive, so
            % interleaved adds yield identical results. Three adds -> [tab,tab,tab],
            % len 3. Value probe only. oracle choice\d_delta_4.json.
            dd = ddOracle();
            tabs = newLoose("w:tabs");
            tabs.addChild("w:tab", strings(1,0));
            tabs.addChild("w:tab", strings(1,0));
            tabs.addChild("w:tab", strings(1,0));
            testCase.verifyEqual(localnames(tabs), string(dd.interleaved_localnames), ...
                'D-delta-4: interleaved add_x/_add_x localnames');
            testCase.verifyEqual(numel(tabs.getChildList("w:tab")), ...
                double(dd.interleaved_len), 'D-delta-4: interleaved length');
        end

        % ===================== choice / ZeroOrOneChoice -- SYNTHETIC parity ===
        % SYNTHETIC: dead code in docx v1.2.0, NO byte-oracle -- value probes vs
        % choice\oracle.json only. See class header CHOICE note.
        function test_choice_add_before_successor(testCase)
            % getOrChangeToChild adds the chosen member BEFORE the present successor
            % (w:s): [gB, s]. oracle add_gB_before_s.
            host = synthHost();
            host.getOrChangeToChild("w:gB", choiceGroup(), "w:s");
            testCase.verifyEqual(localnames(host), string(chOracle().add_gB_before_s), ...
                'choice: member inserted before successor (SYNTHETIC parity)');
        end

        function test_choice_repeat_same_handle(testCase)
            % H5 on a choice member: repeat getOrChangeToChild for the SAME member
            % returns the same handle. oracle repeat_same_handle_is = true.
            host = synthHost();
            b1 = host.getOrChangeToChild("w:gB", choiceGroup(), "w:s");
            b2 = host.getOrChangeToChild("w:gB", choiceGroup(), "w:s");
            testCase.verifyEqual(b1 == b2, logical(chOracle().repeat_same_handle_is), ...
                'choice: repeat same member -> same handle (SYNTHETIC parity)');
        end

        function test_choice_change_to_member(testCase)
            % getOrChangeToChild to a DIFFERENT member removes the prior member and
            % inserts the new one before the successor: [gC, s]. oracle change_to_gC.
            host = synthHost();
            host.getOrChangeToChild("w:gB", choiceGroup(), "w:s");
            host.getOrChangeToChild("w:gC", choiceGroup(), "w:s");
            testCase.verifyEqual(localnames(host), string(chOracle().change_to_gC), ...
                'choice: change-to swaps the group member (SYNTHETIC parity)');
        end

        function test_choice_firstChildFoundIn_argorder(testCase)
            % firstChildFoundIn searches in ARGUMENT order (first NAME wins, not
            % document order): args [gC,gA] on a host holding gA then gC -> gC.
            % oracle firstfound_argorder = "gC".
            host = newLoose("w:x");
            host.append(mat2doc.oxml.OxmlElement("w:gA"));
            host.append(mat2doc.oxml.OxmlElement("w:gC"));
            f = host.firstChildFoundIn(["w:gC","w:gA"]);
            testCase.verifyEqual(localOf(f.tag), string(chOracle().firstfound_argorder), ...
                'choice: firstChildFoundIn arg-order wins (SYNTHETIC parity)');
        end

        function test_choice_removeChildren(testCase)
            % removeChildren strips every present group member -> empty host [].
            % oracle removeChildren_result = "[]".
            host = newLoose("w:x");
            host.append(mat2doc.oxml.OxmlElement("w:gA"));
            host.append(mat2doc.oxml.OxmlElement("w:gC"));
            host.removeChildren(choiceGroup());
            testCase.verifyEqual(localnames(host), string(chOracle().removeChildren_result), ...
                'choice: removeChildren strips the whole group (SYNTHETIC parity)');
        end
    end
end

% ============================= file-local helpers ==========================

function u = wuri()
    u = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
end

function el = newLoose(tag)
    % Loose BaseOxmlElement host carrying the w prefix binding (registry has no
    % CT_* yet; production hosts are registered CT_* subclasses, as in Python).
    el = mat2doc.oxml.BaseOxmlElement(tag, struct('w', wuri()));
end

function p = newPPr()
    p = newLoose("w:pPr");
end

function h = synthHost()
    % Synthetic choice host preloaded with the successor w:s (dead code in docx).
    h = newLoose("w:x");
    h.append(mat2doc.oxml.OxmlElement("w:s"));
end

function g = choiceGroup()
    g = ["w:gA","w:gB","w:gC"];
end

function b = serialize(elm)
    b = uint8(mat2doc.oxml.serialize_part_xml(elm));
    b = b(:)';
end

function s = localnames(el)
    kids = el.to_array();
    names = strings(1, numel(kids));
    for i = 1:numel(kids)
        names(i) = localOf(string(kids(i).tag));
    end
    s = "[" + strjoin(names, ", ") + "]";
end

function t = localOf(tagStr)
    t = regexprep(string(tagStr), "^\{.*\}", "");   % Clark -> local
    t = regexprep(t, "^.*:", "");                   % prefixed -> local
end

function [kind, arg] = opParts(op)
    % cases.json .op is a 2-element JSON array; jsondecode yields a 2x1 cell of
    % char (elements differ in length). Coerce robustly either way.
    if iscell(op)
        kind = string(op{1}); arg = string(op{2});
    else
        kind = string(op(1)); arg = string(op(2));
    end
end

function b = stepBytes(testCase, sid)
    b = stepField(testCase, sid).bytes;
end

function n = stepNames(testCase, sid)
    n = stepField(testCase, sid).names;
end

function s = stepField(testCase, sid)
    ids = string({testCase.H11.id});
    k = find(ids == string(sid), 1);
    testCase.assertNotEmpty(k, sprintf('H11 step %s not captured', sid));
    s = testCase.H11(k);
end

function p = dataPath(varargin)
    here = fileparts(mfilename('fullpath'));   % tests\oxml
    parts = cellfun(@char, varargin, 'UniformOutput', false);
    p = fullfile(here, 'data', 'p1_3b', parts{:});
end

function b = readFixture(p)
    f = fopen(char(p), 'r');
    if f < 0; b = uint8([]); return; end
    b = fread(f, Inf, '*uint8')';
    fclose(f);
end

function e = h11Oracle(sid)
    o = jsondecode(fileread(dataPath('h11','oracle.json')));
    e = o.(char(sid));
end

function o = idOracle()
    o = jsondecode(fileread(dataPath('identity','oracle.json')));
end

function o = chOracle()
    o = jsondecode(fileread(dataPath('choice','oracle.json')));
end

function o = ddOracle()
    o = jsondecode(fileread(dataPath('choice','d_delta_4.json')));
end

function err = grabError(fn)
    err = [];
    try
        fn();
    catch e
        err = e;
    end
end

function h = sha256hex(bytes)
    md = java.security.MessageDigest.getInstance('SHA-256');
    dig = typecast(md.digest(uint8(bytes(:))), 'uint8');
    h = string(lower(reshape(dec2hex(dig, 2).', 1, [])));
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

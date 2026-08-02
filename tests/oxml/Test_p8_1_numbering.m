classdef Test_p8_1_numbering < matlab.unittest.TestCase
% TEST_P8_1_NUMBERING  Gate-4 permanent unit tests for Mat2Doc P8-1 [N] -- the
%   oxml/numbering element tier + NumberingPart finalize.
%   src/docx/oxml/numbering.py -> +mat2doc\+oxml\+numbering\{CT_Num, CT_NumLvl,
%   CT_NumPr, CT_Numbering}; src/docx/parts/numbering.py ->
%   +mat2doc\+parts\{NumberingPart (numbering_definitions un-stubbed),
%   NumberingDefinitions_ (NEW)}; src/docx/parts/document.py::DocumentPart
%   .numbering_part (un-stubbed @lazyproperty); + the 8 numbering registry rows
%   (w:abstractNumId/w:ilvl/w:numId/w:startOverride -> CT_DecimalNumber [REUSED,
%   P4-6]; w:lvlOverride -> CT_NumLvl; w:num -> CT_Num; w:numPr -> CT_NumPr;
%   w:numbering -> CT_Numbering). Completes the numbering oxml layer + registry.
%
%   P8-1 is a REGISTRY-ADDING, byte-critical parse-path WP: registering the 8
%   numbering tags flips the PARSE CLASS of the <w:numbering> root + 9 <w:num> in
%   numbering.xml and the 7 <w:numPr> (6 <w:numId> + 1 <w:ilvl>) in styles.xml
%   from generic XmlElement onto the new CT_* classes. All CT_* exit through the
%   identical +oxml\serialize_part_xml walk, so the flip is byte-neutral --
%   word/numbering.xml stays 5513 B / 70976f19...; word/styles.xml stays 349458 B
%   / 02d71a68... (the styles.xml pin is owned by Test_p4_6 / Test_p1_8; this
%   class adds the numbering.xml-specific SHA pin the brief calls for).
%
%   This class permanently freezes what the prior gates established:
%     * Gate-1 Porter  : audit_P8-1_numbering.md (self-probe 34/34; C1 17/17 L1).
%     * Gate-2 Auditor (Fable): APPROVE -- 34/34 independent probes, ZERO new
%       D-numbers; H11 successor slices re-derived EXACT; VERIFY-1/2/3 approved.
%     * Gate-3 Validator: validate_P8-1_numbering.md -- PASS, ZERO new D-numbers.
%       C1 M1 17/17 byte-identical (numbering.xml 5513 B / 70976f19...). Froze the
%       production-serializer byte oracle references\s0099\ (4 mutation cases) and
%       the API-value probe references\s0100\ (probe.json, 17 tagged facts,
%       probe_diff exit 0). Targeted regression 55 total, 52 pass, 3 EXPECTED
%       byte-neutral stale-pin flips (re-pinned at Gate-4: Test_p4_6:598,
%       Test_p2_2 stub-count, Test_p1_8:411).
%
%   THE HIGHEST-VALUE PERMANENT PINS IN THIS WP (a regression goes RED loudly):
%     * ★ (s0099 BYTE BATTERY) test_s0099_serialize_byte_battery -- THE numbering
%       mutation-serialize FOUNDATION. For each of the 4 frozen s0099 cases the
%       EXACT mutation-API sequence (add_num auto-numId / CT_Num.new /
%       add_lvlOverride / add_startOverride / CT_NumPr get_or_add ordering / the
%       <w:numIdMacAtCleanup> + <w:ins> successor boundaries) is replayed through
%       the port, serialized via the PRODUCTION serialize_part_xml, and asserted
%       BYTE-IDENTICAL (raw uint8 AND SHA-256) to the frozen python-docx oracle.
%       A single perturbed byte of the numbering serializer goes RED here.
%     * ★ (H11 ORDERING) test_ct_numpr_h11_successor_ordering -- numId added
%       FIRST then ilvl -> <w:ilvl> serializes BEFORE <w:numId> (the successor
%       slice re-sort). Wrong order -> Word repair. Byte-pinned.
%     * ★ (M1) test_m1_numbering_xml_byte_identical -- mat2doc.Document().save()
%       -> word/numbering.xml == 5513 B, SHA-256 70976f19... The registry-adding
%       parse-path byte-neutrality guard (registering a CT changes only a parsed
%       node's CLASS, never content/order). L1.
%     * (GAP-FILL) test_next_numId_gapfill -- _next_numId value arithmetic:
%       {1,3}->2, {}->1, {1,2}->3, and the duplicate {1,1}->2 -- DATA arithmetic
%       on numId VALUES, NOT a 0/1 index shift (H1).
%     * (KeyError) test_num_having_numId_miss -- miss -> mat2doc:KeyError with the
%       RAW message "no <w:num> element with numId 99" (VERIFY-2: Python's
%       str(KeyError) repr-quotes are __str__ decoration, NOT the constructed
%       message; the id + raw message is the faithful contract).
%     * (s0100 EQUIVALENCE) test_s0100_api_probe_vs_frozen_oracle -- replays the
%       full 16-fact API probe and diffs each tagged value to the frozen oracle.
%
%   Provenance (all Gate-3 frozen 2026-08-02):
%     * Audit    : validation\mat2doc\audit_P8-1_numbering.md
%     * Validate : validation\mat2doc\validate_P8-1_numbering.md
%     * Scenarios: validation\mat2doc\scenarios\s0099_p8_1_numbering_serialize.{py,m}
%                  (the 4-case production-serialize byte twin; runS0099 below
%                  replays its mutation body VERBATIM),
%                  s0100_p8_1_numbering_api.{py,m} (the API-value probe; runS0100
%                  below replays its probe body VERBATIM).
%     * Frozen refs (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE at Gate-3):
%         references\s0099\ (4 case*.xml + manifest_bytes.json) copied verbatim
%           into tests\oxml\data\s0099\ WITH a co-located `.gitattributes`
%           `* binary` pin (frozen-byte fixtures must not be line-ending mangled
%           on the master checkout -- the Gate-4 byte-fixture lesson).
%         references\s0100\ (probe.json) copied into tests\oxml\data\s0100\ WITH
%           the same `* binary` pin (jsondecode is line-ending agnostic, so the
%           binary pin is belt-and-suspenders here).
%         references\s0001\parts\word\numbering.xml -- the M1 byte reference (SHA
%           of what Document().save() itself emits); NOT copied.
%
%   Coverage taxonomy
%   -----------------
%   * Nominal      -- 8 registry rows resolve; add_num auto-numId 1,2,3;
%                     num_having_numId hit + H5 identity; CT_Num.new attrs;
%                     add_lvlOverride/add_startOverride; NumberingDefinitions_ len;
%                     DocumentPart.numbering_part on default.docx (NumberingPart,
%                     len 9) + lazyproperty identity.
%   * Edge         -- _next_numId gap-fill incl. the duplicate {1,1}->2; the
%                     num_having_numId(99) KeyError path (id + raw message);
%                     RequiredAttribute missing -> mat2doc:InvalidXmlError; the
%                     <w:numIdMacAtCleanup> + <w:ins> successor-boundary inserts
%                     (byte); NumberingPart.new -> mat2doc:NotImplementedError
%                     (faithful upstream raise); w:abstractNum stays GENERIC
%                     (deliberately unregistered upstream).
%   * Equivalence  -- test_s0099_serialize_byte_battery (4 production-serialized
%                     trees byte-identical to the frozen oracle) +
%                     test_s0100_api_probe_vs_frozen_oracle (16 tagged API facts
%                     value-identical to the frozen probe.json; Gate-3 probe_diff
%                     was exit 0).
%   * Regression   -- hard-coded expected serialized-XML strings (ASCII ==
%                     byte-identical L1) + the 4 s0099 SHA-256 pins + the M1
%                     numbering.xml SHA-256 pin.
%   * Upstream     -- the s0099 case*.xml ARE python-docx production serialize_part_xml
%                     output (full Word nsmap, single-quote XML decl); the H11
%                     successor slices are ported VERBATIM from numbering.py.
%
%   Byte-level (L1) note: every s0099 comparison is BOTH raw uint8 byte-equality
%   AND SHA-256 equality vs the frozen oracle -- the ladder demanded L1 and Gate-3
%   delivered 4/4 byte-identical with ZERO new D-numbers, so every s0099/M1 pin
%   here is L1. The only looser-than-byte assertion is the s0100 Equivalence
%   value-compare (tagged strings, not bytes) -- justified because probe_diff
%   already proved value equivalence at Gate-3 (exit 0). The mat2doc:KeyError /
%   mat2doc:InvalidXmlError / mat2doc:NotImplementedError identifiers equal the
%   Python exception-class names and the messages are the verbatim constructed
%   text (the signed exception-model mapping, design.md section 2 -- non-byte,
%   non-output, NOT a D-number).
%
%   Determinism: no network, no absolute paths. The worktree root and the
%   co-located s0099/s0100 fixtures resolve relative to this file via
%   fileparts(mfilename('fullpath')); the M1 save goes to a tempname .docx deleted
%   via onCleanup; every fixture read is binary ('r','n'). The +mat2doc package
%   resolves via the MANDATORY PathFixture(worktree-root) in TestClassSetup
%   (WP9-F4 lesson).

    properties (Constant)
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        % --- registered class names (the 8 P8-1 rows) ---
        CT_NUM       = 'mat2doc.oxml.numbering.CT_Num'
        CT_NUMLVL    = 'mat2doc.oxml.numbering.CT_NumLvl'
        CT_NUMPR     = 'mat2doc.oxml.numbering.CT_NumPr'
        CT_NUMBERING = 'mat2doc.oxml.numbering.CT_Numbering'
        CT_DECNUM    = 'mat2doc.oxml.shared.CT_DecimalNumber'
        GENERIC      = 'mat2doc.oxml.XmlElement'

        % --- frozen s0001 M1 numbering.xml byte reference (registry flip neutral) ---
        NUMBERING_SIZE_M1 = 5513
        NUMBERING_SHA_M1  = "70976f19cbcd896e51890859fe6ecb3467a5a7ad0c040160fedc5a1993cb09ce"
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied verbatim from tests\oxml\Test_p6_3b_ct_tbl.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\oxml
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. the 8 registry rows (registration is what flips the class)    %
        % =============================================================== %

        function test_registry_resolves_eight_rows(testCase)
            % Nominal / Regression: REGISTRATION flips the parse class of a
            % <w:numbering> subtree. The 4 int-attr tags REUSE CT_DecimalNumber
            % (P4-6, NOT re-ported); w:abstractNum is DELIBERATELY unregistered
            % upstream and must stay GENERIC (grep 0 register rows; the
            % numbering.xml <w:abstractNum> round-trips byte-identical as generic).
            pairs = { ...
                "w:abstractNumId", testCase.CT_DECNUM; ...
                "w:ilvl",          testCase.CT_DECNUM; ...
                "w:lvlOverride",   testCase.CT_NUMLVL; ...
                "w:num",           testCase.CT_NUM; ...
                "w:numId",         testCase.CT_DECNUM; ...
                "w:numPr",         testCase.CT_NUMPR; ...
                "w:numbering",     testCase.CT_NUMBERING; ...
                "w:startOverride", testCase.CT_DECNUM };
            for i = 1:size(pairs, 1)
                tag = pairs{i, 1}; cls = pairs{i, 2};
                r = mat2doc.oxml.registry(mat2doc.oxml.qn(tag));
                testCase.verifyEqual(char(r), cls, ...
                    sprintf('registry must resolve %s -> %s (P8-1 row)', tag, cls));
                testCase.verifyEqual(class(mat2doc.oxml.OxmlElement(tag)), cls, ...
                    sprintf('OxmlElement(%s) must be a %s', tag, cls));
            end
            % w:abstractNum stays GENERIC (upstream has no register row for it)
            testCase.verifyEqual(class(mat2doc.oxml.OxmlElement("w:abstractNum")), ...
                testCase.GENERIC, ...
                'w:abstractNum stays generic XmlElement (deliberately unregistered upstream)');
        end

        % =============================================================== %
        % 2. CT_Numbering.add_num -- auto numId 1,2,3                       %
        % =============================================================== %

        function test_add_num_auto_numid(testCase)
            % Nominal: add_num allocates the next free numId (1,2,3...) and returns
            % a CT_Num referencing the given abstractNum_id. The suppressed
            % generated public add_num is replaced by the explicit
            % add_num(abstractNum_id); the returned element is a real CT_Num.
            numbering = mat2doc.oxml.OxmlElement("w:numbering");
            a1 = numbering.add_num(0);
            a2 = numbering.add_num(1);
            a3 = numbering.add_num(2);
            testCase.verifyClass(a1, testCase.CT_NUM, 'add_num returns a CT_Num');
            testCase.verifyEqual(a1.numId, 1, 'first add_num -> numId 1');
            testCase.verifyEqual(a2.numId, 2, 'second add_num -> numId 2');
            testCase.verifyEqual(a3.numId, 3, 'third add_num -> numId 3');
            testCase.verifyEqual(numel(numbering.num_lst), 3, '3 <w:num> children');
            % abstractNumId child carries the passed value
            testCase.verifyEqual(a2.abstractNumId.val, 1, 'add_num(1) -> abstractNumId val 1');
        end

        % =============================================================== %
        % 3. num_having_numId -- hit + H5 identity + KeyError miss          %
        % =============================================================== %

        function test_num_having_numId_hit_and_identity(testCase)
            % Nominal + H5: the <w:num> whose @w:numId matches is returned as the
            % SAME handle already in num_lst (Python xpath()[0] is the live element,
            % not a copy). H1: first match.
            numbering = mat2doc.oxml.OxmlElement("w:numbering");
            numbering.add_num(0);       % numId 1
            numbering.add_num(0);       % numId 2
            hit = numbering.num_having_numId(2);
            testCase.verifyEqual(hit.numId, 2, 'num_having_numId(2) -> the numId==2 <w:num>');
            nl = numbering.num_lst;
            testCase.verifyTrue(hit == nl(2), ...
                'H5: num_having_numId returns the SAME handle as num_lst(2)');
        end

        function test_num_having_numId_miss(testCase)
            % Edge / error path (VERIFY-2): a miss raises mat2doc:KeyError with the
            % RAW constructed message. Python raises KeyError("no <w:num> element
            % with numId %d" % numId); str(KeyError) adds repr single-quotes, but
            % those are __str__ decoration -- the id + RAW message is the faithful
            % contract (consistent with existing mat2doc:KeyError sites). Assert
            % BOTH the identifier AND the exact raw message text.
            numbering = mat2doc.oxml.OxmlElement("w:numbering");
            numbering.add_num(0);       % numId 1 only
            fn = @() numbering.num_having_numId(99);
            testCase.verifyError(fn, 'mat2doc:KeyError', ...
                'num_having_numId miss must raise mat2doc:KeyError');
            testCase.verifyEqual(errmsg(fn), "no <w:num> element with numId 99", ...
                'KeyError raw message verbatim (no repr single-quotes)');
        end

        % =============================================================== %
        % 4. _next_numId -- gap-fill VALUE arithmetic (H1)                  %
        % =============================================================== %

        function test_next_numId_gapfill(testCase)
            % Edge (H1): _next_numId is DATA arithmetic on numId VALUES (fill the
            % first gap, else max+1), NOT a 0/1 index shift. Seed numberings via
            % CT_Num.new (so the numId set is exactly the seeded values) and drive
            % the next allocation through add_num, whose numId == _next_numId.
            %   {1,3} -> 2   (fills the gap)
            %   {}    -> 1   (empty starts at 1)
            %   {1,2} -> 3   (no gap -> max+1)
            %   {1,1} -> 2   (duplicate collapses; 1 is used, 2 is free)
            CT_Num = "mat2doc.oxml.numbering.CT_Num";

            nb = mat2doc.oxml.OxmlElement("w:numbering");
            nb.append(feval(CT_Num + ".new", 1, 0));
            nb.append(feval(CT_Num + ".new", 3, 0));
            testCase.verifyEqual(nb.next_numId_, 2, '{1,3} -> next_numId 2 (gap-fill)');
            testCase.verifyEqual(nb.add_num(0).numId, 2, '{1,3} add_num -> numId 2');

            nb = mat2doc.oxml.OxmlElement("w:numbering");
            testCase.verifyEqual(nb.next_numId_, 1, '{} -> next_numId 1 (empty)');
            testCase.verifyEqual(nb.add_num(0).numId, 1, '{} add_num -> numId 1');

            nb = mat2doc.oxml.OxmlElement("w:numbering");
            nb.append(feval(CT_Num + ".new", 1, 0));
            nb.append(feval(CT_Num + ".new", 2, 0));
            testCase.verifyEqual(nb.next_numId_, 3, '{1,2} -> next_numId 3 (max+1, no gap)');

            nb = mat2doc.oxml.OxmlElement("w:numbering");
            nb.append(feval(CT_Num + ".new", 1, 0));
            nb.append(feval(CT_Num + ".new", 1, 0));
            testCase.verifyEqual(nb.next_numId_, 2, '{1,1} dup -> next_numId 2 (1 used, 2 free)');
        end

        % =============================================================== %
        % 5. CT_Num.new + add_lvlOverride + RequiredAttribute round-trip    %
        % =============================================================== %

        function test_ct_num_new_and_required_attr(testCase)
            % Nominal + RequiredAttribute (ST_DecimalNumber): CT_Num.new(num_id,
            % abstractNum_id) sets @w:numId and appends a <w:abstractNumId w:val=..>
            % child (a CT_DecimalNumber). numId round-trips through
            % ST_DecimalNumber.to_xml -> pyStr int. Also the OneAndOnlyOne
            % abstractNumId getter + a bare <w:num> RequiredAttribute raise.
            n = mat2doc.oxml.numbering.CT_Num.new(12, 42);
            testCase.verifyClass(n, testCase.CT_NUM, 'CT_Num.new returns a CT_Num');
            testCase.verifyEqual(n.numId, 12, 'CT_Num.new num_id -> @w:numId 12');
            testCase.verifyClass(n.abstractNumId, testCase.CT_DECNUM, ...
                'abstractNumId child is a CT_DecimalNumber (OneAndOnlyOne getter)');
            testCase.verifyEqual(n.abstractNumId.val, 42, 'abstractNumId val 42');
            % byte shape (ASCII == L1): @w:numId serializes via ST_DecimalNumber.to_xml.
            % serialize_part_xml (production path) prepends the XML declaration.
            testCase.verifyEqual(ser(n), ...
                xmlDecl() + ...
                "<w:num xmlns:w=""" + testCase.W + """ w:numId=""12"">" + ...
                "<w:abstractNumId w:val=""42""/></w:num>", ...
                'CT_Num.new(12,42) serialized byte shape (RequiredAttribute round-trip)');

            % RequiredAttribute raise: a bare <w:num> has no @w:numId
            bare = mat2doc.oxml.OxmlElement("w:num");
            testCase.verifyError(@() bare.numId, 'mat2doc:InvalidXmlError', ...
                'missing required @w:numId raises mat2doc:InvalidXmlError');

            % ST_DecimalNumber round-trip (the int() analogue, VERIFY-1): from_xml
            % parses the XML integer literal exactly as Python int(); to_xml formats
            % via pyStr int.
            ST = "mat2doc.oxml.simpletypes.ST_DecimalNumber";
            testCase.verifyEqual(feval(ST + ".from_xml", "007"), 7, 'from_xml "007" -> 7');
            testCase.verifyEqual(feval(ST + ".from_xml", "-5"), -5, 'from_xml "-5" -> -5');
            testCase.verifyEqual(feval(ST + ".to_xml", 3), "3", 'to_xml 3 -> "3"');
        end

        function test_ct_num_add_lvlOverride(testCase)
            % Nominal (CT_Num.add_lvlOverride): the explicit adder (which SUPPRESSES
            % the generated public add_lvlOverride) returns a CT_NumLvl with @w:ilvl
            % set. successors=() -> the <w:lvlOverride> APPENDS.
            n = mat2doc.oxml.numbering.CT_Num.new(1, 0);
            lo = n.add_lvlOverride(0);
            testCase.verifyClass(lo, testCase.CT_NUMLVL, 'add_lvlOverride returns a CT_NumLvl');
            testCase.verifyEqual(lo.ilvl, 0, 'add_lvlOverride(0) -> @w:ilvl 0');
            testCase.verifyEqual(numel(n.lvlOverride_lst), 1, 'one <w:lvlOverride> appended');
        end

        function test_ct_numlvl_add_startOverride(testCase)
            % Nominal (CT_NumLvl.add_startOverride): returns a CT_DecimalNumber
            % <w:startOverride w:val=..>; successors=("w:lvl",) -> the startOverride
            % sorts before any <w:lvl>. Assert the byte shape inside the lvlOverride.
            n = mat2doc.oxml.numbering.CT_Num.new(1, 0);
            lo = n.add_lvlOverride(0);
            so = lo.add_startOverride(5);
            testCase.verifyClass(so, testCase.CT_DECNUM, ...
                'add_startOverride returns a CT_DecimalNumber');
            testCase.verifyEqual(so.val, 5, 'add_startOverride(5) -> @w:val 5');
            testCase.verifyTrue(contains(ser(lo), "<w:startOverride w:val=""5""/>"), ...
                '<w:startOverride w:val="5"/> serialized inside the lvlOverride');
        end

        % =============================================================== %
        % 6. ★ CT_NumPr H11 successor ordering (byte-critical)              %
        % =============================================================== %

        function test_ct_numpr_h11_successor_ordering(testCase)
            % ★ (H11) Nominal + byte: get_or_add_numId() added FIRST, then
            % get_or_add_ilvl() -- because ilvl's successor slice includes "w:numId",
            % the <w:ilvl> RE-SORTS to serialize BEFORE <w:numId> despite being
            % added second. This is the exact ordering python-docx emits (xml_numPr
            % match). A wrong successor slice -> Word repair.
            numPr = mat2doc.oxml.OxmlElement("w:numPr");
            testCase.verifyClass(numPr, testCase.CT_NUMPR, 'w:numPr -> CT_NumPr');
            numId = numPr.get_or_add_numId();   % added FIRST
            numId.val = 2;
            ilvl = numPr.get_or_add_ilvl();     % added SECOND, sorts before numId
            ilvl.val = 0;
            testCase.verifyEqual(ser(numPr), ...
                xmlDecl() + ...
                "<w:numPr xmlns:w=""" + testCase.W + """>" + ...
                "<w:ilvl w:val=""0""/><w:numId w:val=""2""/></w:numPr>", ...
                'H11: <w:ilvl> serializes BEFORE <w:numId> (numId added first, successor re-sort)');
            % child parse classes are CT_DecimalNumber (registered this WP)
            testCase.verifyClass(numPr.ilvl, testCase.CT_DECNUM, 'w:ilvl -> CT_DecimalNumber');
            testCase.verifyClass(numPr.numId, testCase.CT_DECNUM, 'w:numId -> CT_DecimalNumber');
        end

        % =============================================================== %
        % 7. NumberingPart.new -- faithful NotImplementedError              %
        % =============================================================== %

        function test_numberingpart_new_notimplemented(testCase)
            % Edge (faithful upstream raise, §7): python-docx v1.2.0 itself declares
            % `def new(cls): raise NotImplementedError` (numbering.py 11-14).
            % NumberingPart.new() reproduces this as mat2doc:NotImplementedError --
            % the ONE notYetPorted-shaped site that is faithful upstream behavior,
            % NOT a port stub. The identifier distinguishes it from notYetPorted.
            testCase.verifyError(@() mat2doc.parts.NumberingPart.new(), ...
                'mat2doc:NotImplementedError', ...
                'NumberingPart.new raises mat2doc:NotImplementedError (faithful upstream)');
        end

        % =============================================================== %
        % 8. NumberingDefinitions_ + numbering_part on default.docx         %
        % =============================================================== %

        function test_numberingdefinitions_len(testCase)
            % Nominal (_NumberingDefinitions.__len__ -> len_()): len == number of
            % <w:num> children of the wrapped <w:numbering>.
            nb = mat2doc.oxml.OxmlElement("w:numbering");
            nb.add_num(0); nb.add_num(0); nb.add_num(0);
            nd = mat2doc.parts.NumberingDefinitions_(nb);
            testCase.verifyEqual(nd.len_(), 3, 'NumberingDefinitions_ len == 3 <w:num>');
        end

        function test_documentpart_numbering_part_default_and_identity(testCase)
            % Nominal + lazyproperty identity: default.docx SHIPS a numbering part,
            % so DocumentPart.numbering_part returns it via part_related_by(NUMBERING)
            % (never reaching the faithful new() raise). numbering_definitions len ==
            % 9 (default.docx's <w:num> count). @lazyproperty: two reads return the
            % SAME handle (both numbering_part and numbering_definitions cached).
            d  = mat2doc.Document();
            dp = d.part;
            np = dp.numbering_part;
            testCase.verifyClass(np, 'mat2doc.parts.NumberingPart', ...
                'DocumentPart.numbering_part -> NumberingPart on default.docx');
            nd = np.numbering_definitions;
            testCase.verifyEqual(nd.len_(), 9, ...
                'default.docx numbering_definitions len == 9 <w:num>');
            % lazyproperty identity (Python `is`): two reads -> same handle
            testCase.verifyTrue(dp.numbering_part == np, ...
                'numbering_part lazyproperty caches (two reads, same handle)');
            testCase.verifyTrue(np.numbering_definitions == nd, ...
                'numbering_definitions lazyproperty caches (two reads, same handle)');
        end

        % =============================================================== %
        % 9. ★ EQUIVALENCE -- s0099 production-serialize byte battery        %
        % =============================================================== %

        function test_s0099_serialize_byte_battery(testCase)
            % ★ (s0099) Equivalence + Regression + Upstream (byte-identical L1): for
            % each of the 4 frozen s0099 cases the EXACT mutation-API sequence
            % (runS0099, the .m twin's body VERBATIM) is replayed through the port,
            % serialized via the PRODUCTION serialize_part_xml, and asserted
            % BYTE-IDENTICAL (raw uint8 AND SHA-256) to the frozen python-docx
            % oracle part shipped in data\s0099\. Also cross-checks the shipped
            % fixture is intact (size + SHA vs manifest). Gate-3 froze 4/4
            % byte-identical with ZERO new D-numbers.
            here = fileparts(mfilename('fullpath'));
            man  = jsondecode(native2unicode( ...
                readBytes(fullfile(here, 'data', 's0099', 'manifest_bytes.json')), 'UTF-8'));
            cand = runS0099();   % struct: caseA_build/caseB_cleanup/caseC_numpr_order/caseD_numpr_ins -> uint8

            testCase.verifyEqual(numel(man.artifacts), 4, 's0099 manifest has 4 artifacts');
            for i = 1:numel(man.artifacts)
                a    = man.artifacts(i);
                stem = erase(string(a.file), ".xml");   % caseA_build, ...

                % -- the shipped frozen fixture is intact (size + SHA) --
                frozen = readBytes(fullfile(here, 'data', 's0099', char(a.file)));
                testCase.verifyEqual(numel(frozen), a.size, ...
                    sprintf('%s frozen fixture size intact (%d B)', a.file, a.size));
                testCase.verifyEqual(sha256hex(frozen), string(a.sha256), ...
                    sprintf('%s frozen fixture SHA-256 intact', a.file));

                % -- the replayed candidate is byte-identical to the oracle --
                out = cand.(stem);
                testCase.verifyEqual(numel(out), a.size, ...
                    sprintf('%s replay serialized size == frozen oracle (%d B)', a.file, a.size));
                testCase.verifyEqual(uint8(out(:)'), uint8(frozen(:)'), ...
                    sprintf('%s mutation-API replay -> serialize must be BYTE-IDENTICAL to the frozen s0099 oracle', a.file));
                testCase.verifyEqual(sha256hex(out), string(a.sha256), ...
                    sprintf('%s replay SHA-256 == frozen manifest (L1)', a.file));
            end
        end

        % =============================================================== %
        % 10. EQUIVALENCE -- s0100 API-value probe vs the frozen oracle      %
        % =============================================================== %

        function test_s0100_api_probe_vs_frozen_oracle(testCase)
            % Equivalence (values, not bytes): replay the ENTIRE s0100 API probe
            % (runS0100 -- the .m twin's body VERBATIM: add_num auto-numId,
            % _next_numId gap-fill, num_having_numId hit/identity/miss, CT_Num.new
            % attrs, NumberingDefinitions_ len, DocumentPart.numbering_part on
            % default.docx, NumberingPart.new raise) and compare EACH tagged fact to
            % the frozen python-docx 1.2.0 oracle copied into data\s0100\probe.json.
            % Gate-3 probe_diff was exit 0, so every fact must be value-identical.
            % Looser-than-byte (tagged-string values); justified by the Gate-3
            % probe_diff exit-0 proof.
            here   = fileparts(mfilename('fullpath'));
            port   = runS0100();
            oracle = jsondecode(native2unicode( ...
                readBytes(fullfile(here, 'data', 's0100', 'probe.json')), 'UTF-8'));

            keys = fieldnames(oracle);
            % Non-triviality floor (guards a silent-empty replay).
            testCase.verifyGreaterThanOrEqual(numel(keys), 16, ...
                'the s0100 oracle must expose all 16 tagged facts');
            testCase.verifyEqual(sort(fieldnames(port)), sort(keys), ...
                'the replayed probe and the frozen oracle expose the same keys');
            for i = 1:numel(keys)
                k = keys{i};
                testCase.verifyEqual(string(port.(k)), string(oracle.(k)), ...
                    sprintf('s0100 fact "%s" must be value-identical to the frozen oracle', k));
            end
        end

        % =============================================================== %
        % 11. ★ M1 numbering.xml byte-pin (neutrality guard, C1)             %
        % =============================================================== %

        function test_m1_numbering_xml_byte_identical(testCase)
            % ★ (M1 / C1) Regression (byte-neutrality, L1): mat2doc.Document().save()
            % emits word/numbering.xml at EXACTLY 5513 B with the frozen s0001
            % SHA-256 70976f19... -- byte-identical DESPITE the <w:numbering> root +
            % 9 <w:num> now parsing onto CT_Numbering/CT_Num and the 9
            % <w:abstractNum> staying GENERIC (upstream unregisters w:abstractNum).
            % Registering a CT changes only a parsed node's CLASS, never
            % content/attr-order/child-order (the P4-6/P6-2 precedent). This is the
            % numbering.xml-specific SHA pin the brief calls for (styles.xml is owned
            % by Test_p4_6/Test_p1_8). SHA == L1.
            bytes = emitDocPart('numbering.xml');
            testCase.verifyEqual(numel(bytes), testCase.NUMBERING_SIZE_M1, ...
                sprintf('word/numbering.xml must be exactly %d B after the numbering registry rows', ...
                    testCase.NUMBERING_SIZE_M1));
            testCase.verifyEqual(sha256hex(bytes), testCase.NUMBERING_SHA_M1, ...
                'word/numbering.xml SHA-256 == frozen s0001 oracle (registry flip byte-neutral, L1)');
        end

    end
end

% ===================== file-local helpers ============================== %

function s = ser(e)
    % Decode serialize_part_xml bytes to a string (ASCII content -> string-equality
    % is a byte-identical L1 assertion; non-ASCII round-trips via UTF-8).
    s = string(native2unicode(mat2doc.oxml.serialize_part_xml(e), "UTF-8"));
end

function s = xmlDecl()
    % The production serialize_part_xml XML declaration prefix (single-quote,
    % standalone='yes', trailing newline) -- the lxml
    % etree.tostring(encoding="UTF-8", standalone=True) analogue.
    s = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>" + newline;
end

function s = errmsg(fn)
    try
        fn();
        s = "NOERR";
    catch e
        s = string(e.message);
    end
end

function b = readBytes(p)
    f = fopen(p, 'r', 'n');            % binary read (no CRLF translation)
    assert(f >= 0, 'could not open for read: %s', p);
    b = fread(f, Inf, '*uint8')';
    fclose(f);
end

function h = sha256hex(bytes)
    % Lowercase hex SHA-256 of a uint8 vector (base MATLAB / Java MessageDigest;
    % matches the python hashlib manifest SHAs).
    md  = java.security.MessageDigest.getInstance('SHA-256');
    dig = md.digest(typecast(uint8(bytes(:)'), 'int8'));
    b   = typecast(dig, 'uint8');
    h   = lower(string(sprintf('%02x', b)));
end

function bytes = emitDocPart(leaf)
    % Document().save() to a temp .docx, unzip once, return the word/<leaf> bytes.
    d = mat2doc.Document();
    tmp = [tempname '.docx'];
    cleanTmp = onCleanup(@() deleteIfExists(tmp)); %#ok<NASGU>
    d.save(tmp);
    exdir = tempname;
    cleanEx = onCleanup(@() rmdirIfExists(exdir)); %#ok<NASGU>
    unzip(tmp, exdir);
    bytes = readBytes(fullfile(exdir, 'word', leaf));
end

function deleteIfExists(p)
    if isfile(p)
        delete(p);
    end
end

function rmdirIfExists(p)
    if isfolder(p)
        rmdir(p, 's');
    end
end

% ---- s0099 mutation-serialize replay (the .m twin body, VERBATIM) ----

function cand = runS0099()
    % Replay the s0099 numbering PRODUCTION-serializer mutation sequences and
    % return a struct of raw uint8 serialize_part_xml bytes per case. Embedded here
    % so the Equivalence leg is self-contained (the validation-folder scenario is
    % NOT on the toolbox path and must not be a dependency). Mirrors
    % validation\mat2doc\scenarios\s0099_p8_1_numbering_serialize.m VERBATIM.
    % Fully-qualified inline (NO specific import -- a `import mat2doc.oxml.X`
    % errors at suite-creation parse time before the PathFixture runs; see the
    % memory on specific-import-fails-in-test-class).
    OxmlElement        = @mat2doc.oxml.OxmlElement;
    serialize_part_xml = @mat2doc.oxml.serialize_part_xml;
    cand = struct();

    % ---- caseA: add_num auto-numId + lvlOverride + startOverride (append) ----
    numbering = OxmlElement("w:numbering");
    n1 = numbering.add_num(0);              % numId auto = 1
    numbering.add_num(3);                   % numId auto = 2
    lo = n1.add_lvlOverride(0);             % <w:lvlOverride w:ilvl="0"> on num #1
    lo.add_startOverride(5);                % <w:startOverride w:val="5"/>
    cand.caseA_build = serialize_part_xml(numbering);

    % ---- caseB: successor insert-before <w:numIdMacAtCleanup> ----
    numbering = OxmlElement("w:numbering");
    cleanup = OxmlElement("w:numIdMacAtCleanup");
    numbering.append(cleanup);
    numbering.add_num(7);                   % numId 1, inserted BEFORE cleanup marker
    cand.caseB_cleanup = serialize_part_xml(numbering);

    % ---- caseC: CT_NumPr H11 ordering -- numId first, ilvl sorts before it ----
    numPr = OxmlElement("w:numPr");
    numId = numPr.get_or_add_numId();       % added FIRST
    numId.val = 2;
    ilvl = numPr.get_or_add_ilvl();         % added SECOND; successors incl w:numId
    ilvl.val = 0;                           % -> <w:ilvl> serializes before <w:numId>
    cand.caseC_numpr_order = serialize_part_xml(numPr);

    % ---- caseD: CT_NumPr with pre-existing <w:ins> (successor boundary) ----
    numPr = OxmlElement("w:numPr");
    numPr.append(OxmlElement("w:ins"));     % pre-existing trailing successor
    numId = numPr.get_or_add_numId();       % inserts before <w:ins>
    numId.val = 5;
    ilvl = numPr.get_or_add_ilvl();         % inserts before <w:numId>
    ilvl.val = 1;
    cand.caseD_numpr_ins = serialize_part_xml(numPr);
end

% ---- s0100 API-value probe replay (the .m twin body, VERBATIM) ----

function P = runS0100()
    % Replay the s0100 numbering API-value probe and return a struct of tagged
    % canonical strings (int|/str|/bool|/err|). Embedded here so the Equivalence
    % leg is self-contained. Mirrors
    % validation\mat2doc\scenarios\s0100_p8_1_numbering_api.m VERBATIM.
    % Fully-qualified inline (NO specific import -- parse-time hazard; see memory).
    OxmlElement = @mat2doc.oxml.OxmlElement;
    CT_Num = @(a, b) mat2doc.oxml.numbering.CT_Num.new(a, b);
    P = struct();

    % ---- add_num sequential auto numId ----
    numbering = OxmlElement("w:numbering");
    a1 = numbering.add_num(0); P.add_num_id1 = ci(a1.numId);   % 1
    a2 = numbering.add_num(1); P.add_num_id2 = ci(a2.numId);   % 2
    a3 = numbering.add_num(2); P.add_num_id3 = ci(a3.numId);   % 3

    % ---- _next_numId gap-fill via CT_Num.new-seeded numberings ----
    nb = OxmlElement("w:numbering");
    nb.append(CT_Num(1, 0)); nb.append(CT_Num(3, 0));
    g = nb.add_num(0);  P.gapfill_13 = ci(g.numId);            % 2
    nb = OxmlElement("w:numbering");
    g = nb.add_num(0);  P.gapfill_empty = ci(g.numId);         % 1
    nb = OxmlElement("w:numbering");
    nb.append(CT_Num(1, 0)); nb.append(CT_Num(2, 0));
    g = nb.add_num(0);  P.gapfill_12 = ci(g.numId);            % 3

    % ---- num_having_numId: hit (identity) + miss (KeyError raw msg) ----
    nb = OxmlElement("w:numbering");
    nb.add_num(0); nb.add_num(0);                              % numIds 1, 2
    hit = nb.num_having_numId(2);
    P.having_hit_numid = ci(hit.numId);                       % 2
    nl = nb.num_lst;
    P.having_hit_identity = cb(hit == nl(2));                 % same handle as num_lst[1]
    P.having_miss = errmsg_(@() nb.num_having_numId(99));

    % ---- CT_Num.new attrs ----
    n = CT_Num(5, 42);
    P.ctnum_numid = ci(n.numId);                              % 5
    anid = n.abstractNumId;
    P.ctnum_abstractnumid_val = ci(anid.val);                % 42

    % ---- NumberingDefinitions_ len ----
    nb = OxmlElement("w:numbering");
    nb.add_num(0); nb.add_num(0); nb.add_num(0);
    P.nd_len = ci(mat2doc.parts.NumberingDefinitions_(nb).len_());  % 3

    % ---- DocumentPart.numbering_part on default.docx ----
    d = mat2doc.Document();
    dp = d.part;
    np = dp.numbering_part;
    nd = np.numbering_definitions;
    cparts = split(string(class(np)), ".");
    P.default_numpart_class = "str|" + cparts(end);          % NumberingPart
    P.default_nd_len = ci(nd.len_());                        % 9
    P.default_numpart_identity = cb(dp.numbering_part == np);
    P.default_numdef_identity = cb(np.numbering_definitions == nd);

    % ---- NumberingPart.new faithful NotImplementedError ----
    P.numpart_new_raises = errtype_(@() mat2doc.parts.NumberingPart.new());
end

function s = ci(v)
    s = "int|" + mat2doc.shared.pyStr(double(v), "int");
end

function s = cb(v)
    if v, s = "bool|true"; else, s = "bool|false"; end
end

function s = errmsg_(fn)
    try
        fn(); s = "NO-ERROR";
    catch e
        p = split(string(e.identifier), ":");
        s = "err|" + p(end) + "|" + string(e.message);
    end
end

function s = errtype_(fn)
    try
        fn(); s = "NO-ERROR";
    catch e
        p = split(string(e.identifier), ":");
        s = "err|" + p(end);
    end
end

classdef Test_p1_4_opc_oxml < matlab.unittest.TestCase
% TEST_P1_4_OPC_OXML  Gate-4 permanent unit tests for Mat2Doc P1-4
%   (opc two-tier oxml + constants/spec/shared).
%
%   The FIRST Mat2Doc tests\opc\ class. Freezes the M1 byte-critical OPC
%   serializer path ported in +mat2doc\+opc\ from python-docx v1.2.0
%   src/docx/opc/{oxml,constants,spec,shared}.py, plus the un-stubbed pretty
%   engine +mat2doc\+oxml\serialize_for_reading.m.
%
%   Provenance (Gate-1..3, all 2026-07-25):
%     * Audit  : validation\mat2doc\audit_P1-4_opc_oxml.md (Porter Gate-1 +
%                Fable Gate-2 APPROVE, fixes F1/F2 to the test-only pretty engine)
%     * Validate: validation\mat2doc\validate_P1-4_opc_oxml.md (Gate-3 PASS --
%                isolated-serializer L1 three-way byte agreement proven 4/4)
%     * Harness : harness\mat2doc\validate_p1_4.m (the three Gate-3 lanes)
%     * Scenario: validation\mat2doc\scenarios\s0002_opc_serializer.m (the exact
%                frozen s0001 rows in frozen document order, reproduced below)
%
%   Coverage taxonomy
%   -----------------
%   * Equivalence / Regression (THE M1 pin) -- rebuild CT_Types (3 Defaults + 11
%     Overrides) and the three CT_Relationships from the frozen s0001 rows and
%     assert BYTE-IDENTICAL (and SHA-256-identical) to the frozen reference
%     parts, copied byte-for-byte into tests\opc\data\ so the suite is
%     self-contained. Sizes 1738 / 734 / 1227 / 295 B. This is L1: the ladder
%     demanded L1 for the whole M1 serializer path (Gate-3 verdict). If anyone
%     perturbs the serializer these go RED.
%   * Regression (F1/F2) -- the 13 pretty-engine vectors (9 pretty + 2 root-tail
%     + 2 .xml routes) frozen against live lxml 5.3.0, folded verbatim from the
%     Gate-3 harness. The core .xml un-stub is exercised via the
%     BaseOxmlElement CONSTRUCTOR route (see test_xml_route_core_baseoxml) --
%     NOT parse_xml("<w:p>"), which is docx-faithful behaviour, not a bug.
%   * Regression (constants) -- a representative subset of CONTENT_TYPE /
%     RELATIONSHIP_TYPE (incl. the anomalous PIVOT_CACHE_RECORDS infix), and
%     ALL of NAMESPACE and RELATIONSHIP_TARGET_MODE, pinned value-for-value.
%   * Regression (spec pair-list) -- default_content_types is a duplicate-key
%     pair-LIST, not a dict: the three `bin` rows must survive (a dict-collapse
%     would drop them). Row count + the multi-valued key pinned.
%   * Behavioral -- TargetMode tri-state, CaseInsensitiveDict fold/casing,
%     defaults/overrides/Relationship_lst counts 3/11/4, registry Clark-name
%     dispatch to CT_Types / CT_Relationships.
%
%   Deviations exercised (adopt-only, ZERO new D-numbers -- Gate-3):
%     D-001 (own OOXML-subset parser) and D-serializer-nsdecl (verbatim-until-
%     moved single-root-xmlns) are the MECHANISM producing the L1 bytes, not L2
%     divergences; every byte pin below re-verifies them byte-neutral.
%
%   Determinism: no network, no absolute paths -- every fixture is resolved
%   relative to this file via fileparts(mfilename('fullpath')).

    properties (Constant)
        % --- frozen s0001 rows, in frozen DOCUMENT order (from
        %     scenarios\s0002_opc_serializer.m -- the Gate-3 twin). ----------
        DEFAULTS = [ ...
            "jpeg", "image/jpeg"; ...
            "rels", "application/vnd.openxmlformats-package.relationships+xml"; ...
            "xml",  "application/xml"];

        OVERRIDES = [ ...
            "/customXml/itemProps1.xml",  "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"; ...
            "/docProps/app.xml",          "application/vnd.openxmlformats-officedocument.extended-properties+xml"; ...
            "/docProps/core.xml",         "application/vnd.openxmlformats-package.core-properties+xml"; ...
            "/word/document.xml",         "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"; ...
            "/word/fontTable.xml",        "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"; ...
            "/word/numbering.xml",        "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"; ...
            "/word/settings.xml",         "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"; ...
            "/word/styles.xml",           "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"; ...
            "/word/stylesWithEffects.xml","application/vnd.ms-word.stylesWithEffects+xml"; ...
            "/word/theme/theme1.xml",     "application/vnd.openxmlformats-officedocument.theme+xml"; ...
            "/word/webSettings.xml",      "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"];
    end

    properties (TestParameter)
        % --- L1 rels parts: {fixture-basename, expected-size} -------------
        relsCase = struct( ...
            'main',      {{'rels_main',       734}}, ...
            'document',  {{'rels_document',   1227}}, ...
            'customxml', {{'rels_customxml',  295}});

        % --- F1/F2 pretty-engine vectors: {input, expected} (from
        %     harness\mat2doc\validate_p1_4.m Lane 2; live lxml 5.3.0). ------
        prettyVec = struct( ...
            'v1_types_container', {{ ...
                '<Types xmlns="http://x/ct"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/p" ContentType="q"/></Types>', ...
                sprintf('<Types xmlns="http://x/ct">\n  <Default Extension="xml" ContentType="application/xml"/>\n  <Override PartName="/p" ContentType="q"/>\n</Types>\n')}}, ...
            'v2_wml_text_leaf', {{ ...
                '<w:p xmlns:w="http://w"><w:r><w:t>hi</w:t></w:r></w:p>', ...
                sprintf('<w:p xmlns:w="http://w">\n  <w:r>\n    <w:t>hi</w:t>\n  </w:r>\n</w:p>\n')}}, ...
            'v3_mixed_inline', {{ ...
                '<a>text<b><c/></b></a>', ...
                sprintf('<a>text<b><c/></b></a>\n')}}, ...
            'v4_deep_mixed_inline', {{ ...
                '<a>text<b><c><d/></c></b></a>', ...
                sprintf('<a>text<b><c><d/></c></b></a>\n')}}, ...
            'v5_mixed_subtree_offpropagate', {{ ...
                '<a><b>t<c><d/></c></b></a>', ...
                sprintf('<a>\n  <b>t<c><d/></c></b>\n</a>\n')}}, ...
            'v6_tail_inline', {{ ...
                '<a><b/>tail<c/></a>', ...
                sprintf('<a><b/>tail<c/></a>\n')}}, ...
            'v7_empty_open_and_selfclose', {{ ...
                '<a><b></b><c/></a>', ...
                sprintf('<a>\n  <b/>\n  <c/>\n</a>\n')}}, ...
            'v8_all_element_deep', {{ ...
                '<a><b><c><d><e/></d></c></b></a>', ...
                sprintf('<a>\n  <b>\n    <c>\n      <d>\n        <e/>\n      </d>\n    </c>\n  </b>\n</a>\n')}}, ...
            'v9_root_text_offpropagate', {{ ...
                '<a>t<b><c><d><e/></d></c></b></a>', ...
                sprintf('<a>t<b><c><d><e/></d></c></b></a>\n')}});

        % --- constants: {constant-NAME, expected-value} ------------------
        %     Field name = the constant symbol (nice sub-test label).
        nsCase = struct( ...
            'DML_WORDPROCESSING_DRAWING', {{'DML_WORDPROCESSING_DRAWING', "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"}}, ...
            'OFC_RELATIONSHIPS',          {{'OFC_RELATIONSHIPS',          "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}}, ...
            'OPC_RELATIONSHIPS',          {{'OPC_RELATIONSHIPS',          "http://schemas.openxmlformats.org/package/2006/relationships"}}, ...
            'OPC_CONTENT_TYPES',          {{'OPC_CONTENT_TYPES',          "http://schemas.openxmlformats.org/package/2006/content-types"}}, ...
            'WML_MAIN',                   {{'WML_MAIN',                   "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}});

        rtmCase = struct( ...
            'INTERNAL', {{'INTERNAL', "Internal"}}, ...
            'EXTERNAL', {{'EXTERNAL', "External"}});

        ctCase = struct( ...
            'BMP',               {{'BMP',               "image/bmp"}}, ...
            'JPEG',              {{'JPEG',              "image/jpeg"}}, ...
            'PNG',               {{'PNG',               "image/png"}}, ...
            'XML',               {{'XML',               "application/xml"}}, ...
            'OPC_RELATIONSHIPS', {{'OPC_RELATIONSHIPS', "application/vnd.openxmlformats-package.relationships+xml"}}, ...
            'SML_SHEET',         {{'SML_SHEET',         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}}, ...
            'WML_DOCUMENT_MAIN', {{'WML_DOCUMENT_MAIN', "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"}});

        % PIVOT_CACHE_RECORDS is the transcription-anomaly guard: its infix is
        % ".../relationships/spreadsheetml/pivotCacheRecords" (constants.py),
        % a shape the auditor flagged for spot-checking (audit sec 10).
        rtCase = struct( ...
            'CORE_PROPERTIES',     {{'CORE_PROPERTIES',     "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"}}, ...
            'OFFICE_DOCUMENT',     {{'OFFICE_DOCUMENT',     "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"}}, ...
            'STYLES',              {{'STYLES',              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"}}, ...
            'NUMBERING',           {{'NUMBERING',           "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"}}, ...
            'SETTINGS',            {{'SETTINGS',            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"}}, ...
            'PIVOT_CACHE_RECORDS', {{'PIVOT_CACHE_RECORDS', "http://schemas.openxmlformats.org/officeDocument/2006/relationships/spreadsheetml/pivotCacheRecords"}});
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into
            % the test folder, so without the worktree root on the path a COLD
            % run cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Copy of the proven idiom from tests\oxml\Test_p1_2_oxml.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\opc
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. Isolated-serializer L1 -- the M1 byte-critical pin           %
        % =============================================================== %

        function test_content_types_L1(testCase)
            % Equivalence + Regression (M1 L1): rebuild CT_Types with the 3
            % frozen Defaults + 11 frozen Overrides in frozen document order,
            % serialize_part_xml, assert BYTE-IDENTICAL and SHA-256-identical
            % to the frozen 1738-byte [Content_Types].xml.
            %   Byte-level L1: Gate-3 ladder demanded L1 for the whole M1 path.
            types = mat2doc.opc.oxml.CT_Types.new();
            D = testCase.DEFAULTS;  O = testCase.OVERRIDES;
            for k = 1:size(D, 1), types.add_default(D(k,1), D(k,2)); end
            for k = 1:size(O, 1), types.add_override(O(k,1), O(k,2)); end
            got = uint8(mat2doc.opc.oxml.serialize_part_xml(types));
            got = got(:)';
            want = testCase.loadFixture('content_types');
            testCase.verifyEqual(numel(got), 1738, ...
                '[Content_Types].xml must be exactly 1738 B');
            verifyByteIdentical(testCase, got, want, '[Content_Types].xml');
            testCase.verifyEqual(sha256hex(got), sha256hex(want), ...
                '[Content_Types].xml SHA-256 must equal the frozen reference');
        end

        function test_rels_L1(testCase, relsCase)
            % Equivalence + Regression (M1 L1): rebuild each CT_Relationships
            % from the frozen rows (frozen order) and assert its xml_file_bytes
            % (== Python CT_Relationships.xml override) BYTE-IDENTICAL and
            % SHA-256-identical to the frozen .rels part. Byte-level L1.
            base = relsCase{1};  expsz = relsCase{2};
            rows = testCase.relsRows(base);
            rels = mat2doc.opc.oxml.CT_Relationships.new();
            for k = 1:size(rows, 1)
                rels.add_rel(rows{k,1}, rows{k,2}, rows{k,3}, rows{k,4});
            end
            got = uint8(rels.xml_file_bytes);  got = got(:)';
            want = testCase.loadFixture(base);
            testCase.verifyEqual(numel(got), expsz, ...
                sprintf('%s must be exactly %d B', base, expsz));
            verifyByteIdentical(testCase, got, want, base);
            testCase.verifyEqual(sha256hex(got), sha256hex(want), ...
                sprintf('%s SHA-256 must equal the frozen reference', base));
        end

        % =============================================================== %
        % 2. Pretty engine (serialize_for_reading) -- F1/F2 vectors       %
        % =============================================================== %

        function test_pretty_vector(testCase, prettyVec)
            % Regression (Gate-2 F1/F2): the 9 pretty vectors captured from live
            % lxml 5.3.0. Vectors v4/v5/v9 are the DEEP MIXED-content cases F1
            % fixed (format-off propagates through the whole mixed subtree, never
            % re-enabled). serialize_for_reading returns a STRING; byte-equal.
            in = string(prettyVec{1});  want = string(prettyVec{2});
            el  = mat2doc.oxml.parse_xml(in);
            got = string(mat2doc.oxml.serialize_for_reading(el));
            testCase.verifyEqual(got, want, ...
                sprintf('pretty vector mismatch for input <%s>', in));
        end

        function test_root_tail_selfclose(testCase)
            % Regression (Gate-2 F2): lxml tostring defaults with_tail=True, so a
            % root element's own tail is emitted before the trailing LF.
            % <a/>TAIL\n for a self-closing root that carries a tail.
            r = mat2doc.oxml.parse_xml("<r><a/>TAIL</r>");
            kids = r.to_array();  a = kids(1);
            got = string(mat2doc.oxml.serialize_for_reading(a));
            testCase.verifyEqual(got, string(sprintf('<a/>TAIL\n')), ...
                'F2 root-tail: self-closing root must emit <a/>TAIL then LF');
        end

        function test_root_tail_after_close(testCase)
            % Regression (Gate-2 F2): tail after a closing tag, </a>TAIL\n, with
            % the (formatted) child subtree indented above it.
            r = mat2doc.oxml.parse_xml("<r><a><b/></a>TAIL</r>");
            kids = r.to_array();  a = kids(1);
            got = string(mat2doc.oxml.serialize_for_reading(a));
            testCase.verifyEqual(got, string(sprintf('<a>\n  <b/>\n</a>TAIL\n')), ...
                'F2 root-tail: tail after </a> must follow the formatted subtree');
        end

        function test_xml_route_core_baseoxml(testCase)
            % Regression (.xml un-stub, CORRECT route -- Gate-3 handoff note):
            % the core BaseOxmlElement.xml un-stub is exercised via the
            % BaseOxmlElement CONSTRUCTOR, *** NOT *** parse_xml("<w:p ...>").
            % An UNREGISTERED w:p faithfully falls back to the bare element class
            % (CT_P is not ported yet), so it correctly LACKS the custom .xml --
            % exactly as python-docx's ElementNamespaceClassLookup() falls back to
            % bare etree._Element for unregistered tags. A parse_xml-route
            % assertion here would WRONGLY appear to fail; do NOT "fix" it to
            % parse_xml. (validate_P1-4_opc_oxml.md, reconciliation note.)
            W  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            wp = mat2doc.oxml.BaseOxmlElement("w:p", struct('w', W));
            wr = mat2doc.oxml.BaseOxmlElement("w:r", struct('w', W));
            wp.append(wr);
            want = sprintf(['<w:p xmlns:w="http://schemas.openxmlformats.org/' ...
                'wordprocessingml/2006/main">\n  <w:r/>\n</w:p>\n']);
            testCase.verifyEqual(string(wp.xml), string(want), ...
                'core BaseOxmlElement.xml un-stub (w:p>w:r) must match live lxml');
        end

        function test_xml_route_opc_ct_types(testCase)
            % Regression (.xml route, opc pretty property): CT_Types inherits the
            % opc BaseOxmlElement pretty .xml; a one-Default Types prints with a
            % 2-space indent and one trailing LF, byte-equal to live python-docx.
            t = mat2doc.opc.oxml.CT_Types.new();
            t.add_default("xml", "application/xml");
            want = sprintf(['<Types xmlns="http://schemas.openxmlformats.org/' ...
                'package/2006/content-types">\n  <Default Extension="xml" ' ...
                'ContentType="application/xml"/>\n</Types>\n']);
            testCase.verifyEqual(string(t.xml), string(want), ...
                'opc BaseOxmlElement.xml pretty (CT_Types) must match live lxml');
        end

        % =============================================================== %
        % 3. Constants integrity                                          %
        % =============================================================== %

        function test_namespace_value(testCase, nsCase)
            % Regression: NAMESPACE constant transcription (all 5 URIs).
            c = mat2doc.opc.NAMESPACE;
            testCase.verifyEqual(c.(nsCase{1}), nsCase{2}, ...
                sprintf('NAMESPACE.%s value mismatch', nsCase{1}));
        end

        function test_rtm_value(testCase, rtmCase)
            % Regression: RELATIONSHIP_TARGET_MODE (both members).
            c = mat2doc.opc.RELATIONSHIP_TARGET_MODE;
            testCase.verifyEqual(c.(rtmCase{1}), rtmCase{2}, ...
                sprintf('RELATIONSHIP_TARGET_MODE.%s value mismatch', rtmCase{1}));
        end

        function test_content_type_value(testCase, ctCase)
            % Regression: CONTENT_TYPE subset transcription.
            c = mat2doc.opc.CONTENT_TYPE;
            testCase.verifyEqual(c.(ctCase{1}), ctCase{2}, ...
                sprintf('CONTENT_TYPE.%s value mismatch', ctCase{1}));
        end

        function test_relationship_type_value(testCase, rtCase)
            % Regression: RELATIONSHIP_TYPE subset incl. the anomalous
            % PIVOT_CACHE_RECORDS infix (.../spreadsheetml/pivotCacheRecords).
            c = mat2doc.opc.RELATIONSHIP_TYPE;
            testCase.verifyEqual(c.(rtCase{1}), rtCase{2}, ...
                sprintf('RELATIONSHIP_TYPE.%s value mismatch', rtCase{1}));
        end

        % =============================================================== %
        % 4. spec pair-list semantics (duplicate-key survival)            %
        % =============================================================== %

        function test_spec_row_count(testCase)
            % Regression: default_content_types is an 18-row Nx2 pair-LIST
            % (spec.py:5-24). A dict-collapse would drop rows -> pin the count
            % and the 2 columns.
            pairs = mat2doc.opc.default_content_types();
            testCase.verifyEqual(size(pairs), [18 2], ...
                'default_content_types must be an 18x2 pair-list');
        end

        function test_spec_bin_is_multivalued(testCase)
            % Regression (trap 4, CRITICAL): the extension `bin` maps to THREE
            % distinct content-types (pptx/xlsx/docx printerSettings). A
            % containers.Map/dictionary keyed by extension would COLLAPSE these
            % to one and silently break default-vs-override precedence. Pin that
            % `bin` appears 3x with 3 DISTINCT content-types.
            pairs = mat2doc.opc.default_content_types();
            binRows = pairs(pairs(:,1) == "bin", 2);
            testCase.verifyEqual(numel(binRows), 3, ...
                'extension `bin` must appear on 3 rows (pair-list, not dict)');
            testCase.verifyEqual(numel(unique(binRows)), 3, ...
                'the 3 `bin` rows must carry 3 DISTINCT content-types');
        end

        function test_spec_bin_content_types_frozen(testCase)
            % Regression: pin the exact 3 `bin` content-type values (the ones a
            % dict-collapse would silently lose). Sorted for order-insensitivity
            % (membership is a set test in Python; row order carries no behavior).
            pairs = mat2doc.opc.default_content_types();
            binRows = sort(pairs(pairs(:,1) == "bin", 2));
            want = sort([ ...
                "application/vnd.openxmlformats-officedocument.presentationml.printerSettings"; ...
                "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"; ...
                "application/vnd.openxmlformats-officedocument.wordprocessingml.printerSettings"]);
            testCase.verifyEqual(binRows, want, ...
                'the 3 frozen `bin` content-types must survive verbatim');
        end

        % =============================================================== %
        % 5. Behavioral pins                                              %
        % =============================================================== %

        function test_targetmode_internal_omits_attr(testCase)
            % Behavioral (H3 tri-state): an Internal relationship OMITS the
            % TargetMode attribute entirely, and target_mode getter DEFAULTS to
            % "Internal" (two-arg get). get("TargetMode") returns [] (absent).
            r = mat2doc.opc.oxml.CT_Relationship.new("rId9", "http://t", "target.xml");
            testCase.verifyEmpty(r.get("TargetMode"), ...
                'Internal rel must NOT carry a TargetMode attribute');
            testCase.verifyEqual(r.target_mode, "Internal", ...
                'target_mode getter must default to "Internal" when absent');
        end

        function test_targetmode_external_emits_attr(testCase)
            % Behavioral (trap 2): an External relationship EMITS
            % TargetMode="External" and the getter returns "External".
            r = mat2doc.opc.oxml.CT_Relationship.new("rId9", "http://t", ...
                "http://x/y", "External");
            testCase.verifyEqual(string(r.get("TargetMode")), "External", ...
                'External rel must carry TargetMode="External"');
            testCase.verifyEqual(r.target_mode, "External", ...
                'target_mode getter must report "External"');
        end

        function test_caseinsensitivedict(testCase)
            % Behavioral (H15): key folded to lowercase on set/get/isKey, but the
            % stored VALUE casing is preserved verbatim. keys() returns the
            % lowercased key.
            d = mat2doc.opc.CaseInsensitiveDict();
            d.set("PNG", "image/PNG");            % value casing intentionally mixed
            testCase.verifyTrue(d.isKey("png"), 'isKey must fold the key');
            testCase.verifyTrue(d.isKey("Png"), 'isKey must fold the key (mixed)');
            testCase.verifyEqual(d.get("pNg"), "image/PNG", ...
                'get must fold the key AND preserve stored value casing');
            testCase.verifyEqual(d.keys(), "png", ...
                'keys() must return the lowercased stored key');
        end

        function test_counts_3_11_4(testCase)
            % Behavioral: the frozen M1 build yields defaults=3, overrides=11
            % (via OPC-local qn ct:Default/ct:Override) and Relationship_lst=4
            % on the main .rels (via pr:Relationship).
            types = mat2doc.opc.oxml.CT_Types.new();
            D = testCase.DEFAULTS;  O = testCase.OVERRIDES;
            for k = 1:size(D, 1), types.add_default(D(k,1), D(k,2)); end
            for k = 1:size(O, 1), types.add_override(O(k,1), O(k,2)); end
            testCase.verifyEqual(numel(types.defaults), 3, 'defaults count');
            testCase.verifyEqual(numel(types.overrides), 11, 'overrides count');

            rows = testCase.relsRows('rels_main');
            rels = mat2doc.opc.oxml.CT_Relationships.new();
            for k = 1:size(rows, 1)
                rels.add_rel(rows{k,1}, rows{k,2}, rows{k,3}, rows{k,4});
            end
            testCase.verifyEqual(numel(rels.Relationship_lst), 4, ...
                'main .rels Relationship_lst count');
        end

        function test_registry_dispatch_types(testCase)
            % Behavioral (H10): parsing the frozen [Content_Types].xml dispatches
            % the root to mat2doc.opc.oxml.CT_Types by raw Clark name (the 5 OPC
            % registry rows). Was plain XmlElement before P1-4.
            b = testCase.loadFixture('content_types');
            elm = mat2doc.oxml.parse_xml(uint8(b(:)'));
            testCase.verifyClass(elm, 'mat2doc.opc.oxml.CT_Types', ...
                'frozen [Content_Types].xml root must dispatch to CT_Types');
        end

        function test_registry_dispatch_rels(testCase)
            % Behavioral (H10): parsing a frozen .rels dispatches the root to
            % mat2doc.opc.oxml.CT_Relationships by raw Clark name.
            b = testCase.loadFixture('rels_main');
            elm = mat2doc.oxml.parse_xml(uint8(b(:)'));
            testCase.verifyClass(elm, 'mat2doc.opc.oxml.CT_Relationships', ...
                'frozen .rels root must dispatch to CT_Relationships');
        end

    end

    % ===================== instance helpers ============================ %
    methods (Access = private)
        function b = loadFixture(~, base)
            here = fileparts(mfilename('fullpath'));   % tests\opc
            p = fullfile(here, 'data', [base '.bin']);
            f = fopen(p, 'r', 'n');
            assert(f >= 0, 'byte fixture missing: %s', p);
            b = fread(f, Inf, '*uint8')';
            fclose(f);
        end

        function rows = relsRows(~, base)
            % The frozen s0001 rows for each .rels part, frozen row order:
            %   {rId, reltype, target, is_external}. From
            %   scenarios\s0002_opc_serializer.m.
            OFC  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            PKG  = "http://schemas.openxmlformats.org/package/2006/relationships";
            CORE = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
            MS07 = "http://schemas.microsoft.com/office/2007/relationships";
            switch base
                case 'rels_main'
                    rows = { ...
                        "rId3", CORE,                         "docProps/core.xml",       false; ...
                        "rId4", OFC + "/extended-properties", "docProps/app.xml",        false; ...
                        "rId1", OFC + "/officeDocument",      "word/document.xml",       false; ...
                        "rId2", PKG + "/metadata/thumbnail",  "docProps/thumbnail.jpeg", false};
                case 'rels_document'
                    rows = { ...
                        "rId3", OFC + "/styles",             "styles.xml",            false; ...
                        "rId4", MS07 + "/stylesWithEffects", "stylesWithEffects.xml", false; ...
                        "rId5", OFC + "/settings",           "settings.xml",          false; ...
                        "rId6", OFC + "/webSettings",        "webSettings.xml",       false; ...
                        "rId7", OFC + "/fontTable",          "fontTable.xml",         false; ...
                        "rId8", OFC + "/theme",              "theme/theme1.xml",      false; ...
                        "rId1", OFC + "/customXml",          "../customXml/item1.xml",false; ...
                        "rId2", OFC + "/numbering",          "numbering.xml",         false};
                case 'rels_customxml'
                    rows = { ...
                        "rId1", OFC + "/customXmlProps",     "itemProps1.xml",        false};
                otherwise
                    error('Test_p1_4_opc_oxml:unknownRels', 'unknown rels base %s', base);
            end
        end
    end
end

% ===================== file-local helpers ============================== %

function verifyByteIdentical(testCase, got, want, label)
    % Byte-level assertion (L1). On mismatch report sizes and first diff offset.
    got = uint8(got(:)');  want = uint8(want(:)');
    if ~isequal(got, want)
        n = min(numel(got), numel(want));
        d = find(got(1:n) ~= want(1:n), 1);
        if isempty(d), d = n + 1; end
        gv = 0; wv = 0;
        if d <= numel(got), gv = double(got(d)); end
        if d <= numel(want), wv = double(want(d)); end
        diag = sprintf(['%s: bytes differ (got %d B, want %d B); first diff @%d ' ...
            '(got 0x%02X, want 0x%02X)'], char(label), numel(got), numel(want), d, gv, wv);
    else
        diag = char(label);
    end
    testCase.verifyEqual(got, want, diag);
end

function h = sha256hex(bytes)
    % Lowercase hex SHA-256 of a uint8 row vector (base MATLAB / Java
    % MessageDigest). typecast(uint8->int8) yields Java's signed byte[] view.
    md  = java.security.MessageDigest.getInstance('SHA-256');
    dig = md.digest(typecast(uint8(bytes(:)'), 'int8'));
    b   = typecast(dig, 'uint8');
    h   = lower(string(sprintf('%02x', b)));
end

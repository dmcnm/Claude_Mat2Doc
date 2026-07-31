classdef Test_p5_1_settings < matlab.unittest.TestCase
% TEST_P5_1_SETTINGS  Gate-4 permanent unit tests for Mat2Doc P5-1
%   (the oxml/settings + settings-API layer): src/docx/oxml/settings.py::
%   CT_Settings -> +mat2doc\+oxml\+settings\CT_Settings; src/docx/settings.py::
%   Settings -> +mat2doc\+settings\Settings; the TWO settings-block registry rows
%   (w:settings->CT_Settings, w:evenAndOddHeaders->CT_OnOff) in
%   +mat2doc\+oxml\registry.m; and the THREE un-stubbed .settings accessors
%   (Document.settings / DocumentPart.settings / SettingsPart.settings).
%
%   P5-1 is the FIRST Phase-5 WP -- a REGISTRY-ADDING, M1-byte-path-touching WP:
%   registering w:settings flips the PARSE CLASS of the word/settings.xml root
%   from generic XmlElement to CT_Settings on every load/save. CT_Settings exits
%   through the identical +oxml\serialize_part_xml walk, so the flip is
%   byte-neutral -- word/settings.xml stays 2535 B, SHA-256 51a0d348... This class
%   permanently freezes the guarantees the prior gates established:
%     * Gate-1 Porter  : audit_P5-1_settings.md (self-probe).
%     * Gate-2 Auditor (Fable): audit_P5-1_settings.md GATE-2 -- APPROVE; the H11
%       successor-slice [48:]->(49:end) proven; H4 identity guard proven.
%     * Gate-3 Validator: validate_P5-1_settings.md -- PASS, ZERO new D-numbers.
%       probe_diff s0036 MATCH exit 0; M1 17/17 (settings.xml 2535 B L1); the
%       novel WRITE path 17/17 byte-identical (settings.xml 2557 B / 66052d2f...);
%       the REMOVE path 17/17 byte-identical back to the M1 default; H11
%       successor-slice insertion byte-identical adversarially (both-neighbor +
%       prefixed); targeted regression clean but for the 2 by-design stub-flip
%       pins (which THIS Gate-4 job re-pins in Test_p2_2 / Test_p2_3).
%
%   THE HIGHEST-VALUE PERMANENT PINS IN THIS WP (a regression goes RED loudly):
%     * ★ H11 SUCCESSOR-SLICE (test_h11_successor_slice_pins): the descriptor uses
%       Python successors=_tag_seq[48:] -> MATLAB TAG_SEQ(49:end). Add
%       w:evenAndOddHeaders to a settings element carrying BOTH a before-@48
%       (w:defaultTableStyle @47) and an after-@48 (w:bookFoldRevPrinting @49)
%       neighbor -> it MUST land EXACTLY between them (serhex byte pin). A slice
%       regression ([48:]->(48:end), off-by-one, or a wrong start) mis-positions
%       the child -> Word repair -> this pin goes RED. The m:/sl:-prefixed
%       successor case (m:mathPr @84) is covered too (H8, nsmap prefix resolves).
%     * ★ M1 settings.xml byte-pin (test_m1_settings_xml_byte_identical): the
%       registry-adding parse-path regression guard -- mat2doc.Document().save()
%       -> word/settings.xml == 2535 B, SHA-256 51a0d348... (frozen s0001 oracle).
%     * ★ WRITE-path byte-pin (test_write_path_byte_identical): the NOVEL emit
%       path -- d.settings.odd_and_even_pages_header_footer = true -> save ->
%       word/settings.xml == 2557 B, SHA-256 66052d2f... (frozen s0037 oracle).
%     * ★ REMOVE-path byte-pin (test_remove_path_byte_identical): True-then-False
%       -> save -> word/settings.xml back to 2535 B, SHA-256 51a0d348...
%       (frozen s0038 oracle == the M1 default).
%     * H4 IDENTITY (test_ct_settings_even_val_tristate, h4_int0): a double 0 is
%       NOT `is False`, so set even_val = 0 CREATES <w:evenAndOddHeaders w:val="0"/>
%       (attribute PRESENT), NOT a removal -- the islogical&&isscalar&&~value guard.
%     * D-delta-1 (set True): True -> the EMPTY <w:evenAndOddHeaders/> (CT_OnOff.val
%       setter removes @val), pinned byte-identically.
%
%   Provenance (all 2026-07-31):
%     * Audit    : validation\mat2doc\audit_P5-1_settings.md
%     * Validate : validation\mat2doc\validate_P5-1_settings.md
%     * Scenario : validation\mat2doc\scenarios\s0036_p5_1_settings_probe.{py,m}
%                  (its probe body is replayed VERBATIM by runProbes() below);
%                  s0037_...write_gscenario / s0038_...remove_gscenario (the
%                  write/remove G-scenario byte oracles).
%     * Frozen refs (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE at Gate-3):
%         references\s0036\probe.json -- copied verbatim (self-contained) into
%           tests\oxml\data\s0036_probe_oracle.json (the Equivalence replay set;
%           value/serhex JSON, jsondecode is line-ending agnostic -> no
%           `* binary` .gitattributes needed, per the s0022/s0023/s0030 precedent).
%         references\s0001\parts\word\settings.xml -- the M1 byte reference (2535 B
%           / 51a0d348...); NOT copied (the byte pin compares SHA-256 of what
%           Document().save() itself emits, so no fixture is needed).
%         references\s0037\parts\word\settings.xml (2557 B / 66052d2f...) and
%           references\s0038\... (back to 51a0d348...) -- the write/remove byte
%           oracles, pinned here by SHA-256 of the MATLAB-emitted part (no fixture).
%
%   Coverage taxonomy
%   -----------------
%   * Nominal      -- CT_Settings.evenAndOddHeaders_val get (absent->False) and set
%                     True (empty element) / False (remove); Settings proxy
%                     get/set delegation; registry resolution; un-stub resolution.
%   * Edge         -- H4 double-0 identity (val="0", NOT removed); set None removes;
%                     parse @val="0"->False / @val="true"->True / present-no-val->
%                     True; non-default insertion positions (H11 4 cases incl.
%                     prefixed m: successor); H5 element-identity eq over the same
%                     element / eq-chain across the three un-stub accessors.
%   * Equivalence  -- test_equivalence_full_battery_vs_frozen_oracle replays the
%                     ENTIRE s0036 battery (runProbes, the .m twin's body verbatim,
%                     incl. the real settings.xml parse->serialize leg) and
%                     flatten-compares every leaf to the frozen oracle (0 diffs).
%   * Regression   -- hard-coded expected serialized-XML strings (ASCII, so
%                     string-equality == byte-identical L1) + UPPERCASE serhex of
%                     the raw UTF-8 bytes vs the frozen oracle + SHA-256 of the M1 /
%                     write / remove settings.xml parts.
%   * Upstream     -- the H11 successor-slice ordering IS the python-docx
%                     settings.py surface; the frozen oracle IS lxml's expected
%                     output for this insertion sequence.
%
%   Byte-level (L1) note: every serialized-XML comparison is either the FULL
%   serialize_part_xml output decoded as an ASCII string (string-equality ==
%   byte-equality L1) or its UPPERCASE hex (serhex) vs the frozen oracle, or the
%   SHA-256 of the emitted settings.xml part. No D-number granted any L2
%   relaxation in this WP (Gate-3: zero new, none at L2), so every pin here is L1.
%   The equivalence key-count guard is the only looser-than-byte check and is
%   commented at its site.
%
%   Determinism: no network, no absolute paths. The worktree root and the
%   co-located oracle resolve relative to this file via
%   fileparts(mfilename('fullpath')); saves go to tempname .docx deleted via
%   onCleanup; every file read is binary ('r','n'). The +mat2doc package resolves
%   via the MANDATORY PathFixture(worktree-root) in TestClassSetup (WP9-F4 lesson).

    properties (Constant)
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        % --- registered class names (P5-1, +oxml\registry.m 213-214) ---
        CT_SETTINGS = 'mat2doc.oxml.settings.CT_Settings'
        CT_ONOFF    = 'mat2doc.oxml.shared.CT_OnOff'

        % --- frozen s0001 M1 word/settings.xml byte reference (the P5-1
        %     registry-adding parse-path risk) ---
        SETTINGS_SIZE_DEFAULT = 2535
        SETTINGS_SHA_DEFAULT  = "51a0d348fe85965c66e4748a03c3c0d055d78455514f03cb121c334af7d73689"

        % --- frozen s0037 WRITE-path oracle (even_val=True inserts
        %     <w:evenAndOddHeaders/> between w:defaultTabStop @39 and
        %     w:characterSpacingControl @61 -- the first successor present) ---
        SETTINGS_SIZE_EVEN = 2557
        SETTINGS_SHA_EVEN  = "66052d2f7e0e9ad06df0db9a1beffc9eddcaf4aa01c72a7ffe303b2b926ce750"
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied verbatim from tests\oxml\Test_p4_6_oxml_styles.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\oxml
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. registry rows (the two P5-1 flips)                            %
        % =============================================================== %

        function test_registry_resolves_settings_and_evenandoddheaders(testCase)
            % Nominal / Regression (registry.m 213-214): the registry maps
            % w:settings -> CT_Settings (the settings.xml root) and
            % w:evenAndOddHeaders -> CT_OnOff (the child descriptor's element).
            % REGISTRATION is what flips the parse class of the settings part.
            rs = mat2doc.oxml.registry(mat2doc.oxml.qn("w:settings"));
            re = mat2doc.oxml.registry(mat2doc.oxml.qn("w:evenAndOddHeaders"));
            testCase.verifyEqual(char(rs), testCase.CT_SETTINGS, ...
                'registry must resolve w:settings -> CT_Settings (P5-1 row)');
            testCase.verifyEqual(char(re), testCase.CT_ONOFF, ...
                'registry must resolve w:evenAndOddHeaders -> CT_OnOff (P5-1 row)');
            % a real OxmlElement("w:settings") is exact-class CT_Settings
            testCase.verifyEqual(class(mat2doc.oxml.OxmlElement("w:settings")), ...
                testCase.CT_SETTINGS, 'OxmlElement("w:settings") must be a CT_Settings');
        end

        % =============================================================== %
        % 2. CT_Settings.evenAndOddHeaders_val tri-state (H3/H4/D-delta-1)  %
        % =============================================================== %

        function test_ct_settings_even_val_tristate(testCase)
            % Nominal + Edge + Regression (H3 tri-state, H4 identity, D-delta-1,
            % s0036 ct_settings): the full get/set/parse truth table with byte pins.
            oracle = loadOracle();

            % -- bare: child ABSENT -> the Python False (settings.py:128-129, NOT
            %    the CT_OnOff own default); the child getter -> [] (None) --
            b = mat2doc.oxml.OxmlElement("w:settings");
            testCase.verifyFalse(b.evenAndOddHeaders_val, ...
                'bare evenAndOddHeaders_val -> False (child absent)');
            testCase.verifyTrue(isequal(b.evenAndOddHeaders, []), ...
                'bare evenAndOddHeaders child -> [] (None)');

            % -- set True -> the EMPTY <w:evenAndOddHeaders/> (D-delta-1: CT_OnOff.val
            %    setter removes @val); localnames [evenAndOddHeaders]; read-back True --
            s = mat2doc.oxml.OxmlElement("w:settings");
            s.evenAndOddHeaders_val = true;
            testCase.verifyEqual(ser(s), decl() + newline + ...
                "<w:settings xmlns:w=""" + testCase.W + """>" + ...
                "<w:evenAndOddHeaders/></w:settings>", ...
                'set True -> the EMPTY <w:evenAndOddHeaders/> (D-delta-1, L1)');
            testCase.verifyEqual(hx_e(s), string(oracle.ct_settings.set_true.serhex));
            testCase.verifyEqual(childLocalnames(s), "evenAndOddHeaders");
            testCase.verifyTrue(s.evenAndOddHeaders_val, 'read-back True');

            % -- set True then False -> child REMOVED -> bare <w:settings/> --
            s = mat2doc.oxml.OxmlElement("w:settings");
            s.evenAndOddHeaders_val = true;
            s.evenAndOddHeaders_val = false;
            testCase.verifyEqual(ser(s), decl() + newline + ...
                "<w:settings xmlns:w=""" + testCase.W + """/>", ...
                'set True->False REMOVES the child -> bare <w:settings/> (L1)');
            testCase.verifyEqual(hx_e(s), string(oracle.ct_settings.set_true_then_false.serhex));
            testCase.verifyEqual(childLocalnames(s), strings(1,0), 'no children after False');
            testCase.verifyFalse(s.evenAndOddHeaders_val, 'read-back False');

            % -- set True then None ([]) -> child REMOVED (identity `is None`) --
            s = mat2doc.oxml.OxmlElement("w:settings");
            s.evenAndOddHeaders_val = true;
            s.evenAndOddHeaders_val = [];      % Python None -> remove
            testCase.verifyEqual(hx_e(s), string(oracle.ct_settings.set_true_then_none.serhex), ...
                'set True->None ([]) REMOVES the child (byte-identical to bare)');
            testCase.verifyEqual(childLocalnames(s), strings(1,0), 'no children after None');
            testCase.verifyFalse(s.evenAndOddHeaders_val);

            % ===================================================================== %
            %  *** H4 IDENTITY PIN -- set even_val = 0 (double) MUST CREATE @val="0" ***
            %  Python `value is None or value is False` is IDENTITY, so a double 0
            %  does NOT match -- the child is CREATED with @w:val="0" (NOT removed).
            %  The islogical&&isscalar&&~value guard reproduces the identity exactly.
            % ===================================================================== %
            s = mat2doc.oxml.OxmlElement("w:settings");
            s.evenAndOddHeaders_val = 0;       % double 0 -- NOT `is False`
            testCase.verifyEqual(ser(s), decl() + newline + ...
                "<w:settings xmlns:w=""" + testCase.W + """>" + ...
                "<w:evenAndOddHeaders w:val=""0""/></w:settings>", ...
                'H4: set even_val = 0 (double) CREATES <w:evenAndOddHeaders w:val="0"/> (NOT removed)');
            testCase.verifyEqual(hx_e(s), string(oracle.ct_settings.h4_int0.serhex));
            testCase.verifyEqual(childLocalnames(s), "evenAndOddHeaders", ...
                'H4: the child is PRESENT (not removed) after set 0');
            testCase.verifyFalse(s.evenAndOddHeaders_val, ...
                'H4: read-back False (the written "0" round-trips to False)');

            % -- parse reads: present-no-@val -> True (CT_OnOff default); @val="0"
            %    -> False; @val="true" -> True --
            p1 = parse("<w:settings " + nsW() + "><w:evenAndOddHeaders/></w:settings>");
            testCase.verifyTrue(p1.evenAndOddHeaders_val, ...
                'parse present-no-@val -> True (CT_OnOff default)');
            p2 = parse("<w:settings " + nsW() + "><w:evenAndOddHeaders w:val=""0""/></w:settings>");
            testCase.verifyFalse(p2.evenAndOddHeaders_val, 'parse @val="0" -> False');
            p3 = parse("<w:settings " + nsW() + "><w:evenAndOddHeaders w:val=""true""/></w:settings>");
            testCase.verifyTrue(p3.evenAndOddHeaders_val, 'parse @val="true" -> True');
        end

        % =============================================================== %
        % 3. ★ H11 successor-slice insertion (THE crux -- [48:]->(49:end))  %
        % =============================================================== %

        function test_h11_successor_slice_pins(testCase)
            % Regression (H11, the highest-value permanent pin, s0036 h11):
            % get_or_add_evenAndOddHeaders() uses successors=_tag_seq[48:] ->
            % TAG_SEQ(49:end). Built on adversarial neighbors, the child MUST land
            % in the exact OOXML schema position. Bytes hard-coded AND serhex vs the
            % frozen oracle. A slice regression mis-positions the child -> Word repair.
            oracle = loadOracle();

            % -- both_neighbors (THE brief's required adversarial case): a before-@48
            %    predecessor (w:defaultTableStyle @47) AND an after-@48 successor
            %    (w:bookFoldRevPrinting @49) -> child lands EXACTLY between them. --
            both = parse("<w:settings " + nsW() + ">" + ...
                "<w:defaultTableStyle w:val=""TableNormal""/>" + ...
                "<w:bookFoldRevPrinting/></w:settings>");
            both.get_or_add_evenAndOddHeaders();
            testCase.verifyEqual(childLocalnames(both), ...
                ["defaultTableStyle" "evenAndOddHeaders" "bookFoldRevPrinting"], ...
                'H11: child inserts EXACTLY between the @47 pred and @49 succ');
            testCase.verifyEqual(ser(both), decl() + newline + ...
                "<w:settings xmlns:w=""" + testCase.W + """>" + ...
                "<w:defaultTableStyle w:val=""TableNormal""/>" + ...
                "<w:evenAndOddHeaders/><w:bookFoldRevPrinting/></w:settings>", ...
                'H11 both_neighbors serialized bytes (L1) hard-coded');
            testCase.verifyEqual(hx_e(both), string(oracle.h11.both_neighbors.serhex));

            % -- pred_only: predecessor only -> APPEND after it --
            pred = parse("<w:settings " + nsW() + ">" + ...
                "<w:defaultTableStyle w:val=""TableNormal""/></w:settings>");
            pred.get_or_add_evenAndOddHeaders();
            testCase.verifyEqual(childLocalnames(pred), ...
                ["defaultTableStyle" "evenAndOddHeaders"], ...
                'H11 pred_only: append after the predecessor');
            testCase.verifyEqual(hx_e(pred), string(oracle.h11.pred_only_append.serhex));

            % -- succ_only: successor only -> INSERT before it (the slice's job) --
            succ = parse("<w:settings " + nsW() + ">" + ...
                "<w:bookFoldRevPrinting/></w:settings>");
            succ.get_or_add_evenAndOddHeaders();
            testCase.verifyEqual(childLocalnames(succ), ...
                ["evenAndOddHeaders" "bookFoldRevPrinting"], ...
                'H11 succ_only: insert BEFORE the successor');
            testCase.verifyEqual(hx_e(succ), string(oracle.h11.succ_only_insert.serhex));

            % -- succ_prefixed_mathpr (H8): the successor slice contains prefixed tags
            %    (m:mathPr @84); the m: prefix resolves via +oxml\nsmap.m -> insert
            %    BEFORE m:mathPr. --
            pfx = parse("<w:settings " + nsWM() + "><m:mathPr/></w:settings>");
            pfx.get_or_add_evenAndOddHeaders();
            testCase.verifyEqual(childLocalnames(pfx), ...
                ["evenAndOddHeaders" "mathPr"], ...
                'H11 prefixed (H8): insert BEFORE the m:-prefixed successor');
            testCase.verifyEqual(hx_e(pfx), string(oracle.h11.succ_prefixed_mathpr.serhex));
        end

        % =============================================================== %
        % 4. ★ M1 settings.xml byte-pin (registry-adding parse-path guard)  %
        % =============================================================== %

        function test_m1_settings_xml_byte_identical(testCase)
            % Regression (byte-neutrality, L1): mat2doc.Document().save() emits
            % word/settings.xml at EXACTLY 2535 B with the frozen s0001 SHA-256 --
            % byte-identical DESPITE the settings.xml root now parsing to CT_Settings
            % on the save path (P5-1's specific parse-path risk). SHA-256 equality is
            % a byte-level (L1) assertion. (Test_p1_8 owns the full 17/17 M1 sweep;
            % this pins the ONE part P5-1 could break.)
            bytes = emitDocPart('settings.xml');
            testCase.verifyEqual(numel(bytes), testCase.SETTINGS_SIZE_DEFAULT, ...
                sprintf('word/settings.xml must be exactly %d B after the settings registry rows', ...
                    testCase.SETTINGS_SIZE_DEFAULT));
            testCase.verifyEqual(sha256hex(bytes), testCase.SETTINGS_SHA_DEFAULT, ...
                'word/settings.xml SHA-256 must equal the frozen s0001 oracle (byte-identical L1)');
        end

        function test_settings_xml_parse_serialize_roundtrip_L1(testCase)
            % Regression (byte-neutrality, L1): parse the emitted settings.xml back
            % through mat2doc.oxml.parse_xml (instantiating CT_Settings at the root)
            % and re-serialize -- must be byte-identical to the input. Directly
            % exercises the new CT_Settings parse->serialize path at part scale.
            inBytes  = emitDocPart('settings.xml');
            root     = mat2doc.oxml.parse_xml(inBytes);
            outBytes = mat2doc.oxml.serialize_part_xml(root);
            testCase.verifyEqual(uint8(outBytes(:)'), uint8(inBytes(:)'), ...
                'settings.xml parse->serialize must be byte-identical (CT_Settings path is byte-neutral)');
            testCase.verifyEqual(class(root), testCase.CT_SETTINGS, ...
                'the settings.xml root must parse as CT_Settings');
            % the default part carries NO w:evenAndOddHeaders -> even_val False
            testCase.verifyFalse(root.evenAndOddHeaders_val, ...
                'default settings.xml has no w:evenAndOddHeaders -> even_val False');
        end

        % =============================================================== %
        % 5. ★ WRITE-path + REMOVE-path byte-pins (novel emit paths)        %
        % =============================================================== %

        function test_write_path_byte_identical(testCase)
            % Regression (novel WRITE path, L1): setting
            % d.settings.odd_and_even_pages_header_footer = true and saving inserts
            % <w:evenAndOddHeaders/> into word/settings.xml, growing it to EXACTLY
            % 2557 B with the frozen s0037 SHA-256 66052d2f... The mutation delegates
            % Settings -> CT_Settings.evenAndOddHeaders_val -> the H11 successor slice
            % (inserted before w:characterSpacingControl, the first present successor).
            bytes = emitDocPartMutated(@(d) setEven(d, true));
            testCase.verifyEqual(numel(bytes), testCase.SETTINGS_SIZE_EVEN, ...
                sprintf('write-path word/settings.xml must be exactly %d B', ...
                    testCase.SETTINGS_SIZE_EVEN));
            testCase.verifyEqual(sha256hex(bytes), testCase.SETTINGS_SHA_EVEN, ...
                'write-path settings.xml SHA-256 must equal the frozen s0037 oracle (byte-identical L1)');
        end

        function test_remove_path_byte_identical(testCase)
            % Regression (REMOVE path, L1): setting true THEN false removes the child
            % and returns word/settings.xml byte-for-byte to the M1 default -- EXACTLY
            % 2535 B, SHA-256 51a0d348... (frozen s0038 == s0001). Proves the remove
            % path leaves no residue.
            bytes = emitDocPartMutated(@(d) setEvenTwice(d));
            testCase.verifyEqual(numel(bytes), testCase.SETTINGS_SIZE_DEFAULT, ...
                sprintf('remove-path word/settings.xml must return to exactly %d B', ...
                    testCase.SETTINGS_SIZE_DEFAULT));
            testCase.verifyEqual(sha256hex(bytes), testCase.SETTINGS_SHA_DEFAULT, ...
                'remove-path settings.xml SHA-256 must return to the M1 default 51a0d348... (L1)');
        end

        % =============================================================== %
        % 6. Settings proxy delegation + H5 element identity               %
        % =============================================================== %

        function test_settings_proxy_delegation_and_h5(testCase)
            % Nominal + Edge + Regression (s0036 proxy): the Settings proxy is a
            % leaf-class Settings wrapping a CT_Settings; odd_and_even_pages_header_
            % footer get/set delegates straight to CT_Settings.evenAndOddHeaders_val
            % (default False; set True -> empty element serhex; set False -> bare
            % serhex); two Settings over the SAME element compare == True (H5,
            % ElementProxy element identity).
            oracle = loadOracle();
            el = mat2doc.oxml.OxmlElement("w:settings");
            st = mat2doc.settings.Settings(el);
            testCase.verifyEqual(class(st), 'mat2doc.settings.Settings', ...
                'proxy leaf class must be Settings');
            testCase.verifyEqual(class(st.element), testCase.CT_SETTINGS, ...
                'proxy .element must be a CT_Settings');
            testCase.verifyFalse(st.odd_and_even_pages_header_footer, ...
                'default odd_and_even_pages_header_footer -> False');

            st.odd_and_even_pages_header_footer = true;
            testCase.verifyTrue(st.odd_and_even_pages_header_footer, 'set True delegates -> True');
            testCase.verifyEqual(hx_e(st.element), string(oracle.proxy.after_true_serhex), ...
                'set True on the proxy -> empty <w:evenAndOddHeaders/> (serhex oracle)');

            st.odd_and_even_pages_header_footer = false;
            testCase.verifyFalse(st.odd_and_even_pages_header_footer, 'set False delegates -> False');
            testCase.verifyEqual(hx_e(st.element), string(oracle.proxy.after_false_serhex), ...
                'set False on the proxy -> bare <w:settings/> (serhex oracle)');

            % H5: two Settings over the SAME element compare == True
            testCase.verifyTrue(mat2doc.settings.Settings(el) == mat2doc.settings.Settings(el), ...
                'H5: two Settings over the same CT_Settings element compare == True');
        end

        % =============================================================== %
        % 7. Un-stub resolution (Document / DocumentPart / SettingsPart)    %
        % =============================================================== %

        function test_unstub_resolution_and_h5_chain(testCase)
            % Regression (un-stub, s0036 unstub): the THREE previously-stubbed
            % .settings accessors now RESOLVE to a live Settings proxy over a
            % CT_Settings (no mat2doc:notYetPorted). H5 eq-chain: Document.settings,
            % DocumentPart.settings and SettingsPart.settings all wrap the SAME
            % element, so they compare == True; setting True on the proxy flips a
            % freshly-fetched getter to True (shared underlying element).
            d = mat2doc.Document();

            % Document.settings (document.py 211-214) -> Settings over CT_Settings
            ds = d.settings;
            testCase.verifyEqual(class(ds), 'mat2doc.settings.Settings', ...
                'Document.settings RESOLVES to a Settings proxy (un-stubbed at P5-1)');
            testCase.verifyEqual(class(ds.element), testCase.CT_SETTINGS, ...
                'Document.settings wraps a CT_Settings');
            testCase.verifyFalse(ds.odd_and_even_pages_header_footer, 'default flag False');

            % DocumentPart.settings (parts/document.py 116-120)
            dps = d.part.settings;
            testCase.verifyEqual(class(dps), 'mat2doc.settings.Settings', ...
                'DocumentPart.settings RESOLVES to a Settings proxy (un-stubbed at P5-1)');

            % SettingsPart.settings (parts/settings.py 36-42)
            sp = d.part.settings_part_();
            sps = sp.settings;
            testCase.verifyEqual(class(sps), 'mat2doc.settings.Settings', ...
                'SettingsPart.settings RESOLVES to a Settings proxy (un-stubbed at P5-1)');

            % H5 eq-chain: all three wrap the SAME CT_Settings element
            testCase.verifyTrue(d.settings == d.part.settings, ...
                'H5: Document.settings == DocumentPart.settings (same element)');
            testCase.verifyTrue(d.part.settings == sp.settings, ...
                'H5: DocumentPart.settings == SettingsPart.settings (same element)');

            % mutation through the proxy is visible on a freshly-fetched getter
            d.settings.odd_and_even_pages_header_footer = true;
            testCase.verifyTrue(d.settings.odd_and_even_pages_header_footer, ...
                'set True through the proxy flips a freshly-fetched getter to True (shared element)');
        end

        % =============================================================== %
        % 8. EQUIVALENCE -- full s0036 battery vs the frozen oracle         %
        % =============================================================== %

        function test_equivalence_full_battery_vs_frozen_oracle(testCase)
            % Equivalence: replay the ENTIRE s0036 battery (runProbes -- the .m twin's
            % body VERBATIM: ct_settings / h11 / proxy / unstub + the real
            % settings.xml parse->serialize leg) and flatten-compare EVERY leaf to
            % the frozen python-docx 1.2.0 oracle copied into
            % data\s0036_probe_oracle.json. Gate-3 found ZERO divergences (probe_diff
            % exit 0), so every leaf must be byte/value-identical. Ties the suite to
            % the Gate-3 output.
            port   = runProbes();
            oracle = loadOracle();
            pMap = containers.Map('KeyType','char','ValueType','char');
            oMap = containers.Map('KeyType','char','ValueType','char');
            flattenLeaves(port,   '', pMap);
            flattenLeaves(oracle, '', oMap);

            pKeys = sort(pMap.keys());
            oKeys = sort(oMap.keys());
            testCase.verifyEqual(pKeys, oKeys, ...
                'the replayed battery and the frozen oracle must have identical leaf keys');
            % Non-trivial size guard (guards a silent-empty replay).
            testCase.verifyGreaterThan(numel(oKeys), 25, ...
                'the flattened oracle must expose the full battery of leaves');
            for i = 1:numel(oKeys)
                k = oKeys{i};
                testCase.verifyTrue(isKey(pMap, k), sprintf('port is missing leaf %s', k));
                testCase.verifyEqual(pMap(k), oMap(k), ...
                    sprintf('leaf %s must be byte/value-identical to the frozen oracle', k));
            end
        end

    end
end

% ===================== file-local helpers ============================== %

function s = decl()
    s = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>";
end

function s = W_()
    s = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
end

function s = nsW()
    s = "xmlns:w=""" + W_() + """";
end

function s = nsWM()
    s = "xmlns:w=""" + W_() + """ " + ...
        "xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math""";
end

function e = parse(xml)
    e = mat2doc.oxml.parse_xml(uint8(unicode2native(char(xml), "UTF-8")));
end

function s = ser(e)
    % Decode serialize_part_xml bytes to a string (ASCII content -> string-equality
    % is a byte-identical L1 assertion; non-ASCII round-trips via UTF-8).
    s = string(native2unicode(mat2doc.oxml.serialize_part_xml(e), "UTF-8"));
end

function h = hx(raw)
    % UPPERCASE hex of raw UTF-8 bytes (matches Python bytes.hex().upper()).
    h = string(sprintf('%02X', uint8(raw)));
end

function h = hx_e(e)
    h = hx(mat2doc.oxml.serialize_part_xml(e));
end

function names = childLocalnames(e)
    kids = e.xpath("./*");
    names = strings(1, numel(kids));
    for k = 1:numel(kids)
        names(k) = string(kids(k).local_part);
    end
end

function C = lnsCell(e)
    % Ordered child localnames as a 1xN cell row of char ({} when no children).
    kids = e.xpath("./*");
    if isempty(kids)
        C = cell(1, 0);
        return
    end
    C = cell(1, numel(kids));
    for k = 1:numel(kids)
        C{k} = char(kids(k).local_part);
    end
end

function s = rv(x)
    % Uniform accessor repr mirroring the s0036 rv(): None->"None", bool->
    % "True"/"False", int->decimal string.
    if isequal(x, [])
        s = "None";
    elseif islogical(x)
        if x, s = "True"; else, s = "False"; end
    elseif isnumeric(x)
        s = string(sprintf('%.0f', double(x)));
    else
        s = string(x);
    end
end

function s = cn(x)
    % Leaf class name (package prefix stripped) so py/ml agree.
    full = string(class(x));
    parts = split(full, ".");
    s = parts(end);
end

function s = tf(b)
    if b, s = "True"; else, s = "False"; end
end

function o = loadOracle()
    % Read the co-located frozen oracle in BINARY mode (no CRLF translation) and
    % decode UTF-8 -> struct. jsondecode is whitespace/line-ending agnostic, so no
    % `* binary` .gitattributes pin is needed (value-based fixture, s0030 precedent).
    here = fileparts(mfilename('fullpath'));
    p = fullfile(here, 'data', 's0036_probe_oracle.json');
    fid = fopen(p, 'r', 'n');
    if fid < 0
        error('mat2doc:test:openFail', 'cannot open frozen oracle %s', p);
    end
    raw = fread(fid, Inf, '*uint8')';
    fclose(fid);
    o = jsondecode(native2unicode(raw, 'UTF-8'));
end

function setEven(d, tf_)
    d.settings.odd_and_even_pages_header_footer = tf_;
end

function setEvenTwice(d)
    d.settings.odd_and_even_pages_header_footer = true;
    d.settings.odd_and_even_pages_header_footer = false;
end

function bytes = emitDocPart(partLeaf)
    % Document().save() to a temp .docx, unzip, return word/<partLeaf> raw bytes.
    % Base-MATLAB unzip (no toolbox) into a temp dir, both cleaned up on exit.
    % (Idiom from Test_p4_6_oxml_styles.m; tempname paths are absolute.)
    d = mat2doc.Document();
    bytes = saveAndExtract(d, partLeaf);
end

function bytes = emitDocPartMutated(mutate)
    % Document().save() after applying a mutation closure to the live document,
    % then extract word/settings.xml raw bytes (the write/remove novel paths).
    d = mat2doc.Document();
    mutate(d);
    bytes = saveAndExtract(d, 'settings.xml');
end

function bytes = saveAndExtract(d, partLeaf)
    tmp = [tempname '.docx'];
    cleanTmp = onCleanup(@() deleteIfExists(tmp)); %#ok<NASGU>
    d.save(tmp);
    exdir = tempname;
    cleanEx = onCleanup(@() rmdirIfExists(exdir)); %#ok<NASGU>
    unzip(tmp, exdir);
    bytes = readBytes(fullfile(exdir, 'word', partLeaf));
end

function b = readBytes(p)
    f = fopen(p, 'r', 'n');            % binary read (no CRLF translation)
    assert(f >= 0, 'could not open for read: %s', p);
    b = fread(f, Inf, '*uint8')';
    fclose(f);
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

function h = sha256hex(bytes)
    % Lowercase hex SHA-256 of a uint8 vector (base MATLAB / Java MessageDigest).
    md  = java.security.MessageDigest.getInstance('SHA-256');
    dig = md.digest(typecast(uint8(bytes(:)'), 'int8'));
    b   = typecast(dig, 'uint8');
    h   = lower(string(sprintf('%02x', b)));
end

function P = runProbes()
    % Replay the s0036 probe sequence (the .m twin's P-struct body, VERBATIM
    % tags/inputs/order) and return the nested struct of tagged canonical values.
    % Embedded here so the Equivalence leg is self-contained (the validation-folder
    % scenario is NOT on the toolbox path and must not be a dependency). Mirrors
    % validation\mat2doc\scenarios\s0036_p5_1_settings_probe.m lines 21-118. The
    % parse-path leg reads the settings.xml that Document().save() itself emits
    % (rather than an external file arg), so no fixture path is needed.
    P = struct();

    % ===================== ct_settings: evenAndOddHeaders_val ===============
    cs = struct();
    b = mat2doc.oxml.OxmlElement("w:settings");
    cs.bare = struct("even_val", rv(b.evenAndOddHeaders_val), ...
        "even_child", rv(b.evenAndOddHeaders));

    s = mat2doc.oxml.OxmlElement("w:settings");
    s.evenAndOddHeaders_val = true;
    cs.set_true = struct("serhex", hx_e(s), "localnames", {lnsCell(s)}, ...
        "even_val", rv(s.evenAndOddHeaders_val));

    s = mat2doc.oxml.OxmlElement("w:settings");
    s.evenAndOddHeaders_val = true;
    s.evenAndOddHeaders_val = false;
    cs.set_true_then_false = struct("serhex", hx_e(s), "localnames", {lnsCell(s)}, ...
        "even_val", rv(s.evenAndOddHeaders_val));

    s = mat2doc.oxml.OxmlElement("w:settings");
    s.evenAndOddHeaders_val = true;
    s.evenAndOddHeaders_val = [];        % Python: None -> remove
    cs.set_true_then_none = struct("serhex", hx_e(s), "localnames", {lnsCell(s)}, ...
        "even_val", rv(s.evenAndOddHeaders_val));

    s = mat2doc.oxml.OxmlElement("w:settings");
    s.evenAndOddHeaders_val = 0;         % H4: double 0 is NOT `is False` -> w:val="0"
    cs.h4_int0 = struct("serhex", hx_e(s), "localnames", {lnsCell(s)}, ...
        "even_val", rv(s.evenAndOddHeaders_val));

    p1 = parse("<w:settings " + nsW() + "><w:evenAndOddHeaders/></w:settings>");
    cs.parse_present_no_val = struct("even_val", rv(p1.evenAndOddHeaders_val));
    p2 = parse("<w:settings " + nsW() + "><w:evenAndOddHeaders w:val=""0""/></w:settings>");
    cs.parse_val_0 = struct("even_val", rv(p2.evenAndOddHeaders_val));
    p3 = parse("<w:settings " + nsW() + "><w:evenAndOddHeaders w:val=""true""/></w:settings>");
    cs.parse_val_true = struct("even_val", rv(p3.evenAndOddHeaders_val));
    P.ct_settings = cs;

    % ===================== h11: successor-slice insertion ===================
    h11 = struct();
    both = parse("<w:settings " + nsW() + "><w:defaultTableStyle w:val=""TableNormal""/>" + ...
        "<w:bookFoldRevPrinting/></w:settings>");
    both.get_or_add_evenAndOddHeaders();
    h11.both_neighbors = struct("localnames", {lnsCell(both)}, "serhex", hx_e(both));

    pred = parse("<w:settings " + nsW() + "><w:defaultTableStyle w:val=""TableNormal""/></w:settings>");
    pred.get_or_add_evenAndOddHeaders();
    h11.pred_only_append = struct("localnames", {lnsCell(pred)}, "serhex", hx_e(pred));

    succ = parse("<w:settings " + nsW() + "><w:bookFoldRevPrinting/></w:settings>");
    succ.get_or_add_evenAndOddHeaders();
    h11.succ_only_insert = struct("localnames", {lnsCell(succ)}, "serhex", hx_e(succ));

    pfx = parse("<w:settings " + nsWM() + "><m:mathPr/></w:settings>");
    pfx.get_or_add_evenAndOddHeaders();
    h11.succ_prefixed_mathpr = struct("localnames", {lnsCell(pfx)}, "serhex", hx_e(pfx));
    P.h11 = h11;

    % ===================== proxy: Settings delegation =======================
    px = struct();
    el = mat2doc.oxml.OxmlElement("w:settings");
    st = mat2doc.settings.Settings(el);
    px.class = cn(st);
    px.element_class = cn(st.element);
    px.default = rv(st.odd_and_even_pages_header_footer);
    st.odd_and_even_pages_header_footer = true;
    px.after_true = rv(st.odd_and_even_pages_header_footer);
    px.after_true_serhex = hx_e(st.element);
    st.odd_and_even_pages_header_footer = false;
    px.after_false = rv(st.odd_and_even_pages_header_footer);
    px.after_false_serhex = hx_e(st.element);
    px.eq_same_element = rv(mat2doc.settings.Settings(el) == mat2doc.settings.Settings(el));
    P.proxy = px;

    % ===================== unstub: through a real Document() ================
    us = struct();
    d = mat2doc.Document();
    dsettings = d.settings;
    us.document_settings_class = cn(dsettings);
    us.settings_element_class = cn(dsettings.element);
    us.default_flag = rv(dsettings.odd_and_even_pages_header_footer);
    us.docpart_settings_class = cn(d.part.settings);
    us.eq_doc_docpart = rv(d.settings == d.part.settings);
    d.settings.odd_and_even_pages_header_footer = true;
    us.flag_after_set_true = rv(d.settings.odd_and_even_pages_header_footer);
    P.unstub = us;

    % ===================== parsepath: real settings.xml L1 ==================
    pp = struct();
    sb = emitDocPart('settings.xml');
    elm = mat2doc.oxml.parse_xml(sb);
    out = mat2doc.oxml.serialize_part_xml(elm);
    pp.roundtrip_byte_identical = tf(isequal(uint8(out(:)'), uint8(sb(:)')));
    pp.root_class = cn(elm);
    pp.even_val = rv(elm.evenAndOddHeaders_val);
    P.parsepath = pp;
end

function flattenLeaves(s, prefix, map)
    % Recursively flatten a (possibly nested) struct / cell / array into
    % map(dotted.path) -> canonical string, so a MATLAB-built probe struct and the
    % jsondecode'd oracle compare leaf-by-leaf regardless of container types.
    if isstruct(s)
        fn = fieldnames(s);
        for i = 1:numel(fn)
            if isempty(prefix)
                child = fn{i};
            else
                child = [prefix '.' fn{i}];
            end
            flattenLeaves(s.(fn{i}), child, map);
        end
    elseif iscell(s)
        if isempty(s)
            map(prefix) = ''; %#ok<NASGU>
        else
            parts = strings(1, numel(s));
            for i = 1:numel(s)
                parts(i) = canonScalar(s{i});
            end
            map(prefix) = char(join(parts, "|")); %#ok<NASGU>
        end
    elseif (isstring(s) && ~isscalar(s))
        if isempty(s), map(prefix) = ''; else, map(prefix) = char(join(s(:).', "|")); end %#ok<NASGU>
    elseif isnumeric(s) && ~isscalar(s)
        if isempty(s), map(prefix) = ''; else, map(prefix) = char(join(string(s(:).'), ",")); end %#ok<NASGU>
    else
        map(prefix) = char(canonScalar(s)); %#ok<NASGU>
    end
end

function s = canonScalar(v)
    if islogical(v)
        if v, s = "true"; else, s = "false"; end
    elseif isnumeric(v)
        s = string(num2str(v));
    else
        s = string(v);
    end
end

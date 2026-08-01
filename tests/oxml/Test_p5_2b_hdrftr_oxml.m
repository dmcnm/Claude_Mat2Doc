classdef Test_p5_2b_hdrftr_oxml < matlab.unittest.TestCase
% TEST_P5_2B_HDRFTR_OXML  Gate-4 permanent unit tests for Mat2Doc P5-2b
%   (oxml CT_HdrFtr + the section block-element iterator): src/docx/oxml/section.py
%   -> +mat2doc\+oxml\+section\ CT_HdrFtr (w:hdr / w:ftr root, ZeroOrMore p/tbl +
%   inner_content_elements) and SectBlockElementIterator_ (the ported
%   _SectBlockElementIterator, section.py 437-537); the CT_SectPr.iter_inner_content
%   un-stub (section.py 263-269, now LIVE, delegating to the iterator); and the two
%   registry rows in +mat2doc\+oxml\registry.m (w:hdr / w:ftr -> CT_HdrFtr,
%   oxml/__init__.py:124-125).
%
%   P5-2b is a REGISTRY-ADDING but M1-NEUTRAL WP. w:hdr / w:ftr are the roots of
%   SEPARATE package parts (word/header1.xml, word/footer1.xml, ...), never present
%   in default.docx (which has 17 parts and NO header/footer part; the strings
%   "header"/"footer" do not occur in its [Content_Types].xml and no M1 part carries
%   a <w:hdr>/<w:ftr> tag). So registering the two rows lights up ONLY when a
%   header/footer part is actually loaded (first materialized at P5-3b): byte-neutral
%   AND flip-neutral. No existing exact-class XmlElement pin can see a w:hdr/w:ftr
%   class. This class permanently freezes the guarantees the prior gates established:
%     * Gate-1 Porter  : audit_P5-2b_hdrftr_oxml.md (self-probe).
%     * Gate-2 Auditor (Fable): audit_P5-2b_hdrftr_oxml.md GATE-2 -- APPROVE; the
%       iterator partition arithmetic proven (H1 slice +1, H5 identity find, the
%       mutually-exclusive p-sect/body-sect xpath shapes, the skip-count).
%     * Gate-3 Validator: validate_P5-2b_hdrftr_oxml.md -- PASS, ZERO new D-numbers,
%       NO re-pin list (M1-neutral / flip-neutral). M1 17/17 (document.xml 1548 B /
%       0e4dd503 UNCHANGED); s0042 probe_diff MATCH exit 0 (CT_HdrFtr ICE surface +
%       iter_inner_content resolution); s0043 header-part round-trip byte-identical
%       (1842 B / e6abe568); s0044 multi-section partition corpus MATCH exit 0
%       (4 docs / 11 sections, every edge covered).
%
%   THE HIGHEST-VALUE PERMANENT PINS IN THIS WP (a regression goes RED loudly):
%     * (P) _SectBlockElementIterator PARTITION pins (test_iterator_partition_pins,
%       the crux): over docA/docE/docC/docD each section's partition equals the
%       frozen s0044 structure, AND the CONCATENATION of all partitions == the body's
%       block children, HANDLE-IDENTICAL, in order (the no-drop / no-dup /
%       no-off-by-one invariant, pinned explicitly per document). Edges covered:
%       EMPTY-FIRST section (docE S1 = [p|]); tbl-IN-section / tbl-FIRST (docE S3 =
%       [tbl|S3-T, p|]); tbl-LAST before a break (docE S4); break-paragraph CARRYING
%       content (docE S2 [.. p|S2-break-text]); EMPTY-BODY (last) section (docE S5 =
%       [] -- a 0-LENGTH typed partition, no boundary error at blocks(n+1:end) with
%       n==numel); body-section ENDING in a tbl (docC S2 = [p, tbl|C-S2-T]).
%     * (I) CT_HdrFtr.inner_content_elements (test_ct_hdrftr_inner_content): over a
%       MIXED header (p + tbl + p) returns [CT_P, CT_Tbl, CT_P] in document
%       order -- the tbl INCLUDED (now CT_Tbl, registered P6-3b; class re-pinned),
%       never dropped (pin the classes + order). A <w:p> nested in <w:ins> is SHADED
%       and EXCLUDED (the fixed two-branch CHILD xpath ./w:p|./w:tbl, not a descendant
%       scan). Non-ASCII text (café-Ñ—end / 中文) UTF-8-exact. H5 two-read identity.
%     * (R) Header-part round-trip byte-pin (test_header_part_roundtrip_byte_pin):
%       the FROZEN s0043 real header1.xml (paragraph + 1x2 table + paragraph, non-ASCII
%       in both a paragraph and a cell) parsed through CT_HdrFtr, iterated, and
%       re-serialized is BYTE-IDENTICAL to the input (1842 B, SHA-256 e6abe568...) ==
%       python-docx's OWN parse->serialize round-trip (D-001 own-parser proof extended
%       to a real header PART). The frozen s0043 bytes are copied into data\s0043\.
%     * (U) CT_SectPr.iter_inner_content un-stub (test_ct_sectpr_iter_inner_content_unstub):
%       resolves (no notYetPorted), yields the correct partition, and is
%       HANDLE-IDENTICAL to SectBlockElementIterator_.iter_sect_block_elements(sectPr)
%       (the delegation is exact).
%     * (V) iterator ValueError contract (test_iterator_valueerror_detached_sectpr):
%       a sectPr NOT in the document's sectPrs list raises the IDENTIFIER
%       mat2doc:ValueError (Python list.index() contract, H5 identity search).
%     * (M) M1 document.xml byte-pin (test_m1_document_xml_byte_identical): the
%       registry-neutrality guard -- mat2doc.Document().save() -> word/document.xml
%       == 1548 B, SHA-256 0e4dd503... UNCHANGED despite the two new registry rows.
%
%   Provenance (all 2026-07-31):
%     * Audit    : validation\mat2doc\audit_P5-2b_hdrftr_oxml.md
%     * Validate : validation\mat2doc\validate_P5-2b_hdrftr_oxml.md
%     * Scenarios: validation\mat2doc\scenarios\s0042_p5_2b_hdrftr_probe.{py,m}
%                  (its ICE + iter_res probe body is replayed VERBATIM below);
%                  s0043_p5_2b_header_roundtrip.{py,m} (header byte oracle);
%                  s0044_p5_2b_multisect_partition.{py,m} (partition corpus).
%     * Frozen refs (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE at Gate-3):
%         references\s0042\probe.json -> copied verbatim (self-contained) into
%           tests\oxml\data\s0042_probe_oracle.json (the ICE/iter Equivalence set).
%         references\s0044\probe.json -> tests\oxml\data\s0044_probe_oracle.json
%           (the partition-structure Equivalence set).
%         references\s0043\probe.json -> tests\oxml\data\s0043_probe_oracle.json
%           (round-trip SHA + ice_sigs). All three are value/serhex JSON; jsondecode
%           is line-ending agnostic -> NO `* binary` .gitattributes pin needed
%           (s0030/s0036/s0039 precedent).
%         references\s0043\header1.xml (1842 B / e6abe568) -> copied verbatim to
%           tests\oxml\data\s0043\header1.xml WITH a co-located `.gitattributes`
%           `* binary` (byte fixture -- MANDATORY pin so the master checkout does not
%           mangle it; msoffice-byte-fixture-gitattributes lesson).
%         references\s0001\parts\word\document.xml -- the M1 byte reference (1548 /
%           0e4dd503); NOT copied (SHA of what Document().save() emits).
%
%   Coverage taxonomy
%   -----------------
%   * Nominal      -- CT_HdrFtr root class + ICE over p_only / tbl_only / mixed;
%                     registry w:hdr/w:ftr resolution; CT_SectPr.iter_inner_content
%                     resolves and partitions docA.
%   * Edge         -- ICE empty (0-length typed), w:ins-SHADED exclusion, non-ASCII
%                     (é/Ñ/—/CJK) UTF-8-exact; iterator EMPTY-FIRST / EMPTY-BODY
%                     (0-length partition, no boundary error) / tbl-FIRST / tbl-LAST /
%                     break-with-content; the mat2doc:ValueError error path on a
%                     detached sectPr; single-section (docD) body-sect.
%   * Equivalence  -- test_equivalence_hdrftr_vs_frozen_oracle (replays the s0042 ICE
%                     + iter_res battery, flatten-compares every leaf to
%                     data\s0042_probe_oracle.json) and
%                     test_equivalence_partitions_vs_frozen_oracle (replays the s0044
%                     4-doc corpus vs data\s0044_probe_oracle.json), 0 diffs.
%   * Regression   -- hard-coded expected localname|text signature arrays (transcribed
%                     from the frozen s0044/s0042 oracle) + SHA-256 of the M1 and the
%                     header-part round-trip bytes + the fixed ICE element classes.
%   * Upstream     -- the CT_HdrFtr ICE ordering, the w:ins shading exclusion, and the
%                     _SectBlockElementIterator boundary partition ARE the python-docx
%                     section.py surface; the frozen oracle IS lxml's expected output.
%
%   Byte-level (L1) note: the header-part round-trip and M1 pins are SHA-256 of the
%   emitted/round-tripped bytes (L1). The ICE/partition pins are localname|descendant-
%   text signature arrays + handle identity + element class -- the equivalence-relevant
%   order/text (the CLASS of a <w:tbl> differs by design: python-docx CT_Tbl vs the
%   port's generic XmlElement, CT_Tbl being P6-unregistered; that class-name difference
%   is asserted, NOT treated as a divergence, exactly as Gate-3 dispositioned). NO
%   D-number granted any L2 relaxation in this WP (Gate-3: zero new). The equivalence
%   leaf/partition comparisons are the only looser-than-byte checks and are commented
%   at their sites (signature-level, floor-guarded on leaf count).
%
%   Determinism: no network, no absolute paths. Fixtures resolve relative to this file
%   via fileparts(mfilename('fullpath')); the header byte fixture and JSON oracles live
%   in data\; the M1 save goes to tempname .docx deleted via onCleanup; every file read
%   is binary ('r','n'). The +mat2doc package resolves via the MANDATORY
%   PathFixture(worktree-root) in TestClassSetup (WP9-F4 lesson).

    properties (Constant)
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        % --- registered class names (P5-2b + prior) ---
        CT_HDRFTR = 'mat2doc.oxml.section.CT_HdrFtr'
        CT_SECTPR = 'mat2doc.oxml.section.CT_SectPr'
        CT_P      = 'mat2doc.oxml.text.CT_P'
        CT_TBL    = 'mat2doc.oxml.table.CT_Tbl'  % a <w:tbl> parses here (registered P6-3b)

        % --- frozen s0001 M1 word/document.xml byte reference (registry-neutrality) ---
        DOC_SIZE_M1 = 1548
        DOC_SHA_M1  = "0e4dd503bc095219cd24e8312b372c588ca56bb834aacac53eafb7bf47836327"

        % --- frozen s0043 real header-part round-trip byte reference ---
        HDR_SIZE = 1842
        HDR_SHA  = "e6abe5685afb0e53eb1d80b0747ffe09b102c486947524a0a62ba487b397fcb2"
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run cannot
            % resolve the +mat2doc package (MATLAB:undefinedVarOrClass). Idiom copied
            % verbatim from tests\oxml\Test_p5_2a_sectpr.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\oxml
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. registry rows (the 2 P5-2b flips)                             %
        % =============================================================== %

        function test_registry_resolves_hdr_ftr_rows(testCase)
            % Nominal / Regression (registry.m, oxml/__init__.py:124-125): BOTH
            % w:hdr AND w:ftr resolve to CT_HdrFtr, and OxmlElement builds the exact
            % class. Registration is what flips the parse class of a header/footer
            % PART root (M1-neutral until such a part is loaded, P5-3b).
            for tag = ["w:hdr" "w:ftr"]
                r = mat2doc.oxml.registry(mat2doc.oxml.qn(tag));
                testCase.verifyEqual(char(r), testCase.CT_HDRFTR, ...
                    sprintf('registry must resolve %s -> CT_HdrFtr (P5-2b row)', tag));
                testCase.verifyEqual(class(mat2doc.oxml.OxmlElement(tag)), testCase.CT_HDRFTR, ...
                    sprintf('OxmlElement(%s) must be a CT_HdrFtr', tag));
            end
        end

        % =============================================================== %
        % 2. (I) CT_HdrFtr.inner_content_elements                          %
        % =============================================================== %

        function test_ct_hdrftr_inner_content(testCase)
            % (I) Nominal + Edge + Regression (s0042 hdrftr): ICE over the 6 header
            % shapes -- root class, count, localname|text signatures, and the H5
            % two-read identity. The MIXED case pins the element CLASSES
            % [CT_P, CT_Tbl, CT_P] (tbl INCLUDED, now CT_Tbl after the P6-3b re-pin,
            % in position); ins_shaded pins the w:ins EXCLUSION; nonascii pins
            % UTF-8-exact text.

            % ---- empty: 0-length TYPED array (never [] / None) ----
            h = parseXml(hdr("", nsW()));
            testCase.verifyEqual(class(h), testCase.CT_HDRFTR, 'w:hdr root is CT_HdrFtr');
            ice = h.inner_content_elements;
            testCase.verifyEqual(numel(ice), 0, 'empty header -> 0 inner-content elements');
            testCase.verifyEqual(sigRow(ice), strings(1,0), 'empty header -> no signatures');

            % ---- p_only / tbl_only ----
            h = parseXml(hdr(pp("H1"), nsW()));
            testCase.verifyEqual(sigRow(h.inner_content_elements), "p|H1", 'p_only signature');
            h = parseXml(hdr(tbl("T"), nsW()));
            ice = h.inner_content_elements;
            testCase.verifyEqual(sigRow(ice), "tbl|T", 'tbl_only signature');
            testCase.verifyEqual(class(ice(1)), testCase.CT_TBL, ...
                'a lone <w:tbl> now resolves to CT_Tbl (registered P6-3b; registry-flip re-pin), still INCLUDED');

            % ---- mixed: [CT_P, CT_Tbl, CT_P] in document order ----
            h = parseXml(hdr(pp("H-P1") + tbl("H-T00") + pp("H-P2"), nsW()));
            ice = h.inner_content_elements;
            testCase.verifyEqual(numel(ice), 3, 'mixed -> 3 inner-content elements');
            testCase.verifyEqual(sigRow(ice), ["p|H-P1" "tbl|H-T00" "p|H-P2"], ...
                'mixed signatures in document order (tbl INCLUDED in position)');
            testCase.verifyEqual( ...
                {class(ice(1)), class(ice(2)), class(ice(3))}, ...
                {testCase.CT_P, testCase.CT_TBL, testCase.CT_P}, ...
                'mixed classes: [CT_P, CT_Tbl, CT_P] -- tbl now CT_Tbl (P6-3b re-pin), still included in position, never dropped');
            % H5 two-read identity: two reads return the SAME live handles
            ice2 = h.inner_content_elements;
            testCase.verifyTrue(numel(ice) == numel(ice2) && all(ice == ice2), ...
                'H5: two ICE reads yield the SAME element handles');

            % ---- ins_shaded: a <w:p> nested in <w:ins> is EXCLUDED (fixed CHILD xpath) ----
            h = parseXml(hdr(pp("H1") + "<w:ins>" + pp("INSIDE") + "</w:ins>" + tbl("T"), nsW()));
            ice = h.inner_content_elements;
            testCase.verifyEqual(numel(ice), 2, 'ins_shaded -> 2 (the w:ins-nested <w:p> is SHADED)');
            testCase.verifyEqual(sigRow(ice), ["p|H1" "tbl|T"], ...
                'ins_shaded: the INSIDE <w:p> nested in <w:ins> is EXCLUDED (./w:p|./w:tbl is a CHILD union, not a descendant scan)');

            % ---- nonascii: UTF-8-exact (café-Ñ—end / 中文), built by codepoint ----
            paraTxt = char([99 97 102 233 45 209 8212 101 110 100]);   % café-Ñ—end
            cellTxt = char([20013 25991]);                             % 中文
            h = parseXml(hdr(pp(string(paraTxt)) + tbl(string(cellTxt)), nsW()));
            ice = h.inner_content_elements;
            testCase.verifyEqual(numel(ice), 2, 'nonascii -> 2 elements');
            testCase.verifyEqual(sigRow(ice), ["p|" + string(paraTxt), "tbl|" + string(cellTxt)], ...
                'nonascii signatures UTF-8-exact (é/Ñ/—/CJK round-trip through parse+xpath)');
        end

        % =============================================================== %
        % 3. (P) _SectBlockElementIterator partition pins (THE crux)       %
        % =============================================================== %

        function test_iterator_partition_pins(testCase)
            % (P) Regression (the highest-value ordering pins, s0044): each section's
            % partition equals the frozen structure, AND per document the concatenation
            % of ALL partitions == the body's block children HANDLE-IDENTICAL, in order
            % (no-drop / no-dup / no-off-by-one). Hard-coded expected partitions are
            % transcribed VERBATIM from the frozen references\s0044\probe.json oracle.
            %   docE covers EMPTY-FIRST / break-with-content / tbl-FIRST / tbl-LAST /
            %   EMPTY-BODY (0-length partition, no boundary error). docC ends in a tbl.
            %   docD is a single body-sect section.

            % ---- docA: realistic 3-section, tbl in S1 ----
            expA = { ["p|A-S1-P1" "p|A-S1-P2" "tbl|A-S1-T" "p|"], ...
                     ["p|A-S2-P1" "p|"], ...
                     ["p|A-S3-P1" "p|A-S3-P2"] };
            testCase.assertPartition(docA(), expA, 'docA');

            % ---- docE: 5-section edge stack (empty-first / content-break / tbl-first
            %      / tbl-last / EMPTY-BODY) ----
            expE = { "p|", ...                                     % S1 EMPTY-FIRST
                     ["p|S2-P1" "tbl|S2-T" "p|S2-break-text"], ... % S2 break carries content
                     ["tbl|S3-T" "p|"], ...                        % S3 tbl FIRST
                     ["p|S4-P1" "tbl|S4-T" "p|"], ...              % S4 tbl LAST before break
                     strings(1,0) };                               % S5 EMPTY-BODY (0-length, no boundary error)
            testCase.assertPartition(docE(), expE, 'docE');

            % ---- docC: 2-section, body section ENDING in a tbl (tbl-in-section) ----
            expC = { ["p|C-S1-P1" "p|"], ...
                     ["p|C-S2-P1" "tbl|C-S2-T"] };
            testCase.assertPartition(docC(), expC, 'docC');

            % ---- docD: single body-sect section ----
            expD = { ["p|D-P1" "p|D-P2"] };
            testCase.assertPartition(docD(), expD, 'docD');
        end

        % =============================================================== %
        % 4. (U) CT_SectPr.iter_inner_content un-stub + delegation          %
        % =============================================================== %

        function test_ct_sectpr_iter_inner_content_unstub(testCase)
            % (U) Nominal + Regression (s0042 iter_res): iter_inner_content RESOLVES
            % (no notYetPorted), yields the correct partition, and is HANDLE-IDENTICAL
            % to SectBlockElementIterator_.iter_sect_block_elements(sectPr) (the
            % delegation is exact -- section.py 263-269).
            root = parseXml(docA());
            sectPrs = root.xpath("/w:document/w:body/w:p/w:pPr/w:sectPr | /w:document/w:body/w:sectPr");
            testCase.verifyEqual(numel(sectPrs), 3, 'docA has 3 sectPrs');
            testCase.verifyEqual(class(sectPrs(1)), testCase.CT_SECTPR, 'a break sectPr parses as CT_SectPr');

            % resolves (no error) and yields the S1 partition
            items = sectPrs(1).iter_inner_content();
            testCase.verifyEqual(sigRow(items), ["p|A-S1-P1" "p|A-S1-P2" "tbl|A-S1-T" "p|"], ...
                'iter_inner_content yields the S1 partition (un-stubbed, live)');

            % delegation: identical to the static iterator entry, handle-for-handle
            direct = mat2doc.oxml.section.SectBlockElementIterator_ ...
                .iter_sect_block_elements(sectPrs(1));
            testCase.verifyTrue(numel(items) == numel(direct) && all(items == direct), ...
                'iter_inner_content delegates HANDLE-IDENTICALLY to the iterator helper');
        end

        function test_iterator_valueerror_detached_sectpr(testCase)
            % (V) Edge / error path: a sectPr NOT among the document's sectPrs (a
            % detached loose <w:sectPr>, its own root, no w:document ancestor) makes
            % the H5 identity search find(sectPrs == sectPr) empty -> the iterator
            % raises the IDENTIFIER mat2doc:ValueError (Python list.index() contract,
            % section.py 462 / SectBlockElementIterator_ line 119). Unreachable in a
            % real document; pinned here to lock the contract.
            loose = parseXml("<w:sectPr " + nsW() + "></w:sectPr>");
            testCase.verifyEqual(class(loose), testCase.CT_SECTPR, 'loose <w:sectPr> is CT_SectPr');
            testCase.verifyError(@() loose.iter_inner_content(), 'mat2doc:ValueError', ...
                'a sectPr not in the document raises mat2doc:ValueError (list.index contract)');
        end

        % =============================================================== %
        % 5. (R) Header-part round-trip byte-pin (frozen s0043)             %
        % =============================================================== %

        function test_header_part_roundtrip_byte_pin(testCase)
            % (R) Regression (byte-identical L1): the FROZEN real header1.xml parsed
            % through CT_HdrFtr, iterated (ICE exercised), re-serialized, is
            % BYTE-IDENTICAL to the input (1842 B, SHA-256 e6abe568...). Extends the
            % D-001 own-parser round-trip proof to a real header PART (fresh input
            % class). The frozen s0043 bytes are the co-located data\s0043\header1.xml.
            inBytes = readBytes(fullfile(dataDir(), 's0043', 'header1.xml'));
            testCase.verifyEqual(numel(inBytes), testCase.HDR_SIZE, ...
                'frozen header1.xml must be exactly 1842 B');
            testCase.verifyEqual(sha256hex(inBytes), testCase.HDR_SHA, ...
                'frozen header1.xml SHA-256 == the s0043 oracle (fixture integrity)');

            root = mat2doc.oxml.parse_xml(inBytes);
            testCase.verifyEqual(class(root), testCase.CT_HDRFTR, 'header part root parses as CT_HdrFtr');

            ice = root.inner_content_elements;           % exercise the ICE path
            testCase.verifyEqual(sigRow(ice), ...
                ["p|S0043-HDR-café" "tbl|C00-中文/C01" "p|S0043-HDR-P2"], ...
                'header ICE signatures (non-ASCII in a paragraph AND a cell) == s0043 oracle');

            outBytes = mat2doc.oxml.serialize_part_xml(root);
            testCase.verifyEqual(numel(outBytes), testCase.HDR_SIZE, ...
                'round-tripped header must be exactly 1842 B');
            testCase.verifyEqual(uint8(outBytes(:)'), uint8(inBytes(:)'), ...
                'header-part parse->serialize BYTE-IDENTICAL to the input (L1)');
            testCase.verifyEqual(sha256hex(outBytes), testCase.HDR_SHA, ...
                'round-tripped header SHA-256 == the frozen input == python-docx own round-trip (L1)');
        end

        % =============================================================== %
        % 6. (M) M1 document.xml byte-pin (registry-neutrality guard)       %
        % =============================================================== %

        function test_m1_document_xml_byte_identical(testCase)
            % (M) Regression (byte-neutrality, L1): registering w:hdr/w:ftr must NOT
            % perturb M1. mat2doc.Document().save() -> word/document.xml == 1548 B with
            % the frozen s0001 SHA-256 0e4dd503... -- UNCHANGED. SHA-256 equality is an
            % L1 assertion. default.docx has no header/footer part and no M1 part
            % carries a <w:hdr>/<w:ftr>, so the two new rows never touch an M1 path.
            bytes = emitDocumentXml();
            testCase.verifyEqual(numel(bytes), testCase.DOC_SIZE_M1, ...
                sprintf('word/document.xml must be exactly %d B after the hdr/ftr registry rows', ...
                    testCase.DOC_SIZE_M1));
            testCase.verifyEqual(sha256hex(bytes), testCase.DOC_SHA_M1, ...
                'word/document.xml SHA-256 must equal the frozen s0001 oracle (byte-identical L1, registry-neutral)');
        end

        % =============================================================== %
        % 7. EQUIVALENCE -- replay vs the frozen oracle (self-contained)     %
        % =============================================================== %

        function test_equivalence_hdrftr_vs_frozen_oracle(testCase)
            % Equivalence: replay the s0042 ICE battery over the 6 header shapes plus
            % the iter_res resolution, and compare EVERY leaf to the frozen python-docx
            % 1.2.0 oracle copied into data\s0042_probe_oracle.json. Gate-3 found ZERO
            % divergences (probe_diff exit 0). Signature-level (localname|text) rather
            % than byte -- the tbl CLASS difference is by design (Gate-3 dispositioned);
            % order/text carry the equivalence and must match exactly.
            o = loadOracleJson('s0042_probe_oracle.json');

            % ---- hdrftr: 6 shapes ----
            shapes = struct( ...
                'empty',      hdr("", nsW()), ...
                'p_only',     hdr(pp("H1"), nsW()), ...
                'tbl_only',   hdr(tbl("T"), nsW()), ...
                'mixed',      hdr(pp("H-P1") + tbl("H-T00") + pp("H-P2"), nsW()), ...
                'ins_shaded', hdr(pp("H1") + "<w:ins>" + pp("INSIDE") + "</w:ins>" + tbl("T"), nsW()), ...
                'nonascii',   hdr(pp(string(char([99 97 102 233 45 209 8212 101 110 100]))) + ...
                                  tbl(string(char([20013 25991]))), nsW()));
            keys = fieldnames(shapes);
            for i = 1:numel(keys)
                k = keys{i};
                root = parseXml(shapes.(k));
                ice  = root.inner_content_elements;
                ice2 = root.inner_content_elements;
                ok = o.hdrftr.(k);
                testCase.verifyTrue(isa(root, 'mat2doc.oxml.section.CT_HdrFtr'), ...
                    sprintf('%s: is_hdrftr', k));
                testCase.verifyEqual(numel(ice), double(ok.n), sprintf('%s: n == oracle', k));
                testCase.verifyEqual(sigRow(ice), rowStr(ok.sigs), sprintf('%s: signatures == oracle', k));
                testCase.verifyTrue(numel(ice) == numel(ice2) && (isempty(ice) || all(ice == ice2)), ...
                    sprintf('%s: two_read_identity', k));
            end

            % ---- iter_res: partition of docA via iter_inner_content ----
            root = parseXml(docA());
            sectPrs = root.xpath("/w:document/w:body/w:p/w:pPr/w:sectPr | /w:document/w:body/w:sectPr");
            testCase.verifyEqual(numel(sectPrs), double(o.iter_res.n_sections), ...
                'iter_res: n_sections == oracle');
            testCase.verifyTrue(logical(o.iter_res.resolves), 'iter_res oracle resolves == true');
            expParts = oraclePartitions(o.iter_res);
            for i = 1:numel(sectPrs)
                testCase.verifyEqual(sigRow(sectPrs(i).iter_inner_content()), expParts{i}, ...
                    sprintf('iter_res: section %d partition == oracle', i));
            end
        end

        function test_equivalence_partitions_vs_frozen_oracle(testCase)
            % Equivalence: replay the s0044 4-document / 11-section partition corpus
            % and compare each section partition, the section count, the
            % concat-eq-body invariant, and the empty-partition flag to the frozen
            % oracle copied into data\s0044_probe_oracle.json (probe_diff exit 0 at
            % Gate-3). Floor-guarded on document count.
            o = loadOracleJson('s0044_probe_oracle.json');
            docs = struct('docA', docA(), 'docE', docE(), 'docC', docC(), 'docD', docD());
            names = fieldnames(docs);
            testCase.verifyGreaterThanOrEqual(numel(names), 4, ...
                'the corpus must expose all 4 adversarial documents');
            for n = 1:numel(names)
                nm = names{n};
                od = o.(nm);
                [parts, concatEq, hasEmpty] = partitionDoc(docs.(nm));
                testCase.verifyEqual(numel(parts), double(od.n_sections), ...
                    sprintf('%s: n_sections == oracle', nm));
                testCase.verifyTrue(concatEq, ...
                    sprintf('%s: concat of partitions == body block children (handle-identical)', nm));
                testCase.verifyEqual(logical(hasEmpty), logical(od.has_empty_partition), ...
                    sprintf('%s: has_empty_partition == oracle', nm));
                testCase.verifyTrue(logical(od.concat_eq_body), ...
                    sprintf('%s: oracle asserts concat_eq_body', nm));
                exp = oraclePartitions(od);
                for i = 1:numel(parts)
                    testCase.verifyEqual(parts{i}, exp{i}, ...
                        sprintf('%s: section %d partition == oracle', nm, i));
                end
            end
        end

    end

    % ===================== instance-scoped assertion ==================== %
    methods (Access = private)
        function assertPartition(testCase, xml, expected, label)
            % Assert the per-section partitions of a document XML equal `expected`
            % (a 1xK cell of string rows), AND the no-drop/no-dup concat invariant:
            % the concatenation of all partitions == the body's block children,
            % HANDLE-IDENTICAL, in document order.
            [parts, concatEq, ~] = partitionDoc(xml);
            testCase.verifyEqual(numel(parts), numel(expected), ...
                sprintf('%s: section count', label));
            for i = 1:numel(expected)
                testCase.verifyEqual(parts{i}, expected{i}, ...
                    sprintf('%s: section %d partition', label, i));
            end
            % *** the concat invariant (no-drop / no-dup / no-off-by-one) ***
            testCase.verifyTrue(concatEq, ...
                sprintf('%s: concat of all partitions == body block children, handle-identical, in order', label));
        end
    end
end

% ===================== file-local helpers ============================== %

function s = W_()
    s = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
end

function s = nsW()
    s = "xmlns:w=""" + W_() + """";
end

function e = parseXml(xml)
    e = mat2doc.oxml.parse_xml(uint8(unicode2native(char(xml), "UTF-8")));
end

% ---- XML builders (mirror the s0042/s0044 scenario twins VERBATIM) ----

function s = pp(text)
    if strlength(text) == 0
        s = "<w:p></w:p>";
    else
        s = "<w:p><w:r><w:t>" + text + "</w:t></w:r></w:p>";
    end
end

function s = pbrk()
    s = "<w:p><w:pPr><w:sectPr/></w:pPr></w:p>";
end

function s = tbl(text)
    s = "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>" + text + ...
        "</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";
end

function s = hdr(inner, W)
    s = "<w:hdr " + W + ">" + inner + "</w:hdr>";
end

function s = document(inner, W)
    s = "<w:document " + W + "><w:body>" + inner + "</w:body></w:document>";
end

% ---- the frozen corpus documents (s0044 twin, VERBATIM) ----

function x = docA()
    x = document( ...
        pp("A-S1-P1") + pp("A-S1-P2") + tbl("A-S1-T") + pbrk() + ...
        pp("A-S2-P1") + pbrk() + ...
        pp("A-S3-P1") + pp("A-S3-P2") + "<w:sectPr/>", nsW());
end

function x = docE()
    x = document( ...
        pbrk() + ...                                                            % break1 -> S1 EMPTY
        pp("S2-P1") + tbl("S2-T") + ...
        "<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>S2-break-text</w:t></w:r></w:p>" + ... % break2 WITH content
        tbl("S3-T") + ...                                                       % tbl FIRST in S3
        pbrk() + ...                                                            % break3
        pp("S4-P1") + tbl("S4-T") + ...                                         % tbl LAST before break4
        pbrk() + ...                                                            % break4
        "<w:sectPr/>", nsW());                                                  % body sectPr -> S5 EMPTY
end

function x = docC()
    x = document( ...
        pp("C-S1-P1") + pbrk() + ...
        pp("C-S2-P1") + tbl("C-S2-T") + "<w:sectPr/>", nsW());
end

function x = docD()
    x = document(pp("D-P1") + pp("D-P2") + "<w:sectPr/>", nsW());
end

% ---- partition + signature machinery (mirror partitionDoc / sigOf) ----

function [parts, concatEq, hasEmpty] = partitionDoc(xml)
    % Per-section partitions (via the un-stubbed CT_SectPr.iter_inner_content ->
    % SectBlockElementIterator_), the concat-eq-body invariant, and the
    % empty-partition flag. Mirrors the s0044 scenario twin partitionDoc.
    root = parseXml(xml);
    sectPrs = root.xpath("/w:document/w:body/w:p/w:pPr/w:sectPr | /w:document/w:body/w:sectPr");
    parts = cell(1, numel(sectPrs));
    yielded = mat2doc.oxml.XmlElement.empty(1, 0);
    hasEmpty = false;
    for i = 1:numel(sectPrs)
        items = sectPrs(i).iter_inner_content();
        parts{i} = sigRow(items);
        yielded = [yielded, reshape(items, 1, [])]; %#ok<AGROW>
        if isempty(items)
            hasEmpty = true;
        end
    end
    body = root.xpath("/w:document/w:body");
    bkids = body.xpath("./*");
    bblocks = bkids(arrayfun(@(e) any(localOf(e) == ["p" "tbl"]), bkids));
    concatEq = numel(yielded) == numel(bblocks) && ...
               (isempty(yielded) || all(yielded == bblocks));
end

function r = sigRow(elems)
    % 1xN string row of localname|descendant-w:t-texts (mirrors the scenario sigOf;
    % 1x0 string for an empty array).
    n = numel(elems);
    if n == 0
        r = strings(1, 0);
        return
    end
    r = strings(1, n);
    for k = 1:n
        r(k) = sigOf(elems(k));
    end
end

function s = sigOf(e)
    s = localOf(e);
    ts = e.xpath(".//w:t");
    parts = strings(1, numel(ts));
    for j = 1:numel(ts)
        t = ts(j).text;
        if isequal(t, [])
            t = "";
        end
        parts(j) = string(t);
    end
    s = s + "|" + strjoin(parts, "/");
end

function s = localOf(e)
    tg = string(e.tag);
    if contains(tg, "}")
        s = extractAfter(tg, "}");
    elseif contains(tg, ":")
        s = extractAfter(tg, ":");
    else
        s = tg;
    end
end

% ---- oracle loading + partition normalization ----

function d = dataDir()
    d = fullfile(fileparts(mfilename('fullpath')), 'data');
end

function o = loadOracleJson(name)
    % Read a co-located frozen oracle in BINARY mode (no CRLF translation) and decode
    % UTF-8 -> struct. jsondecode is whitespace/line-ending agnostic (value-based
    % fixture; s0030/s0036/s0039 precedent -- no `* binary` pin needed for JSON).
    p = fullfile(dataDir(), name);
    b = readBytes(p);
    o = jsondecode(native2unicode(b, 'UTF-8'));
end

function P = oraclePartitions(docS)
    % Normalize a jsondecode'd `partitions` field into a 1xK cell of 1xL string rows,
    % robust to jsondecode's shape heuristics (cell for ragged inner lengths; a KxL
    % string matrix when all inner lengths are equal; a bare column vector when the
    % outer length is 1). n_sections (K) disambiguates the outer count.
    K = double(docS.n_sections);
    pf = docS.partitions;
    P = cell(1, K);
    if iscell(pf)
        for i = 1:K
            P{i} = rowStr(pf{i});
        end
    elseif isempty(pf)
        for i = 1:K
            P{i} = strings(1, 0);
        end
    else
        S = string(pf);
        if K == 1
            P{1} = reshape(S, 1, []);              % single partition (outer dim dropped)
        else
            for i = 1:K
                P{i} = reshape(S(i, :), 1, []);    % rows are partitions (KxL matrix)
            end
        end
    end
end

function r = rowStr(x)
    % jsondecode leaf array -> 1xN string row; 1x0 for [] / empty.
    if isempty(x)
        r = strings(1, 0);
    else
        r = reshape(string(x), 1, []);
    end
end

% ---- M1 document.xml emit + byte utilities ----

function bytes = emitDocumentXml()
    % Document().save() to a temp .docx, unzip, return word/document.xml raw bytes.
    d = mat2doc.Document();
    tmp = [tempname '.docx'];
    cleanTmp = onCleanup(@() deleteIfExists(tmp)); %#ok<NASGU>
    d.save(tmp);
    exdir = tempname;
    cleanEx = onCleanup(@() rmdirIfExists(exdir)); %#ok<NASGU>
    unzip(tmp, exdir);
    bytes = readBytes(fullfile(exdir, 'word', 'document.xml'));
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

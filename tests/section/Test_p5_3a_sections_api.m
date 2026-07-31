classdef Test_p5_3a_sections_api < matlab.unittest.TestCase
% TEST_P5_3A_SECTIONS_API  Gate-4 permanent unit tests for Mat2Doc P5-3a
%   (the Section / Sections proxy API + Document.add_section): src/docx/section.py
%   -> +mat2doc\+section\{Section, Sections}, src/docx/document.py::Document
%   {sections, add_section, _block_width} -> +mat2doc\+document\Document, and
%   src/docx/oxml/document.py::CT_Body.add_section_break ->
%   +mat2doc\+oxml\+document\CT_Body.
%
%   P5-3a is a PURE API/PROXY-TIER WP layered on the P5-2a byte-critical CT_SectPr
%   core: Section delegates ONE-TO-ONE to CT_SectPr accessors, Sections is a plain
%   Sequence over CT_Document.sectPr_lst, and add_section clones the sentinel
%   sectPr into a new trailing paragraph (CT_Body.add_section_break). It adds NO
%   oxml registry rows and NO new serialization code, so equivalence is BEHAVIORAL
%   (probe value parity, proven at Gate-3 by probe_diff s0045 MATCH exit 0 over the
%   full surface) PLUS serialized-bytes parity on the output-visible whole-part
%   add_section / property-write paths PLUS the M1 neutrality guard. This class
%   permanently FREEZES that surface -- byte/value-identical to python-docx 1.2.0.
%
%   THE HIGHEST-VALUE PERMANENT PINS IN THIS WP (a regression goes RED loudly):
%     * add_section(NEW_PAGE) whole-part byte pin (test_add_section_byte_pins):
%       Document(); add_section(NEW_PAGE); save() -> word/document.xml == EXACTLY
%       1792 B, SHA-256 4d49edc1...d2bda, byte-identical to python-docx (frozen
%       s0046). All five WD_SECTION start_types pinned against the frozen s0047
%       member SHAs; the ODD_PAGE+EVEN_PAGE THREE-section chain -> 2087 B /
%       f6cc256d...5f6eb (frozen s0048, the multi-section nesting guard).
%     * Section-property-write byte pins (test_section_property_write_byte_pins):
%       sections[0] geometry/orientation/page_width writes -> 1569 B / 5f3a0d58...
%       (frozen s0049); the full-surface write on sections[0]+the new sentinel
%       section -> 1893 B / b6c6af5e... (frozen s0050).
%     * M1 document.xml byte pin (test_m1_document_xml_byte_pin): a bare
%       Document().save() STILL emits word/document.xml at EXACTLY 1548 B /
%       0e4dd503... -- the un-stub-neutrality guard (un-stubbing sections /
%       add_section touches no default-save byte).
%     * 0-based Sequence indexing (test_sections_sequence): getitem_(0) is the
%       FIRST section and getitem_(-1) is the LAST (the TabStops 0-based
%       convention); out-of-range -> mat2doc:IndexError "list index out of range";
%       a zero slice step -> mat2doc:ValueError "slice step cannot be zero".
%     * iter_inner_content w:tbl guard (test_iter_inner_content): a <w:tbl> block
%       raises mat2doc:notYetPorted (owner P6-4a) and is NEVER silently dropped; a
%       <w:p> resolves to a Paragraph.
%     * P5-3b hdr/ftr stubs (test_p5_3b_stubs): all SIX header/footer members stay
%       mat2doc:notYetPorted (owner P5-3b).
%
%   Provenance (Gate-1..3, all 2026-07-31):
%     * Audit    : validation\mat2doc\audit_P5-3a_sections_api.md (Porter Gate-1
%                  self-probe + Fable/mso-auditor Gate-2 APPROVE, add_section 12/12).
%     * Validate : validation\mat2doc\validate_P5-3a_sections_api.md (Gate-3 PASS --
%                  probe_diff s0045 MATCH exit 0 over the full Section/Sections
%                  surface; add_section s0046/s0047/s0048 byte packages 17/17;
%                  property-write s0049/s0050 17/17; M1 17/17; multi-section
%                  round-trip s0051; ZERO new D-numbers).
%     * Scenarios: validation\mat2doc\scenarios\s0045_p5_3a_sections_probe.{py,m}
%                  (the full-surface behavioral probe, its .m body replayed VERBATIM
%                  by runProbes() below); s0046..s0050 the byte-package twins
%                  (their build bodies replayed VERBATIM by the emit helpers below).
%     * Frozen refs (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE at Gate-3):
%         references\s0045\probe.json -- copied verbatim (self-contained) into
%           tests\section\data\s0045_probe_oracle.json (value JSON; jsondecode is
%           line-ending agnostic -> no `* binary` pin, s0030/s0039 precedent).
%         references\s0046..s0050\parts\word\document.xml -- NOT copied; pinned here
%           by SHA-256 (+ size) of what the MATLAB Document().save() ITSELF emits
%           (the same pattern as Test_p5_2a's M1/geometry/landscape byte pins).
%
%   Coverage taxonomy
%   -----------------
%   * Nominal      -- all 12 Section geometry/type accessors get/set; the Sections
%                     Sequence (len_/to_array/getitem_ int); add_section over each
%                     start_type; Document.sections / block_width_.
%   * Edge         -- H3 None ([]) round-trips on every nullable accessor; the
%                     page_width/height NO-orientation-swap invariant; NEGATIVE
%                     getitem_ (sections[-1] -> last); the empty / clamped / reversed
%                     SLICE battery; non-ASCII paragraph text through
%                     iter_inner_content; the IndexError / ValueError error paths
%                     (verbatim identifier + message); the w:tbl -> notYetPorted
%                     guard; the 6 P5-3b stubs.
%   * Equivalence  -- test_equivalence_full_battery_vs_frozen_oracle replays the
%                     ENTIRE s0045 battery (runProbes, the .m twin's body verbatim,
%                     incl. section_props / sequence(12-slice) / iter_inner_content /
%                     block_width) and flatten-compares every leaf to the frozen
%                     python-docx 1.2.0 oracle (Gate-3 found ZERO divergences).
%   * Regression   -- hard-coded expected property values (EMU) + the add_section /
%                     property-write / M1 whole-part SHA-256 + size byte pins +
%                     block_width_ 5486400/8229600.
%   * Upstream     -- the CPython slice.indices semantics, the negative-index wrap,
%                     the "list index out of range" / "slice step cannot be zero"
%                     messages, and the add_section_break sentinel-clone ordering
%                     ARE the python-docx section.py / oxml/document.py contract;
%                     the frozen oracle IS lxml's expected output for this sequence.
%
%   Byte-level (L1) note: every serialized-bytes assertion is a SHA-256 (+ size) pin
%   of the raw UTF-8 shipping bytes of word/document.xml extracted from a real
%   Document().save() .docx. SHA-256 equality == byte identity (L1). NO D-number
%   granted any L2 relaxation in this WP (Gate-3: zero new), so every byte pin is
%   L1. The equivalence leaf-key-count guard is the only looser-than-byte check and
%   is commented at its site.
%
%   Determinism: no network, no absolute paths. The co-located oracle resolves
%   relative to this file via fileparts(mfilename('fullpath')); saves go to
%   tempname .docx / tempname dirs deleted via onCleanup; every file read is binary
%   ('r','n'). The +mat2doc package resolves via the MANDATORY
%   PathFixture(worktree-root) in TestClassSetup (WP9-F4 lesson).

    properties (Constant)
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        SECTION  = 'mat2doc.section.Section'
        SECTIONS = 'mat2doc.section.Sections'

        % --- frozen s0001 M1 word/document.xml (un-stub neutrality guard) ---
        DOC_SIZE_M1 = 1548
        DOC_SHA_M1  = "0e4dd503bc095219cd24e8312b372c588ca56bb834aacac53eafb7bf47836327"

        % --- frozen s0046 add_section(NEW_PAGE) document.xml ---
        DOC_SIZE_NEWPAGE = 1792
        DOC_SHA_NEWPAGE  = "4d49edc14a0c282fd9eaa91dc8271df74f7b4c43cfdf80cfc88b58df189d2bda"

        % --- frozen s0048 add_section(ODD_PAGE)+add_section(EVEN_PAGE) 3-section chain ---
        DOC_SIZE_CHAIN = 2087
        DOC_SHA_CHAIN  = "f6cc256d22eb0a73aeb87d059c4c32d069c67af84897be3900b5ac6b3875f6eb"

        % --- frozen s0049 sections[0] property-write document.xml ---
        DOC_SIZE_PROPWRITE0 = 1569
        DOC_SHA_PROPWRITE0  = "5f3a0d5818019120bbe0291f3425f3dcf00398b1365c8e614e3427185a94dc86"

        % --- frozen s0050 full-surface sections[0]+new-section write document.xml ---
        DOC_SIZE_PROPWRITEFULL = 1893
        DOC_SHA_PROPWRITEFULL  = "b6c6af5e56b6ac6382b386b179825d0bd0633d77ae90271b0466be51236ab715"

        % --- frozen s0047 per-start-type document.xml SHA-256 (+ size) battery.
        %     NEW_PAGE == s0046 (removes w:type, identity default); the other four
        %     write <w:type w:val="..."/>. ---
        START_TYPE_SHAS = { ...
            "CONTINUOUS", 1820, "9d8f6d74b79d13f5d7be0a5fe3beb7c734c79407625f57ac0dafaf2e1efbafe5"; ...
            "NEW_COLUMN", 1820, "1fec6403fbe57d53cc5402fa7cee7c91494d4a0da3bf90670a95b8de2b5e50b1"; ...
            "NEW_PAGE",   1792, "4d49edc14a0c282fd9eaa91dc8271df74f7b4c43cfdf80cfc88b58df189d2bda"; ...
            "EVEN_PAGE",  1818, "8cc56f421eb1ee03a3b816cb809e619f0024e68a2f0ab84c766d2b871f553a9a"; ...
            "ODD_PAGE",   1817, "531c2a637ab98082d1213ed9866a7e6fb1e33be091d1f71b962f861f936f0c76" }

        % --- block_width_ (document.py _block_width) frozen probe values (EMU) ---
        BLOCK_WIDTH_DEFAULT  = 5486400
        BLOCK_WIDTH_MODIFIED = 8229600
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied from tests\oxml\Test_p5_2a_sectpr.m. here is
            % tests\section; the worktree root is two levels up.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\section
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. Section geometry / type -- all 12 accessors (H3/H6/H10)       %
        % =============================================================== %

        function test_section_geometry_type_all_props(testCase)
            % Nominal + Edge + Regression (section.py 35-253): all 12 Section
            % accessors delegate to the same-named CT_SectPr member. Read the frozen
            % default (M1 body sectPr geometry, EMU); write each; None-clear the
            % nullable ones (H3). page_width/height do NOT swap on orientation
            % change (H10 no-swap). part() == doc.part() (H5 StoryPart contract).
            IN  = @(v) mat2doc.shared.Inches(v);
            TW  = @(v) mat2doc.shared.Twips(v);
            PT  = @(v) mat2doc.shared.Pt(v);
            WDO = @(n) mat2doc.enum.section.WD_ORIENTATION.(n);
            WDS = @(n) mat2doc.enum.section.WD_SECTION_START.(n);

            d = mat2doc.Document();
            sec = d.sections.getitem_(0);              % Python: doc.sections[0]
            testCase.verifyClass(sec, testCase.SECTION, 'sections[0] -> a Section');

            % --- frozen M1 default geometry (probe s0045 section_props.default) ---
            testCase.verifyEqual(double(sec.bottom_margin), 914400,  'default bottom_margin EMU');
            testCase.verifyEqual(double(sec.top_margin),    914400,  'default top_margin EMU');
            testCase.verifyEqual(double(sec.left_margin),   1143000, 'default left_margin EMU');
            testCase.verifyEqual(double(sec.right_margin),  1143000, 'default right_margin EMU');
            testCase.verifyEqual(double(sec.gutter),        0,       'default gutter EMU');
            testCase.verifyEqual(double(sec.header_distance), 457200, 'default header_distance EMU');
            testCase.verifyEqual(double(sec.footer_distance), 457200, 'default footer_distance EMU');
            testCase.verifyEqual(double(sec.page_width),    7772400,  'default page_width EMU');
            testCase.verifyEqual(double(sec.page_height),   10058400, 'default page_height EMU');
            testCase.verifyEqual(sec.orientation, WDO("PORTRAIT"),   'default orientation PORTRAIT');
            testCase.verifyEqual(sec.start_type,  WDS("NEW_PAGE"),   'default start_type NEW_PAGE');
            testCase.verifyFalse(sec.different_first_page_header_footer, ...
                'default different_first_page_header_footer False');

            % H5: part() is the DocumentPart, identical handle to doc.part()
            testCase.verifyTrue(sec.part() == d.part(), ...
                'H5: sec.part() == doc.part() (same DocumentPart handle, StoryPart contract)');

            % --- write every accessor ---
            sec.top_margin    = IN(1.25);
            sec.bottom_margin = IN(0.5);
            sec.left_margin   = TW(1000);
            sec.gutter        = PT(12);
            sec.header_distance = IN(0.4);
            sec.page_width    = IN(11);
            sec.page_height   = IN(8.5);
            sec.orientation   = WDO("LANDSCAPE");
            sec.start_type    = WDS("EVEN_PAGE");
            sec.different_first_page_header_footer = true;
            testCase.verifyEqual(double(sec.top_margin),    1143000, 'set top_margin round-trip');
            testCase.verifyEqual(double(sec.bottom_margin), 457200,  'set bottom_margin round-trip');
            testCase.verifyEqual(double(sec.left_margin),   635000,  'set left_margin (Twips) round-trip');
            testCase.verifyEqual(double(sec.gutter),        152400,  'set gutter (Pt) round-trip');
            testCase.verifyEqual(double(sec.header_distance), 365760, 'set header_distance round-trip');
            testCase.verifyEqual(sec.orientation, WDO("LANDSCAPE"), 'set orientation round-trip');
            testCase.verifyEqual(sec.start_type,  WDS("EVEN_PAGE"), 'set start_type round-trip');
            testCase.verifyTrue(sec.different_first_page_header_footer, ...
                'set different_first_page_header_footer round-trip');

            % *** NO-SWAP GUARD: page_width/height are NOT swapped by orientation ***
            testCase.verifyEqual(double(sec.page_width),  10058400, ...
                'page_width stays 11in after LANDSCAPE (NO orientation swap)');
            testCase.verifyEqual(double(sec.page_height), 7772400, ...
                'page_height stays 8.5in after LANDSCAPE (NO orientation swap)');

            % --- H3 None ([]) clears the nullable accessors ---
            sec.right_margin = [];
            testCase.verifyTrue(isequal(sec.right_margin, []), 'right_margin = [] -> [] (None)');
            sec.footer_distance = [];
            testCase.verifyTrue(isequal(sec.footer_distance, []), 'footer_distance = [] -> [] (None)');

            % orientation None ([]) -> PORTRAIT default; start_type None -> NEW_PAGE
            sec.orientation = [];
            testCase.verifyEqual(sec.orientation, WDO("PORTRAIT"), ...
                'orientation = [] -> reads PORTRAIT (default)');
            sec.start_type = [];
            testCase.verifyEqual(sec.start_type, WDS("NEW_PAGE"), ...
                'start_type = [] -> reads NEW_PAGE (default)');
            sec.different_first_page_header_footer = false;
            testCase.verifyFalse(sec.different_first_page_header_footer, ...
                'different_first_page_header_footer = false -> False');
        end

        % =============================================================== %
        % 2. Sections Sequence -- 0-based getitem / slice / iter / errors  %
        % =============================================================== %

        function test_sections_sequence(testCase)
            % Nominal + Edge + Error path (section.py 273-286): a 3-section document
            % (add_section NEW_PAGE + ODD_PAGE). len_ == 3; to_array in document
            % order; getitem_ 0-based int + NEGATIVE wrap (sections[-1] -> last);
            % a representative slice battery; the verbatim mat2doc:IndexError and
            % mat2doc:ValueError.
            WDSEC = @(n) mat2doc.enum.section.WD_SECTION.(n);
            WDS   = @(n) mat2doc.enum.section.WD_SECTION_START.(n);

            d = mat2doc.Document();
            d.add_section(WDSEC("NEW_PAGE"));   % -> section 2 (sentinel)
            d.add_section(WDSEC("ODD_PAGE"));   % -> section 3 (sentinel)
            secs = d.sections;
            testCase.verifyClass(secs, testCase.SECTIONS, 'd.sections -> a Sections');

            % len_
            testCase.verifyEqual(secs.len_(), 3, 'len_ == 3 after two add_section');

            % to_array order
            arr = secs.to_array();
            testCase.verifyEqual(numel(arr), 3, 'to_array length 3');
            testCase.verifyEqual([arr.start_type], ...
                [WDS("NEW_PAGE"), WDS("NEW_PAGE"), WDS("ODD_PAGE")], ...
                'to_array start_types in document order');

            % *** 0-based int indexing (the TabStops convention) ***
            testCase.verifyEqual(secs.getitem_(0).start_type, WDS("NEW_PAGE"), ...
                'getitem_(0) is the FIRST section (0-based)');
            testCase.verifyEqual(secs.getitem_(1).start_type, WDS("NEW_PAGE"), 'getitem_(1)');
            testCase.verifyEqual(secs.getitem_(2).start_type, WDS("ODD_PAGE"), ...
                'getitem_(2) is the THIRD section');

            % *** NEGATIVE wrap: sections[-1] -> last ***
            testCase.verifyEqual(secs.getitem_(-1).start_type, WDS("ODD_PAGE"), ...
                'getitem_(-1) is the LAST section (negative wrap)');
            testCase.verifyEqual(secs.getitem_(-3).start_type, WDS("NEW_PAGE"), ...
                'getitem_(-3) is the FIRST section');
            testCase.verifyTrue(secs.getitem_(-1).start_type == secs.getitem_(2).start_type, ...
                'sections[-1] == sections[n-1]');

            % representative slice battery (full battery in the equivalence leg)
            testCase.verifyEqual(sliceTypes(secs, 0, 2, []), ["NEW_PAGE" "NEW_PAGE"], ...
                'slice [0:2]');
            testCase.verifyEqual(sliceTypes(secs, [], [], -1), ["ODD_PAGE" "NEW_PAGE" "NEW_PAGE"], ...
                'slice [::-1] reversed');
            testCase.verifyEqual(sliceTypes(secs, 2, 1, []), strings(1, 0), ...
                'slice [2:1] empty');
            testCase.verifyEqual(sliceTypes(secs, 5, [], []), strings(1, 0), ...
                'slice [5:] clamped empty');

            % error paths -- identifier + verbatim message
            ME = captureError(@() secs.getitem_(3));
            testCase.verifyEqual(string(ME.identifier), "mat2doc:IndexError", ...
                'getitem_(3) out-of-range -> mat2doc:IndexError');
            testCase.verifyEqual(string(ME.message), "list index out of range", ...
                'IndexError verbatim message');
            MEn = captureError(@() secs.getitem_(-4));
            testCase.verifyEqual(string(MEn.identifier), "mat2doc:IndexError", ...
                'getitem_(-4) out-of-range -> mat2doc:IndexError');
            MEs = captureError(@() secs.getitem_(mkslice(0, 2, 0)));
            testCase.verifyEqual(string(MEs.identifier), "mat2doc:ValueError", ...
                'zero slice step -> mat2doc:ValueError');
            testCase.verifyEqual(string(MEs.message), "slice step cannot be zero", ...
                'ValueError verbatim message');
        end

        % =============================================================== %
        % 3. iter_inner_content -- Paragraph value + w:tbl -> P6-4a raise   %
        % =============================================================== %

        function test_iter_inner_content(testCase)
            % Nominal + Edge (section.py 157-163): a <w:p> block resolves to a
            % Paragraph (value: non-ASCII text survives); a <w:tbl> block RAISES
            % mat2doc:notYetPorted (owner P6-4a) and is NEVER silently dropped.
            % --- CT_P -> Paragraph (value) ---
            d = mat2doc.Document();
            d.add_paragraph("IIC-caf" + string(char(233)));   % "IIC-café" (U+00E9)
            sec = d.sections.getitem_(0);
            items = sec.iter_inner_content();
            testCase.verifyClass(items, 'cell', 'iter_inner_content -> a cell array');
            testCase.verifyEqual(numel(items), 1, 'one inner block');
            testCase.verifyClass(items{1}, 'mat2doc.text.Paragraph', 'w:p -> Paragraph');
            testCase.verifyEqual(string(items{1}.text), "IIC-caf" + string(char(233)), ...
                'Paragraph text survives non-ASCII (é)');

            % --- w:tbl -> mat2doc:notYetPorted (P6-4a); NEVER silently dropped ---
            % A body carrying a <w:tbl> before the sentinel <w:sectPr>; the tbl is a
            % preceding-sibling of the body sectPr so the iterator INCLUDES it (as a
            % generic XmlElement), and Section wraps it via the Table branch -> raise.
            xml = "<w:document xmlns:w=""" + testCase.W + """><w:body>" + ...
                "<w:tbl/><w:sectPr/></w:body></w:document>";
            root = mat2doc.oxml.parse_xml(uint8(unicode2native(char(xml), "UTF-8")));
            sectPr = root.xpath("./w:body/w:sectPr");
            secT = mat2doc.section.Section(sectPr(1), []);   % document_part unused on the tbl path
            ME = captureError(@() secT.iter_inner_content());
            testCase.verifyEqual(string(ME.identifier), "mat2doc:notYetPorted", ...
                'w:tbl in iter_inner_content -> mat2doc:notYetPorted (P6-4a), NOT a silent drop');
            testCase.verifyTrue(contains(string(ME.message), "Table"), ...
                'notYetPorted message names the Table dependency');
        end

        % =============================================================== %
        % 4. add_section whole-part byte pins (s0046 / s0047 / s0048)       %
        % =============================================================== %

        function test_add_section_byte_pins(testCase)
            % Regression (L1): Document(); add_section(NEW_PAGE); save() ->
            % word/document.xml == EXACTLY 1792 B, SHA-256 4d49edc1... (frozen s0046).
            % Each of the five WD_SECTION start_types pinned vs the frozen s0047
            % member SHAs; the ODD_PAGE+EVEN_PAGE three-section chain -> 2087 B /
            % f6cc256d... (frozen s0048, the multi-section nesting guard).

            % s0046 -- add_section(NEW_PAGE)
            bytes = emitAddSection("NEW_PAGE");
            testCase.verifyEqual(numel(bytes), testCase.DOC_SIZE_NEWPAGE, ...
                sprintf('add_section(NEW_PAGE) document.xml must be exactly %d B', testCase.DOC_SIZE_NEWPAGE));
            testCase.verifyEqual(sha256hex(bytes), testCase.DOC_SHA_NEWPAGE, ...
                'add_section(NEW_PAGE) SHA-256 == frozen s0046 oracle (byte-identical L1)');

            % s0047 -- all five start_types
            tbl = testCase.START_TYPE_SHAS;
            for i = 1:size(tbl, 1)
                name = tbl{i, 1}; sz = tbl{i, 2}; sha = string(tbl{i, 3});
                b = emitAddSection(name);
                testCase.verifyEqual(numel(b), sz, ...
                    sprintf('add_section(%s) document.xml must be exactly %d B', name, sz));
                testCase.verifyEqual(sha256hex(b), sha, ...
                    sprintf('add_section(%s) SHA-256 == frozen s0047 member (L1)', name));
            end

            % s0048 -- ODD_PAGE then EVEN_PAGE (THREE sections)
            d = mat2doc.Document();
            d.add_section(mat2doc.enum.section.WD_SECTION.ODD_PAGE);
            d.add_section(mat2doc.enum.section.WD_SECTION.EVEN_PAGE);
            chain = saveAndExtract(d, 'document.xml');
            testCase.verifyEqual(numel(chain), testCase.DOC_SIZE_CHAIN, ...
                sprintf('3-section chain document.xml must be exactly %d B', testCase.DOC_SIZE_CHAIN));
            testCase.verifyEqual(sha256hex(chain), testCase.DOC_SHA_CHAIN, ...
                '3-section chain SHA-256 == frozen s0048 oracle (multi-section guard, L1)');
        end

        % =============================================================== %
        % 5. Section-property-write whole-part byte pins (s0049 / s0050)    %
        % =============================================================== %

        function test_section_property_write_byte_pins(testCase)
            % Regression (L1): the section-authoring write paths.
            %   * s0049: sections[0] top_margin/orientation/page_width -> 1569 B /
            %            5f3a0d58...
            %   * s0050: a FULL geometry/type/titlePg write on BOTH sections[0] and
            %            the newly-added sections[1] -> 1893 B / b6c6af5e...
            IN  = @(v) mat2doc.shared.Inches(v);
            WDO = @(n) mat2doc.enum.section.WD_ORIENTATION.(n);
            WDS = @(n) mat2doc.enum.section.WD_SECTION_START.(n);

            % ---- s0049 (sections[0] partial write) ----
            d = mat2doc.Document();
            sec = d.sections.getitem_(0);
            sec.top_margin  = IN(1.5);
            sec.orientation = WDO("LANDSCAPE");
            sec.page_width  = IN(11);
            b49 = saveAndExtract(d, 'document.xml');
            testCase.verifyEqual(numel(b49), testCase.DOC_SIZE_PROPWRITE0, ...
                sprintf('s0049 document.xml must be exactly %d B', testCase.DOC_SIZE_PROPWRITE0));
            testCase.verifyEqual(sha256hex(b49), testCase.DOC_SHA_PROPWRITE0, ...
                's0049 SHA-256 == frozen oracle (section property-write, L1)');

            % ---- s0050 (full-surface write on sections[0] + the new section) ----
            d = mat2doc.Document();
            d.add_section(mat2doc.enum.section.WD_SECTION.ODD_PAGE);
            secs = d.sections;
            s0 = secs.getitem_(0);
            s1 = secs.getitem_(1);
            s0.top_margin = IN(1);    s0.bottom_margin = IN(1);
            s0.left_margin = IN(1.25); s0.right_margin = IN(1.25);
            s0.gutter = IN(0.5);      s0.page_width = IN(8.5);
            s0.page_height = IN(11);  s0.orientation = WDO("PORTRAIT");
            s0.start_type = WDS("CONTINUOUS");
            s0.different_first_page_header_footer = true;
            s1.top_margin = IN(2);    s1.bottom_margin = IN(2);
            s1.left_margin = IN(1.5); s1.right_margin = IN(1.5);
            s1.page_width = IN(11);   s1.page_height = IN(8.5);
            s1.orientation = WDO("LANDSCAPE");
            s1.start_type = WDS("EVEN_PAGE");
            s1.different_first_page_header_footer = true;
            b50 = saveAndExtract(d, 'document.xml');
            testCase.verifyEqual(numel(b50), testCase.DOC_SIZE_PROPWRITEFULL, ...
                sprintf('s0050 document.xml must be exactly %d B', testCase.DOC_SIZE_PROPWRITEFULL));
            testCase.verifyEqual(sha256hex(b50), testCase.DOC_SHA_PROPWRITEFULL, ...
                's0050 SHA-256 == frozen oracle (full-surface property-write, L1)');
        end

        % =============================================================== %
        % 6. M1 document.xml byte-pin (un-stub neutrality guard)           %
        % =============================================================== %

        function test_m1_document_xml_byte_pin(testCase)
            % Regression (byte-neutrality, L1): un-stubbing sections / add_section /
            % add_section_break touches NO default-save byte -- a bare
            % Document().save() STILL emits word/document.xml at EXACTLY 1548 B with
            % the frozen s0001 SHA-256.
            d = mat2doc.Document();
            bytes = saveAndExtract(d, 'document.xml');
            testCase.verifyEqual(numel(bytes), testCase.DOC_SIZE_M1, ...
                sprintf('bare Document().save() document.xml must be exactly %d B', testCase.DOC_SIZE_M1));
            testCase.verifyEqual(sha256hex(bytes), testCase.DOC_SHA_M1, ...
                'M1 document.xml SHA-256 unchanged (un-stub neutrality, byte-identical L1)');
        end

        % =============================================================== %
        % 7. block_width_ (Document._block_width) frozen values             %
        % =============================================================== %

        function test_block_width(testCase)
            % Regression (document.py 232-239): block_width_ over the LIVE last
            % section. default -> 5486400 EMU (8.5in page - 1in - 1in margins on the
            % M1 default sectPr); after setting page_width=11in and 1in margins on
            % the last section -> 8229600 EMU (frozen probe s0045 block_width).
            IN = @(v) mat2doc.shared.Inches(v);

            d = mat2doc.Document();
            testCase.verifyEqual(double(d.block_width_()), testCase.BLOCK_WIDTH_DEFAULT, ...
                'default block_width_ == 5486400 EMU');

            d2 = mat2doc.Document();
            last = d2.sections.getitem_(-1);   % Python: doc.sections[-1]
            last.page_width  = IN(11);
            last.left_margin = IN(1);
            last.right_margin = IN(1);
            testCase.verifyEqual(double(d2.block_width_()), testCase.BLOCK_WIDTH_MODIFIED, ...
                'modified block_width_ == 8229600 EMU (11in - 1in - 1in)');
        end

        % =============================================================== %
        % 8. P5-3b hdr/ftr stubs (all six stay notYetPorted)               %
        % =============================================================== %

        function test_p5_3b_stubs(testCase)
            % Regression (error path, section.py 61-142): the SIX header/footer
            % members return _Header/_Footer objects (P5-3b) and stay STUBBED as
            % mat2doc:notYetPorted. A regression that un-stubs early (or silently
            % no-ops) goes RED.
            d = mat2doc.Document();
            sec = d.sections.getitem_(0);
            stubs = { ...
                'even_page_footer',  @() sec.even_page_footer(); ...
                'even_page_header',  @() sec.even_page_header(); ...
                'first_page_footer', @() sec.first_page_footer(); ...
                'first_page_header', @() sec.first_page_header(); ...
                'footer',            @() sec.footer(); ...
                'header',            @() sec.header() };
            for i = 1:size(stubs, 1)
                ME = captureError(stubs{i, 2});
                testCase.verifyEqual(string(ME.identifier), "mat2doc:notYetPorted", ...
                    sprintf('Section.%s must raise mat2doc:notYetPorted (P5-3b stub)', stubs{i, 1}));
            end
        end

        % =============================================================== %
        % 9. EQUIVALENCE -- full s0045 battery vs the frozen oracle         %
        % =============================================================== %

        function test_equivalence_full_battery_vs_frozen_oracle(testCase)
            % Equivalence: replay the ENTIRE s0045 battery (runProbes -- the .m twin's
            % body VERBATIM: section_props / sequence (12-slice + int + neg + errors)
            % / iter_inner_content / block_width) and flatten-compare EVERY leaf to
            % the frozen python-docx 1.2.0 oracle copied into
            % data\s0045_probe_oracle.json. Gate-3 found ZERO divergences (probe_diff
            % MATCH exit 0), so every leaf must be byte/value-identical.
            port   = runProbes();
            oracle = loadOracle();
            pMap = containers.Map('KeyType', 'char', 'ValueType', 'char');
            oMap = containers.Map('KeyType', 'char', 'ValueType', 'char');
            flattenLeaves(port,   '', pMap);
            flattenLeaves(oracle, '', oMap);

            pKeys = sort(pMap.keys());
            oKeys = sort(oMap.keys());
            testCase.verifyEqual(pKeys, oKeys, ...
                'the replayed battery and the frozen oracle must have identical leaf keys');
            % Non-trivial size guard (guards a silent-empty replay). The only looser-
            % than-byte assertion in this class; justified as a floor on leaf count.
            testCase.verifyGreaterThan(numel(oKeys), 40, ...
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

function types = sliceTypes(secs, a, b, c)
    % start_type NAMES of the Section array selected by the Python slice a:b:c.
    arr = secs.getitem_(mkslice(a, b, c));
    types = strings(1, numel(arr));
    for j = 1:numel(arr)
        types(j) = string(arr(j).start_type);
    end
end

function s = mkslice(a, b, c)
    % Scalar struct with start/stop/step ([]=None); mirrors Python slice(a,b,c).
    s = struct('start', a, 'stop', b, 'step', c);
end

function bytes = emitAddSection(memberName)
    % Document(); add_section(WD_SECTION.<memberName>); save() -> word/document.xml
    % raw bytes (the s0046/s0047 build body, VERBATIM).
    d = mat2doc.Document();
    d.add_section(mat2doc.enum.section.WD_SECTION.(memberName));
    bytes = saveAndExtract(d, 'document.xml');
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
    if isfile(p), delete(p); end
end

function rmdirIfExists(p)
    if isfolder(p), rmdir(p, 's'); end
end

function h = sha256hex(bytes)
    % Lowercase hex SHA-256 of a uint8 vector (base MATLAB / Java MessageDigest).
    md  = java.security.MessageDigest.getInstance('SHA-256');
    dig = md.digest(typecast(uint8(bytes(:)'), 'int8'));
    b   = typecast(dig, 'uint8');
    h   = lower(string(sprintf('%02x', b)));
end

function ME = captureError(fn)
    % Run fn and RETURN the caught MException (asserts it actually raised).
    raised = true;
    try
        fn();
        raised = false;
    catch ME
        return
    end
    if ~raised
        error('mat2doc:test:noRaise', 'expected an error but none was raised');
    end
end

function o = loadOracle()
    % Read the co-located frozen s0045 oracle in BINARY mode (no CRLF translation)
    % and decode UTF-8 -> struct. jsondecode is whitespace/line-ending agnostic, so
    % no `* binary` pin is needed for this value fixture (s0030/s0039 precedent).
    here = fileparts(mfilename('fullpath'));
    p = fullfile(here, 'data', 's0045_probe_oracle.json');
    fid = fopen(p, 'r', 'n');
    if fid < 0
        error('mat2doc:test:openFail', 'cannot open frozen oracle %s', p);
    end
    raw = fread(fid, Inf, '*uint8')';
    fclose(fid);
    o = jsondecode(native2unicode(raw, 'UTF-8'));
end

% ---- s0045 probe replay (for the Equivalence leg) --------------------- %

function P = runProbes()
    % Replay the s0045 probe sequence (the .m twin's P-struct body, VERBATIM
    % tags/inputs/order) and return the nested struct of tagged canonical values.
    % Embedded so the Equivalence leg is self-contained (the validation-folder
    % scenario is NOT on the toolbox path and must not be a dependency). Mirrors
    % validation\mat2doc\scenarios\s0045_p5_3a_sections_probe.m lines 27-128.
    WDO = @(n) mat2doc.enum.section.WD_ORIENTATION.(n);
    WDS = @(n) mat2doc.enum.section.WD_SECTION_START.(n);
    WDSEC = @(n) mat2doc.enum.section.WD_SECTION.(n);
    IN  = @(v) mat2doc.shared.Inches(v);
    TW  = @(v) mat2doc.shared.Twips(v);
    PT  = @(v) mat2doc.shared.Pt(v);

    P = struct();

    % ======================= section_props ==============================
    d = mat2doc.Document();
    sec = d.sections.getitem_(0);                    % Python: doc.sections[0]
    sp = struct();
    sp.default = read_all(sec);
    sp.part_is_doc_part = rv(sec.part() == d.part()); % Python: sec.part is doc.part

    sec.top_margin = IN(1.25);
    sec.bottom_margin = IN(0.5);
    sec.left_margin = TW(1000);
    sec.right_margin = [];                            % Python: None (removes)
    sec.gutter = PT(12);
    sec.header_distance = IN(0.4);
    sec.footer_distance = [];                         % Python: None (removes)
    sec.page_width = IN(11);
    sec.page_height = IN(8.5);
    sec.orientation = WDO("LANDSCAPE");
    sec.start_type = WDS("EVEN_PAGE");
    sec.different_first_page_header_footer = true;
    sp.after_set = read_all(sec);

    sec.orientation = [];                             % falsy -> PORTRAIT default
    sec.start_type = [];                              % removes -> NEW_PAGE default
    sec.different_first_page_header_footer = false;
    an = struct();
    an.orientation = rv(sec.orientation);
    an.start_type = rv(sec.start_type);
    an.different_first_page_header_footer = rv(sec.different_first_page_header_footer);
    sp.after_none = an;
    P.section_props = sp;

    % ============================ sequence ==============================
    d2 = mat2doc.Document();
    d2.add_section(WDSEC("NEW_PAGE"));                % -> section 2 (sentinel)
    d2.add_section(WDSEC("ODD_PAGE"));               % -> section 3 (sentinel)
    secs = d2.sections;
    n = secs.len_();
    seq = struct();
    seq.len = rv(n);
    arr = secs.to_array();
    seq.iter_start_types = st_names(arr);
    seq.getitem_int = {rv(secs.getitem_(0).start_type), ...
                       rv(secs.getitem_(1).start_type), ...
                       rv(secs.getitem_(2).start_type)};
    seq.getitem_neg = {rv(secs.getitem_(-1).start_type), ...
                       rv(secs.getitem_(-2).start_type), ...
                       rv(secs.getitem_(-3).start_type)};
    seq.neg1_is_last = rv(secs.getitem_(-1).start_type == secs.getitem_(n - 1).start_type);

    % 12-case slice battery (keys == python SLICES; fields []=None).
    sl = struct();
    sl.sl_0_2_None          = st_names(secs.getitem_(mkslice(0, 2, [])));
    sl.sl_None_None_None    = st_names(secs.getitem_(mkslice([], [], [])));
    sl.sl_None_None_2       = st_names(secs.getitem_(mkslice([], [], 2)));
    sl.sl_neg2_None_None    = st_names(secs.getitem_(mkslice(-2, [], [])));
    sl.sl_2_1_None_empty    = st_names(secs.getitem_(mkslice(2, 1, [])));
    sl.sl_None_None_neg1    = st_names(secs.getitem_(mkslice([], [], -1)));
    sl.sl_neg1_None_neg1    = st_names(secs.getitem_(mkslice(-1, [], -1)));
    sl.sl_1_None_neg2       = st_names(secs.getitem_(mkslice(1, [], -2)));
    sl.sl_0_3_2             = st_names(secs.getitem_(mkslice(0, 3, 2)));
    sl.sl_5_None_None_clamp = st_names(secs.getitem_(mkslice(5, [], [])));
    sl.sl_neg5_2_None_clamp = st_names(secs.getitem_(mkslice(-5, 2, [])));
    sl.sl_None_1_neg1       = st_names(secs.getitem_(mkslice([], 1, -1)));
    seq.slices = sl;

    seq.err_index_pos = errmsg(@() secs.getitem_(3));
    seq.err_index_neg = errmsg(@() secs.getitem_(-4));
    seq.err_slice_step0 = errmsg(@() secs.getitem_(mkslice(0, 2, 0)));
    P.sequence = seq;

    % ====================== iter_inner_content ==========================
    d3 = mat2doc.Document();
    d3.add_paragraph("IIC-caf" + string(char(233)));  % Python: "IIC-café" (U+00E9)
    sec3 = d3.sections.getitem_(0);
    items = sec3.iter_inner_content();                % cell of Paragraph|Table
    types = cell(1, numel(items));
    texts = {};
    for k = 1:numel(items)
        short = shortcls(items{k});
        types{k} = short;
        if short == "Paragraph"
            texts{end+1} = string(items{k}.text); %#ok<AGROW>
        end
    end
    iic = struct();
    iic.types = types;
    iic.para_texts = texts;
    P.iter_inner_content = iic;

    % ============================ block_width ===========================
    bw = struct();
    d4 = mat2doc.Document();
    bw.default = rv(d4.block_width_());               % Python: int(doc._block_width)
    d5 = mat2doc.Document();
    last = d5.sections.getitem_(-1);                  % Python: doc.sections[-1]
    last.page_width = IN(11);
    last.left_margin = IN(1);
    last.right_margin = IN(1);
    bw.modified = rv(d5.block_width_());
    P.block_width = bw;
end

function r = read_all(sec)
    r = struct();
    r.bottom_margin = rv(sec.bottom_margin);
    r.top_margin = rv(sec.top_margin);
    r.left_margin = rv(sec.left_margin);
    r.right_margin = rv(sec.right_margin);
    r.gutter = rv(sec.gutter);
    r.header_distance = rv(sec.header_distance);
    r.footer_distance = rv(sec.footer_distance);
    r.page_width = rv(sec.page_width);
    r.page_height = rv(sec.page_height);
    r.orientation = rv(sec.orientation);
    r.start_type = rv(sec.start_type);
    r.different_first_page_header_footer = rv(sec.different_first_page_header_footer);
end

function c = st_names(arr)
    % 1xN cell of rv(start_type) over a Section array (1x0 -> {} -> JSON []).
    c = cell(1, numel(arr));
    for j = 1:numel(arr)
        c{j} = rv(arr(j).start_type);
    end
end

function s = shortcls(x)
    parts = split(string(class(x)), ".");
    s = parts(end);
end

function s = rv(x)
    % Uniform accessor repr mirroring the s0045 rv(): None->"None",
    % enum->member NAME, bool->"True"/"False", Length/int->EMU decimal.
    if isequal(x, [])
        s = "None";
    elseif isenum(x)
        s = string(x);
    elseif islogical(x)
        if x, s = "True"; else, s = "False"; end
    elseif isa(x, 'mat2doc.shared.Length')
        s = string(sprintf('%.0f', double(x)));
    elseif isnumeric(x)
        s = string(sprintf('%.0f', double(x)));
    else
        s = string(x);
    end
end

function s = errmsg(fn)
    try
        fn();
        s = "NOERR";
    catch e
        s = string(e.message);
    end
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

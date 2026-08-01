classdef Test_p2_3_document_shell < matlab.unittest.TestCase
% TEST_P2_3_DOCUMENT_SHELL  Gate-4 permanent unit tests for Mat2Doc P2-3
%   (document shell + blkcntnr + CT_Document/CT_Body) -- the FINAL P2 WP that
%   completes the document object graph.
%
%   Surface under test (ported from python-docx v1.2.0):
%     * src/docx/oxml/document.py  -> +mat2doc\+oxml\+document\CT_Document.m,
%       CT_Body.m  (registered for w:document / w:body in +oxml\registry.m);
%     * src/docx/document.py       -> +mat2doc\+document\Document.m (body_ cache),
%       Body_.m;
%     * src/docx/blkcntnr.py       -> +mat2doc\BlockItemContainer.m.
%
%   THE P2-3 byte-risk is the ONE step that touches the document.xml parse path:
%   registering w:document->CT_Document and w:body->CT_Body means word/document.xml
%   now PARSES to CT_Document/CT_Body (BaseOxmlElement subclasses) and RESERIALIZES
%   through the BaseOxmlElement path. Both tiers exit the identical
%   +oxml\serialize_part_xml walk and neither class overrides any serialization
%   member, so the flip is BYTE-NEUTRAL -- pinned here specifically for
%   word/document.xml (Test_p1_8_skeleton_m1 owns the full 17/17 M1 sweep; this
%   class adds the P2-3-specific document.xml pin, since P2-3 is what could move it).
%
%   Provenance (Gate-1..3, all 2026-07-26):
%     * Audit    : validation\mat2doc\audit_P2-3_document_shell.md
%                  (Porter Gate-1 24/24 sanity + Opus Gate-2 [R] APPROVE, 58/58
%                  hazard probes, 2 inline doc fixes; document.xml byte-neutral
%                  independently re-confirmed A2/A3).
%     * Validate : validation\mat2doc\validate_P2-3_document_shell.md
%                  (Gate-3 PASS on equivalence -- all 6 bars MATCH, 17/17
%                  byte-neutrality, word/document.xml 1548 B byte-identical; 0 new
%                  D-numbers; the ONLY blocker was 4 stale exact-class pins, fixed
%                  in JOB A alongside this class).
%     * Scenario : validation\mat2doc\scenarios\s0015_p2_3_document_shell.{py,m}
%                  (the probe sequences mirrored below).
%     * Frozen refs (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE):
%         references\s0015\probe.json     -- CT behavior (bars 2-4);
%         references\s0015\matlab_only.json -- forward-dep + 21-stub battery (bars 5-6);
%         references\s0001\ (word/document.xml 1548 B, sha 0e4dd503...) -- the M1
%           byte reference. The frozen oracle VALUES are embedded below as Constants
%           (ORACLE_*) so this suite is self-contained.
%
%   Coverage taxonomy
%   -----------------
%   * Nominal / Regression (CT registration + dispatch) -- registry resolves
%     w:document->CT_Document and w:body->CT_Body; a REAL opened document's
%     element() isa CT_Document and .body isa CT_Body.
%   * Regression (document.xml byte-neutrality, THE P2-3 byte-risk) --
%     mat2doc.Document().save() emits word/document.xml at EXACTLY 1548 B with the
%     frozen s0001 SHA-256 (byte-identical L1) AFTER CT_Document/CT_Body
%     registration.
%   * Equivalence (frozen s0015 oracle) -- CT_Document.sectPr_lst union order,
%     CT_Body.inner_content_elements (shading excluded), clear_content (with/without
%     sectPr), the H11 successor child-order -- each value-compared to the frozen
%     python-docx probe.json.
%   * Regression (H5/H9 body_ cache) -- repeated Document.body_() returns the SAME
%     Body_ handle (manual __body cache, document.py 244-246).
%   * Edge / Regression (P2 EXIT stub-safety, error path) -- the 21 content-authoring
%     stubs (Document + Body_ + CT_Body.add_section_break) each raise the error
%     IDENTIFIER mat2doc:notYetPorted; a clean Document().save() fires ZERO stubs.
%
%   Deviations exercised (adopt-only, ZERO new -- Gate-3): D-001 (own OOXML parser,
%   via parse/serialize -- byte-proven), D-serializer-nsdecl (unreachable at P2-3:
%   the open->save path creates/mutates no elements). The forward-dep class
%   divergence (created w:p/w:tbl/w:sectPr resolve to generic XmlElement, since
%   CT_P/CT_Tbl/CT_SectPr land at P4/P6/P5) is byte-neutral and scope-driven -- NOT
%   a deviation -- and is pinned as an intentional-difference regression.
%
%   Determinism: no network, no absolute paths -- the worktree root and template
%   resolve relative to this file via fileparts(mfilename('fullpath')); every parsed
%   blob is an in-memory literal; every file read is binary ('r','n'); saves go to
%   tempname .docx deleted via onCleanup (no 'wt').

    properties (Constant)
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        % --- registered class names (P2-3) ---
        CT_DOCUMENT = 'mat2doc.oxml.document.CT_Document'
        CT_BODY     = 'mat2doc.oxml.document.CT_Body'

        % --- frozen s0015 oracle (references\s0015\probe.json) ---
        ORACLE_SECTPR_PARENTS   = ["pPr" "body"]          % sectPr_lst union order
        ORACLE_INNER_CONTENT    = ["p" "tbl"]             % w:ins/w:p shaded out
        ORACLE_CLEAR_WITH_SECT  = "sectPr"                % remaining after clear
        ORACLE_H11_BEFORE_SECT  = ["p" "tbl" "p" "sectPr"]% add_p/add_tbl before sectPr
        ORACLE_H11_SECT_LAST    = ["p" "tbl" "sectPr"]    % sectPr (succ=()) appends last

        % --- frozen s0001 word/document.xml byte reference (M1) ---
        DOCXML_SIZE = 1548
        DOCXML_SHA  = "0e4dd503bc095219cd24e8312b372c588ca56bb834aacac53eafb7bf47836327"

        % --- frozen matlab_only.json: the stub battery names (bar 6). Registry-flip
        %     stale-pins as the un-stubs land (each drops out of this notYetPorted
        %     battery and is re-pinned to its resolved behavior in
        %     test_content_stubs_raise_notYetPorted):
        %       P4-7a un-stubbed Document.styles                    (21 -> 20)
        %       P4-7b un-stubbed Document.add_paragraph/add_heading +
        %             Body_.add_paragraph/Body_.paragraphs          (20 -> 16)
        %       P5-1  un-stubbed Document.settings                  (16 -> 15)
        %       P5-3a un-stubbed Document.add_section / Document.sections /
        %             CT_Body.add_section_break                     (15 -> 12)
        %       P6-4a un-stubbed Document.add_table /
        %             Document.iter_inner_content / Body_.add_table /
        %             Body_.iter_inner_content                      (12 -> 8)
        %     Document.add_page_break / Document.add_picture / Document.add_comment /
        %     Document.paragraphs / Document.tables / Document.comments /
        %     Document.inline_shapes / Body_.tables and the P6/P7/P8 adders
        %     REMAIN genuinely stubbed. ---
        STUB_COUNT = 8
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied from tests\oxml\Test_p2_2_xpath_hoist.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\oxml
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. CT_Document / CT_Body registration + dispatch                %
        % =============================================================== %

        function test_registry_resolves_document_and_body(testCase)
            % Nominal: the registry maps w:document -> CT_Document and w:body ->
            % CT_Body (the two P2-3 rows). REGISTRATION is what flips the parse
            % class (and, incidentally, the 4 JOB-A pins).
            rd = mat2doc.oxml.registry(mat2doc.oxml.qn("w:document"));
            rb = mat2doc.oxml.registry(mat2doc.oxml.qn("w:body"));
            testCase.verifyEqual(char(rd), testCase.CT_DOCUMENT, ...
                'registry must resolve w:document -> CT_Document (P2-3 row)');
            testCase.verifyEqual(char(rb), testCase.CT_BODY, ...
                'registry must resolve w:body -> CT_Body (P2-3 row)');
        end

        function test_real_document_dispatches_ct_classes(testCase)
            % Nominal / Regression: a REAL opened document (default template) has
            % element() exact-class CT_Document and .body exact-class CT_Body --
            % registration is live end-to-end on a genuine document.xml.
            % (s0015 probe.json dispatch: root_class CT_Document, body_class CT_Body.)
            d = mat2doc.Document();
            root = d.element();
            testCase.verifyEqual(class(root), testCase.CT_DOCUMENT, ...
                'opened document root must be CT_Document');
            testCase.verifyEqual(class(root.body), testCase.CT_BODY, ...
                'opened document .body must be CT_Body');
        end

        % =============================================================== %
        % 2. document.xml byte-neutrality reinforcement (THE byte-risk)    %
        % =============================================================== %

        function test_document_xml_byte_identical_after_registration(testCase)
            % Regression (byte-neutrality, L1): mat2doc.Document().save() emits
            % word/document.xml at EXACTLY 1548 B with the frozen s0001 SHA-256 --
            % byte-identical DESPITE the CT_Document/CT_Body parse-class flip.
            % SHA-256 equality is a byte-level (L1) assertion. RED on any single-byte
            % drift in the reparsed-through-CT part. (validate_P2-3 Bar 1;
            % Test_p1_8_skeleton_m1 owns the full 17/17 sweep.)
            docxml = testCase.emitDocumentXml();
            testCase.verifyEqual(numel(docxml), testCase.DOCXML_SIZE, ...
                sprintf('word/document.xml must be exactly %d B after CT registration', ...
                    testCase.DOCXML_SIZE));
            testCase.verifyEqual(sha256hex(docxml), testCase.DOCXML_SHA, ...
                'word/document.xml SHA-256 must equal the frozen s0001 oracle (byte-identical L1)');
        end

        % =============================================================== %
        % 3. CT_Document / CT_Body xpath methods (frozen s0015 oracle)     %
        % =============================================================== %

        function test_sectPr_lst_union_document_order(testCase)
            % Equivalence (s0015): CT_Document.sectPr_lst is the union
            % ./w:body/w:p/w:pPr/w:sectPr | ./w:body/w:sectPr returned in DOCUMENT
            % order -- a paragraph-level pPr/sectPr THEN the body-level sectPr.
            % Pinned by the ordered parent localnames [pPr, body].
            doc = testCase.parseDoc("<w:p><w:pPr><w:sectPr/></w:pPr></w:p><w:tbl/><w:sectPr/>");
            sl = doc.sectPr_lst;
            testCase.verifyEqual(numel(sl), 2, 'sectPr_lst count');
            testCase.verifyEqual(testCase.parentLocalnames(sl), testCase.ORACLE_SECTPR_PARENTS, ...
                'sectPr_lst union must be in document order (pPr-level then body-level)');
        end

        function test_inner_content_elements_excludes_shading(testCase)
            % Equivalence (s0015): CT_Body.inner_content_elements = ./w:p | ./w:tbl
            % -- a two-branch CHILD-axis union, so a w:p SHADED inside w:ins is
            % EXCLUDED (not a descendant scan). Result [p, tbl].
            body = testCase.parseBodyOf("<w:p/><w:tbl/><w:ins><w:p/></w:ins><w:sectPr/>");
            ice = body.inner_content_elements;
            testCase.verifyEqual(numel(ice), 2, 'inner_content_elements count');
            testCase.verifyEqual(testCase.localnames(ice), testCase.ORACLE_INNER_CONTENT, ...
                'inner_content_elements must be [p, tbl] (w:ins/w:p shaded out)');
        end

        function test_clear_content_keeps_sectPr(testCase)
            % Equivalence (s0015): CT_Body.clear_content = remove
            % ./*[not(self::w:sectPr)] -- with a w:sectPr present only the sectPr
            % survives (materialized-then-remove, H9-safe). Remaining [sectPr].
            body = testCase.parseBodyOf("<w:p/><w:tbl/><w:ins><w:p/></w:ins><w:sectPr/>");
            body.clear_content();
            rem = body.getchildren();
            testCase.verifyEqual(numel(rem), 1, 'clear_content with sectPr: one child remains');
            testCase.verifyEqual(testCase.localnames(rem), testCase.ORACLE_CLEAR_WITH_SECT, ...
                'clear_content must preserve the w:sectPr');
        end

        function test_clear_content_empties_sectprless_body(testCase)
            % Edge / Equivalence (s0015): clear_content on a sectPr-less body empties
            % it (remaining_count 0) -- the not(self::w:sectPr) predicate matches
            % every child.
            body = testCase.parseBodyOf("<w:p/><w:tbl/>");
            body.clear_content();
            rem = body.getchildren();
            testCase.verifyEqual(numel(rem), 0, ...
                'clear_content on a sectPr-less body must remove every child');
        end

        % =============================================================== %
        % 4. H11 successors (THE WP hazard)                               %
        % =============================================================== %

        function test_h11_add_p_add_tbl_insert_before_sectPr(testCase)
            % Equivalence (s0015, H11): p/tbl carry successors=("w:sectPr",), so with
            % a sentinel w:sectPr present add_p/add_tbl INSERT BEFORE it. Sequence
            % get_or_add_sectPr -> add_p -> add_tbl -> add_p yields
            % [p, tbl, p, sectPr] (sectPr always last).
            body = testCase.parseBody();
            body.get_or_add_sectPr();   % [sectPr]
            body.add_p();               % [p, sectPr]
            body.add_tbl();             % [p, tbl, sectPr]
            body.add_p();               % [p, tbl, p, sectPr]
            testCase.verifyEqual(testCase.localnames(body.getchildren()), ...
                testCase.ORACLE_H11_BEFORE_SECT, ...
                'add_p/add_tbl must insert BEFORE the sentinel w:sectPr');
        end

        function test_h11_sectPr_appends_last(testCase)
            % Equivalence (s0015, H11): sectPr carries successors=() (NO_SUCCESSORS),
            % so on a fresh body it APPENDS LAST. Sequence add_p -> add_tbl ->
            % get_or_add_sectPr yields [p, tbl, sectPr].
            body = testCase.parseBody();
            body.add_p();               % [p]
            body.add_tbl();             % [p, tbl]
            body.get_or_add_sectPr();   % [p, tbl, sectPr]
            testCase.verifyEqual(testCase.localnames(body.getchildren()), ...
                testCase.ORACLE_H11_SECT_LAST, ...
                'get_or_add_sectPr (successors=()) must append the w:sectPr LAST');
        end

        % =============================================================== %
        % 5. body_ manual cache (H5 / H9)                                 %
        % =============================================================== %

        function test_body_cache_same_handle_on_repeat(testCase)
            % Regression (H5/H9, faithful): Document.body_() twice returns the SAME
            % Body_ handle -- python-docx v1.2.0 _body is a plain @property with a
            % manual self.__body is None cache (document.py 244-246); the port's
            % body__=[] + isequal(body__,[]) mirrors it. (s0015 body_cache: true.)
            d = mat2doc.Document();
            b1 = d.body_();
            b2 = d.body_();
            testCase.verifyClass(b1, 'mat2doc.document.Body_', ...
                'body_() must return a mat2doc.document.Body_ proxy');
            testCase.verifyTrue(b1 == b2, ...
                'repeat body_() must return the SAME cached Body_ handle (H5)');
        end

        % =============================================================== %
        % 6. Forward-dep scope + P2 EXIT stub-safety                      %
        % =============================================================== %

        function test_created_children_class_reflects_registration(testCase)
            % Regression (registration-driven parse class, s0015 matlab_only): the
            % w:p/w:tbl/w:sectPr created by add_p/add_tbl/get_or_add_sectPr resolve to
            % whichever class is registered at the time. As of P4-2, w:p->CT_P is now
            % registered, so add_p() returns a mat2doc.oxml.text.CT_P (this test's own
            % prior comment predicted the flip when "P4" landed -- P4-2 IS "P4"). w:tbl
            % FLIPPED to CT_Tbl at P6-3b (registry-flip lesson; the pin below was
            % re-pinned at Gate-4 from XmlElement -> CT_Tbl when w:tbl was registered);
            % w:sectPr FLIPPED to CT_SectPr at P5-2a (registry-flip lesson; the pin
            % below was re-pinned at Gate-4 from XmlElement -> CT_SectPr). The class
            % difference is byte-neutral (identical serialization + insertion order)
            % and scope-driven; the CT_P / CT_SectPr pins are now the exact-correct
            % classes (not a defect).
            body = testCase.parseBody();
            cSect = body.get_or_add_sectPr();
            cP    = body.add_p();
            cTbl  = body.add_tbl();
            testCase.verifyEqual(class(cP), 'mat2doc.oxml.text.CT_P', ...
                'created w:p now resolves to CT_P (registered in P4-2)');
            testCase.verifyEqual(class(cTbl), 'mat2doc.oxml.table.CT_Tbl', ...
                'created w:tbl now resolves to CT_Tbl (registered in P6-3b; byte-neutral class flip)');
            testCase.verifyEqual(class(cSect), 'mat2doc.oxml.section.CT_SectPr', ...
                'created w:sectPr now resolves to CT_SectPr (registered in P5-2a; byte-neutral class flip)');
        end

        function test_content_stubs_raise_notYetPorted(testCase)
            % Edge / Regression (error path, P2 EXIT): every content-authoring stub
            % raises the IDENTIFIER mat2doc:notYetPorted (not a silent no-op). Covers
            % the Document tier, the canonical BlockItemContainer/Body_ tier, and
            % CT_Body.add_section_break. REGISTRY-FLIP RE-PINS (the un-stub moves the
            % behavior, so the pin moves with it):
            %   * P4-7a: Document.styles un-stubbed -> REMOVED (asserted resolved below).
            %   * P4-7b (M2 milestone WP): Document.add_paragraph / Document.add_heading
            %     / Body_.add_paragraph / Body_.paragraphs un-stubbed -> REMOVED (all
            %     asserted resolved below). (20 -> 16.)
            %   * P5-1: Document.settings un-stubbed -> REMOVED (asserted resolved
            %     below). (16 -> 15.)
            %   * P5-3a: Document.add_section / Document.sections /
            %     CT_Body.add_section_break un-stubbed -> REMOVED (all asserted
            %     resolved below). (15 -> 12.)
            %   * P6-4a: Document.add_table / Document.iter_inner_content /
            %     Body_.add_table / Body_.iter_inner_content un-stubbed -> REMOVED
            %     (all asserted resolved below -- add_table returns a Table;
            %     iter_inner_content returns a cell). (12 -> 8.)
            % Document.add_page_break / Document.add_picture / Document.add_comment /
            % Document.paragraphs / Document.tables / Document.comments /
            % Document.inline_shapes / Body_.tables REMAIN genuinely pinned (deps not
            % all live and/or deliberately left stubbed for the next content WP; the
            % P6/P7/P8 adders land later).
            d  = mat2doc.Document();
            b  = d.body_();
            ct = testCase.parseBody();
            calls = { ...
                'Document.add_page_break',     @() d.add_page_break(); ...
                'Document.add_picture',        @() d.add_picture("x"); ...
                'Document.add_comment',        @() d.add_comment([], "t", "a", "i"); ...
                'Document.paragraphs',         @() d.paragraphs(); ...
                'Document.tables',             @() d.tables(); ...
                'Document.comments',           @() d.comments(); ...
                'Document.inline_shapes',      @() d.inline_shapes(); ...
                'Body_.tables',                @() b.tables()};
            testCase.verifyEqual(size(calls, 1), testCase.STUB_COUNT, ...
                'the stub battery must cover exactly 8 members (after Document.styles P4-7a + 4 adders P4-7b + Document.settings P5-1 + 3 section members P5-3a + 4 table members P6-4a un-stubbed)');
            for k = 1:size(calls, 1)
                caught = [];
                try
                    calls{k, 2}();
                catch ME
                    caught = ME;
                end
                testCase.assertNotEmpty(caught, ...
                    sprintf('stub %s must raise', calls{k, 1}));
                testCase.verifyEqual(caught.identifier, 'mat2doc:notYetPorted', ...
                    sprintf('stub %s must raise mat2doc:notYetPorted', calls{k, 1}));
            end
            % Document.styles now RESOLVES (un-stubbed at P4-7a).
            testCase.verifyClass(d.styles(), 'mat2doc.styles.Styles', ...
                'Document.styles RESOLVES to a Styles proxy (P4-7a un-stub)');
            % Document.settings now RESOLVES (un-stubbed at P5-1).
            testCase.verifyClass(d.settings(), 'mat2doc.settings.Settings', ...
                'Document.settings RESOLVES to a Settings proxy (P5-1 un-stub)');

            % The P4-7b (M2) adder un-stubs now RESOLVE end-to-end. Use fresh
            % handles so these positive checks do not interfere with the stub loop's
            % document (add_* mutates the body). add_paragraph/add_heading return a
            % Paragraph; Body_.paragraphs returns a (possibly empty) Paragraph array.
            dr = mat2doc.Document();
            br = dr.body_();
            testCase.verifyClass(dr.add_paragraph("x"), 'mat2doc.text.Paragraph', ...
                'Document.add_paragraph RESOLVES to a Paragraph (P4-7b un-stub, M2)');
            testCase.verifyClass(dr.add_heading("x", 1), 'mat2doc.text.Paragraph', ...
                'Document.add_heading RESOLVES to a Paragraph (P4-7b un-stub, M2)');
            testCase.verifyClass(br.add_paragraph("x"), 'mat2doc.text.Paragraph', ...
                'Body_.add_paragraph RESOLVES to a Paragraph (P4-7b un-stub)');
            testCase.verifyClass(br.paragraphs(), 'mat2doc.text.Paragraph', ...
                'Body_.paragraphs RESOLVES to a Paragraph array (P4-7b un-stub)');

            % The P5-3a section un-stubs now RESOLVE (registry-flip re-pins:
            % Document.sections / Document.add_section / CT_Body.add_section_break
            % moved out of the notYetPorted battery to their resolved behavior). Use
            % fresh handles so add_section's body mutation does not perturb others.
            testCase.verifyClass(dr.sections(), 'mat2doc.section.Sections', ...
                'Document.sections RESOLVES to a Sections proxy (P5-3a un-stub)');
            testCase.verifyClass(dr.add_section(), 'mat2doc.section.Section', ...
                'Document.add_section RESOLVES to a Section (P5-3a un-stub)');
            % CT_Body.add_section_break returns the (sentinel) CT_SectPr. `ct` is the
            % standalone CT_Body from parseBody() (untouched by the stub loop).
            testCase.verifyClass(ct.add_section_break(), 'mat2doc.oxml.section.CT_SectPr', ...
                'CT_Body.add_section_break RESOLVES to a CT_SectPr (P5-3a un-stub)');

            % The P6-4a table un-stubs now RESOLVE (registry-flip re-pins:
            % Document.add_table / Document.iter_inner_content / Body_.add_table /
            % Body_.iter_inner_content moved out of the notYetPorted battery to their
            % resolved behavior). Use FRESH handles so add_table's body mutation does
            % not perturb the others. Document.add_table(rows,cols,style=[]) and
            % BlockItemContainer.add_table(rows,cols,width) both return a Table;
            % iter_inner_content returns a (heterogeneous Paragraph|Table) cell.
            dt = mat2doc.Document();
            bt = dt.body_();
            testCase.verifyClass(dt.add_table(1, 1, []), 'mat2doc.table.Table', ...
                'Document.add_table RESOLVES to a Table (P6-4a un-stub)');
            testCase.verifyClass(dt.iter_inner_content(), 'cell', ...
                'Document.iter_inner_content RESOLVES to a cell array (P6-4a un-stub)');
            bt2 = mat2doc.Document().body_();
            testCase.verifyClass(bt2.add_table(1, 1, mat2doc.shared.Inches(1)), 'mat2doc.table.Table', ...
                'Body_.add_table RESOLVES to a Table (P6-4a un-stub)');
            testCase.verifyClass(bt.iter_inner_content(), 'cell', ...
                'Body_.iter_inner_content RESOLVES to a cell array (P6-4a un-stub)');
        end

        function test_clean_save_fires_zero_stubs(testCase)
            % Regression (P2 EXIT proof): a clean mat2doc.Document().save() completes
            % with ZERO mat2doc:notYetPorted stubs on the open/save path. If any stub
            % were on the path save would throw; the file being written IS the
            % mechanical no-stub-fired proof (also the byte leg of Bar 1).
            d = mat2doc.Document();
            tmp = [tempname '.docx'];
            cleanup = onCleanup(@() deleteIfExists(tmp)); %#ok<NASGU>
            caught = [];
            try
                d.save(tmp);
            catch ME
                caught = ME;
            end
            testCase.verifyEmpty(caught, ...
                'clean Document().save must NOT fire any mat2doc:notYetPorted stub');
            testCase.verifyTrue(isfile(tmp), ...
                'clean Document().save must write the output file');
        end

    end

    % ===================== instance helpers ============================ %
    methods (Access = private)

        function root = parseDoc(testCase, inner)
            % Parse a <w:document><w:body>INNER</w:body></w:document> blob -> the
            % registered CT_Document root.
            xml = "<w:document xmlns:w=""" + testCase.W + """><w:body>" + ...
                inner + "</w:body></w:document>";
            root = mat2doc.oxml.parse_xml(uint8(unicode2native(char(xml), "UTF-8")));
        end

        function body = parseBodyOf(testCase, inner)
            % The CT_Body of a parsed document blob.
            body = testCase.parseDoc(inner).body;
        end

        function body = parseBody(testCase)
            % Parse a bare <w:body/> blob -> a standalone CT_Body (for H11 mutations).
            xml = "<w:body xmlns:w=""" + testCase.W + """/>";
            body = mat2doc.oxml.parse_xml(uint8(unicode2native(char(xml), "UTF-8")));
        end

        function names = localnames(~, elarr)
            % Ordered local-name string row for a (possibly empty) element array.
            n = numel(elarr);
            if n == 0
                names = strings(1, 0);
                return
            end
            names = strings(1, n);
            for i = 1:n
                names(i) = elarr(i).local_part;
            end
        end

        function names = parentLocalnames(~, elarr)
            % Ordered parent local-name string row for an element array.
            n = numel(elarr);
            if n == 0
                names = strings(1, 0);
                return
            end
            names = strings(1, n);
            for i = 1:n
                names(i) = elarr(i).getparent().local_part;
            end
        end

        function bytes = emitDocumentXml(~)
            % mat2doc.Document().save() to a temp .docx, extract word/document.xml,
            % return its raw bytes. Uses base-MATLAB unzip (no toolbox) into a temp
            % dir, both cleaned up on exit.
            d = mat2doc.Document();
            tmp = [tempname '.docx'];
            cleanTmp = onCleanup(@() deleteIfExists(tmp)); %#ok<NASGU>
            d.save(tmp);
            exdir = tempname;
            cleanEx = onCleanup(@() rmdirIfExists(exdir)); %#ok<NASGU>
            unzip(tmp, exdir);
            bytes = readBytes(fullfile(exdir, 'word', 'document.xml'));
        end
    end
end

% ===================== file-local helpers ============================== %

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

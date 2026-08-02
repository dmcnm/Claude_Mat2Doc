classdef Test_p8_3_unstub < matlab.unittest.TestCase
% TEST_P8_3_UNSTUB  Gate-4 permanent unit tests for Mat2Doc P8-3 -- the FINAL
%   port WP (the zero-stub / C4 exit sweep + the new Drawing tier).
%
%   Surface under test (ported from python-docx v1.2.0):
%     * src/docx/drawing/__init__.py::Drawing              -> +mat2doc\+drawing\Drawing.m
%         (NEW class -- has_picture, image, image ValueError path)
%     * src/docx/oxml/text/run.py::CT_R.inner_content_items -> +oxml\+text\CT_R.m
%     * src/docx/text/run.py::Run.iter_inner_content        -> +text\Run.m
%     * src/docx/document.py::Document.{add_page_break,paragraphs,tables,
%         inline_shapes}                                    -> +document\Document.m
%     * src/docx/parts/document.py::DocumentPart.inline_shapes -> +parts\DocumentPart.m
%     * src/docx/opc/package.py::OpcPackage._core_properties_part (KeyError branch)
%                                                           -> +opc\OpcPackage.m
%
%   P8-3 is the LAST port WP: after it the toolbox raises ZERO live
%   mat2doc:notYetPorted (exit condition C4). Every touched member is a read-side
%   accessor except Document.add_page_break, which composes already-live
%   serializers; the port is therefore OUTPUT-NEUTRAL and Gate-3 proved it L1
%   (18/18 byte-identical). This class pins the BEHAVIORAL surface + the one new
%   serialization substring (the page break) + the C4 zero-stub guard.
%
%   Provenance (Gate-1..3, all 2026-08-02):
%     * Audit    : validation\mat2doc\audit_P8-3_unstub.md
%                  (Porter Gate-1 126/126 + Opus Gate-2 APPROVE, zero defects, C4
%                  zero-stub grep independently confirmed, paired runtime probes
%                  character-identical to the python-docx v1.2.0 oracle).
%     * Validate : validation\mat2doc\validate_P8-3_unstub.md
%                  (Gate-3 PASS -- s0109 byte-neutrality 18/18 L1, s0110 API probe
%                  30/30 MATCH, s0111 reopen probe 14/14 MATCH incl. reopened-tree
%                  Drawing.image sha1 identity; ZERO new D-number).
%     * Scenarios: validation\mat2doc\scenarios\s0110_p8_3_api_probe.{py,m},
%                  s0111_p8_3_reopen_probe.{py,m} (the probe sequences mirrored below).
%     * Frozen refs (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE):
%         references\s0110\probe.json  -- the 30 API-probe oracle values (ORACLE_*);
%         references\s0111\probe.json  -- the 14 reopen-probe oracle values;
%         references\s0111\input.docx  -- the reopened rich-doc package. The oracle
%       VALUES are embedded below as Constants (ORACLE_*) and the reopen package +
%       the picture source are copied byte-for-byte into tests\drawing\data\
%       (p8_3_reopen.docx == references\s0111\input.docx; python-powered.png, the
%       embedded media whose sha1 IS b0a1e6cf..., co-located `* binary`
%       .gitattributes) so this suite is self-contained relative to the worktree.
%
%   Coverage taxonomy
%   -----------------
%   * Nominal  -- Drawing.has_picture over a real inline picture (build side); the
%     Document paragraphs/tables/inline_shapes delegators; add_page_break returns a
%     Paragraph.
%   * Edge     -- empty run -> 0 content items (both layers); a bare <w:drawing>
%     has_picture false; Drawing.image error path (verify IDENTIFIER
%     mat2doc:ValueError, not merely that it throws); non-ASCII `café 中 🎉`
%     paragraph text round-trips through parse+read (reopen fixture).
%   * Equivalence (frozen s0110/s0111 oracle) -- CT_R.inner_content_items /
%     Run.iter_inner_content document-order coalescing (str|.|str|drawing) with the
%     coalesced `"abc\ndef"` / `"ghi\tjkl"` spans; Drawing.image content_type +
%     sha1 media identity (build side AND reopened parsed tree); Document
%     paragraphs/tables/inline_shapes counts + order; InlineShape type/width EMU.
%   * Regression -- add_page_break emits `<w:br w:type="page"/>` byte-exact inside
%     the serialized document (L1 substring); the two coalesced text spans are the
%     hard-coded expected strings.
%   * C4 guard -- none of the four formerly-stubbed Document accessors
%     (add_page_break/paragraphs/tables/inline_shapes) raises mat2doc:notYetPorted
%     (positive resolution, mirroring Test_p2_3_document_shell's battery -> 0).
%
%   Deviations exercised: NONE new (Gate-3 confirmed zero new D-number). Adopt-only
%   D-001 (own OOXML parser/serializer, via parse_xml / serialize_part_xml).
%
%   Determinism: no network, no absolute paths -- the reopen package and the
%   picture source resolve relative to this file via
%   fileparts(mfilename('fullpath')); every parsed blob is an in-memory literal;
%   the reopen fixture is opened read-only; the fresh save (page-break serialize)
%   goes to serialize_part_xml IN MEMORY (no temp file). The picture-add build uses
%   the co-located PNG only.

    properties (Constant)
        % --- registered leaf classes (P4-3 / P7-3) the content-items resolve to ---
        CT_LRPB   = 'mat2doc.oxml.text.CT_LastRenderedPageBreak'
        CT_DRAW   = 'mat2doc.oxml.drawing.CT_Drawing'
        RPB_PROXY = 'mat2doc.text.RenderedPageBreak'
        DRAW_PROXY = 'mat2doc.drawing.Drawing'

        % --- frozen s0110 API-probe oracle (references\s0110\probe.json) ---
        ORACLE_IC_COUNT        = 4
        ORACLE_IIC_TYPES       = ["str" "CT_LastRenderedPageBreak" "str" "CT_Drawing"]
        ORACLE_IC_TYPES        = ["str" "RenderedPageBreak" "str" "Drawing"]
        ORACLE_IMG_CONTENT_TYPE = "image/png"
        ORACLE_IMG_SHA1        = "b0a1e6cf904691e6fa42bd9e72acc2b05280dc86"
        ORACLE_NONPIC_ERR_MSG  = "drawing does not contain a picture"
        ORACLE_INLINE0_TYPE    = "PICTURE"
        ORACLE_INLINE0_WIDTH   = 1778000

        % --- frozen s0111 reopen-probe oracle (references\s0111\probe.json) ---
        ORACLE_REOPEN_PARA_COUNT = 4
        ORACLE_REOPEN_PARA_TEXTS = ["Alpha" "café 中 🎉" "" ""]   % non-ASCII round-trip
        ORACLE_REOPEN_TBL_COUNT  = 1
        ORACLE_REOPEN_TBL_ROWS   = 2
        ORACLE_REOPEN_TBL_COLS   = 3
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied from tests\coreprops\Test_p1_7_coreprops.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\drawing
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. Drawing tier (drawing/__init__.py::Drawing) -- Gate-2         %
        %    VERIFY-TEST-REPIN explicitly asked for permanent Drawing tests %
        % =============================================================== %

        function test_drawing_has_picture_false_over_empty_drawing(testCase)
            % Edge: a bare <w:drawing/> (no wp:inline/pic:pic) -> has_picture FALSE.
            % Python bool(self._drawing.xpath(...)) over an empty match list is False
            % (H4). (s0110 nonpic_has_picture: false.)
            bare = testCase.parseFrag("<w:drawing " + testCase.nsW() + "/>");
            d = mat2doc.drawing.Drawing(bare, []);
            testCase.verifyFalse(d.has_picture, ...
                'bare <w:drawing/> must have has_picture == false');
        end

        function test_drawing_image_raises_ValueError_when_no_picture(testCase)
            % Edge / error path: Drawing.image over a non-picture drawing raises the
            % IDENTIFIER mat2doc:ValueError (NOT merely that it throws) with the
            % VERBATIM python-docx message. (s0110 nonpic_err. This is the
            % mat2ppt:<PyExceptionName> family = ValueError.)
            bare = testCase.parseFrag("<w:drawing " + testCase.nsW() + "/>");
            d = mat2doc.drawing.Drawing(bare, []);
            caught = [];
            try
                tmp = d.image; %#ok<NASGU>
            catch ME
                caught = ME;
            end
            testCase.assertNotEmpty(caught, 'Drawing.image on a non-picture must raise');
            testCase.verifyEqual(caught.identifier, 'mat2doc:ValueError', ...
                'Drawing.image must raise the IDENTIFIER mat2doc:ValueError');
            testCase.verifyEqual(string(caught.message), testCase.ORACLE_NONPIC_ERR_MSG, ...
                'Drawing.image ValueError message must be VERBATIM (byte-identical)');
        end

        function test_drawing_has_picture_true_and_image_over_real_picture(testCase)
            % Nominal + Equivalence (build side, s0110 G3): a REAL inline picture
            % (add_picture) reached via Run.iter_inner_content{1} is a Drawing with
            % has_picture TRUE; .image is an Image with content_type image/png and
            % sha1 == the embedded media identity b0a1e6cf... (the co-located PNG's
            % own sha1). Value-compared to the frozen s0110 oracle.
            d = mat2doc.Document();
            run = d.add_paragraph().add_run();
            run.add_picture(testCase.pngPath());
            items = run.iter_inner_content();
            testCase.verifyEqual(numel(items), 1, ...
                'a run holding one <w:drawing> yields exactly one content item');
            draw = items{1};
            testCase.verifyClass(draw, testCase.DRAW_PROXY, ...
                'the picture content item must be a mat2doc.drawing.Drawing');
            testCase.verifyTrue(draw.has_picture, ...
                'a real inline picture must have has_picture == true');
            img = draw.image;
            testCase.verifyClass(img, 'mat2doc.image.Image', ...
                'Drawing.image must return a mat2doc.image.Image');
            testCase.verifyEqual(string(img.content_type), testCase.ORACLE_IMG_CONTENT_TYPE, ...
                'Drawing.image content_type must be image/png (s0110 oracle)');
            testCase.verifyEqual(string(img.sha1), testCase.ORACLE_IMG_SHA1, ...
                'Drawing.image sha1 must equal the embedded media identity (s0110 oracle)');
        end

        % =============================================================== %
        % 2. CT_R.inner_content_items / Run.iter_inner_content -- Gate-2   %
        %    explicitly requested document-order + coalescing tests        %
        % =============================================================== %

        function test_inner_content_items_document_order_coalescing_oxml(testCase)
            % Equivalence / Regression (s0110 G1, oxml layer): CT_R.inner_content_items
            % over a run mixing w:t / w:cr / w:lastRenderedPageBreak / w:tab /
            % w:drawing returns items in DOCUMENT ORDER with adjacent text-runs
            % COALESCED (the w:cr renders "\n" INSIDE a span, the w:tab "\t"): 4 items
            % [ "abc\ndef", CT_LastRenderedPageBreak, "ghi\tjkl", CT_Drawing ].
            r = testCase.interleavedRun();
            iic = r.inner_content_items();
            testCase.verifyEqual(numel(iic), testCase.ORACLE_IC_COUNT, ...
                'inner_content_items must yield 4 items (2 coalesced spans + break + drawing)');
            testCase.verifyEqual(testCase.typeNames(iic), testCase.ORACLE_IIC_TYPES, ...
                'inner_content_items types must be [str, CT_LastRenderedPageBreak, str, CT_Drawing] in order');
            % Hard-coded expected coalesced spans (Regression).
            testCase.verifyEqual(iic{1}, "abc" + newline + "def", ...
                'span 1 must coalesce w:t/w:cr/w:t to "abc\ndef"');
            testCase.verifyEqual(iic{3}, "ghi" + sprintf('\t') + "jkl", ...
                'span 2 must coalesce w:t/w:tab/w:t to "ghi\tjkl"');
        end

        function test_iter_inner_content_document_order_coalescing_proxy(testCase)
            % Equivalence / Regression (s0110 G1, proxy layer): Run.iter_inner_content
            % maps the oxml items to proxies IN ORDER -- str stays str, the break ->
            % RenderedPageBreak, the drawing -> Drawing: 4 items
            % [ "abc\ndef", RenderedPageBreak, "ghi\tjkl", Drawing ].
            r = testCase.interleavedRun();
            ic = mat2doc.text.Run(r, []).iter_inner_content();
            testCase.verifyEqual(numel(ic), testCase.ORACLE_IC_COUNT, ...
                'iter_inner_content must yield 4 items');
            testCase.verifyEqual(testCase.typeNames(ic), testCase.ORACLE_IC_TYPES, ...
                'iter_inner_content types must be [str, RenderedPageBreak, str, Drawing] in order');
            testCase.verifyEqual(ic{1}, "abc" + newline + "def", ...
                'proxy span 1 coalesced text');
            testCase.verifyEqual(ic{3}, "ghi" + sprintf('\t') + "jkl", ...
                'proxy span 2 coalesced text');
        end

        function test_empty_run_zero_content_items(testCase)
            % Edge (s0110 G2): an empty <w:r/> yields ZERO content items at BOTH
            % layers -- oxml inner_content_items and proxy iter_inner_content.
            r0 = testCase.parseFrag("<w:r " + testCase.nsW() + "/>");
            testCase.verifyEqual(numel(r0.inner_content_items()), 0, ...
                'empty run: inner_content_items length 0');
            testCase.verifyEqual(numel(mat2doc.text.Run(r0, []).iter_inner_content()), 0, ...
                'empty run: iter_inner_content length 0');
        end

        % =============================================================== %
        % 3. Document delegators (add_page_break/paragraphs/tables/        %
        %    inline_shapes) + DocumentPart.inline_shapes parity            %
        % =============================================================== %

        function test_add_page_break_returns_paragraph_with_one_run(testCase)
            % Nominal (s0110 G5): Document.add_page_break composes
            % add_paragraph().add_run().add_break(WD_BREAK.PAGE) and returns the new
            % Paragraph, which holds exactly one run.
            d = mat2doc.Document();
            pb = d.add_page_break();
            testCase.verifyClass(pb, 'mat2doc.text.Paragraph', ...
                'add_page_break must return a mat2doc.text.Paragraph');
            testCase.verifyEqual(numel(pb.runs), 1, ...
                'the page-break paragraph must hold exactly one run');
        end

        function test_add_page_break_emits_page_break_xml(testCase)
            % Regression (L1 substring, Gate-3 VERIFY discharged): add_page_break
            % serializes `<w:br w:type="page"/>` byte-exact inside word/document.xml.
            % Serialized IN MEMORY via serialize_part_xml over the document root (no
            % temp file). String containment over the UTF-8 decode is a byte-level
            % assertion on the emitted run. (validate_P8-3 s1: document.xml L1.)
            d = mat2doc.Document();
            d.add_page_break();
            xml = string(native2unicode(mat2doc.oxml.serialize_part_xml(d.element()), "UTF-8"));
            testCase.verifyTrue(contains(xml, "<w:br w:type=""page""/>"), ...
                'add_page_break must emit <w:br w:type="page"/> byte-exact');
        end

        function test_paragraphs_and_tables_delegation(testCase)
            % Nominal (s0110 G5): Document.paragraphs / Document.tables delegate to
            % the body and return the expected proxy arrays in order. Two authored
            % paragraphs + a page-break paragraph = 3 paragraphs; one 2x3 table.
            d = mat2doc.Document();
            d.add_paragraph("Alpha");
            d.add_paragraph("Beta");
            d.add_page_break();
            d.add_table(2, 3);
            paras = d.paragraphs;
            testCase.verifyClass(paras, 'mat2doc.text.Paragraph', ...
                'Document.paragraphs must return a Paragraph array');
            testCase.verifyEqual(numel(paras), 3, ...
                'Document.paragraphs count (Alpha, Beta, page-break)');
            ptx = strings(1, numel(paras));
            for k = 1:numel(paras), ptx(k) = paras(k).text; end
            testCase.verifyEqual(ptx, ["Alpha" "Beta" ""], ...
                'Document.paragraphs text + order (page-break paragraph empty-text)');
            tbls = d.tables;
            testCase.verifyClass(tbls, 'mat2doc.table.Table', ...
                'Document.tables must return a Table array');
            testCase.verifyEqual(numel(tbls), 1, 'Document.tables count');
            testCase.verifyEqual(tbls(1).rows.len_(), testCase.ORACLE_REOPEN_TBL_ROWS, ...
                'the delegated table has 2 rows');
            testCase.verifyEqual(tbls(1).columns.len_(), testCase.ORACLE_REOPEN_TBL_COLS, ...
                'the delegated table has 3 columns');
        end

        function test_inline_shapes_document_vs_part_parity(testCase)
            % Equivalence (s0110 G5): Document.inline_shapes <-> DocumentPart.inline_shapes
            % length parity, and InlineShape[0] type/width match the frozen oracle
            % over a real add_picture.
            d = mat2doc.Document();
            d.add_picture(testCase.pngPath());
            doc_inline  = d.inline_shapes;
            part_inline = d.part.inline_shapes;
            testCase.verifyClass(doc_inline, 'mat2doc.shape.InlineShapes', ...
                'Document.inline_shapes must return an InlineShapes collection');
            testCase.verifyEqual(doc_inline.len_(), 1, 'Document.inline_shapes length 1');
            testCase.verifyEqual(part_inline.len_(), doc_inline.len_(), ...
                'Document.inline_shapes and DocumentPart.inline_shapes must have equal length (parity)');
            testCase.verifyEqual(string(doc_inline.getitem_(0).type), testCase.ORACLE_INLINE0_TYPE, ...
                'InlineShape[0].type must be PICTURE (s0110 oracle)');
            testCase.verifyEqual(double(doc_inline.getitem_(0).width), testCase.ORACLE_INLINE0_WIDTH, ...
                'InlineShape[0].width EMU must equal the s0110 oracle');
        end

        % =============================================================== %
        % 4. Reopened-package read surface (parsed tree, self-contained)   %
        %    -- the more stringent registry-populated path (s0111)         %
        % =============================================================== %

        function test_reopen_read_surface_equivalence(testCase)
            % Equivalence (frozen s0111): open the SAME on-disk rich-doc package as
            % the python-docx oracle and read the P8-3 accessor surface over the
            % PARSED tree. paragraphs (incl. non-ASCII `café 中 🎉` round-trip),
            % tables (2x3), inline_shapes/part.inline_shapes parity, and the picture
            % run's iter_inner_content -> Drawing with the reopened-tree image sha1.
            d = mat2doc.Document(testCase.reopenDocxPath());

            paras = d.paragraphs;
            testCase.verifyEqual(numel(paras), testCase.ORACLE_REOPEN_PARA_COUNT, ...
                'reopen: paragraph count');
            ptx = strings(1, numel(paras));
            for k = 1:numel(paras), ptx(k) = paras(k).text; end
            testCase.verifyEqual(ptx, testCase.ORACLE_REOPEN_PARA_TEXTS, ...
                'reopen: paragraph text + order (non-ASCII café 中 🎉 round-trips)');

            tbls = d.tables;
            testCase.verifyEqual(numel(tbls), testCase.ORACLE_REOPEN_TBL_COUNT, ...
                'reopen: table count');
            testCase.verifyEqual(tbls(1).rows.len_(), testCase.ORACLE_REOPEN_TBL_ROWS, ...
                'reopen: table rows');
            testCase.verifyEqual(tbls(1).columns.len_(), testCase.ORACLE_REOPEN_TBL_COLS, ...
                'reopen: table columns');

            testCase.verifyEqual(d.inline_shapes.len_(), 1, 'reopen: Document.inline_shapes len');
            testCase.verifyEqual(d.part.inline_shapes.len_(), d.inline_shapes.len_(), ...
                'reopen: inline_shapes parity');

            % Picture run: the last body paragraph holds the add_picture run.
            items = paras(end).runs(1).iter_inner_content();
            testCase.verifyEqual(numel(items), 1, 'reopen: picture run yields one Drawing');
            draw = items{1};
            testCase.verifyClass(draw, testCase.DRAW_PROXY, 'reopen: content item is a Drawing');
            testCase.verifyTrue(draw.has_picture, 'reopen: has_picture true');
            img = draw.image;
            testCase.verifyEqual(string(img.content_type), testCase.ORACLE_IMG_CONTENT_TYPE, ...
                'reopen: Drawing.image content_type');
            testCase.verifyEqual(string(img.sha1), testCase.ORACLE_IMG_SHA1, ...
                'reopen: Drawing.image sha1 == media identity on the PARSED tree (Gate-2 VERIFY)');
        end

        % =============================================================== %
        % 5. OpcPackage.core_properties_part_ default-creation branch      %
        % =============================================================== %

        function test_core_properties_found_branch_on_real_document(testCase)
            % Nominal (found branch): the default template ships docProps/core.xml,
            % so OpcPackage._core_properties_part returns the RELATED CorePropertiesPart
            % (part_related_by succeeds -- the KeyError branch is NOT taken). Reached
            % via the public Document core_properties accessor.
            d = mat2doc.Document();
            cp = d.core_properties;
            testCase.verifyClass(cp, 'mat2doc.opc.CoreProperties', ...
                'a real document core_properties resolves via the found branch');
        end

        function test_core_properties_default_creation_branch(testCase)
            % Equivalence / error-recovery (the P8-3 un-stub, opc/package.py 168-179):
            % a package with NO core-properties relationship makes part_related_by
            % raise mat2doc:KeyError; the un-stubbed branch CATCHES it, builds
            % CorePropertiesPart.default(self), relates it, and returns it. A bare
            % OpcPackage() (no parts/rels loaded) drives exactly that KeyError branch.
            % The created default carries the docx-faithful title "Word Document"
            % (CorePropertiesPart.default, coreprops.py) -- proving the branch ran.
            pkg = mat2doc.opc.OpcPackage();
            cp = pkg.core_properties;
            testCase.verifyClass(cp, 'mat2doc.opc.CoreProperties', ...
                'default-creation branch must return CoreProperties');
            testCase.verifyEqual(cp.title, "Word Document", ...
                'the default-created core.xml carries title "Word Document" (branch ran)');
            testCase.verifyEqual(cp.last_modified_by, "python-docx", ...
                'the default-created core.xml carries last_modified_by "python-docx"');
        end

        % =============================================================== %
        % 6. C4 zero-stub guard (positive resolution)                     %
        % =============================================================== %

        function test_no_stub_accessor_raises_notYetPorted(testCase)
            % C4 guard (mirrors Test_p2_3_document_shell's battery -> 0): the four
            % formerly-stubbed Document accessors now RESOLVE -- none raises the
            % IDENTIFIER mat2doc:notYetPorted. This is the whole-port C4 exit
            % condition asserted positively at the accessor surface.
            d = mat2doc.Document();
            accessors = { ...
                'paragraphs',     @() d.paragraphs; ...
                'tables',         @() d.tables; ...
                'inline_shapes',  @() d.inline_shapes; ...
                'add_page_break', @() d.add_page_break };
            for k = 1:size(accessors, 1)
                caught = [];
                try
                    accessors{k, 2}();
                catch ME
                    caught = ME;
                end
                if ~isempty(caught)
                    testCase.verifyNotEqual(caught.identifier, 'mat2doc:notYetPorted', ...
                        sprintf('Document.%s must NOT raise mat2doc:notYetPorted (C4 exit)', ...
                        accessors{k, 1}));
                end
            end
            % And each resolves to its documented type (belt-and-braces positive pin).
            testCase.verifyClass(d.paragraphs,     'mat2doc.text.Paragraph', 'paragraphs resolved');
            testCase.verifyClass(d.tables,         'mat2doc.table.Table',    'tables resolved');
            testCase.verifyClass(d.inline_shapes,  'mat2doc.shape.InlineShapes', 'inline_shapes resolved');
            testCase.verifyClass(d.add_page_break, 'mat2doc.text.Paragraph', 'add_page_break resolved');
        end

    end

    % ===================== instance helpers ============================ %
    methods (Access = private)

        function s = nsW(~)
            % The w namespace declaration (xmlns:w="...").
            s = mat2doc.oxml.nsdecls("w");
        end

        function root = parseFrag(~, frag)
            % Parse a loose-element XML fragment (string) -> registered root.
            root = mat2doc.oxml.parse_xml(frag);
        end

        function r = interleavedRun(testCase)
            % The s0110 G1 interleaved run: w:t/w:cr/w:t + w:lastRenderedPageBreak +
            % w:t/w:tab/w:t + w:drawing, mirroring s0110_p8_3_api_probe.m exactly.
            frag = "<w:r " + testCase.nsW() + ">" + ...
                "<w:t>abc</w:t><w:cr/><w:t>def</w:t>" + ...
                "<w:lastRenderedPageBreak/>" + ...
                "<w:t>ghi</w:t><w:tab/><w:t>jkl</w:t>" + ...
                "<w:drawing/>" + ...
                "</w:r>";
            r = mat2doc.oxml.parse_xml(frag);
        end

        function names = typeNames(~, items)
            % Per-item Python-equivalent short type name row: str stays "str"
            % (isinstance str), otherwise the leaf class name.
            names = strings(1, numel(items));
            for k = 1:numel(items)
                x = items{k};
                if isstring(x) || ischar(x)
                    names(k) = "str";
                else
                    parts = split(string(class(x)), ".");
                    names(k) = parts(end);
                end
            end
        end

        function p = pngPath(~)
            % The co-located picture source (sha1 b0a1e6cf..., == the s0110 media).
            here = fileparts(mfilename('fullpath'));   % tests\drawing
            p = char(fullfile(here, 'data', 'python-powered.png'));
        end

        function p = reopenDocxPath(~)
            % The co-located reopen package (== references\s0111\input.docx).
            here = fileparts(mfilename('fullpath'));   % tests\drawing
            p = char(fullfile(here, 'data', 'p8_3_reopen.docx'));
        end
    end
end

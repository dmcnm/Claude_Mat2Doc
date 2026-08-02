classdef Test_p7_3_drawing_shape < matlab.unittest.TestCase
% TEST_P7_3_DRAWING_SHAPE  Gate-4 permanent unit tests for Mat2Doc P7-3 [N]
%   (DrawingML inline-picture oxml + InlineShape API): src/docx/oxml/shape.py +
%   src/docx/oxml/drawing.py -> +mat2doc\+oxml\+shape\ (~17 CT classes) +
%   +mat2doc\+oxml\+drawing\CT_Drawing, and src/docx/shape.py -> +mat2doc\+shape\
%   InlineShapes / InlineShape, plus the SIXTEEN P7-3 registry rows in
%   +mat2doc\+oxml\registry.m (a:blip..wp:inline, __init__.py:47-62).
%
%   P7-3 is the DrawingML-inline-picture FOUNDATION WP: it ports the oxml
%   *builder* CT_Inline.new_pic_inline (+ its sub-builder CT_Picture.new), the
%   whole picture accessor chain (CT_Inline/CT_Picture/CT_Blip/CT_PositiveSize2D/
%   CT_Transform2D/CT_Point2D/CT_ShapeProperties/...), and the InlineShapes /
%   InlineShape proxy API. It is registry-adding and M1-NEUTRAL (default.docx has
%   no <w:drawing>; none of the 16 tags occurs in any of the 17 default.docx
%   parts, so nothing transits the new CT classes on the save path).
%
%   ==== WHY new_pic_inline IS THE HEADLINE (the P7-4 foundation guard) ====
%   CT_Inline.new_pic_inline(shape_id, rId, filename, cx, cy) builds the COMPLETE
%   wp:inline / a:graphic / pic:pic tree that P7-4's Document.add_picture emits.
%   If a single byte of this builder drifts, every add_picture document diverges.
%   test_new_pic_inline_byte_pins freezes all 5 param sets (A porter / B XML-
%   escaping / C UTF-8 CJK+emoji / D tiny-EMU / E large) byte-identical to the
%   python-docx oracle, plus CT_Picture.new (test_new_pic_byte_pins). A regression
%   goes RED loudly here.
%
%   This class permanently freezes the guarantees the prior gates established:
%     * Gate-1 Porter  : audit_P7-3_drawing_shape.md (self-probe).
%     * Gate-2 Auditor (Fable): APPROVE.
%     * Gate-3 Validator: validate_P7-3_drawing_shape.md -- PASS, ZERO new
%       D-numbers, NO re-pin list (0 flips). new_pic_inline 10/10 byte-identical
%       (5 sets x {inline, pic}); the real add_picture drawing round-trip +
%       width/height mutation byte-identical; M1 17/17 (styles.xml 349458 B /
%       02d71a68 AND document.xml 1548 B / 0e4dd503 byte-identical); the 33-probe
%       CT-accessor + InlineShapes/InlineShape API MATCH; targeted regression
%       127/127, EXACTLY 0 flips (registry addition flip-neutral).
%
%   THE HIGHEST-VALUE PERMANENT PINS IN THIS WP (a regression goes RED loudly):
%     * (foundation) test_new_pic_inline_byte_pins -- the 5 new_pic_inline byte
%       sets == frozen s0087 (A ce13bdb0/958 B ; B 7e070b29 escaping ; C c3818617
%       UTF-8 4-byte emoji F0 9F 9A 80 ; D c9611ad1 tiny-EMU cx=cy=1 ; E 3b078bbd).
%       THIS is what P7-4's add_picture must emit.
%     * (foundation) test_new_pic_byte_pins -- CT_Picture.new pic:pic 5 sets.
%     * (H8) test_h8_pic_ns_suppression -- the emitted pic:pic inside the inline
%       tree carries ZERO xmlns decls (the already-signed D-serializer-nsdecl
%       verbatim-until-moved manifestation); the only xmlns decls are the 4 on the
%       wp:inline root. A regression that re-adds xmlns:pic/a/r on the moved pic
%       fails here.
%     * (R) test_drawing_roundtrip_byte_pin / test_drawing_mutation_byte_pin --
%       a REAL add_picture word/document.xml parses->serializes byte-identical
%       (d172f803, 2334 B); setting InlineShape.width/height (which writes BOTH
%       wp:extent AND pic:spPr extent) re-serializes byte-identical to python-docx
%       (bf69e34d, 2332 B).
%     * (H1) test_inline_shapes_api -- getitem_ 0-based + negative-wrap; the
%       out-of-range IndexError carries the VERBATIM python-docx message
%       "inline shape index [%d] out of range" with the ORIGINAL idx (5 and -3),
%       under identifier mat2doc:IndexError.
%     * (H10) test_inline_shape_type_5way -- WD_INLINE_SHAPE by graphicData @uri:
%       PICTURE / LINKED_PICTURE (blip has r:link) / CHART / SMART_ART /
%       NOT_IMPLEMENTED, all by enum NAME.
%     * (M) test_m1_styles_and_document -- Document().save() -> styles.xml
%       349458 B / 02d71a68 AND document.xml 1548 B / 0e4dd503 (the 16-row
%       registry addition is byte-neutral). SHA equality is L1.
%
%   Provenance (all Gate-3 frozen 2026-08-01):
%     * Audit    : validation\mat2doc\audit_P7-3_drawing_shape.md
%     * Validate : validation\mat2doc\validate_P7-3_drawing_shape.md
%     * Scenarios: s0087_p7_3_new_pic_inline.{py,m} (10 byte fixtures);
%                  s0088_p7_3_drawing_roundtrip.{py,m} (round-trip + mutation);
%                  s0089_p7_3_api_probe.{py,m} (its probe body is replayed
%                  VERBATIM by runApiProbes() below for the Equivalence leg).
%     * Frozen refs (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE at Gate-3):
%         references\s0087\ (10 XML byte-fixtures + manifest_bytes.json) copied
%           verbatim into tests\shape\data\s0087\ WITH a co-located `.gitattributes`
%           `* binary` pin (frozen-byte fixtures must not be line-ending mangled on
%           the master checkout -- the Gate-4 byte-fixture lesson).
%         references\s0088\ (doc_roundtrip.xml, doc_mutated.xml, probe.json,
%           manifest_bytes.json) copied into tests\shape\data\s0088\ (`* binary`).
%         references\s0089\probe.json copied into tests\shape\data\s0089\ (the
%           33-probe API oracle; jsondecode is line-ending agnostic).
%         references\s0001\parts\word\{styles,document}.xml -- the M1 byte
%           references (SHA of what Document().save() emits); NOT copied.
%
%   Coverage taxonomy
%   -----------------
%   * Nominal      -- new_pic_inline / new_pic set A; the CT-accessor deep chains
%                     (extent/docPr/graphic; nvPicPr/blipFill/spPr; blip embed;
%                     xfrm off/ext; spPr cx/cy); InlineShapes len/getitem/type;
%                     the 16 registry rows.
%   * Edge         -- B XML-escaping filename (&/"/< in docPr/cNvPr @name); C
%                     non-ASCII (é + CJK 图片 + emoji 🚀 built from DECIMAL DOUBLES,
%                     surrogate-safe -> 4-byte UTF-8); D tiny-EMU cx=cy=1; the H3
%                     tri-state None (blip link None; xfrm.cx None with no a:ext;
%                     spPr.cx None with no a:xfrm); getitem_ negative-wrap;
%                     getitem_(5)/(-3) IndexError verbatim (error path).
%   * Equivalence  -- test_equivalence_s0089_full_battery replays the ENTIRE s0089
%                     probe sequence (runApiProbes, the .m twin's body verbatim)
%                     and compares EVERY tagged leaf to the frozen s0089 oracle
%                     copied into data\s0089\probe.json (0 diffs at Gate-3).
%   * Regression   -- hard-coded set-A serialized XML string (ASCII == byte-
%                     identical L1) + SHA-256 of the 10 s0087 fixtures + the
%                     s0088 round-trip/mutation docs + the M1 styles/document parts.
%   * Upstream     -- the s0088 doc_roundtrip.xml IS a real python-docx
%                     Document().add_picture("monty-truth.png") document.xml; its
%                     bytes ARE lxml's expected serialization for that operation.
%
%   Byte-level (L1) note: every serialized-XML comparison is either the FULL
%   serialize_part_xml output compared byte-for-byte (uint8) against a frozen
%   fixture, or its SHA-256 vs the frozen oracle, or (set A) an ASCII string
%   equality (string-equality == byte-equality for pure ASCII). NO D-number
%   granted any L2 relaxation in this WP (Gate-3: zero new, none at L2), so every
%   pin here is L1. The only looser-than-byte check is the equivalence leaf-count
%   floor, commented at its site.
%
%   Determinism: no network, no absolute paths. The worktree root and the
%   co-located s0087/s0088/s0089 fixtures resolve relative to this file via
%   fileparts(mfilename('fullpath')); saves go to tempname .docx deleted via
%   onCleanup; every file read is binary ('r','n'). The +mat2doc package resolves
%   via the MANDATORY PathFixture(worktree-root) in TestClassSetup (WP9-F4 lesson).

    properties (Constant)
        % --- the 16 P7-3 registry rows (+oxml\registry.m 345-360) ---
        %     {tag, expected fully-qualified class}
        REG_ROWS = { ...
            "a:blip",        "mat2doc.oxml.shape.CT_Blip"; ...
            "a:ext",         "mat2doc.oxml.shape.CT_PositiveSize2D"; ...
            "a:graphic",     "mat2doc.oxml.shape.CT_GraphicalObject"; ...
            "a:graphicData", "mat2doc.oxml.shape.CT_GraphicalObjectData"; ...
            "a:off",         "mat2doc.oxml.shape.CT_Point2D"; ...
            "a:xfrm",        "mat2doc.oxml.shape.CT_Transform2D"; ...
            "pic:blipFill",  "mat2doc.oxml.shape.CT_BlipFillProperties"; ...
            "pic:cNvPr",     "mat2doc.oxml.shape.CT_NonVisualDrawingProps"; ...
            "pic:nvPicPr",   "mat2doc.oxml.shape.CT_PictureNonVisual"; ...
            "pic:pic",       "mat2doc.oxml.shape.CT_Picture"; ...
            "pic:spPr",      "mat2doc.oxml.shape.CT_ShapeProperties"; ...
            "w:drawing",     "mat2doc.oxml.drawing.CT_Drawing"; ...
            "wp:anchor",     "mat2doc.oxml.shape.CT_Anchor"; ...
            "wp:docPr",      "mat2doc.oxml.shape.CT_NonVisualDrawingProps"; ...
            "wp:extent",     "mat2doc.oxml.shape.CT_PositiveSize2D"; ...
            "wp:inline",     "mat2doc.oxml.shape.CT_Inline" }

        % --- frozen s0001 M1 byte references (P7-3 is M1-neutral; the 16 rows are
        %     byte-neutral -- validate_P7-3 section 1) ---
        STYLES_SIZE_M1 = 349458
        STYLES_SHA_M1  = "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"
        DOC_SIZE_M1    = 1548
        DOC_SHA_M1     = "0e4dd503bc095219cd24e8312b372c588ca56bb834aacac53eafb7bf47836327"

        % --- frozen s0088 drawing round-trip / mutation byte references ---
        RT_SIZE  = 2334
        RT_SHA   = "d172f803e4a6cf63d7dc85129efd7553572f16e6ce3dec83e37343df18b1c99b"
        MUT_SIZE = 2332
        MUT_SHA  = "bf69e34d516d4503201fa7f9052ede0771498e4e682878235f962f29730d9a7c"
        MUT_CX   = 1000000
        MUT_CY   = 750000
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied verbatim from tests\oxml\Test_p6_2_table_props.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\shape
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. registry -- the 16 P7-3 rows resolve to their CT classes       %
        % =============================================================== %

        function test_registry_resolves_sixteen_rows(testCase)
            % Nominal / Regression (registry.m 345-360): the 16 P7-3 DrawingML
            % inline-picture rows. REGISTRATION is what flips the parse class of a
            % picture subtree; OxmlElement(tag) confirms the parser instantiates it.
            rows = testCase.REG_ROWS;
            for i = 1:size(rows, 1)
                tag = rows{i, 1}; cls = char(rows{i, 2});
                r = mat2doc.oxml.registry(mat2doc.oxml.qn(tag));
                testCase.verifyEqual(char(r), cls, ...
                    sprintf('registry must resolve %s -> %s (P7-3 row)', tag, cls));
                testCase.verifyEqual(class(mat2doc.oxml.OxmlElement(tag)), cls, ...
                    sprintf('OxmlElement(%s) must be a %s', tag, cls));
            end
        end

        % =============================================================== %
        % 2. ★★ new_pic_inline byte-pins -- the P7-4 add_picture FOUNDATION  %
        % =============================================================== %

        function test_new_pic_inline_byte_pins(testCase)
            % ★★ FOUNDATION GUARD (byte-identical L1): CT_Inline.new_pic_inline
            % builds the EXACT wp:inline tree P7-4's add_picture emits. For each of
            % the 5 frozen param sets, the built bytes must be byte-identical to the
            % s0087 fixture AND SHA-256 == the frozen manifest. A single-byte drift
            % here diverges every future add_picture document.
            sets = paramSets();
            for i = 1:numel(sets)
                s = sets(i);
                inline = mat2doc.oxml.shape.CT_Inline.new_pic_inline( ...
                    s.shape_id, s.rId, s.filename, ...
                    mat2doc.shared.Emu(s.cx), mat2doc.shared.Emu(s.cy));
                outBytes = mat2doc.oxml.serialize_part_xml(inline);

                fx = readFixture('s0087', "new_pic_inline_" + s.label + ".xml");
                testCase.verifyEqual(numel(fx), s.inline_size, ...
                    sprintf('new_pic_inline %s frozen fixture size', s.label));
                testCase.verifyEqual(sha256hex(fx), s.inline_sha, ...
                    sprintf('new_pic_inline %s frozen fixture SHA-256 (intact)', s.label));
                testCase.verifyEqual(uint8(outBytes(:)'), uint8(fx(:)'), ...
                    sprintf('new_pic_inline %s built bytes must be byte-identical to frozen (P7-4 foundation)', s.label));
                testCase.verifyEqual(sha256hex(outBytes), s.inline_sha, ...
                    sprintf('new_pic_inline %s built SHA-256 == frozen oracle (L1)', s.label));
            end
        end

        function test_new_pic_byte_pins(testCase)
            % ★ FOUNDATION GUARD (byte-identical L1): the CT_Picture.new pic:pic
            % sub-builder new_pic_inline feeds. 5 sets byte-identical + SHA.
            sets = paramSets();
            for i = 1:numel(sets)
                s = sets(i);
                pic = mat2doc.oxml.shape.CT_Picture.new(0, s.filename, s.rId, ...
                    mat2doc.shared.Emu(s.cx), mat2doc.shared.Emu(s.cy));
                outBytes = mat2doc.oxml.serialize_part_xml(pic);

                fx = readFixture('s0087', "new_pic_" + s.label + ".xml");
                testCase.verifyEqual(numel(fx), s.pic_size, ...
                    sprintf('new_pic %s frozen fixture size', s.label));
                testCase.verifyEqual(sha256hex(fx), s.pic_sha, ...
                    sprintf('new_pic %s frozen fixture SHA-256 (intact)', s.label));
                testCase.verifyEqual(uint8(outBytes(:)'), uint8(fx(:)'), ...
                    sprintf('new_pic %s built bytes must be byte-identical to frozen', s.label));
                testCase.verifyEqual(sha256hex(outBytes), s.pic_sha, ...
                    sprintf('new_pic %s built SHA-256 == frozen oracle (L1)', s.label));
            end
        end

        function test_new_pic_inline_setA_hardcoded(testCase)
            % Regression (hard-coded expected XML, ASCII == byte-identical L1): set
            % A's full serialized wp:inline. This is the exact string P7-4's
            % add_picture(python-image.png, 2438400x1828800, rId7, shape_id 1)
            % must produce. Loudly explicit so a structural drift is human-readable.
            inline = mat2doc.oxml.shape.CT_Inline.new_pic_inline( ...
                1, "rId7", "python-image.png", ...
                mat2doc.shared.Emu(2438400), mat2doc.shared.Emu(1828800));
            got = string(native2unicode(mat2doc.oxml.serialize_part_xml(inline), "UTF-8"));

            A = "http://schemas.openxmlformats.org/drawingml/2006/main";
            PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
            R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
            expected = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>" + newline + ...
                "<wp:inline xmlns:wp=""" + WP + """ xmlns:a=""" + A + """ xmlns:pic=""" + PIC + """ xmlns:r=""" + R + """>" + ...
                "<wp:extent cx=""2438400"" cy=""1828800""/>" + ...
                "<wp:docPr id=""1"" name=""Picture 1""/>" + ...
                "<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect=""1""/></wp:cNvGraphicFramePr>" + ...
                "<a:graphic><a:graphicData uri=""" + PIC + """>" + ...
                "<pic:pic><pic:nvPicPr><pic:cNvPr id=""0"" name=""python-image.png""/><pic:cNvPicPr/></pic:nvPicPr>" + ...
                "<pic:blipFill><a:blip r:embed=""rId7""/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>" + ...
                "<pic:spPr><a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""2438400"" cy=""1828800""/></a:xfrm>" + ...
                "<a:prstGeom prst=""rect""/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline>";
            testCase.verifyEqual(got, expected, ...
                'set-A new_pic_inline hard-coded serialized bytes (L1) -- the P7-4 add_picture oracle');
            % H14: the docPr @name is "Picture %d" % shape_id -> "Picture 1"
            testCase.verifyEqual(string(inline.docPr.name), "Picture 1", ...
                'H14: wp:docPr name is "Picture 1" (shape_id 1)');
        end

        % =============================================================== %
        % 3. ★ H8 -- pic:pic ns-suppression (D-serializer-nsdecl manifest)   %
        % =============================================================== %

        function test_h8_pic_ns_suppression(testCase)
            % ★ (H8) Regression: after insert_pic_ MOVES the freshly-built pic:pic
            % into the wp:inline tree, the emitted <pic:pic> subtree carries ZERO
            % xmlns declarations -- the pic/a/r prefixes resolve from the wp:inline
            % root's fresh nsdecls (verbatim-until-moved reconciliation, the
            % already-signed D-serializer-nsdecl behaviour). The ONLY xmlns decls in
            % the serialized inline are the 4 on the wp:inline root. A regression
            % that re-adds xmlns:pic/a/r on the moved pic FAILS here. Checked on
            % every one of the 5 sets.
            sets = paramSets();
            for i = 1:numel(sets)
                s = sets(i);
                inline = mat2doc.oxml.shape.CT_Inline.new_pic_inline( ...
                    s.shape_id, s.rId, s.filename, ...
                    mat2doc.shared.Emu(s.cx), mat2doc.shared.Emu(s.cy));
                xml = string(native2unicode(mat2doc.oxml.serialize_part_xml(inline), "UTF-8"));

                % exactly 4 xmlns decls total (all on wp:inline: wp, a, pic, r)
                testCase.verifyEqual(count(xml, "xmlns"), 4, ...
                    sprintf('%s: exactly 4 xmlns decls total (only on wp:inline root)', s.label));
                % the pic:pic open tag is BARE <pic:pic> (no attributes, no xmlns)
                testCase.verifyTrue(contains(xml, "<pic:pic>"), ...
                    sprintf('%s: pic:pic open tag is bare <pic:pic> (H8)', s.label));
                % the entire pic subtree (from <pic:pic> onward) carries no xmlns
                picSubtree = extractAfter(xml, "<pic:pic>");
                testCase.verifyFalse(contains(picSubtree, "xmlns"), ...
                    sprintf('%s: the moved pic:pic subtree emits NO xmlns decls (D-serializer-nsdecl)', s.label));
            end
        end

        % =============================================================== %
        % 4. Fixture-drift guard -- s0087 manifest == the hard-coded pins    %
        % =============================================================== %

        function test_s0087_manifest_matches_pinned(testCase)
            % Regression (fixture-drift guard): the shipped data\s0087\
            % manifest_bytes.json SHAs must equal the pin table paramSets() uses.
            % A silent re-freeze would flip the manifest but not this constant.
            here = fileparts(mfilename('fullpath'));
            man = loadJson(fullfile(here, 'data', 's0087', 'manifest_bytes.json'));
            sets = paramSets();
            for i = 1:numel(sets)
                s = sets(i);
                inl = pickArtifact(man.artifacts, 'new_pic_inline', s.label);
                testCase.verifyEqual(inl.size, s.inline_size, ...
                    sprintf('manifest new_pic_inline %s size == pin', s.label));
                testCase.verifyEqual(string(inl.sha256), s.inline_sha, ...
                    sprintf('manifest new_pic_inline %s SHA == pin', s.label));
                pc = pickArtifact(man.artifacts, 'new_pic', s.label);
                testCase.verifyEqual(pc.size, s.pic_size, ...
                    sprintf('manifest new_pic %s size == pin', s.label));
                testCase.verifyEqual(string(pc.sha256), s.pic_sha, ...
                    sprintf('manifest new_pic %s SHA == pin', s.label));
            end
        end

        % =============================================================== %
        % 5. CT accessors -- the picture chain + the H3 tri-state None       %
        % =============================================================== %

        function test_ct_accessors(testCase)
            % Nominal + Edge (s0089 accessor probes): the CT_Inline / CT_Picture
            % deep accessor chains + the H3 tri-state None, each tied to the frozen
            % s0089 oracle field.
            o = loadJson(fullfile(fileparts(mfilename('fullpath')), 'data', 's0089', 'probe.json'));

            % ---- CT_Inline chain (extent / docPr / graphic.graphicData.uri) ----
            inline = mat2doc.oxml.shape.CT_Inline.new_pic_inline(1, "rId7", ...
                "python-image.png", mat2doc.shared.Emu(2438400), mat2doc.shared.Emu(1828800));
            testCase.verifyEqual(ci(inline.extent.cx), string(o.inline_extent_cx), 'inline.extent.cx');
            testCase.verifyEqual(ci(inline.extent.cy), string(o.inline_extent_cy), 'inline.extent.cy');
            testCase.verifyEqual(ci(inline.docPr.id),  string(o.inline_docPr_id),  'inline.docPr.id');
            testCase.verifyEqual(cs(inline.docPr.name), string(o.inline_docPr_name), 'inline.docPr.name');
            testCase.verifyEqual(cs(inline.graphic.graphicData.uri), ...
                string(o.inline_graphicData_uri), 'inline.graphic.graphicData.uri');

            % ---- CT_Picture deep chain (nvPicPr/blipFill/spPr + xfrm off/ext) ----
            pic = mat2doc.oxml.shape.CT_Picture.new(0, "python-image.png", "rId7", ...
                mat2doc.shared.Emu(2438400), mat2doc.shared.Emu(1828800));
            testCase.verifyEqual(ci(pic.nvPicPr.cNvPr.id),   string(o.pic_cNvPr_id),   'pic.nvPicPr.cNvPr.id');
            testCase.verifyEqual(cs(pic.nvPicPr.cNvPr.name), string(o.pic_cNvPr_name), 'pic.nvPicPr.cNvPr.name');
            testCase.verifyEqual(cn(pic.blipFill.blip.embed), string(o.pic_blip_embed), 'pic.blipFill.blip.embed (r:embed)');
            testCase.verifyEqual(cn(pic.blipFill.blip.link),  string(o.pic_blip_link),  'pic.blipFill.blip.link None (H3)');
            testCase.verifyEqual(ci(pic.spPr.cx), string(o.pic_spPr_cx), 'pic.spPr.cx (CT_ShapeProperties)');
            testCase.verifyEqual(ci(pic.spPr.cy), string(o.pic_spPr_cy), 'pic.spPr.cy');
            testCase.verifyEqual(ci(pic.spPr.xfrm.off.x), string(o.pic_xfrm_off_x), 'pic.spPr.xfrm.off.x (CT_Point2D)');
            testCase.verifyEqual(ci(pic.spPr.xfrm.off.y), string(o.pic_xfrm_off_y), 'pic.spPr.xfrm.off.y');
            testCase.verifyEqual(ci(pic.spPr.xfrm.ext.cx), string(o.pic_xfrm_ext_cx), 'pic.spPr.xfrm.ext.cx (CT_PositiveSize2D)');
            testCase.verifyEqual(ci(pic.spPr.xfrm.ext.cy), string(o.pic_xfrm_ext_cy), 'pic.spPr.xfrm.ext.cy');

            % ---- CT_PresetGeometry2D: generic @prst read (empty, unregistered) ----
            prstgeom = mat2doc.oxml.parse_xml("<a:prstGeom " + mat2doc.oxml.nsdecls("a") + " prst='rect'/>");
            testCase.verifyEqual(cn(prstgeom.get("prst")), string(o.prstGeom_prst), 'a:prstGeom @prst generic read -> rect');

            % ---- CT_Blip tri-state (H3): both present / embed-only ----
            blip_both = mat2doc.oxml.parse_xml("<a:blip " + mat2doc.oxml.nsdecls("a","r") + " r:embed='rId3' r:link='rId9'/>");
            testCase.verifyEqual(cn(blip_both.embed), string(o.blip_both_embed), 'blip both embed');
            testCase.verifyEqual(cn(blip_both.link),  string(o.blip_both_link),  'blip both link');
            blip_eo = mat2doc.oxml.parse_xml("<a:blip " + mat2doc.oxml.nsdecls("a","r") + " r:embed='rId3'/>");
            testCase.verifyEqual(cn(blip_eo.link), string(o.blip_embedonly_link), 'blip embed-only -> link None (H3)');

            % ---- CT_Transform2D: cx None when no a:ext (H3); off.x present ----
            xfrm_noext = mat2doc.oxml.parse_xml("<a:xfrm " + mat2doc.oxml.nsdecls("a") + "><a:off x='5' y='6'/></a:xfrm>");
            testCase.verifyEqual(ci(xfrm_noext.off.x), string(o.xfrm_off_x), 'xfrm.off.x = 5');
            testCase.verifyEqual(cn(xfrm_noext.cx), string(o.xfrm_cx_noext), 'xfrm.cx None when a:ext absent (H3)');

            % ---- CT_ShapeProperties: cx None when no a:xfrm (H3) ----
            sppr_noxfrm = mat2doc.oxml.parse_xml("<pic:spPr " + mat2doc.oxml.nsdecls("pic","a") + ...
                "><a:prstGeom prst='rect'/></pic:spPr>");
            testCase.verifyEqual(cn(sppr_noxfrm.cx), string(o.sppr_cx_noxfrm), 'spPr.cx None when a:xfrm absent (H3)');
        end

        % =============================================================== %
        % 6. ★ drawing round-trip byte-pin (a REAL add_picture doc)          %
        % =============================================================== %

        function test_drawing_roundtrip_byte_pin(testCase)
            % ★ (R) Regression + Upstream (byte-identical L1): a REAL python-docx
            % Document().add_picture("monty-truth.png") word/document.xml parses
            % through the registered CT_* classes and re-serializes BYTE-IDENTICAL.
            rtBytes = readFixture('s0088', 'doc_roundtrip.xml');
            testCase.verifyEqual(numel(rtBytes), testCase.RT_SIZE, 'doc_roundtrip.xml frozen size');
            testCase.verifyEqual(sha256hex(rtBytes), testCase.RT_SHA, 'doc_roundtrip.xml frozen SHA (intact)');

            docroot = mat2doc.oxml.parse_xml(rtBytes);
            outBytes = mat2doc.oxml.serialize_part_xml(docroot);
            testCase.verifyEqual(uint8(outBytes(:)'), uint8(rtBytes(:)'), ...
                'real add_picture document.xml parse->serialize must be byte-identical');
            testCase.verifyEqual(sha256hex(outBytes), testCase.RT_SHA, ...
                'round-trip re-serialized SHA == frozen oracle (L1)');
        end

        function test_drawing_mutation_byte_pin(testCase)
            % ★ (R) Regression (byte-identical L1): setting InlineShape.width/height
            % writes BOTH wp:extent AND the picture pic:spPr extent (the two must
            % stay in sync); re-serializing the whole document is byte-identical to
            % python-docx applying the same mutation+save. Also probes the API over
            % the real doc against the frozen s0088 probe.json.
            rtBytes = readFixture('s0088', 'doc_roundtrip.xml');
            docroot = mat2doc.oxml.parse_xml(rtBytes);
            bodies = docroot.xpath("//w:body");
            body = bodies(1);
            shapes = mat2doc.shape.InlineShapes(body, []);   % parent [] (VERIFY-3 unused)
            sh = shapes.getitem_(0);                          % inline_shapes[0]

            % --- API probe over the real doc vs frozen s0088/probe.json ---
            pj = loadJson(fullfile(fileparts(mfilename('fullpath')), 'data', 's0088', 'probe.json'));
            testCase.verifyEqual(ci(shapes.len_()), string(pj.len), 's0088 len');
            t = sh.type;
            testCase.verifyEqual("str|" + string(t), string(pj.type_name), 's0088 type name PICTURE');
            testCase.verifyEqual(ci(double(t.value)), string(pj.type_value), 's0088 type value 3');
            testCase.verifyEqual(ci(double(sh.width)),  string(pj.width_before),  's0088 width_before 1905000');
            testCase.verifyEqual(ci(double(sh.height)), string(pj.height_before), 's0088 height_before 2717800');

            % --- mutate width/height, re-serialize whole document ---
            sh.width  = mat2doc.shared.Emu(testCase.MUT_CX);
            sh.height = mat2doc.shared.Emu(testCase.MUT_CY);
            testCase.verifyEqual(ci(double(sh.width)),  string(pj.width_after),  's0088 width_after 1000000');
            testCase.verifyEqual(ci(double(sh.height)), string(pj.height_after), 's0088 height_after 750000');

            outBytes = mat2doc.oxml.serialize_part_xml(docroot);
            mutBytes = readFixture('s0088', 'doc_mutated.xml');
            testCase.verifyEqual(numel(mutBytes), testCase.MUT_SIZE, 'doc_mutated.xml frozen size');
            testCase.verifyEqual(sha256hex(mutBytes), testCase.MUT_SHA, 'doc_mutated.xml frozen SHA (intact)');
            testCase.verifyEqual(uint8(outBytes(:)'), uint8(mutBytes(:)'), ...
                'width/height mutation whole-document re-serialize must be byte-identical');
            testCase.verifyEqual(sha256hex(outBytes), testCase.MUT_SHA, ...
                'mutated re-serialized SHA == frozen oracle (L1)');
        end

        % =============================================================== %
        % 7. InlineShapes -- len / getitem / negative-wrap / IndexError      %
        % =============================================================== %

        function test_inline_shapes_api(testCase)
            % Nominal + Edge (H1 + error path): getitem_ is 0-based with Python
            % negative-index wrap; the out-of-range IndexError carries the VERBATIM
            % python-docx message with the ORIGINAL idx, under mat2doc:IndexError.
            rtBytes = readFixture('s0088', 'doc_roundtrip.xml');
            docroot = mat2doc.oxml.parse_xml(rtBytes);
            body = firstOf(docroot.xpath("//w:body"));
            shapes = mat2doc.shape.InlineShapes(body, []);

            testCase.verifyEqual(shapes.len_(), 1, 'len_ = 1 (one inline picture)');
            % getitem_(0) -> a PICTURE; width read
            sh0 = shapes.getitem_(0);
            testCase.verifyClass(sh0, 'mat2doc.shape.InlineShape', 'getitem_(0) is an InlineShape');
            testCase.verifyEqual(string(sh0.type), "PICTURE", 'getitem_(0).type PICTURE');
            testCase.verifyEqual(double(sh0.width), 1905000, 'getitem_(0).width');
            % H1 negative wrap: getitem_(-1) -> last (== index 0 here)
            testCase.verifyEqual(double(shapes.getitem_(-1).width), 1905000, ...
                'H1: getitem_(-1) negative-wrap -> last element');
            % to_array materializes 1x1
            arr = shapes.to_array();
            testCase.verifyEqual(numel(arr), 1, 'to_array materializes 1 InlineShape');

            % ★ error path: getitem_(5) -> mat2doc:IndexError, verbatim message
            testCase.verifyError(@() shapes.getitem_(5), 'mat2doc:IndexError', ...
                'getitem_(5) out of range -> mat2doc:IndexError');
            testCase.verifyEqual(errmsg(@() shapes.getitem_(5)), ...
                "inline shape index [5] out of range", ...
                'IndexError message verbatim (original idx 5)');
            % ★ error path: getitem_(-3) -> the ORIGINAL (negative) idx in the message
            testCase.verifyError(@() shapes.getitem_(-3), 'mat2doc:IndexError', ...
                'getitem_(-3) out of range -> mat2doc:IndexError');
            testCase.verifyEqual(errmsg(@() shapes.getitem_(-3)), ...
                "inline shape index [-3] out of range", ...
                'IndexError message verbatim carries the ORIGINAL negative idx (-3)');
        end

        % =============================================================== %
        % 8. (H10) InlineShape.type -- the 5-way graphicData-uri dispatch     %
        % =============================================================== %

        function test_inline_shape_type_5way(testCase)
            % (H10) Nominal + Edge: WD_INLINE_SHAPE member implied by the a:graphicData
            % @uri (and, for a picture, whether the a:blip carries r:link). All by
            % enum NAME (matching the frozen s0089 oracle). PICTURE / LINKED_PICTURE
            % / CHART / SMART_ART / NOT_IMPLEMENTED.
            o = loadJson(fullfile(fileparts(mfilename('fullpath')), 'data', 's0089', 'probe.json'));
            m = mat2doc.oxml.nsmap();
            testCase.verifyEqual(cs(string(inlineType(m.pic, false))), string(o.type_PICTURE), ...
                'pic uri, no link -> PICTURE');
            testCase.verifyEqual(cs(string(inlineType(m.pic, true))), string(o.type_LINKED_PICTURE), ...
                'pic uri + r:link -> LINKED_PICTURE');
            testCase.verifyEqual(cs(string(inlineType(m.c, false))), string(o.type_CHART), ...
                'chart uri -> CHART');
            testCase.verifyEqual(cs(string(inlineType(m.dgm, false))), string(o.type_SMART_ART), ...
                'diagram uri -> SMART_ART');
            testCase.verifyEqual(cs(string(inlineType("http://example.com/other", false))), ...
                string(o.type_NOT_IMPLEMENTED), 'foreign uri -> NOT_IMPLEMENTED');
        end

        % =============================================================== %
        % 9. (M) M1 styles.xml + document.xml byte-pins (registry-neutral)   %
        % =============================================================== %

        function test_m1_styles_and_document(testCase)
            % (M) Regression (byte-neutrality, L1): mat2doc.Document().save() emits
            % word/styles.xml at EXACTLY 349458 B / 02d71a68 and word/document.xml at
            % 1548 B / 0e4dd503 even though the 16 drawing rows are now registered --
            % default.docx has no <w:drawing> and none of the 16 tags transits any of
            % the 17 parts, so the addition is byte-neutral. SHA equality is L1.
            % A single save() emits both parts.
            [styBytes, docBytes] = emitTwoParts('styles.xml', 'document.xml');

            testCase.verifyEqual(numel(styBytes), testCase.STYLES_SIZE_M1, ...
                sprintf('word/styles.xml must be exactly %d B', testCase.STYLES_SIZE_M1));
            testCase.verifyEqual(sha256hex(styBytes), testCase.STYLES_SHA_M1, ...
                'word/styles.xml SHA-256 == frozen s0001 oracle (16-row registry byte-neutral, L1)');

            testCase.verifyEqual(numel(docBytes), testCase.DOC_SIZE_M1, ...
                sprintf('word/document.xml must be exactly %d B', testCase.DOC_SIZE_M1));
            testCase.verifyEqual(sha256hex(docBytes), testCase.DOC_SHA_M1, ...
                'word/document.xml SHA-256 == frozen s0001 oracle (byte-identical L1)');
        end

        % =============================================================== %
        % 10. EQUIVALENCE -- the full s0089 battery vs the frozen oracle       %
        % =============================================================== %

        function test_equivalence_s0089_full_battery(testCase)
            % Equivalence: replay the ENTIRE s0089 probe sequence (runApiProbes --
            % the .m twin's body VERBATIM) and compare EVERY tagged leaf to the
            % frozen python-docx 1.2.0 oracle copied into data\s0089\probe.json.
            % Gate-3 found ZERO divergences (probe_diff 33/33 MATCH), so every leaf
            % must be byte/value-identical.
            port   = runApiProbes();
            oracle = loadJson(fullfile(fileparts(mfilename('fullpath')), 'data', 's0089', 'probe.json'));

            pKeys = sort(fieldnames(port));
            oKeys = sort(fieldnames(oracle));
            testCase.verifyEqual(pKeys, oKeys, ...
                'replayed battery and frozen oracle must have identical probe keys');
            % Non-trivial floor guard (guards a silent-empty replay). The only
            % looser-than-byte assertion in this class; justified as a leaf-count floor.
            testCase.verifyGreaterThanOrEqual(numel(oKeys), 30, ...
                'the s0089 oracle must expose the full 33-probe battery');
            for i = 1:numel(oKeys)
                k = oKeys{i};
                testCase.verifyEqual(string(port.(k)), string(oracle.(k)), ...
                    sprintf('probe %s must be byte/value-identical to the frozen oracle', k));
            end
        end

    end
end

% ===================== file-local helpers ============================== %

function sets = paramSets()
    % The 5 frozen s0087 param sets + their pinned sizes/SHAs
    % (manifest_bytes.json / validate_P7-3 section 2). The C-set filename is built
    % from DECIMAL DOUBLES (surrogate-safe): "café 图片 <U+1F680>.png" -- NOT 0x hex
    % literals, which MATLAB types as uint8 and saturates codepoints > 255 to 255
    % (the Gate-3 harness lesson). char() of a double array is the correct path.
    fn_C = string(char([99 97 102 233 32 22270 29255 32 55357 56960 46 112 110 103]));

    sets = struct( ...
        'label',      {"A_porter", "B_escaping", "C_utf8", "D_tiny_emu", "E_distinct"}, ...
        'shape_id',   {1, 43, 7, 999, 12345}, ...
        'rId',        {"rId7", "rId99", "rId1", "rId12", "rId2048"}, ...
        'filename',   {"python-image.png", 'Amp & "quote" <tag>.png', fn_C, "spaces in name.png", "a.png"}, ...
        'cx',         {2438400, 914400, 1000000, 1, 6858000}, ...
        'cy',         {1828800, 914400, 500000, 1, 9144000}, ...
        'inline_size',{958, 984, 961, 941, 958}, ...
        'inline_sha', { ...
            "ce13bdb0a75f41ebb57942c205ec2039d69b69933f1b55f98b50963494771490", ...
            "7e070b294beef9abe8ab2af9d09c5e61b860838500f749ff751d61f2b438da56", ...
            "c38186175e7d18ded0333681013340e7ab576d363ee0f4360e09dbc81b2e716e", ...
            "c9611ad125c4eaef16b456d1c0253a4ca9339168c521e395e5f100864921268f", ...
            "3b078bbd49033dd28ba8b9d304502ec95451a40c9350ddbb9cdc92c0fdd6dfa4"}, ...
        'pic_size',   {577, 603, 581, 568, 569}, ...
        'pic_sha',    { ...
            "9d28d220df9007bcd7d78bca66a2d60376d046dab8314a732ace1d608bbc431f", ...
            "7ff5d154949c45e99f2e185e31778294dd87f57bc7c8de80e895c4bd976be5b3", ...
            "f69bd193590eee143b01d8c1727c49f2799444e0128c6ad4954ddf483266c94f", ...
            "ca5d952615d414a74a9c27dce9c7aac984dc87568ce45de6acfaa894f3db5b29", ...
            "3aa2ed724e3f4f5be2fe4db3bf4c730df1a1d4db5a660c1c79c8c376334ff564"});
end

function P = runApiProbes()
    % Replay the s0089 probe sequence (the .m twin's P-struct body, VERBATIM
    % tags/inputs/order) and return the struct of tagged canonical strings.
    % Embedded here so the Equivalence leg is self-contained (the validation-folder
    % scenario is NOT on the toolbox path and must not be a dependency). Mirrors
    % validation\mat2doc\scenarios\s0089_p7_3_api_probe.m lines 20-82.
    %
    % NB: NO `import mat2doc...` here -- a specific (non-.* ) import is resolved at
    % suite-CREATION PARSE time, BEFORE TestClassSetup's PathFixture puts +mat2doc
    % on the path, so it errors MATLAB:undefinedVarOrClass at suite build (the
    % "specific import fails in test class" lesson). Everything is fully qualified,
    % which resolves at RUN time (after PathFixture).
    Emu = @(v) mat2doc.shared.Emu(v);
    parse_xml = @(x) mat2doc.oxml.parse_xml(x);
    nsdecls = @(varargin) mat2doc.oxml.nsdecls(varargin{:});

    P = struct();

    inline = mat2doc.oxml.shape.CT_Inline.new_pic_inline(1, "rId7", "python-image.png", Emu(2438400), Emu(1828800));
    P.inline_extent_cx = ci(inline.extent.cx);
    P.inline_extent_cy = ci(inline.extent.cy);
    P.inline_docPr_id  = ci(inline.docPr.id);
    P.inline_docPr_name = cs(inline.docPr.name);
    P.inline_graphicData_uri = cs(inline.graphic.graphicData.uri);

    pic = mat2doc.oxml.shape.CT_Picture.new(0, "python-image.png", "rId7", Emu(2438400), Emu(1828800));
    P.pic_cNvPr_id   = ci(pic.nvPicPr.cNvPr.id);
    P.pic_cNvPr_name = cs(pic.nvPicPr.cNvPr.name);
    P.pic_blip_embed = cn(pic.blipFill.blip.embed);
    P.pic_blip_link  = cn(pic.blipFill.blip.link);
    P.pic_spPr_cx = ci(pic.spPr.cx);
    P.pic_spPr_cy = ci(pic.spPr.cy);
    P.pic_xfrm_off_x = ci(pic.spPr.xfrm.off.x);
    P.pic_xfrm_off_y = ci(pic.spPr.xfrm.off.y);
    P.pic_xfrm_ext_cx = ci(pic.spPr.xfrm.ext.cx);
    P.pic_xfrm_ext_cy = ci(pic.spPr.xfrm.ext.cy);

    prstgeom = parse_xml("<a:prstGeom " + nsdecls("a") + " prst='rect'/>");
    P.prstGeom_prst = cn(prstgeom.get("prst"));

    blip_both = parse_xml("<a:blip " + nsdecls("a","r") + " r:embed='rId3' r:link='rId9'/>");
    P.blip_both_embed = cn(blip_both.embed);
    P.blip_both_link  = cn(blip_both.link);
    blip_embed_only = parse_xml("<a:blip " + nsdecls("a","r") + " r:embed='rId3'/>");
    P.blip_embedonly_link = cn(blip_embed_only.link);

    xfrm_noext = parse_xml("<a:xfrm " + nsdecls("a") + "><a:off x='5' y='6'/></a:xfrm>");
    P.xfrm_off_x = ci(xfrm_noext.off.x);
    P.xfrm_cx_noext = cn(xfrm_noext.cx);

    sppr_noxfrm = parse_xml("<pic:spPr " + nsdecls("pic","a") + "><a:prstGeom prst='rect'/></pic:spPr>");
    P.sppr_cx_noxfrm = cn(sppr_noxfrm.cx);

    rt_path = fullfile(fileparts(mfilename('fullpath')), 'data', 's0088', 'doc_roundtrip.xml');
    body = parse_xml(readBytes(rt_path));
    shapes = mat2doc.shape.InlineShapes(body, []);
    P.shapes_len        = ci(shapes.len_());
    P.shapes_0_type     = cs(string(shapes.getitem_(0).type));
    P.shapes_0_width    = ci(shapes.getitem_(0).width);
    P.shapes_neg1_width = ci(shapes.getitem_(-1).width);
    P.shapes_idx5_err     = errcap(@() shapes.getitem_(5));
    P.shapes_idxneg3_err  = errcap(@() shapes.getitem_(-3));

    m = mat2doc.oxml.nsmap();
    P.type_PICTURE         = cs(string(inlineType(m.pic, false)));
    P.type_LINKED_PICTURE  = cs(string(inlineType(m.pic, true)));
    P.type_CHART           = cs(string(inlineType(m.c, false)));
    P.type_SMART_ART       = cs(string(inlineType(m.dgm, false)));
    P.type_NOT_IMPLEMENTED = cs(string(inlineType("http://example.com/other", false)));
end

function t = inlineType(uri, set_link)
    % InlineShape.type for a given graphicData uri (+ optional r:link) -- mirrors
    % the s0089 .m twin inline_type_ helper.
    el = mat2doc.oxml.shape.CT_Inline.new_pic_inline(1, "rId7", "x.png", ...
        mat2doc.shared.Emu(1), mat2doc.shared.Emu(1));
    el.graphic.graphicData.uri = uri;
    if set_link
        el.graphic.graphicData.pic.blipFill.blip.link = "rId5";
    end
    t = mat2doc.shape.InlineShape(el).type;
end

function s = ci(v)
    % canonical int tag "int|<decimal>" (matches the s0089 oracle format)
    s = "int|" + mat2doc.shared.pyStr(double(v), "int");
end

function s = cs(v)
    s = "str|" + string(v);
end

function s = cn(v)
    % None ([] / <missing>) -> "none|", else "str|<value>"
    if isequal(v, []) || (isstring(v) && isscalar(v) && ismissing(v))
        s = "none|";
    else
        s = "str|" + string(v);
    end
end

function s = errcap(fn)
    try
        fn();
        s = "NO-ERROR";
    catch e
        s = "err|" + string(e.message);
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

function x = firstOf(arr)
    x = arr(1);
end

function a = pickArtifact(artifacts, kind, label)
    % Select the manifest artifact with the given kind + label. jsondecode returns
    % a CELL array here (the per-artifact `args` sub-objects have heterogeneous
    % fields: new_pic_inline has shape_id, new_pic has pic_id), so index with {}.
    for i = 1:numel(artifacts)
        if iscell(artifacts)
            e = artifacts{i};
        else
            e = artifacts(i);
        end
        if strcmp(e.kind, kind) && strcmp(e.label, label)
            a = e; return
        end
    end
    error('mat2doc:test:noArtifact', 'no %s artifact labelled %s', kind, label);
end

function o = loadJson(p)
    % Read a co-located JSON file in BINARY mode (no CRLF translation) and decode
    % UTF-8 -> struct. jsondecode is whitespace/line-ending agnostic.
    fid = fopen(p, 'r', 'n');
    if fid < 0
        error('mat2doc:test:openFail', 'cannot open %s', p);
    end
    raw = fread(fid, Inf, '*uint8')';
    fclose(fid);
    o = jsondecode(native2unicode(raw, 'UTF-8'));
end

function b = readFixture(scenario, name)
    here = fileparts(mfilename('fullpath'));
    b = readBytes(fullfile(here, 'data', scenario, char(name)));
end

function b = readBytes(p)
    f = fopen(p, 'r', 'n');            % binary read (no CRLF translation)
    assert(f >= 0, 'could not open for read: %s', p);
    b = fread(f, Inf, '*uint8')';
    fclose(f);
end

function [aBytes, bBytes] = emitTwoParts(leafA, leafB)
    % Document().save() to a temp .docx, unzip once, return two word/<leaf> parts.
    d = mat2doc.Document();
    tmp = [tempname '.docx'];
    cleanTmp = onCleanup(@() deleteIfExists(tmp)); %#ok<NASGU>
    d.save(tmp);
    exdir = tempname;
    cleanEx = onCleanup(@() rmdirIfExists(exdir)); %#ok<NASGU>
    unzip(tmp, exdir);
    aBytes = readBytes(fullfile(exdir, 'word', leafA));
    bBytes = readBytes(fullfile(exdir, 'word', leafB));
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

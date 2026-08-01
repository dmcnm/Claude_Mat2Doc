classdef Test_p7_1b_png_gif_bmp < matlab.unittest.TestCase
% TEST_P7_1B_PNG_GIF_BMP  Gate-4 permanent unit tests for Mat2Doc P7-1b
%   (the PNG / GIF / BMP image-header parsers, WITH the PIL->docx dpi reversion).
%
%   Surface under test (ported from python-docx v1.2.0 src/docx/image/{png,gif,
%   bmp}.py; re-ported off the Mat2Ppt +image parsers with the PIL contract
%   STRIPPED):
%     - mat2doc.image.Png / PngParser_ / Chunks_ / ChunkParser_ / ChunkFactory_ /
%       Chunk_ / IHDRChunk_ / pHYsChunk_  (png.py) -- the PNG chunk-walk + dpi.
%     - mat2doc.image.Gif  (gif.py)  -- little-endian "<HH" px dims; dpi 72/72.
%     - mat2doc.image.Bmp  (bmp.py)  -- little-endian BITMAPINFOHEADER; dpi via
%       Bmp.dpi_ (the docx px_per_meter*0.0254 formula, 0->96).
%   These un-stub the P7-1a notYetPorted Png/Gif/Bmp rows; the P7-1a factory
%   (ImageHeaderFactory_) now dispatches PNG/GIF/BMP signatures to the real
%   parsers. Jfif/Exif/Tiff remain notYetPorted (P7-2).
%
%   Provenance (Gate-1..3, all 2026-08-01):
%     * Audit    : validation\mat2doc\audit_P7-1b_png_gif_bmp.md
%                  (Porter Gate-1 + Fable Gate-2 APPROVE -- dpi reversion proven).
%     * Validate : validation\mat2doc\validate_P7-1b_png_gif_bmp.md
%                  (Gate-3 PASS, 143/143 value-identical vs python-docx 1.2.0,
%                  ZERO new D-numbers; pure-parsing WP, no package bytes).
%     * Scenario : validation\mat2doc\scenarios\s0084_p7_1b_png_gif_bmp_probe.{py,m}
%     * Frozen ref (python-docx 1.2.0 oracle, frozen ONCE):
%         references\s0084\probe.json -- the 143-value oracle, copied verbatim
%         into tests\image\data\s0084_probe.json (co-located `s0084_probe.json
%         text eol=lf` .gitattributes) so this suite is self-contained; the 7 real
%         test images (150-dpi.png, 300-dpi.png, monty-truth.png, python-icon.png,
%         python-powered.png, python.bmp, sonic.gif) copied into tests\image\data\
%         (co-located `* binary` .gitattributes -- they are the sha1 END-TO-END
%         equivalence blobs; all 7 sha1 re-verified byte-identical to the clone).
%     * M1 byte-pin values reuse the frozen s0001 manifest (word/styles.xml
%       02d71a68..., word/document.xml 0e4dd503...) already owned by
%       Test_p1_8_skeleton_m1; image parsing is M1-neutral (touches no oxml
%       registry row, no PartFactory registration, nothing on the save path).
%
%   Coverage taxonomy
%   -----------------
%   * ★ Equivalence (headline) -- test_replay_all_143_records replays the ENTIRE
%     frozen Gate-3 battery (the exact probe sequence of s0084_..._probe.m,
%     rebuilt inline through mat2doc.image.*) and asserts the port reproduces
%     every one of the 143 records' tagged canon value-identical. Self-contained:
%     the 7 images + oracle are copied into tests\image\data\.
%   * ★ Image equivalence (regression) -- Image.from_file on the real corpus:
%     300-dpi.png->300/300, 150-dpi.png->150/150, no-pHYs PNGs->72/72,
%     python.bmp->96/96, sonic.gif->72/72; sha1/content_type/ext/px/dpi/EMU all
%     hard-pinned to the frozen s0084 oracle. sha1 is END-TO-END on the real blob.
%   * ★★ dpi TIE-DISCRIMINATOR pins (the reversion guard -- LOUD) -- Bmp.dpi_ /
%     PngParser_.dpi_ direct static calls on the four decisive discriminators
%     (0->96, 2500->64, 7500->190, 17500->444) plus the sanity/odd-tie battery.
%     THESE ARE THE ONLY CASES THAT CATCH A SILENT REVERT to MATLAB round() or the
%     PIL formula -- see the LOUD comment on test_bmp_dpi_tie_discriminators.
%   * Reverted-dpi->EMU (regression) -- crafted BMP blobs (make_bmp) through
%     Image.from_blob: the reverted dpi flows into width_EMU/height_EMU (ppm 0 ->
%     dpi 96 -> width_EMU 381000, NOT PIL's 72->508000).
%   * Chunk-parse fidelity (regression) -- 150-dpi.png (IHDR|pHYs|iCCP|cHRM|IDAT|
%     IEND, 901x1350, pHYs 5906|5906|1->150 dpi) + monty-truth.png (IHDR|tEXt|IDAT|
%     IEND, 150x214, no pHYs->72), a BMP header (211x71, 96/96), a GIF header
%     (290x360, little-endian).
%   * Factory dispatch (edge) -- PNG/GIF/BMP blobs dispatch to the REAL parsers
%     (return an Image with the right px/dpi); the Jfif/Exif/Tiff stubs still raise
%     mat2doc:notYetPorted (P7-2 boundary).
%   * ★ M1 byte-pin (regression) -- mat2doc.Document().save() -> word/styles.xml
%     sha256 02d71a68... + word/document.xml sha256 0e4dd503... (image parsing is
%     M1-neutral; RED here would mean the re-port perturbed the save path).
%
%   Deviations exercised (adopt-only, ZERO new D-numbers -- Gate-3): the dpi is the
%   sole deviation risk and it is the python-docx formula EXACTLY (the Mat2Ppt
%   D-bmp-dpi PIL contract is REVERTED); no serialization occurs in the parsers so
%   no L1 byte surface is touched (equivalence is value-level; validate section 7).
%   Error paths pin the mat2doc: IDENTIFIER (UnexpectedEndOfFileError / notYetPorted).
%
%   Determinism: no network, no hard-coded absolute paths -- the frozen oracle and
%   the 7 test images resolve relative to this file via
%   fileparts(mfilename('fullpath')). Every image read is BINARY (no CRLF xlate).

    properties (Constant)
        % ★ Image-core equivalence oracle (references\s0084\probe.json, IMG_* rows).
        % Cols: tag, filename, content_type, ext, px_w, px_h, hdpi, vdpi,
        % width_EMU, height_EMU, sha1 (END-TO-END on the real blob).
        IMG_ORACLE = { ...
            "PNG150",  "150-dpi.png",        "image/png", "png", 901, 1350, 150, 150, 5492496, 8229600, "0531f61e611277d04eb2def465cb4197b06d2f68"; ...
            "PNG300",  "300-dpi.png",        "image/png", "png", 860,  579, 300, 300, 2621280, 1764792, "ca11f589af080a26b619b76823897001d34c7e44"; ...
            "PNGMONTY","monty-truth.png",    "image/png", "png", 150,  214,  72,  72, 1905000, 2717800, "79769f1e202add2e963158b532e36c2c0f76a70c"; ...
            "PNGICON", "python-icon.png",    "image/png", "png",  24,   24,  72,  72,  304800,  304800, "e4bf438f958a276cb5f57d0600194edebd853cfb"; ...
            "PNGPWR",  "python-powered.png", "image/png", "png", 140,   56,  72,  72, 1778000,  711200, "b0a1e6cf904691e6fa42bd9e72acc2b05280dc86"; ...
            "BMP",     "python.bmp",         "image/bmp", "bmp", 211,   71,  96,  96, 2009774,  676275, "438d5cf4a5fc27799d98f658234e0aa6d5d80181"; ...
            "GIF",     "sonic.gif",          "image/gif", "gif", 290,  360,  72,  72, 3683000, 4572000, "328eefd6c4a25b1311065d4aff8305be43d9ee6c"};

        % ★★ Bmp.dpi_ tie-discriminator oracle (BDPI_* rows). {px_per_meter, dpi}.
        % 0->96 and 2500->64 discriminate the PIL formula; 7500->190 and
        % 17500->444 are floor-EVEN ties that discriminate MATLAB native round().
        BDPI_ORACLE = { ...
            0, 96; 1, 0; 2500, 64; 3780, 96; 7500, 190; ...
            11811, 300; 12500, 318; 17500, 444; 22500, 572; 27500, 698};

        % PngParser_.dpi_ oracle (PDPI_* rows). {units_specifier, px_per_unit, dpi}.
        PDPI_ORACLE = { ...
            1, 2500, 64; 1, 7500, 190; 1, 12500, 318; 1, 17500, 444; ...
            1, 11811, 300; 0, 2835, 72; 1, 0, 72; 2, 2835, 72; 3, 2835, 72};

        % Crafted-BMP reverted-dpi->EMU oracle (CBMP_* rows). Cols: tag, w, h,
        % hppm, vppm, px_w, px_h, hdpi, vdpi, width_EMU, height_EMU.
        CBMP_ORACLE = { ...
            "bmp0",     40, 30,     0,     0, 40, 30,  96,  96, 381000, 285750; ...
            "bmp2500",  40, 30,  2500,  2500, 40, 30,  64,  64, 571500, 428625; ...
            "bmp3780",  40, 30,  3780,  3780, 40, 30,  96,  96, 381000, 285750; ...
            "bmp7500", 100, 50,  7500,  7500,100, 50, 190, 190, 481263, 240631; ...
            "bmp17500",100, 50, 17500, 17500,100, 50, 444, 444, 205945, 102972};

        % ★ M1 byte-pin (== Test_p1_8_skeleton_m1 rows; python-docx 1.2.0 s0001
        % oracle). SHA-256 equality IS byte-identity (L1). P7-1b is M1-neutral.
        M1_STYLES_SHA   = "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"
        M1_DOCUMENT_SHA = "0e4dd503bc095219cd24e8312b372c588ca56bb834aacac53eafb7bf47836327"
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied from the proven sibling tests\image\Test_p7_1a_image_core.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\image
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % ★ 1. Image equivalence -- END-TO-END real corpus (headline)     %
        % =============================================================== %

        function test_image_equivalence_corpus(testCase)
            % ★ Regression: Image.from_file on the 7 REAL docx test images (read
            % from the self-contained tests\image\data\ copies) reproduces every
            % emitted field value-identical to the frozen s0084 oracle. The
            % px/dpi/content_type are parsed END-TO-END by the newly-live PNG/GIF/
            % BMP parsers (no seeded test-double); sha1 hashes the real blob; the
            % EMU values run the genuine Inches(px/dpi) Length math. Covers the WP
            % dpi checklist: 300->300, 150->150, no-pHYs->72, BMP->96, GIF->72.
            here = fileparts(mfilename('fullpath'));
            for r = 1:size(testCase.IMG_ORACLE, 1)
                tag      = testCase.IMG_ORACLE{r, 1};
                fn       = testCase.IMG_ORACLE{r, 2};
                wantCT   = testCase.IMG_ORACLE{r, 3};
                wantExt  = testCase.IMG_ORACLE{r, 4};
                wantPxW  = testCase.IMG_ORACLE{r, 5};
                wantPxH  = testCase.IMG_ORACLE{r, 6};
                wantHDpi = testCase.IMG_ORACLE{r, 7};
                wantVDpi = testCase.IMG_ORACLE{r, 8};
                wantWEmu = testCase.IMG_ORACLE{r, 9};
                wantHEmu = testCase.IMG_ORACLE{r, 10};
                wantSha1 = testCase.IMG_ORACLE{r, 11};

                path = fullfile(here, 'data', char(fn));
                img  = mat2doc.image.Image.from_file(path);

                testCase.verifyEqual(string(img.content_type), string(wantCT), sprintf('%s content_type', tag));
                testCase.verifyEqual(string(img.ext),          string(wantExt), sprintf('%s ext', tag));
                testCase.verifyEqual(string(img.filename),     string(fn),      sprintf('%s filename', tag));
                testCase.verifyEqual(double(img.px_width),  wantPxW,  sprintf('%s px_width', tag));
                testCase.verifyEqual(double(img.px_height), wantPxH,  sprintf('%s px_height', tag));
                testCase.verifyEqual(double(img.horz_dpi),  wantHDpi, sprintf('%s horz_dpi', tag));
                testCase.verifyEqual(double(img.vert_dpi),  wantVDpi, sprintf('%s vert_dpi', tag));
                testCase.verifyEqual(double(img.width),  wantWEmu, sprintf('%s width EMU', tag));
                testCase.verifyEqual(double(img.height), wantHEmu, sprintf('%s height EMU', tag));
                % ★ sha1 END-TO-END (hashes the real blob).
                testCase.verifyEqual(string(img.sha1), string(wantSha1), ...
                    sprintf('%s sha1 END-TO-END must equal the frozen oracle', tag));
            end
        end

        % =============================================================== %
        % ★★ 2. dpi tie discriminators -- THE REVERSION GUARD (LOUD)      %
        % =============================================================== %

        function test_bmp_dpi_tie_discriminators(testCase)
            % ★★★ LOUD REVERSION GUARD. Bmp.dpi_(px_per_meter) is the python-docx
            % formula: px_per_meter==0 -> 96; else int(round(px_per_meter*0.0254))
            % with CPython half-to-EVEN rounding (pyRound), NOT MATLAB's half-away
            % round() and NOT the PIL /39.3701 formula.
            %
            %   ==> THESE TIE DISCRIMINATORS ARE THE ONLY CASES THAT CATCH A SILENT
            %   ==> REVERT TO MATLAB round() OR THE PIL FORMULA. A REGRESSION HERE
            %   ==> MEANS THE PIL->docx dpi REVERSION BROKE. <==
            %
            %   * 0 -> 96        (docx default; PIL would give 72)
            %   * 2500 -> 64     (round(63.5)=64 floor-ODD; PIL /39.3701 gives 63)
            %   * 7500 -> 190    (190.5 EXACT tie, floor-EVEN: CPython rounds DOWN;
            %                     MATLAB native round() would give 191)
            %   * 17500 -> 444   (444.5 EXACT tie, floor-EVEN: 445 if native round)
            %   * 12500 -> 318   (317.5 tie, floor-ODD contrast: CPython rounds UP)
            %   * 22500 -> 572, 27500 -> 698 (further floor-EVEN ties: 573/699 if
            %                     native round)
            %   * 3780 -> 96, 11811 -> 300 (non-tie sanity)  ; 1 -> 0 (tiny)
            % (audit section 3a / validate section 3, frozen s0084 BDPI_* rows.)
            for i = 1:size(testCase.BDPI_ORACLE, 1)
                ppm  = testCase.BDPI_ORACLE{i, 1};
                want = testCase.BDPI_ORACLE{i, 2};
                got  = double(mat2doc.image.Bmp.dpi_(ppm));
                testCase.verifyEqual(got, want, sprintf( ...
                    ['Bmp.dpi_(%d) must be %d (python-docx half-to-even; a %d here ' ...
                     'means a silent revert to MATLAB round() or the PIL formula)'], ...
                    ppm, want, want + 1));
            end
        end

        function test_png_dpi(testCase)
            % PngParser_.dpi_(units_specifier, px_per_unit): units==1 (meters) &&
            % ppu~=0 -> int(round(ppu*0.0254)) (same half-to-even tie battery as
            % BMP); units~=1 -> 72; ppu==0 -> 72 (H4 truthiness). Frozen PDPI_* rows.
            for i = 1:size(testCase.PDPI_ORACLE, 1)
                u    = testCase.PDPI_ORACLE{i, 1};
                ppu  = testCase.PDPI_ORACLE{i, 2};
                want = testCase.PDPI_ORACLE{i, 3};
                got  = double(mat2doc.image.PngParser_.dpi_(u, ppu));
                testCase.verifyEqual(got, want, ...
                    sprintf('PngParser_.dpi_(%d,%d) must be %d', u, ppu, want));
            end
        end

        % =============================================================== %
        % 3. Reverted dpi flows into EMU (crafted BMP, from_blob)         %
        % =============================================================== %

        function test_reverted_dpi_flows_into_emu(testCase)
            % Regression: byte-identical crafted BMP blobs (make_bmp, the s0084
            % make_bmp bytes) driven END-TO-END through Image.from_blob -- the
            % reverted dpi lands in px/dpi AND in the derived width_EMU/height_EMU
            % (Inches(px/dpi)). The ppm-0 blob -> dpi 96 -> width_EMU 381000 (NOT
            % PIL's 72 -> 508000). Frozen s0084 CBMP_* rows.
            for r = 1:size(testCase.CBMP_ORACLE, 1)
                tag   = testCase.CBMP_ORACLE{r, 1};
                blob  = make_bmp(testCase.CBMP_ORACLE{r, 2}, testCase.CBMP_ORACLE{r, 3}, ...
                                 testCase.CBMP_ORACLE{r, 4}, testCase.CBMP_ORACLE{r, 5});
                wantPxW  = testCase.CBMP_ORACLE{r, 6};
                wantPxH  = testCase.CBMP_ORACLE{r, 7};
                wantHDpi = testCase.CBMP_ORACLE{r, 8};
                wantVDpi = testCase.CBMP_ORACLE{r, 9};
                wantWEmu = testCase.CBMP_ORACLE{r, 10};
                wantHEmu = testCase.CBMP_ORACLE{r, 11};

                img = mat2doc.image.Image.from_blob(blob);
                testCase.verifyEqual(string(img.content_type), "image/bmp", sprintf('%s content_type', tag));
                testCase.verifyEqual(string(img.ext), "bmp", sprintf('%s ext', tag));
                testCase.verifyEqual(double(img.px_width),  wantPxW,  sprintf('%s px_width', tag));
                testCase.verifyEqual(double(img.px_height), wantPxH,  sprintf('%s px_height', tag));
                testCase.verifyEqual(double(img.horz_dpi),  wantHDpi, sprintf('%s horz_dpi (reverted)', tag));
                testCase.verifyEqual(double(img.vert_dpi),  wantVDpi, sprintf('%s vert_dpi (reverted)', tag));
                testCase.verifyEqual(double(img.width),  wantWEmu, sprintf('%s width_EMU (reverted dpi flows in)', tag));
                testCase.verifyEqual(double(img.height), wantHEmu, sprintf('%s height_EMU (reverted dpi flows in)', tag));
            end
        end

        % =============================================================== %
        % 4. Chunk-parse fidelity (big-endian walk + IHDR/pHYs dispatch)  %
        % =============================================================== %

        function test_chunk_parse_fidelity(testCase)
            % Regression: real PNGs parsed via ChunkParser_/Chunks_ -- the chunk
            % type sequence (proves the big-endian chunk-walk arithmetic), IHDR px
            % dims, and pHYs resolution all value-identical to the frozen oracle.
            % Plus the raw BMP/GIF header primitives.
            here = fileparts(mfilename('fullpath'));

            % -- 150-dpi.png: 6 chunks, pHYs 5906/5906/units=1 (=> 150 dpi) --
            blob150 = readBlobBin(fullfile(here, 'data', '150-dpi.png'));
            parser150 = mat2doc.image.ChunkParser_.from_stream(mat2doc.image.BytesIO(blob150));
            testCase.verifyEqual(chunkTypes(parser150), "IHDR|pHYs|iCCP|cHRM|IDAT|IEND", ...
                '150-dpi.png chunk type sequence');
            chunks150 = mat2doc.image.Chunks_.from_stream(mat2doc.image.BytesIO(blob150));
            testCase.verifyEqual(double(chunks150.IHDR.px_width),  901,  '150-dpi.png IHDR px_width');
            testCase.verifyEqual(double(chunks150.IHDR.px_height), 1350, '150-dpi.png IHDR px_height');
            testCase.verifyEqual(physStr(chunks150), "5906|5906|1", '150-dpi.png pHYs');

            % -- monty-truth.png: 4 chunks, no pHYs (=> None) --
            blobMonty = readBlobBin(fullfile(here, 'data', 'monty-truth.png'));
            parserMonty = mat2doc.image.ChunkParser_.from_stream(mat2doc.image.BytesIO(blobMonty));
            testCase.verifyEqual(chunkTypes(parserMonty), "IHDR|tEXt|IDAT|IEND", ...
                'monty-truth.png chunk type sequence');
            chunksMonty = mat2doc.image.Chunks_.from_stream(mat2doc.image.BytesIO(blobMonty));
            testCase.verifyEqual(double(chunksMonty.IHDR.px_width),  150, 'monty-truth.png IHDR px_width');
            testCase.verifyEqual(double(chunksMonty.IHDR.px_height), 214, 'monty-truth.png IHDR px_height');
            testCase.verifyEqual(physStr(chunksMonty), "none", 'monty-truth.png pHYs (absent -> none)');

            % -- BMP header primitive (little-endian BITMAPINFOHEADER) --
            bmpHdr = mat2doc.image.Bmp.from_stream( ...
                mat2doc.image.BytesIO(readBlobBin(fullfile(here, 'data', 'python.bmp'))));
            testCase.verifyEqual(double(bmpHdr.px_width),  211, 'python.bmp px_width');
            testCase.verifyEqual(double(bmpHdr.px_height),  71, 'python.bmp px_height');
            testCase.verifyEqual(double(bmpHdr.horz_dpi),   96, 'python.bmp horz_dpi');
            testCase.verifyEqual(double(bmpHdr.vert_dpi),   96, 'python.bmp vert_dpi');

            % -- GIF header primitive (little-endian "<HH" at offset 6) --
            gifStream = mat2doc.image.BytesIO(readBlobBin(fullfile(here, 'data', 'sonic.gif')));
            [gw, gh] = mat2doc.image.Gif.dimensions_from_stream_(gifStream);
            testCase.verifyEqual(double(gw), 290, 'sonic.gif px_width (little-endian)');
            testCase.verifyEqual(double(gh), 360, 'sonic.gif px_height (little-endian)');
        end

        % =============================================================== %
        % 5. Factory dispatch -> REAL parsers; stubs still notYetPorted   %
        % =============================================================== %

        function test_factory_dispatch_real_parsers(testCase)
            % Edge: the P7-1a factory now dispatches PNG/GIF/BMP signatures to the
            % REAL parsers (Image.from_blob returns a parsed Image with the right
            % px/dpi), while the Jfif/Exif/Tiff signatures STILL raise
            % mat2doc:notYetPorted (the P7-2 boundary).
            here = fileparts(mfilename('fullpath'));

            % PNG blob -> Png parses (300-dpi.png -> 860x579 @ 300).
            pngImg = mat2doc.image.Image.from_blob(readBlobBin(fullfile(here, 'data', '300-dpi.png')));
            testCase.verifyEqual(string(pngImg.content_type), "image/png", 'PNG dispatch -> real Png');
            testCase.verifyEqual(double(pngImg.px_width), 860, 'PNG dispatch parsed px_width');
            testCase.verifyEqual(double(pngImg.horz_dpi), 300, 'PNG dispatch parsed dpi');

            % GIF blob -> Gif parses (sonic.gif -> 290x360 @ 72).
            gifImg = mat2doc.image.Image.from_blob(readBlobBin(fullfile(here, 'data', 'sonic.gif')));
            testCase.verifyEqual(string(gifImg.content_type), "image/gif", 'GIF dispatch -> real Gif');
            testCase.verifyEqual(double(gifImg.px_width), 290, 'GIF dispatch parsed px_width');
            testCase.verifyEqual(double(gifImg.horz_dpi), 72, 'GIF dispatch dpi 72');

            % BMP blob -> Bmp parses (python.bmp -> 211x71 @ 96).
            bmpImg = mat2doc.image.Image.from_blob(readBlobBin(fullfile(here, 'data', 'python.bmp')));
            testCase.verifyEqual(string(bmpImg.content_type), "image/bmp", 'BMP dispatch -> real Bmp');
            testCase.verifyEqual(double(bmpImg.px_width), 211, 'BMP dispatch parsed px_width');
            testCase.verifyEqual(double(bmpImg.horz_dpi), 96, 'BMP dispatch dpi 96 (reverted)');

            % Jfif / Exif / Tiff signatures STILL notYetPorted (P7-2).
            jfif = uint8([255 216 255 224 0 16, double('JFIF'), 0, zeros(1, 21)]);
            exif = uint8([255 216 255 225 0 16, double('Exif'), 0, zeros(1, 21)]);
            tiffMM = uint8([77 77 0 42, zeros(1, 28)]);
            tiffII = uint8([73 73 42 0, zeros(1, 28)]);
            verifyRaises(testCase, @() mat2doc.image.Image.from_blob(jfif), ...
                'mat2doc:notYetPorted', 'JFIF still notYetPorted (P7-2)');
            verifyRaises(testCase, @() mat2doc.image.Image.from_blob(exif), ...
                'mat2doc:notYetPorted', 'Exif still notYetPorted (P7-2)');
            verifyRaises(testCase, @() mat2doc.image.Image.from_blob(tiffMM), ...
                'mat2doc:notYetPorted', 'TIFF(MM) still notYetPorted (P7-2)');
            verifyRaises(testCase, @() mat2doc.image.Image.from_blob(tiffII), ...
                'mat2doc:notYetPorted', 'TIFF(II) still notYetPorted (P7-2)');
        end

        % =============================================================== %
        % ★ 6. Full frozen-battery equivalence replay (self-contained)    %
        % =============================================================== %

        function test_replay_all_143_records(testCase)
            % ★ Equivalence (headline): replay the ENTIRE frozen Gate-3 battery
            % (the exact 143-probe sequence of s0084_..._probe.m, rebuilt inline
            % through mat2doc.image.*) and assert the port reproduces every record's
            % tagged canon value-identical to the frozen python-docx 1.2.0 oracle.
            % Fixture is references\s0084\probe.json copied verbatim into
            % tests\image\data\s0084_probe.json (self-contained). Same tagged-canon
            % scheme (OK|int|, OK|str|, ...) the Gate-3 comparator used.
            here = fileparts(mfilename('fullpath'));
            refPath = fullfile(here, 'data', 's0084_probe.json');
            testCase.assertTrue(isfile(refPath), ...
                sprintf('frozen probe fixture missing: %s', refPath));
            ref = jsondecode(readTextUtf8(refPath));

            obs = buildObservedBattery(here);

            keys = fieldnames(ref);
            testCase.verifyEqual(numel(keys), 143, ...
                'frozen s0084 battery must hold 143 records');
            testCase.verifyEqual(double(obs.Count), 143, ...   % Map.Count is uint64
                'rebuilt battery must reproduce exactly 143 records');
            for i = 1:numel(keys)
                k = keys{i};
                want = char(string(ref.(k)));
                testCase.assertTrue(isKey(obs, k), ...
                    sprintf('record not reproduced by port: %s', k));
                testCase.verifyEqual(obs(k), want, ...
                    sprintf('canon mismatch for %s (want %s, got %s)', k, want, obs(k)));
            end
        end

        % =============================================================== %
        % ★ 7. M1 byte-pin (P7-1b is M1-neutral)                          %
        % =============================================================== %

        function test_m1_neutral_styles_document_bytes(testCase)
            % ★ Regression: mat2doc.Document().save() emits word/styles.xml and
            % word/document.xml byte-identical to the frozen s0001 oracle (SHA-256
            % == byte-identity, L1). Image parsing touches no oxml registry row, no
            % PartFactory registration, nothing on the save path -- so M1 must stay
            % green. RED here would mean the P7-1b re-port perturbed the save path.
            d = mat2doc.Document();                     % no arg -> default template
            tmp = [tempname '.docx'];
            cleanup = onCleanup(@() deleteIfExists(tmp)); %#ok<NASGU>
            d.save(tmp);
            zipBytes = readBlobBin(tmp);
            [blobs, names] = zipEntryList(zipBytes);

            styles = entryBlobByName(blobs, names, "word/styles.xml");
            docx   = entryBlobByName(blobs, names, "word/document.xml");
            testCase.verifyEqual(sha256hex(styles), testCase.M1_STYLES_SHA, ...
                'word/styles.xml must be byte-identical to the frozen s0001 oracle (M1-neutral)');
            testCase.verifyEqual(sha256hex(docx), testCase.M1_DOCUMENT_SHA, ...
                'word/document.xml must be byte-identical to the frozen s0001 oracle (M1-neutral)');
        end

    end
end

% ===================== file-local helpers ============================== %

function b = readBlobBin(path)
    % Binary read (no CRLF translation): MATLAB 'r' is binary by default.
    fid = fopen(char(path), 'r', 'n');
    assert(fid >= 0, 'could not open for read: %s', char(path));
    c = onCleanup(@() fclose(fid)); %#ok<NASGU>
    b = fread(fid, inf, '*uint8')';
end

function txt = readTextUtf8(path)
    % Explicit UTF-8 text read (the frozen probe.json is UTF-8, ensure_ascii=False).
    fid = fopen(char(path), 'r', 'n', 'UTF-8');
    assert(fid >= 0, 'could not open for read: %s', char(path));
    c = onCleanup(@() fclose(fid)); %#ok<NASGU>
    txt = fread(fid, inf, '*char')';
end

function e = le32(x)
    % little-endian unsigned long -> 4 bytes (matches the s0084 scenario le32).
    e = uint8([mod(x, 256), mod(floor(x / 256), 256), ...
        mod(floor(x / 65536), 256), mod(floor(x / 16777216), 256)]);
end

function b = make_bmp(w, h, hppm, vppm)
    % 54-byte BMP: 'BM' + BITMAPINFOHEADER fields at the docx offsets
    % (px_width@0x12, px_height@0x16, horz_ppm@0x26, vert_ppm@0x2A), LE longs.
    % Byte-identical to the s0084 scenario's make_bmp.
    b = zeros(1, 54, 'uint8');
    b(1) = 66; b(2) = 77;          % 'B','M'
    b(19:22) = le32(w);            % 0x12 px_width
    b(23:26) = le32(h);            % 0x16 px_height
    b(39:42) = le32(hppm);         % 0x26 horz_px_per_meter
    b(43:46) = le32(vppm);         % 0x2A vert_px_per_meter
end

function s = chunkTypes(parser)
    chunks = parser.iter_chunks();
    parts = cellfun(@(c) string(c.type_name), chunks);
    s = strjoin(parts, "|");
end

function s = physStr(chunks)
    phys = chunks.pHYs;
    if isequal(phys, [])                        % None (H3)
        s = "none";
    else
        s = sprintf("%d|%d|%d", phys.horz_px_per_unit, ...
            phys.vert_px_per_unit, phys.units_specifier);
    end
end

function verifyRaises(testCase, fn, expId, label)
    % Assert fn() raises with the exact mat2doc: identifier.
    caught = [];
    try
        fn();
    catch ME
        caught = ME;
    end
    testCase.assertNotEmpty(caught, sprintf('%s: must raise', label));
    testCase.verifyEqual(caught.identifier, expId, ...
        sprintf('%s: expected %s, got %s', label, expId, caught.identifier));
end

% ------------------------------------------------------------ M1 zip helpers
function [blobs, names] = zipEntryList(zipBytes)
    % Enumerate an in-memory zip in stream (write) order -> ordered {names, blobs}.
    % (Copied from Test_p7_1a_image_core.m so the order pin is independent of the
    % reader under test.)
    bais = java.io.ByteArrayInputStream(int8(typecast(uint8(zipBytes(:)'), 'int8')));
    zis  = java.util.zip.ZipInputStream(bais);
    cleanup = onCleanup(@() zis.close()); %#ok<NASGU>
    copier = com.mathworks.mlwidgets.io.InterruptibleStreamCopier.getInterruptibleStreamCopier;
    names = strings(1, 0);
    blobs = {};
    while true
        ze = zis.getNextEntry();
        if isempty(ze)          % Java null -> no more entries
            break
        end
        names(end + 1) = string(ze.getName()); %#ok<AGROW>
        baos = java.io.ByteArrayOutputStream;
        copier.copyStream(zis, baos);
        blobs{end + 1} = typecast(int8(baos.toByteArray()), 'uint8')'; %#ok<AGROW>
        zis.closeEntry();
    end
end

function blob = entryBlobByName(blobs, names, membername)
    i = find(names == string(membername), 1);
    assert(~isempty(i), 'zip entry not found: %s', membername);
    blob = blobs{i};
end

function h = sha256hex(bytes)
    % Lowercase hex SHA-256 of a uint8 vector (base MATLAB / Java MessageDigest).
    md  = java.security.MessageDigest.getInstance('SHA-256');
    dig = md.digest(typecast(uint8(bytes(:)'), 'int8'));
    b   = typecast(dig, 'uint8');
    h   = lower(string(sprintf('%02x', b)));
end

function deleteIfExists(p)
    if isfile(p)
        delete(p);
    end
end

% ============================ frozen-battery replay ======================
% A faithful inline mirror of validation\mat2doc\scenarios\
% s0084_p7_1b_png_gif_bmp_probe.m (same probe order, same section map, same
% canon scheme), rebuilt through mat2doc.image.* into a containers.Map keyed by
% probe name -> tagged canon string. The 7 test images are read from
% tests\image\data\ (self-contained; the scenario read them from the READ-ONLY
% python-docx clone -- byte-identical, sha1-verified).

function obs = buildObservedBattery(here)
    obs = containers.Map('KeyType', 'char', 'ValueType', 'char');
    dataDir = fullfile(here, 'data');

    % ===================================================== IMG  end-to-end corpus
    corpus = {"PNG150", "150-dpi.png"; "PNG300", "300-dpi.png"; ...
        "PNGMONTY", "monty-truth.png"; "PNGICON", "python-icon.png"; ...
        "PNGPWR", "python-powered.png"; "BMP", "python.bmp"; "GIF", "sonic.gif"};
    for r = 1:size(corpus, 1)
        tag = corpus{r, 1};
        img = mat2doc.image.Image.from_file(fullfile(dataDir, char(corpus{r, 2})));
        put(obs, "IMG_" + tag + "_content_type", P(@() img.content_type));
        put(obs, "IMG_" + tag + "_px_width",     P(@() img.px_width));
        put(obs, "IMG_" + tag + "_px_height",    P(@() img.px_height));
        put(obs, "IMG_" + tag + "_horz_dpi",     P(@() img.horz_dpi));
        put(obs, "IMG_" + tag + "_vert_dpi",     P(@() img.vert_dpi));
        put(obs, "IMG_" + tag + "_width_emu",    P(@() img.width));
        put(obs, "IMG_" + tag + "_height_emu",   P(@() img.height));
        put(obs, "IMG_" + tag + "_ext",          P(@() img.ext));
        put(obs, "IMG_" + tag + "_filename",     P(@() img.filename));
        put(obs, "IMG_" + tag + "_sha1",         P(@() img.sha1));
    end

    % ===================================================== CBMP  crafted BMP e2e
    cbmp = {"bmp0", 40, 30, 0, 0; "bmp2500", 40, 30, 2500, 2500; ...
        "bmp3780", 40, 30, 3780, 3780; "bmp7500", 100, 50, 7500, 7500; ...
        "bmp17500", 100, 50, 17500, 17500};
    for r = 1:size(cbmp, 1)
        tag = cbmp{r, 1};
        blob = make_bmp(cbmp{r, 2}, cbmp{r, 3}, cbmp{r, 4}, cbmp{r, 5});
        img = mat2doc.image.Image.from_blob(blob);
        put(obs, "CBMP_" + tag + "_px_width",     P(@() img.px_width));
        put(obs, "CBMP_" + tag + "_px_height",    P(@() img.px_height));
        put(obs, "CBMP_" + tag + "_horz_dpi",     P(@() img.horz_dpi));
        put(obs, "CBMP_" + tag + "_vert_dpi",     P(@() img.vert_dpi));
        put(obs, "CBMP_" + tag + "_width_emu",    P(@() img.width));
        put(obs, "CBMP_" + tag + "_height_emu",   P(@() img.height));
        put(obs, "CBMP_" + tag + "_content_type", P(@() img.content_type));
        put(obs, "CBMP_" + tag + "_ext",          P(@() img.ext));
    end

    % ===================================================== BDPI  Bmp.dpi_ guard
    bdpi = [0 1 2500 3780 7500 11811 12500 17500 22500 27500];
    for i = 1:numel(bdpi)
        ppm = bdpi(i);
        put(obs, sprintf("BDPI_%d", ppm), P(@() mat2doc.image.Bmp.dpi_(ppm)));
    end

    % ===================================================== PDPI  PngParser_.dpi_
    pdpi = [1 2500; 1 7500; 1 12500; 1 17500; 1 11811; 0 2835; 1 0; 2 2835; 3 2835];
    for i = 1:size(pdpi, 1)
        u = pdpi(i, 1); ppu = pdpi(i, 2);
        put(obs, sprintf("PDPI_u%d_p%d", u, ppu), P(@() mat2doc.image.PngParser_.dpi_(u, ppu)));
    end

    % ===================================================== CHUNK  parse fidelity
    chk = {"PNG150", "150-dpi.png"; "PNGMONTY", "monty-truth.png"};
    for r = 1:size(chk, 1)
        tag = chk{r, 1};
        blob = readBlobBin(fullfile(dataDir, char(chk{r, 2})));
        parser = mat2doc.image.ChunkParser_.from_stream(mat2doc.image.BytesIO(blob));
        put(obs, "CHUNK_" + tag + "_types", P(@() chunkTypes(parser)));
        chunks = mat2doc.image.Chunks_.from_stream(mat2doc.image.BytesIO(blob));
        put(obs, "CHUNK_" + tag + "_ihdr_w", P(@() ihdrW(chunks)));
        put(obs, "CHUNK_" + tag + "_ihdr_h", P(@() ihdrH(chunks)));
        put(obs, "CHUNK_" + tag + "_phys",   P(@() physStr(chunks)));
    end

    % ===================================================== HDR  BMP/GIF header parse
    bmpHdr = mat2doc.image.Bmp.from_stream( ...
        mat2doc.image.BytesIO(readBlobBin(fullfile(dataDir, 'python.bmp'))));
    put(obs, "HDR_bmp_px_w", P(@() bmpHdr.px_width));
    put(obs, "HDR_bmp_px_h", P(@() bmpHdr.px_height));
    put(obs, "HDR_bmp_hdpi", P(@() bmpHdr.horz_dpi));
    put(obs, "HDR_bmp_vdpi", P(@() bmpHdr.vert_dpi));

    gifStream = mat2doc.image.BytesIO(readBlobBin(fullfile(dataDir, 'sonic.gif')));
    [gw, gh] = mat2doc.image.Gif.dimensions_from_stream_(gifStream);
    put(obs, "HDR_gif_px_w", P(@() gw));
    put(obs, "HDR_gif_px_h", P(@() gh));
end

function put(map, key, canonStr)
    map(char(key)) = char(canonStr);
end

function v = ihdrW(chunks), ihdr = chunks.IHDR; v = ihdr.px_width;  end
function v = ihdrH(chunks), ihdr = chunks.IHDR; v = ihdr.px_height; end

% ------------------------------------------------------------ probe wrappers
% (verbatim from the Gate-3 scenario s0084_..._probe.m so the canon is identical)

function s = P(fn)
% Full canon; on error "err|<Name>|<message>".
try
    s = "OK|" + canon(fn());
catch e
    s = "err|" + errName(e) + "|" + string(e.message);
end
end

function nm = errName(e)
id = string(e.identifier);
if startsWith(id, "mat2doc:")
    nm = extractAfter(id, "mat2doc:");
else
    nm = id;
end
end

function s = canon(v)
if isa(v, "uint8")
    s = "hex|" + string(upperHex(v));
elseif islogical(v)
    if v, s = "bool|True"; else, s = "bool|False"; end
elseif isstring(v) || ischar(v)
    s = "str|" + string(v);
elseif isnumeric(v)
    dv = double(v);
    if mod(dv, 1) == 0
        s = "int|" + string(sprintf("%.0f", dv));
    else
        s = "float|" + string(num2str(dv, 17));
    end
else
    s = "obj|" + string(class(v));
end
end

function h = upperHex(v)
if isempty(v)
    h = "";
else
    h = sprintf("%02X", v);   % uppercase, matches python bytes.hex().upper()
end
end

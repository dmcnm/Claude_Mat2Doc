classdef Test_p7_2a_tiff < matlab.unittest.TestCase
% TEST_P7_2A_TIFF  Gate-4 permanent unit tests for Mat2Doc P7-2a
%   (the TIFF image-header parser, WITH the PIL->docx dpi reversion + the
%   D-tiff-den0 error-path re-litigation).
%
%   Surface under test (ported from python-docx v1.2.0 src/docx/image/tiff.py;
%   re-ported off the Mat2Ppt +image TIFF/IFD classes with the PIL contract
%   STRIPPED -- no int_dpi [1,2048] clamp, no PIL joint-axes guard, no CLASS-E
%   exif accessor):
%     - mat2doc.image.Tiff            (tiff.py Tiff)            -- un-stubs P7-1a.
%     - mat2doc.image.TiffParser_     (_TiffParser)            -- the _dpi formula.
%     - mat2doc.image.IfdEntries_     (_IfdEntries)            -- {tag:value}, H11.
%     - mat2doc.image.IfdParser_      (_IfdParser)             -- IFD walk.
%     - mat2doc.image.IfdEntryFactory_(_IfdEntryFactory)       -- H10 dispatch.
%     - mat2doc.image.IfdEntry_       (_IfdEntry base/default) -- UNIMPLEMENTED.
%     - mat2doc.image.{Ascii,Short,Long,Rational}IfdEntry_     -- typed values.
%   These un-stub the P7-1a notYetPorted Tiff row; the P7-1a factory
%   (ImageHeaderFactory_) already resolved @mat2doc.image.Tiff.from_stream, so the
%   un-stub alone flips MM/II dispatch from a stub to the real parser. Jfif/Exif
%   remain notYetPorted (P7-2b jpeg).
%
%   Provenance (Gate-1..3, all 2026-08-01):
%     * Audit    : validation\mat2doc\audit_P7-2a_tiff.md
%                  (Porter Gate-1 + Fable Gate-2 REVISE->F-1 fixed->re-verified
%                   14/14; dpi reversion + den0 error-path match PROVEN).
%     * Validate : validation\mat2doc\validate_P7-2a_tiff.md
%                  (Gate-3 PASS, 191/191 value-identical vs python-docx 1.2.0,
%                   ZERO new D-numbers; pure-parsing WP, no package bytes; 16/16
%                   crafted blobs byte-identical between the two crafters).
%     * Scenario : validation\mat2doc\scenarios\s0085_p7_2a_tiff_probe.{py,m}
%     * Frozen ref (python-docx 1.2.0 oracle, frozen ONCE):
%         references\s0085\probe.json  -- the 191-value oracle, copied verbatim
%           into tests\image\data\s0085_probe.json (co-located
%           `s0085_probe.json text eol=lf`) so this suite is self-contained.
%         references\s0085\blobs\*.tif -- the 16 validator-crafted TIFF blobs
%           (the dpi-branch + den0 equivalence corpus), copied verbatim into
%           tests\image\data\s0085_*.tif (co-located `* binary` .gitattributes;
%           each blob's sha1 re-verified byte-identical to the frozen reference).
%         The 2 real docx TIFFs (72-dpi.tiff, little-endian.tif) copied into
%           tests\image\data\ (binary-pinned; both sha1 re-verified vs the clone).
%     * M1 byte-pin values reuse the frozen s0001 manifest (word/styles.xml
%       02d71a68..., word/document.xml 0e4dd503...) already owned by
%       Test_p1_8_skeleton_m1; TIFF parsing is M1-neutral (touches no oxml
%       registry row, no PartFactory registration, nothing on the save path).
%
%   Coverage taxonomy
%   -----------------
%   * NOMINAL / Image equivalence (headline) -- test_image_equivalence_corpus:
%     Image.from_file on the 2 REAL TIFFs (both endians) + Image.from_blob on the
%     10 crafted resolution-branch TIFFs; sha1/content_type/ext/px/dpi/EMU all
%     hard-pinned to the frozen s0085 oracle. sha1 is END-TO-END on the real blob.
%   * ★★ dpi-REVERSION guards (the docx-not-PIL guards -- LOUD):
%       - test_int_dpi_clamp_reversal   -- big_ii dpi 3000/5000 UNCAPPED (NOT
%         PIL's 2048). A regression to the PIL [1,2048] clamp fails HERE.
%       - test_half_to_even_ties        -- 190.5->190, 192.5->192, 254.5->254
%         (floor-EVEN DOWN) + 63.5->64. Native MATLAB round() gives 191/193/255;
%         these catch a silent revert to round().
%       - test_f1_multivalue_resolution_unit -- RESOLUTION_UNIT count=2 -> 254
%         (docx RETURNS, does NOT throw). The Gate-2 F-1 regression guard.
%       - test_resolution_unit_branches -- the 4 unit branches: inch(2)->x1,
%         cm(3)->x2.54, aspect(1)->72, tag-absent->72, unit-absent->default 2.
%   * ★ D-tiff-den0 (error path) -- test_den0_zero_division: a 0-denominator
%     rational (both endians, from_blob AND direct parse) -> mat2doc:ZeroDivisionError
%     "division by zero" (verbatim message), firing at PARSE time. Error-path
%     match to docx's ZeroDivisionError; NO new D-number.
%   * IFD / endian fidelity -- test_ifd_endian_fidelity (types_ii/types_mm ->
%     px 120x90, dpi 300/72 both endians) + test_entry_type_dispatch (the H10
%     factory: BYTE->base IfdEntry_, SHORT/LONG/RATIONAL/ASCII->typed subclasses,
%     value-exact both endians) + test_pxnone_none_dims (H3: absent WIDTH/LENGTH
%     tags -> None ([]) while dpi still resolves).
%   * factory dispatch (edge) -- test_factory_dispatch_and_stubs: a TIFF MM/II
%     blob dispatches through the P7-1a factory to the REAL Tiff parser; the
%     Jfif/Exif rows RE-PINNED at P7-2b (real parsers now -> mat2doc:Exception
%     EOF on the SOS-less signature blob, no longer notYetPorted).
%   * ★ Equivalence (headline replay) -- test_replay_all_191_records: replays the
%     ENTIRE frozen Gate-3 battery (the exact 191-probe sequence of
%     s0085_..._probe.m, rebuilt inline through mat2doc.image.* from the
%     self-contained data\ copies) and asserts the port reproduces every record's
%     tagged canon value-identical to the frozen python-docx 1.2.0 oracle.
%   * ★ M1 byte-pin (regression) -- test_m1_neutral_styles_document_bytes:
%     mat2doc.Document().save() -> word/styles.xml sha256 02d71a68... +
%     word/document.xml sha256 0e4dd503... (SHA-256 == byte-identity, L1). RED
%     here would mean the TIFF re-port perturbed the save path.
%
%   Deviations exercised (adopt-only, ZERO new D-numbers -- Gate-3): the dpi is the
%   sole value-deviation risk and it is the python-docx formula EXACTLY (the
%   Mat2Ppt PIL int_dpi/joint-axes/CLASS-E contract is REVERTED); den0 is an
%   error-path match to docx's ZeroDivisionError (no output either side, nothing to
%   deviate). No serialization occurs in the parser, so no L1 byte surface is
%   touched (equivalence is value-level; validate section 7).
%
%   Determinism: no network, no hard-coded absolute paths -- the frozen oracle,
%   the 16 crafted blobs, and the 2 real TIFFs resolve relative to this file via
%   fileparts(mfilename('fullpath')). Every image read is BINARY (no CRLF xlate).

    properties (Constant)
        % ★ Image-equivalence oracle (references\s0085\probe.json IMG_* rows).
        % Cols: tag, fixture-file, isReal (from_file vs from_blob), content_type,
        % ext, filename (""=from_blob, not probed), px_w, px_h, hdpi, vdpi,
        % width_EMU, height_EMU, sha1 (END-TO-END on the real blob).
        IMG_ORACLE = { ...
            "SHIP72",   "72-dpi.tiff",        true,  "image/tiff","tiff","72-dpi.tiff",       48,   48,   72,   72,  609600,  609600, "ced06dc01f74e4e1df4ddc4e8d126d22391b8405"; ...
            "SHIPLE",   "little-endian.tif",  true,  "image/tiff","tif", "little-endian.tif", 1600, 2100, 200,  200, 7315200, 9601200, "b3ea598919f336c7b316adab825bbe03e413dd47"; ...
            "inch_ii",  "s0085_inch_ii.tif",  false, "image/tiff","tiff","",                  120,  90,   200,  200, 548640,  411480, "6502362416ceea3e09f7ee99051f6ff4330cb883"; ...
            "inch_mm",  "s0085_inch_mm.tif",  false, "image/tiff","tiff","",                  120,  90,   220,  170, 498763,  484094, "140ebe520ae015f9b68965a3649de0ab44781fe2"; ...
            "cm_ii",    "s0085_cm_ii.tif",    false, "image/tiff","tiff","",                  120,  90,   254,  254, 432000,  324000, "631d54485b2cbb1ce877cb0a144e1eb9d07edcb5"; ...
            "aspect_ii","s0085_aspect_ii.tif",false, "image/tiff","tiff","",                  120,  90,   72,   72,  1524000, 1143000,"bbdb62c15d959fc38ddce6731d4a57ea5eda4647"; ...
            "nounit_ii","s0085_nounit_ii.tif",false, "image/tiff","tiff","",                  120,  90,   150,  150, 731520,  548640, "d9a423d3eb2160eec581d505eaeba4006cbf77db"; ...
            "nores_ii", "s0085_nores_ii.tif", false, "image/tiff","tiff","",                  120,  90,   72,   72,  1524000, 1143000,"85b19ea66a21f248fd19666a89251131b6145ae7"; ...
            "big_ii",   "s0085_big_ii.tif",   false, "image/tiff","tiff","",                  120,  90,   3000, 5000,36576,   16459,  "b4460d7c572519294046d8b7260ed17fa0d151fb"; ...
            "tie_ii",   "s0085_tie_ii.tif",   false, "image/tiff","tiff","",                  120,  90,   64,   190, 1714500, 433136, "1181fcb287aea0f9515764593736f7e70261aea2"; ...
            "tie2_ii",  "s0085_tie2_ii.tif",  false, "image/tiff","tiff","",                  120,  90,   192,  254, 571500,  324000, "94f163e3f7fe5a1f1794b860cb079a79ee643622"; ...
            "tie_cm_ii","s0085_tie_cm_ii.tif",false, "image/tiff","tiff","",                  120,  90,   64,   190, 1714500, 433136, "5c643c13f9ddabf233cd7db53605f6c2a0217317"};

        % ★★ dpi-branch discriminator oracle (probe.json DPI_* rows). Cols:
        % tag, fixture-file, want_h, want_v.  The rows carry the four unit
        % branches AND the reversion discriminators (big_ii uncapped; the ties).
        DPI_ORACLE = { ...
            "inch_ii",  "s0085_inch_ii.tif",  200,  200; ...   % unit 2 (inch) x1
            "inch_mm",  "s0085_inch_mm.tif",  220,  170; ...   % MM, asymmetric (X/Y not swapped)
            "cm_ii",    "s0085_cm_ii.tif",    254,  254; ...   % unit 3 (cm)  100*2.54
            "aspect_ii","s0085_aspect_ii.tif",72,   72;  ...   % unit 1 (aspect ratio only) -> 72
            "nounit_ii","s0085_nounit_ii.tif",150,  150; ...   % RESOLUTION_UNIT absent -> default 2
            "nores_ii", "s0085_nores_ii.tif", 72,   72;  ...   % resolution tag absent -> 72
            "big_ii",   "s0085_big_ii.tif",   3000, 5000;...   % > 2048 UNCAPPED (int_dpi-clamp reversal)
            "tie_ii",   "s0085_tie_ii.tif",   64,   190; ...   % 63.5->64, 190.5->190 (half-to-even)
            "tie2_ii",  "s0085_tie2_ii.tif",  192,  254; ...   % 192.5->192, 254.5->254 (half-to-even DOWN)
            "tie_cm_ii","s0085_tie_cm_ii.tif",64,   190; ...   % cm-factor ties
            "types_ii", "s0085_types_ii.tif", 300,  72};       % X only -> 300; Y absent -> 72

        % Entry-type dispatch oracle (probe.json DISP_* rows). Cols: idx, tag,
        % dispatch class (normalized), value canon ("int|N" or "str|...").
        DISP_ORACLE = { ...
            0, 254, "IfdEntry",         "str|UNIMPLEMENTED FIELD TYPE"; ...
            1, 256, "ShortIfdEntry",    "int|120"; ...
            2, 257, "LongIfdEntry",     "int|90";  ...
            3, 282, "RationalIfdEntry", "int|300"; ...
            4, 296, "ShortIfdEntry",    "int|2";   ...
            5, 305, "AsciiIfdEntry",    "str|docx"};

        % ★ M1 byte-pin (== Test_p1_8_skeleton_m1 rows; python-docx 1.2.0 s0001
        % oracle). SHA-256 equality IS byte-identity (L1). P7-2a is M1-neutral.
        M1_STYLES_SHA   = "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"
        M1_DOCUMENT_SHA = "0e4dd503bc095219cd24e8312b372c588ca56bb834aacac53eafb7bf47836327"
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied from the proven sibling tests\image\Test_p7_1b_png_gif_bmp.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\image
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. NOMINAL: image equivalence -- real + crafted corpus (headline)%
        % =============================================================== %

        function test_image_equivalence_corpus(testCase)
            % ★ Regression/Nominal: Image.from_file on the 2 REAL docx TIFFs (both
            % endians) and Image.from_blob on the 10 crafted resolution-branch
            % TIFFs (all read from the self-contained tests\image\data\ copies)
            % reproduce every emitted field value-identical to the frozen s0085
            % oracle. px/dpi/content_type are parsed END-TO-END by the newly-live
            % TIFF parser; sha1 hashes the real blob; the EMU values run the
            % genuine Inches(px/dpi) Length math.
            here = fileparts(mfilename('fullpath'));
            for r = 1:size(testCase.IMG_ORACLE, 1)
                tag      = testCase.IMG_ORACLE{r, 1};
                fn       = testCase.IMG_ORACLE{r, 2};
                isReal   = testCase.IMG_ORACLE{r, 3};
                wantCT   = testCase.IMG_ORACLE{r, 4};
                wantExt  = testCase.IMG_ORACLE{r, 5};
                wantFn   = testCase.IMG_ORACLE{r, 6};
                wantPxW  = testCase.IMG_ORACLE{r, 7};
                wantPxH  = testCase.IMG_ORACLE{r, 8};
                wantHDpi = testCase.IMG_ORACLE{r, 9};
                wantVDpi = testCase.IMG_ORACLE{r, 10};
                wantWEmu = testCase.IMG_ORACLE{r, 11};
                wantHEmu = testCase.IMG_ORACLE{r, 12};
                wantSha1 = testCase.IMG_ORACLE{r, 13};

                path = fullfile(here, 'data', char(fn));
                if isReal
                    img = mat2doc.image.Image.from_file(path);       % filename-derived ext
                else
                    img = mat2doc.image.Image.from_blob(readBlobBin(path));
                end

                testCase.verifyEqual(string(img.content_type), string(wantCT), sprintf('%s content_type', tag));
                testCase.verifyEqual(string(img.ext),          string(wantExt), sprintf('%s ext', tag));
                testCase.verifyEqual(double(img.px_width),  wantPxW,  sprintf('%s px_width', tag));
                testCase.verifyEqual(double(img.px_height), wantPxH,  sprintf('%s px_height', tag));
                testCase.verifyEqual(double(img.horz_dpi),  wantHDpi, sprintf('%s horz_dpi', tag));
                testCase.verifyEqual(double(img.vert_dpi),  wantVDpi, sprintf('%s vert_dpi', tag));
                testCase.verifyEqual(double(img.width),  wantWEmu, sprintf('%s width EMU', tag));
                testCase.verifyEqual(double(img.height), wantHEmu, sprintf('%s height EMU', tag));
                % ★ sha1 END-TO-END (hashes the real blob).
                testCase.verifyEqual(string(img.sha1), string(wantSha1), ...
                    sprintf('%s sha1 END-TO-END must equal the frozen oracle', tag));
                if isReal
                    testCase.verifyEqual(string(img.filename), string(wantFn), ...
                        sprintf('%s filename', tag));
                end
            end
        end

        % =============================================================== %
        % ★★ 2. dpi-REVERSION guard: int_dpi-clamp REVERSAL (>2048 uncapped)%
        % =============================================================== %

        function test_int_dpi_clamp_reversal(testCase)
            % ★★★ LOUD REVERSION GUARD -- the docx-not-PIL guard.
            % big_ii carries X_RESOLUTION = 3000/1 and Y_RESOLUTION = 5000/1 with
            % RESOLUTION_UNIT = 2 (inch). docx's _dpi returns int(round(3000*1)) =
            % 3000 and 5000 -- UNCAPPED. The Mat2Ppt PIL mirror carried an int_dpi
            % [1,2048] clamp (a pptx parts/image seam); docx tiff.py HAS NO SUCH
            % CLAMP, and it was STRIPPED in the re-port.
            %
            %   ==> docx has no int_dpi clamp; a regression to PIL's [1,2048]
            %   ==> clamp FAILS HERE (both 3000 and 5000 would collapse to 2048).
            %   ==> TWO independent values (3000 and 5000) both pass uncapped. <==
            %
            % The uncapped dpi threads through Inches(px/dpi): width_EMU 36576 /
            % height_EMU 16459 (a 2048 clamp would give ~53578 / ~27432). (audit
            % section 3 / validate section 3a, frozen s0085 big_ii row.)
            here = fileparts(mfilename('fullpath'));
            blob = readBlobBin(fullfile(here, 'data', 's0085_big_ii.tif'));

            parser = mat2doc.image.TiffParser_.parse(mat2doc.image.BytesIO(blob));
            testCase.verifyEqual(double(parser.horz_dpi), 3000, ...
                'big_ii horz_dpi must be 3000 UNCAPPED (docx has no int_dpi [1,2048] clamp; a 2048 here = PIL regression)');
            testCase.verifyEqual(double(parser.vert_dpi), 5000, ...
                'big_ii vert_dpi must be 5000 UNCAPPED (a 2048 here = PIL int_dpi-clamp regression)');

            img = mat2doc.image.Image.from_blob(blob);
            testCase.verifyEqual(double(img.horz_dpi), 3000, 'big_ii Image.horz_dpi uncapped');
            testCase.verifyEqual(double(img.vert_dpi), 5000, 'big_ii Image.vert_dpi uncapped');
            testCase.verifyEqual(double(img.width),  36576, 'big_ii width_EMU (uncapped dpi threads through Inches(px/dpi))');
            testCase.verifyEqual(double(img.height), 16459, 'big_ii height_EMU (uncapped dpi threads through)');
        end

        % =============================================================== %
        % ★★ 3. dpi-REVERSION guard: half-to-even ties (pyRound not round)  %
        % =============================================================== %

        function test_half_to_even_ties(testCase)
            % ★★★ LOUD REVERSION GUARD. docx computes int(round(dots*upi)) with
            % CPython round-HALF-to-EVEN (pyRound), NOT MATLAB's half-away round().
            % The decisive floor-EVEN DOWN ties (190.5, 192.5, 254.5) round DOWN;
            % native MATLAB round() would round them UP to 191/193/255.
            %
            %   ==> THESE CATCH A SILENT REVERT TO NATIVE round(). <==
            %
            %   * tie_ii   : X 127/2  = 63.5  -> 64  (floor 63 ODD; both agree -- sanity)
            %                Y 381/2  = 190.5 -> 190 (floor 190 EVEN; native round()->191)
            %   * tie2_ii  : X 385/2  = 192.5 -> 192 (floor 192 EVEN; native round()->193)
            %                Y 509/2  = 254.5 -> 254 (floor 254 EVEN; native round()->255)
            %   * tie_cm_ii: X 25*2.54= 63.5  -> 64  ; Y 75*2.54=190.5 -> 190 (cm-factor ties)
            % (audit section 3 / validate section 3b, frozen s0085 tie rows.)
            ties = { ...
                's0085_tie_ii.tif',    64,  190; ...
                's0085_tie2_ii.tif',   192, 254; ...
                's0085_tie_cm_ii.tif', 64,  190};
            here = fileparts(mfilename('fullpath'));
            for r = 1:size(ties, 1)
                fn = ties{r, 1}; wantH = ties{r, 2}; wantV = ties{r, 3};
                parser = mat2doc.image.TiffParser_.parse( ...
                    mat2doc.image.BytesIO(readBlobBin(fullfile(here, 'data', fn))));
                testCase.verifyEqual(double(parser.horz_dpi), wantH, sprintf( ...
                    '%s horz_dpi must be %d (CPython half-to-even; a %d here = silent revert to MATLAB round())', ...
                    fn, wantH, wantH + 1));
                testCase.verifyEqual(double(parser.vert_dpi), wantV, sprintf( ...
                    '%s vert_dpi must be %d (floor-EVEN half-to-even DOWN; a %d here = native round() regression)', ...
                    fn, wantV, wantV + 1));
            end
        end

        % =============================================================== %
        % ★★ 4. dpi-REVERSION guard: F-1 multi-value RESOLUTION_UNIT        %
        % =============================================================== %

        function test_f1_multivalue_resolution_unit(testCase)
            % ★★★ THE GATE-2 F-1 REGRESSION GUARD. A TIFF whose RESOLUTION_UNIT
            % IFD entry has count=2 makes ifd_entries.get(RESOLUTION_UNIT, 2) the
            % multi-value SHORT placeholder (non-scalar/non-numeric). In Python a
            % cross-type `placeholder == 1` / `== 2` yields False SILENTLY, so docx
            % falls through to units_per_inch = 2.54 and RETURNS a dpi: multiunit_ii
            % (X/Y_RES = 100/1) -> 100*2.54 = 254 / 254 (docx does NOT throw).
            %
            %   ==> The Gate-2 port THREW MATLAB:string:ComparisonNotDefined here;
            %   ==> the F-1 fix guards each `==` with isnumeric && isscalar so the
            %   ==> comparison evaluates as False (like Python) instead of throwing.
            %   ==> A regression that removes the guard THROWS instead of RETURNING
            %   ==> 254 -- and this test's from_blob would error. <==
            % (audit section 5/13 F-1 / validate section 3c, frozen s0085 MUNIT rows.)
            here = fileparts(mfilename('fullpath'));
            img = mat2doc.image.Image.from_blob( ...
                readBlobBin(fullfile(here, 'data', 's0085_multiunit_ii.tif')));
            testCase.verifyEqual(double(img.horz_dpi), 254, ...
                'multiunit_ii horz_dpi must RETURN 254 (docx silent-False cross-type ==; F-1 guard, not a throw)');
            testCase.verifyEqual(double(img.vert_dpi), 254, ...
                'multiunit_ii vert_dpi must RETURN 254 (multi-value RESOLUTION_UNIT falls through to x2.54)');
            testCase.verifyEqual(double(img.px_width),  120, 'multiunit_ii px_width');
            testCase.verifyEqual(double(img.px_height), 90,  'multiunit_ii px_height');
        end

        % =============================================================== %
        % 5. dpi resolution-unit branches (the 4 unit arms)               %
        % =============================================================== %

        function test_resolution_unit_branches(testCase)
            % Nominal/edge: the four docx _dpi RESOLUTION_UNIT branches, each
            % proven value-identical to the frozen oracle:
            %   * unit 2 (inch)      -> units_per_inch x1  (inch_ii 200/200)
            %   * unit 3 (cm)        -> units_per_inch x2.54 (cm_ii 100*2.54=254)
            %   * unit 1 (aspect)    -> 72 (aspect ratio only; aspect_ii 72/72)
            %   * RES_UNIT absent    -> default 2 (inch) (nounit_ii 150/150)
            %   * resolution tag absent -> 72 (nores_ii 72/72)
            % Plus the MM asymmetric case (inch_mm 220/170) proving X/Y not swapped.
            here = fileparts(mfilename('fullpath'));
            for r = 1:size(testCase.DPI_ORACLE, 1)
                fn    = testCase.DPI_ORACLE{r, 2};
                wantH = testCase.DPI_ORACLE{r, 3};
                wantV = testCase.DPI_ORACLE{r, 4};
                parser = mat2doc.image.TiffParser_.parse( ...
                    mat2doc.image.BytesIO(readBlobBin(fullfile(here, 'data', char(fn)))));
                testCase.verifyEqual(double(parser.horz_dpi), wantH, sprintf('%s horz_dpi', fn));
                testCase.verifyEqual(double(parser.vert_dpi), wantV, sprintf('%s vert_dpi', fn));
            end
        end

        % =============================================================== %
        % ★ 6. D-tiff-den0 -- ZeroDivisionError error-path (NO new D)      %
        % =============================================================== %

        function test_den0_zero_division(testCase)
            % ★ Error path (D-tiff-den0 re-litigation): a TIFF whose X_RESOLUTION
            % rational has denominator == 0 (both endians). docx's
            % _RationalIfdEntry._parse_value does `numerator / denominator` -> a
            % Python ZeroDivisionError("division by zero") that propagates out of
            % IfdEntries_.from_stream at PARSE time (before any dpi is computed).
            % MATLAB's n/0 silently yields Inf, so the port guards at the division
            % site: error("mat2doc:ZeroDivisionError","%s","division by zero").
            %
            % This is an ERROR-PATH MATCH (no output on this path in either
            % implementation) -> NO new D-number. The pptx D-tiff-den0 ledger row
            % does NOT transfer (its int_dpi-clamp rationale is absent in docx).
            % Pinned: the exact identifier + the VERBATIM CPython message + that it
            % fires at BOTH from_blob (end-to-end) AND direct parse (parse time).
            % (audit section 4 / validate section 3d, frozen s0085 DEN0 rows.)
            here = fileparts(mfilename('fullpath'));
            files = {'s0085_den0_ii.tif', 's0085_den0_mm.tif'};
            for i = 1:numel(files)
                fn = files{i};
                blob = readBlobBin(fullfile(here, 'data', fn));

                % (a) end-to-end via Image.from_blob
                [id1, msg1] = catchErr(@() mat2doc.image.Image.from_blob(blob));
                testCase.verifyEqual(id1, 'mat2doc:ZeroDivisionError', ...
                    sprintf('%s from_blob must raise mat2doc:ZeroDivisionError', fn));
                testCase.verifyEqual(msg1, 'division by zero', ...
                    sprintf('%s from_blob message must be the verbatim CPython "division by zero"', fn));

                % (b) direct parse (proves it fires at PARSE / IFD-construction time)
                [id2, msg2] = catchErr(@() mat2doc.image.TiffParser_.parse( ...
                    mat2doc.image.BytesIO(blob)));
                testCase.verifyEqual(id2, 'mat2doc:ZeroDivisionError', ...
                    sprintf('%s parse must raise mat2doc:ZeroDivisionError at parse time', fn));
                testCase.verifyEqual(msg2, 'division by zero', ...
                    sprintf('%s parse message must be verbatim "division by zero"', fn));
            end
        end

        % =============================================================== %
        % 7. IFD / endian fidelity (H2) -- both endians end-to-end        %
        % =============================================================== %

        function test_ifd_endian_fidelity(testCase)
            % Regression: types_ii (little-endian) and types_mm (big-endian) both
            % parse end-to-end to px 120x90, horz_dpi 300 (X_RESOLUTION only),
            % vert_dpi 72 (Y_RESOLUTION absent -> default), content_type image/tiff.
            % Proves the endian dispatch + the byte-address IFD arithmetic (H1/H2)
            % on both orders. (validate section 4b, frozen s0085 IFD rows.)
            here = fileparts(mfilename('fullpath'));
            for fn = ["s0085_types_ii.tif", "s0085_types_mm.tif"]
                img = mat2doc.image.Image.from_blob( ...
                    readBlobBin(fullfile(here, 'data', char(fn))));
                testCase.verifyEqual(double(img.px_width),  120, sprintf('%s px_width', fn));
                testCase.verifyEqual(double(img.px_height), 90,  sprintf('%s px_height', fn));
                testCase.verifyEqual(double(img.horz_dpi),  300, sprintf('%s horz_dpi (X only)', fn));
                testCase.verifyEqual(double(img.vert_dpi),  72,  sprintf('%s vert_dpi (Y absent -> default)', fn));
                testCase.verifyEqual(string(img.content_type), "image/tiff", sprintf('%s content_type', fn));
            end
        end

        % =============================================================== %
        % 8. Entry-type dispatch (H10) -- factory + typed values          %
        % =============================================================== %

        function test_entry_type_dispatch(testCase)
            % Regression: the crafted `types` TIFF exercises every field type. The
            % IfdEntryFactory_ dispatch class AND the parsed value are value-
            % identical both endians:
            %   BYTE -> base IfdEntry_ ("UNIMPLEMENTED FIELD TYPE", the dict
            %           .get(field_type,_IfdEntry) `otherwise` default);
            %   SHORT -> ShortIfdEntry_; LONG -> LongIfdEntry_;
            %   RATIONAL -> RationalIfdEntry_; ASCII -> AsciiIfdEntry_
            %           (read_str(value_count-1, value_offset), Python's own NUL
            %            arithmetic, not an H1 shift).
            % entry_count read = 6 both endians. (validate section 4a.)
            here = fileparts(mfilename('fullpath'));
            variants = { ...
                's0085_types_ii.tif', mat2doc.image.StreamReader.LITTLE_ENDIAN; ...
                's0085_types_mm.tif', mat2doc.image.StreamReader.BIG_ENDIAN};
            for v = 1:size(variants, 1)
                fn     = variants{v, 1};
                endian = variants{v, 2};
                blob   = readBlobBin(fullfile(here, 'data', fn));
                rdr    = mat2doc.image.StreamReader(mat2doc.image.BytesIO(blob), endian);

                n = rdr.read_short(8);                 % entry count at IFD offset 8
                testCase.verifyEqual(double(n), 6, sprintf('%s entry_count', fn));

                for r = 1:size(testCase.DISP_ORACLE, 1)
                    idx      = testCase.DISP_ORACLE{r, 1};
                    wantTag  = testCase.DISP_ORACLE{r, 2};
                    wantCls  = testCase.DISP_ORACLE{r, 3};
                    wantVal  = testCase.DISP_ORACLE{r, 4};

                    off   = 8 + 2 + idx * 12;          % byte address of dir entry idx
                    entry = mat2doc.image.IfdEntryFactory_(rdr, off);

                    testCase.verifyEqual(double(entry.tag), wantTag, ...
                        sprintf('%s entry %d tag', fn, idx));
                    testCase.verifyEqual(normCls(entry), wantCls, ...
                        sprintf('%s entry %d dispatch class', fn, idx));
                    testCase.verifyEqual(canon(entry.value), wantVal, ...
                        sprintf('%s entry %d value', fn, idx));
                end
            end
        end

        % =============================================================== %
        % 9. H3 None dims (Exif-embedded case)                            %
        % =============================================================== %

        function test_pxnone_none_dims(testCase)
            % Edge (H3): a TIFF with NO IMAGE_WIDTH/IMAGE_LENGTH tags (the case for
            % a TIFF embedded in an Exif JPEG, where dims come from the JPEG SOF
            % marker instead): px_width/px_height -> None ([]) via get(tag,[]),
            % while horz_dpi still resolves to 200. (validate section 4c, PXNONE.)
            here = fileparts(mfilename('fullpath'));
            parser = mat2doc.image.TiffParser_.parse( ...
                mat2doc.image.BytesIO(readBlobBin(fullfile(here, 'data', 's0085_pxnone_ii.tif'))));
            testCase.verifyEmpty(parser.px_width,  'pxnone px_width must be None ([])');
            testCase.verifyEmpty(parser.px_height, 'pxnone px_height must be None ([])');
            testCase.verifyEqual(double(parser.horz_dpi), 200, 'pxnone horz_dpi still resolves to 200');
        end

        % =============================================================== %
        % 10. Factory dispatch -> REAL Tiff; Jfif/Exif still notYetPorted %
        % =============================================================== %

        function test_factory_dispatch_and_stubs(testCase)
            % Edge: the P7-1a factory now dispatches a TIFF MM/II signature to the
            % REAL Tiff parser (Image.from_blob returns a parsed Image with the
            % right px/dpi). RE-PINNED at P7-2b: the Jfif/Exif signatures now
            % dispatch to their REAL parsers too (mat2doc:Exception EOF on the
            % SOS-less signature blob, no longer notYetPorted).
            here = fileparts(mfilename('fullpath'));

            % A well-formed TIFF MM + II both dispatch to the real Tiff parser.
            iiImg = mat2doc.image.Image.from_blob( ...
                readBlobBin(fullfile(here, 'data', 's0085_types_ii.tif')));
            testCase.verifyEqual(string(iiImg.content_type), "image/tiff", 'TIFF(II) dispatch -> real Tiff');
            testCase.verifyEqual(double(iiImg.px_width), 120, 'TIFF(II) dispatch parsed px_width');

            mmImg = mat2doc.image.Image.from_blob( ...
                readBlobBin(fullfile(here, 'data', 's0085_types_mm.tif')));
            testCase.verifyEqual(string(mmImg.content_type), "image/tiff", 'TIFF(MM) dispatch -> real Tiff');
            testCase.verifyEqual(double(mmImg.horz_dpi), 300, 'TIFF(MM) dispatch parsed dpi');

            % Jfif / Exif signatures -- RE-PINNED at P7-2b: the real Jfif/Exif
            % parsers (un-stubbed) dispatch through JfifMarkers_.from_stream; the
            % SOS-less signature-only blob makes the marker walk hit EOF ->
            % mat2doc:Exception "unexpected end of file" (was notYetPorted).
            jfif = uint8([255 216 255 224 0 16, double('JFIF'), 0, zeros(1, 21)]);
            exif = uint8([255 216 255 225 0 16, double('Exif'), 0, zeros(1, 21)]);
            [idJ, msgJ] = catchErr(@() mat2doc.image.Image.from_blob(jfif));
            testCase.verifyEqual(idJ, 'mat2doc:Exception', 'JFIF now dispatches to the real Jfif parser (P7-2b)');
            testCase.verifyEqual(msgJ, 'unexpected end of file', 'JFIF EOF message verbatim');
            [idE, msgE] = catchErr(@() mat2doc.image.Image.from_blob(exif));
            testCase.verifyEqual(idE, 'mat2doc:Exception', 'Exif now dispatches to the real Exif parser (P7-2b)');
            testCase.verifyEqual(msgE, 'unexpected end of file', 'Exif EOF message verbatim');
        end

        % =============================================================== %
        % ★ 11. Full frozen-battery equivalence replay (self-contained)    %
        % =============================================================== %

        function test_replay_all_191_records(testCase)
            % ★ Equivalence (headline): replay the ENTIRE frozen Gate-3 battery
            % (the exact 191-probe sequence of s0085_..._probe.m, rebuilt inline
            % through mat2doc.image.* from the self-contained data\ copies) and
            % assert the port reproduces every record's tagged canon value-identical
            % to the frozen python-docx 1.2.0 oracle. Fixture is
            % references\s0085\probe.json copied verbatim into
            % tests\image\data\s0085_probe.json. Same tagged-canon scheme
            % (OK|int|, OK|str|, OK|none|, err|Name|msg) the Gate-3 comparator used.
            here = fileparts(mfilename('fullpath'));
            refPath = fullfile(here, 'data', 's0085_probe.json');
            testCase.assertTrue(isfile(refPath), ...
                sprintf('frozen probe fixture missing: %s', refPath));
            ref = jsondecode(readTextUtf8(refPath));

            obs = buildObservedBattery(here);

            keys = fieldnames(ref);
            testCase.verifyEqual(numel(keys), 191, ...
                'frozen s0085 battery must hold 191 records');
            testCase.verifyEqual(double(obs.Count), 191, ...   % Map.Count is uint64
                'rebuilt battery must reproduce exactly 191 records');
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
        % ★ 12. M1 byte-pin (P7-2a is M1-neutral)                         %
        % =============================================================== %

        function test_m1_neutral_styles_document_bytes(testCase)
            % ★ Regression: mat2doc.Document().save() emits word/styles.xml and
            % word/document.xml byte-identical to the frozen s0001 oracle (SHA-256
            % == byte-identity, L1). TIFF parsing touches no oxml registry row, no
            % PartFactory registration, nothing on the save path -- so M1 must stay
            % green. RED here would mean the P7-2a re-port perturbed the save path.
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

function [id, msg] = catchErr(fn)
    % Run fn(); return the caught error identifier + message ('' / '' if none).
    id = ''; msg = '';
    try
        fn();
    catch ME
        id = ME.identifier;
        msg = ME.message;
    end
end

function nm = normCls(entry)
    % "mat2doc.image.AsciiIfdEntry_" -> "AsciiIfdEntry"; base -> "IfdEntry".
    c = string(class(entry));
    c = extractAfter(c, "mat2doc.image.");
    if endsWith(c, "_"), c = extractBefore(c, strlength(c)); end
    nm = c;
end

% ------------------------------------------------------------ M1 zip helpers
function [blobs, names] = zipEntryList(zipBytes)
    % Enumerate an in-memory zip in stream (write) order -> ordered {names, blobs}.
    % (Copied from Test_p7_1b_png_gif_bmp.m so the pin is independent of the
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
% s0085_p7_2a_tiff_probe.m (same probe order, same section map: IMG / DPI /
% MUNIT / DEN0 / PXNONE / DISP / IFD; same canon scheme), rebuilt through
% mat2doc.image.* into a containers.Map keyed by probe name -> tagged canon
% string. The 2 real TIFFs + 16 crafted blobs are read from tests\image\data\
% (self-contained; the scenario crafted them in-memory / read the reals from the
% READ-ONLY python-docx clone -- byte-identical, sha1-verified).

function obs = buildObservedBattery(here)
    obs = containers.Map('KeyType', 'char', 'ValueType', 'char');
    dataDir = fullfile(here, 'data');

    % ===================================================== IMG  end-to-end surface
    ship = {"SHIP72", "72-dpi.tiff"; "SHIPLE", "little-endian.tif"};
    for r = 1:size(ship, 1)
        img = mat2doc.image.Image.from_file(fullfile(dataDir, char(ship{r, 2})));
        putImgSurface(obs, "IMG_" + ship{r, 1}, img, true);
    end
    crafted = {"inch_ii", "inch_mm", "cm_ii", "aspect_ii", "nounit_ii", ...
        "nores_ii", "big_ii", "tie_ii", "tie2_ii", "tie_cm_ii"};
    for i = 1:numel(crafted)
        img = mat2doc.image.Image.from_blob(readCrafted(dataDir, crafted{i}));
        putImgSurface(obs, "IMG_" + crafted{i}, img, false);
    end

    % ===================================================== DPI  guard battery
    dpiset = {"inch_ii", "inch_mm", "cm_ii", "aspect_ii", "nounit_ii", ...
        "nores_ii", "big_ii", "tie_ii", "tie2_ii", "tie_cm_ii", "types_ii"};
    for i = 1:numel(dpiset)
        nm = dpiset{i};
        parser = mat2doc.image.TiffParser_.parse(mat2doc.image.BytesIO(readCrafted(dataDir, nm)));
        put(obs, "DPI_" + nm + "_h", P(@() parser.horz_dpi));
        put(obs, "DPI_" + nm + "_v", P(@() parser.vert_dpi));
    end

    % ===================================================== MUNIT  F-1 multi-value unit
    img = mat2doc.image.Image.from_blob(readCrafted(dataDir, "multiunit_ii"));
    put(obs, "MUNIT_horz_dpi",  P(@() img.horz_dpi));
    put(obs, "MUNIT_vert_dpi",  P(@() img.vert_dpi));
    put(obs, "MUNIT_px_width",  P(@() img.px_width));
    put(obs, "MUNIT_px_height", P(@() img.px_height));

    % ===================================================== DEN0  ZeroDivisionError
    den0 = {"den0_ii", "den0_mm"};
    for i = 1:numel(den0)
        nm = den0{i};
        blob = readCrafted(dataDir, nm);
        put(obs, "DEN0_" + nm + "_from_blob", P(@() mat2doc.image.Image.from_blob(blob)));
        put(obs, "DEN0_" + nm + "_parse", ...
            P(@() mat2doc.image.TiffParser_.parse(mat2doc.image.BytesIO(blob))));
    end

    % ===================================================== PXNONE  H3 None dims
    parser = mat2doc.image.TiffParser_.parse(mat2doc.image.BytesIO(readCrafted(dataDir, "pxnone_ii")));
    put(obs, "PXNONE_px_width",  P(@() parser.px_width));
    put(obs, "PXNONE_px_height", P(@() parser.px_height));
    put(obs, "PXNONE_horz_dpi",  P(@() parser.horz_dpi));

    % ===================================================== DISP  entry-type dispatch
    disp = {"types_ii", mat2doc.image.StreamReader.LITTLE_ENDIAN; ...
            "types_mm", mat2doc.image.StreamReader.BIG_ENDIAN};
    for r = 1:size(disp, 1)
        nm = disp{r, 1};
        blob = readCrafted(dataDir, nm);
        rdr = mat2doc.image.StreamReader(mat2doc.image.BytesIO(blob), disp{r, 2});
        n = rdr.read_short(8);
        put(obs, "DISP_" + nm + "_count", P(@() n));
        for idx = 0:n-1
            off = 8 + 2 + idx * 12;
            entry = mat2doc.image.IfdEntryFactory_(rdr, off);
            put(obs, sprintf("DISP_%s_%d_tag", nm, idx), P(@() entry.tag));
            put(obs, sprintf("DISP_%s_%d_cls", nm, idx), P(@() normCls(entry)));
            put(obs, sprintf("DISP_%s_%d_val", nm, idx), P(@() entry.value));
        end
    end

    % ===================================================== IFD  endian fidelity
    ifd = {"types_ii", "types_mm"};
    for i = 1:numel(ifd)
        nm = ifd{i};
        img = mat2doc.image.Image.from_blob(readCrafted(dataDir, nm));
        put(obs, "IFD_" + nm + "_px_width",     P(@() img.px_width));
        put(obs, "IFD_" + nm + "_px_height",    P(@() img.px_height));
        put(obs, "IFD_" + nm + "_horz_dpi",     P(@() img.horz_dpi));
        put(obs, "IFD_" + nm + "_vert_dpi",     P(@() img.vert_dpi));
        put(obs, "IFD_" + nm + "_content_type", P(@() img.content_type));
    end
end

function blob = readCrafted(dataDir, name)
    % Read a crafted s0085 blob (data\s0085_<name>.tif). The scenario keyed the
    % corpus by bare name; the self-contained fixtures carry the s0085_ prefix.
    blob = readBlobBin(fullfile(dataDir, "s0085_" + string(name) + ".tif"));
end

function putImgSurface(obs, key, img, withFilename)
    put(obs, key + "_content_type", P(@() img.content_type));
    put(obs, key + "_px_width",     P(@() img.px_width));
    put(obs, key + "_px_height",    P(@() img.px_height));
    put(obs, key + "_horz_dpi",     P(@() img.horz_dpi));
    put(obs, key + "_vert_dpi",     P(@() img.vert_dpi));
    put(obs, key + "_width_emu",    P(@() img.width));
    put(obs, key + "_height_emu",   P(@() img.height));
    put(obs, key + "_ext",          P(@() img.ext));
    put(obs, key + "_sha1",         P(@() img.sha1));
    if withFilename
        put(obs, key + "_filename", P(@() img.filename));
    end
end

function put(map, key, canonStr)
    map(char(key)) = char(canonStr);
end

% ------------------------------------------------------------ probe wrappers
% (verbatim from the Gate-3 scenario s0085_..._probe.m so the canon is identical)

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
if isnumeric(v) && isempty(v)     % H3 None idiom ([]) -> Python None
    s = "none|";
elseif isa(v, "uint8")
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

classdef Test_p7_2b_jpeg < matlab.unittest.TestCase
% TEST_P7_2B_JPEG  Gate-4 permanent unit tests for Mat2Doc P7-2b
%   (the JPEG image-header parser -- Jfif/Exif + the marker classes -- WITH the
%   App1-via-Tiff dpi reversion. COMPLETES the image-parser tier.)
%
%   Surface under test (ported from python-docx v1.2.0 src/docx/image/jpeg.py;
%   the Exif/APP1 dpi comes from the docx Tiff.from_stream parser [P7-2a], NOT
%   the Mat2Ppt PIL CLASS-E reader; the App0/JFIF unit-0 aspect fallback to the
%   Exif dpi -- Mat2Ppt CLASS-J -- is REVERTED, docx uses the APP0 dpi
%   unconditionally):
%     - mat2doc.image.Jpeg       (jpeg.py Jpeg base; content_type/default_ext)
%     - mat2doc.image.Jfif       (jpeg.py Jfif)     -- un-stubs P7-1a Jfif row.
%     - mat2doc.image.Exif       (jpeg.py Exif)     -- un-stubs P7-1a Exif row.
%     - mat2doc.image.JfifMarkers_   (_JfifMarkers) -- app0/app1/sof, KeyError.
%     - mat2doc.image.MarkerParser_  (_MarkerParser)-- iter_markers cursor.
%     - mat2doc.image.MarkerFinder_  (_MarkerFinder)-- byte-address FF scan, EOF.
%     - mat2doc.image.MarkerIterator_(generator body)-- H9 laziness (break at SOS).
%     - mat2doc.image.MarkerFactory_ (_MarkerFactory)-- H10 marker dispatch.
%     - mat2doc.image.Marker_        (_Marker)      -- base/generic marker.
%     - mat2doc.image.App0Marker_    (_App0Marker)  -- JFIF dpi (inch/cm/aspect).
%     - mat2doc.image.App1Marker_    (_App1Marker)  -- Exif dpi via Tiff.from_stream.
%     - mat2doc.image.SofMarker_     (_SofMarker)   -- px dimensions.
%   These un-stub the P7-1a notYetPorted Jfif/Exif rows; with Tiff already live
%   (P7-2a), the image-parser tier is now COMPLETE -- NO parser stub remains.
%
%   Provenance (Gate-1..3, all 2026-08-01):
%     * Audit    : validation\mat2doc\audit_P7-2b_jpeg.md
%                  (Porter Gate-1 + Fable Gate-2 APPROVE -- 151/151, reversion
%                   total; docx-exact dpi, App1-via-Tiff not PIL PROVEN).
%     * Validate : validation\mat2doc\validate_P7-2b_jpeg.md
%                  (Gate-3 PASS, 170/170 value-identical vs python-docx 1.2.0,
%                   ZERO new D-numbers; pure-parsing WP, no package bytes; 9/9
%                   crafted blobs byte-identical between the two crafters).
%     * Scenario : validation\mat2doc\scenarios\s0086_p7_2b_jpeg_probe.{py,m}
%     * Frozen ref (python-docx 1.2.0 oracle, frozen ONCE):
%         references\s0086\probe.json -- the 170-value oracle, copied verbatim
%           into tests\image\data\s0086_probe.json (co-located
%           `s0086_probe.json text eol=lf`) so this suite is self-contained.
%         references\s0086\blobs\*.jpg -- the 7 validator-crafted JPEGs (the
%           App0/App1-via-Tiff dpi-branch corpus) + \*.bin -- the 2 raw APP1
%           segments (marker-level dpi pins), copied verbatim into
%           tests\image\data\s0086_*.jpg / s0086_*.bin (co-located `* binary`
%           .gitattributes; each blob's sha1 re-verified byte-identical to the
%           frozen reference).
%         The 3 real docx JPEGs (jfif-iguana.jpg, python-icon.jpeg,
%           exif-420-dpi.jpg) copied into tests\image\data\ (binary-pinned; all
%           sha1 re-verified vs the READ-ONLY python-docx clone). 300-dpi.jpg was
%           already binary-pinned in the P7-1a fixture block (sha1 4040957e...).
%     * M1 byte-pin values reuse the frozen s0001 manifest (word/styles.xml
%       02d71a68..., word/document.xml 0e4dd503...) already owned by
%       Test_p1_8_skeleton_m1; JPEG parsing is M1-neutral (touches no oxml
%       registry row, no PartFactory registration, nothing on the save path).
%
%   Coverage taxonomy
%   -----------------
%   * NOMINAL / Image equivalence (headline) -- test_image_equivalence_corpus:
%     Image.from_file on the 4 REAL docx JPEGs (JFIF inch/aspect + Exif) +
%     Image.from_blob on the 7 crafted dpi-branch JPEGs; content_type / ext /
%     px / dpi / width_EMU / height_EMU / sha1 [+ filename] all hard-pinned to the
%     frozen s0086 oracle. sha1 is END-TO-END on the real blob.
%   * ★★ App1/Exif-via-Tiff pins (the docx-not-PIL guards -- LOUD) --
%     test_app1_exif_via_tiff_pins:
%       - the exif-420-dpi.jpg RED HERRING -> 72/72 (embedded Exif TIFF holds
%         XRes=72/unit=2 -> docx reads 72, NOT the filename's "420"). A PIL
%         exif-reader regression fails HERE.
%       - exif_xres300 -> 300/300 (== a standalone Tiff on the segment).
%       - the ASYMMETRIC exif_asym (XRes=96, YRes=300) -> 96/300 (X->horz,
%         Y->vert; kills any PIL "XResolution for both axes" leftover, which would
%         give 96/96).
%       - exif_cm (ResolutionUnit=3) -> 183/183 = round(72*2.54) inside the Tiff.
%       - marker-level: App1Marker_.from_stream on the raw Exif APP1 segment
%         (exif_app1_x300.bin) -> 300/300 (routes to Tiff.from_stream); on a
%         NON-Exif APP1 segment (nonexif_app1.bin, sig 'http:/') -> 72/72.
%   * ★ App0 / JFIF dpi pins -- test_app0_jfif_dpi_pins:
%       units=inch(1) -> density (jfif_inch 150); units=cm(2) ->
%       round(density*2.54) (jfif_cm 72/cm -> 183); aspect/units=0 -> 72
%       UNCONDITIONALLY (jfif_aspect; the CLASS-J JFIF->Exif fallback REVERTED);
%       the half-to-even DOWN tie 190.5 -> 190 (jfif_cm_tie horz; NOT native
%       round() 191) alongside the odd-floor 63.5 -> 64 (vert).
%   * ★ V-A dpi -> EMU pin -- test_dpi_to_emu_pin: exif_xres300 -> width_emu
%     365760 / height_emu 274320 (the value P7-4 will place into wp:extent).
%   * marker/SOF fidelity -- test_marker_walk_and_sof: the full marker walk
%     (code/offset/segment_length + count) on the 2 real JPEGs, hard-pinned
%     row-for-row (300-dpi.jpg 12 markers incl. TWO APP1s at idx2/idx4;
%     exif-420-dpi.jpg 6 markers), + SOF px dims (1504x1936 / 2048x1536).
%   * error paths -- test_error_paths: KeyError "no APP0/APP1 marker in image"
%     on an absent marker (both directions) + truncated stream (FF D8 FF) ->
%     mat2doc:Exception "unexpected end of file" (VERIFY-2/-3, verbatim message).
%   * factory dispatch / ALL-6-LIVE -- test_factory_dispatch_all_6_live: a JFIF
%     blob dispatches to mat2doc.image.Jfif, an Exif blob to mat2doc.image.Exif,
%     and every one of the 6 format parsers (Png/Jfif/Exif/Gif/Tiff/Bmp) is LIVE
%     (NO from_stream raises mat2doc:notYetPorted -- the image-parser tier is
%     COMPLETE).
%   * ★ Equivalence (headline replay) -- test_replay_all_170_records: replays the
%     ENTIRE frozen Gate-3 battery (the exact 170-probe sequence of
%     s0086_..._probe.m, rebuilt inline through mat2doc.image.* from the
%     self-contained data\ copies) and asserts the port reproduces every record's
%     tagged canon value-identical to the frozen python-docx 1.2.0 oracle.
%   * ★ M1 byte-pin (regression) -- test_m1_neutral_styles_document_bytes:
%     mat2doc.Document().save() -> word/styles.xml sha256 02d71a68... +
%     word/document.xml sha256 0e4dd503... (SHA-256 == byte-identity, L1). RED
%     here would mean the JPEG port perturbed the save path.
%
%   Deviations exercised (adopt-only, ZERO new D-numbers -- Gate-3): the dpi is the
%   sole value-deviation risk and it is the python-docx formula EXACTLY -- App0 dpi
%   is the docx _dpi formula (inch/cm/aspect + half-to-even), App1 dpi is the docx
%   Tiff.from_stream path (the Mat2Ppt PIL CLASS-E/CLASS-J contracts are REVERTED).
%   Error paths are error-path matches to docx (KeyError / bare Exception; verbatim
%   messages). No serialization occurs in the parser, so no L1 byte surface is
%   touched (equivalence is value-level; validate section 2-5).
%
%   Determinism: no network, no hard-coded absolute paths -- the frozen oracle, the
%   9 crafted blobs, and the 4 real JPEGs resolve relative to this file via
%   fileparts(mfilename('fullpath')). Every image read is BINARY (no CRLF xlate).

    properties (Constant)
        % ★ Image-equivalence oracle (references\s0086\probe.json IMG_* rows).
        % Cols: tag, fixture-file, isReal (from_file vs from_blob), content_type,
        % ext, filename (""=from_blob, not probed), px_w, px_h, hdpi, vdpi,
        % width_EMU, height_EMU, sha1 (END-TO-END on the real blob).
        IMG_ORACLE = { ...
            "REAL300",  "300-dpi.jpg",            true,  "image/jpeg","jpg", "300-dpi.jpg",     1504, 1936, 300, 300, 4584192,  5900928,  "4040957ee43594f8eaa6abe4cd5c1fe0d6fc6a55"; ...
            "REALIGU",  "jfif-iguana.jpg",        true,  "image/jpeg","jpg", "jfif-iguana.jpg", 100,  68,   72,  72,  1270000,  863600,   "c3d98686223ad69ea29c811aaab35d343ff1ae9e"; ...
            "REALICO",  "python-icon.jpeg",       true,  "image/jpeg","jpeg","python-icon.jpeg",204,  204,  72,  72,  2590800,  2590800,  "1be010ea47803b00e140b852765cdf84f491da47"; ...
            "REALEXIF", "exif-420-dpi.jpg",       true,  "image/jpeg","jpg", "exif-420-dpi.jpg",2048, 1536, 72,  72,  26009600, 19507200, "04c3cadfd5661d34a46a588d201861ec823d000c"; ...
            "jfif_inch",    "s0086_jfif_inch.jpg",    false, "image/jpeg","jpg","",             120, 90,  150, 150, 731520,   548640,   "3774be9a40e4886b65306b4f80a524d4c1a32bd8"; ...
            "jfif_cm",      "s0086_jfif_cm.jpg",      false, "image/jpeg","jpg","",             120, 90,  183, 183, 599606,   449704,   "88b029134676ada4fc2b9102ba2495b64275765d"; ...
            "jfif_cm_tie",  "s0086_jfif_cm_tie.jpg",  false, "image/jpeg","jpg","",             120, 90,  190, 64,  577515,   1285875,  "bf1d448a59653a2afd272e7c6e6f167325cf1103"; ...
            "jfif_aspect",  "s0086_jfif_aspect.jpg",  false, "image/jpeg","jpg","",             120, 90,  72,  72,  1524000,  1143000,  "59fd1d7a43ebcadd09ee3bb4b15dc5d39ea9b2b9"; ...
            "exif_xres300", "s0086_exif_xres300.jpg", false, "image/jpeg","jpg","",             120, 90,  300, 300, 365760,   274320,   "3e9679718f60eba88a37ca47482440aceb4c9dcb"; ...
            "exif_cm",      "s0086_exif_cm.jpg",      false, "image/jpeg","jpg","",             120, 90,  183, 183, 599606,   449704,   "0f13d767a2fcd299e1d91f7d13f4be70a9d53228"; ...
            "exif_asym",    "s0086_exif_asym.jpg",    false, "image/jpeg","jpg","",             120, 90,  96,  300, 1143000,  274320,   "8ca7d9b63a5059be81c772667b48c72a7450c2f0"};

        % ★★ App1/Exif-via-Tiff discriminator oracle (the docx-not-PIL guards).
        % Cols: tag, fixture, want_h, want_v.
        APP1_ORACLE = { ...
            "REALEXIF",     "exif-420-dpi.jpg",       72,  72;  ...  % RED HERRING: filename "420" ignored; Tiff XRes=72
            "exif_xres300", "s0086_exif_xres300.jpg", 300, 300; ...  % positive control == standalone Tiff
            "exif_asym",    "s0086_exif_asym.jpg",    96,  300; ...  % ASYMMETRIC: X->horz, Y->vert (NOT 96/96)
            "exif_cm",      "s0086_exif_cm.jpg",      183, 183};     % cm-in-Exif -> round(72*2.54)

        % ★ App0/JFIF dpi discriminator oracle. Cols: tag, fixture, want_h, want_v.
        APP0_ORACLE = { ...
            "jfif_inch",   "s0086_jfif_inch.jpg",   150, 150; ...  % unit 1 (inch) pass-through
            "jfif_cm",     "s0086_jfif_cm.jpg",     183, 183; ...  % unit 2 (cm) round(72*2.54)
            "jfif_aspect", "s0086_jfif_aspect.jpg", 72,  72;  ...  % unit 0 (aspect) -> 72 UNCONDITIONAL (CLASS-J revert)
            "jfif_cm_tie", "s0086_jfif_cm_tie.jpg", 190, 64};      % 190.5->190 (half-to-even DOWN), 63.5->64

        % Marker-walk oracle (probe.json MARKWALK_* rows). Per row: idx, code(dec),
        % offset, segment_length. Hard-pinned row-for-row.
        MARKWALK_R300 = { ...
            0,  216, 2,     0;    ...   % D8 SOI
            1,  224, 4,     16;   ...   % E0 APP0
            2,  225, 22,    6150; ...   % E1 APP1  (first Exif)
            3,  237, 6174,  1508; ...   % ED APP13
            4,  225, 7684,  7903; ...   % E1 APP1  (second)
            5,  226, 15589, 576;  ...   % E2 APP2
            6,  238, 16167, 14;   ...   % EE APPE
            7,  219, 16183, 132;  ...   % DB DQT
            8,  192, 16317, 17;   ...   % C0 SOF0
            9,  221, 16336, 4;    ...   % DD DRI
            10, 196, 16342, 418;  ...   % C4 DHT
            11, 218, 16762, 12};        % DA SOS
        MARKWALK_REXIF = { ...
            0,  216, 2,    0;    ...     % D8 SOI
            1,  225, 4,    8817; ...     % E1 APP1 (Exif)
            2,  219, 8823, 132;  ...     % DB DQT
            3,  196, 8957, 418;  ...     % C4 DHT
            4,  192, 9377, 17;   ...     % C0 SOF0
            5,  218, 9396, 12};          % DA SOS

        % SOF px-dim oracle (probe.json SOF_* rows). Cols: file, px_w, px_h.
        SOF_ORACLE = { ...
            "300-dpi.jpg",      1504, 1936; ...
            "exif-420-dpi.jpg", 2048, 1536};

        % ★ M1 byte-pin (== Test_p1_8_skeleton_m1 rows; python-docx 1.2.0 s0001
        % oracle). SHA-256 equality IS byte-identity (L1). P7-2b is M1-neutral.
        M1_STYLES_SHA   = "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"
        M1_DOCUMENT_SHA = "0e4dd503bc095219cd24e8312b372c588ca56bb834aacac53eafb7bf47836327"
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied from the proven sibling tests\image\Test_p7_2a_tiff.m.
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
            % ★ Regression/Nominal: Image.from_file on the 4 REAL docx JPEGs (JFIF
            % inch/aspect + Exif) and Image.from_blob on the 7 crafted dpi-branch
            % JPEGs (all read from the self-contained tests\image\data\ copies)
            % reproduce every emitted field value-identical to the frozen s0086
            % oracle. px/dpi/content_type are parsed END-TO-END by the newly-live
            % JPEG parser; sha1 hashes the real blob; the EMU values run the
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
        % ★★ 2. App1/Exif-via-Tiff pins (the docx-not-PIL guards -- LOUD)  %
        % =============================================================== %

        function test_app1_exif_via_tiff_pins(testCase)
            % ★★★ LOUD REVERSION GUARD -- the docx-not-PIL Exif dpi guard.
            % docx reads the Exif TIFF resolution via App1Marker_ -> Tiff.from_stream
            % (P7-2a). A PIL/CLASS-E exif-reader regression (XResolution for both
            % axes, ResolutionUnit-only-3 conversion) fails HERE.
            %
            %   ==> exif-420-dpi.jpg is the RED HERRING: the filename says "420" but
            %   ==> the embedded Exif TIFF holds XRes=72/1, ResolutionUnit=2, so the
            %   ==> docx Tiff dpi is 72/72 -- NOT 420. A PIL exif-reader that read a
            %   ==> different tag (or the filename) would go RED. <==
            %
            %   ==> exif_asym (XRes=96, YRes=300) -> 96/300: X->horz, Y->vert, NOT
            %   ==> swapped, NOT XRes-for-both-axes. Any residual PIL "XResolution
            %   ==> for both axes" leftover would have produced 96/96. <==
            % (audit section 3 / validate section 2-3, frozen s0086 rows.)
            here = fileparts(mfilename('fullpath'));
            for r = 1:size(testCase.APP1_ORACLE, 1)
                tag   = testCase.APP1_ORACLE{r, 1};
                fn    = testCase.APP1_ORACLE{r, 2};
                wantH = testCase.APP1_ORACLE{r, 3};
                wantV = testCase.APP1_ORACLE{r, 4};
                path  = fullfile(here, 'data', char(fn));
                if tag == "REALEXIF"
                    img = mat2doc.image.Image.from_file(path);       % real red herring
                else
                    img = mat2doc.image.Image.from_blob(readBlobBin(path));
                end
                testCase.verifyEqual(double(img.horz_dpi), wantH, sprintf( ...
                    '%s horz_dpi must be %d (docx reads the Exif TIFF resolution via Tiff.from_stream; a PIL exif-reader regression fails here)', ...
                    tag, wantH));
                testCase.verifyEqual(double(img.vert_dpi), wantV, sprintf( ...
                    '%s vert_dpi must be %d (App1-via-Tiff; NOT the PIL CLASS-E reader)', tag, wantV));
            end

            % ★ Explicit red-herring callout (docx reads 72, NOT the filename's 420).
            realExif = mat2doc.image.Image.from_file(fullfile(here, 'data', 'exif-420-dpi.jpg'));
            testCase.verifyEqual(double(realExif.horz_dpi), 72, ...
                'exif-420-dpi.jpg RED HERRING: docx reads the embedded Exif TIFF XRes=72, NOT the filename''s 420');

            % ★ Marker-level structural proof (App1Marker_.from_stream directly):
            %   raw Exif APP1 segment routes to Tiff.from_stream -> 300/300; a
            %   NON-Exif APP1 segment (sig 'http:/' != 'Exif\x00\x00') -> 72/72.
            %   The non-Exif branch is unreachable through from_blob (format-detect
            %   needs 'Exif' at offset 6), so it is pinned at the marker entry point.
            APP1 = mat2doc.image.JPEG_MARKER_CODE.APP1;
            exifSeg = readBin(here, 'exif_app1_x300');
            rdr = mat2doc.image.StreamReader( ...
                mat2doc.image.BytesIO(exifSeg), mat2doc.image.StreamReader.BIG_ENDIAN);
            m = mat2doc.image.App1Marker_.from_stream(rdr, APP1, 0);
            testCase.verifyEqual(double(m.horz_dpi), 300, 'raw Exif APP1 segment routes to Tiff.from_stream -> 300');
            testCase.verifyEqual(double(m.vert_dpi), 300, 'raw Exif APP1 segment -> 300 (Tiff dpi)');

            nonexifSeg = readBin(here, 'nonexif_app1');
            rdr2 = mat2doc.image.StreamReader( ...
                mat2doc.image.BytesIO(nonexifSeg), mat2doc.image.StreamReader.BIG_ENDIAN);
            m2 = mat2doc.image.App1Marker_.from_stream(rdr2, APP1, 0);
            testCase.verifyEqual(double(m2.horz_dpi), 72, 'non-Exif APP1 segment defaults to 72');
            testCase.verifyEqual(double(m2.vert_dpi), 72, 'non-Exif APP1 segment defaults to 72');
        end

        % =============================================================== %
        % ★ 3. App0 / JFIF dpi pins (inch/cm/aspect + half-to-even tie)    %
        % =============================================================== %

        function test_app0_jfif_dpi_pins(testCase)
            % ★ App0 dpi is the docx _dpi formula VERBATIM (jpeg.py 305-313):
            %   unit 1 (inch) -> density; unit 2 (cm) -> int(round(density*2.54));
            %   else (aspect-only / unknown) -> 72 UNCONDITIONALLY.
            %
            %   ==> jfif_aspect (units=0) -> 72/72: the Mat2Ppt CLASS-J JFIF->Exif
            %   ==> aspect fallback is REVERTED -- docx uses the APP0 dpi always. <==
            %
            %   ==> jfif_cm_tie: x_density=75 -> 75*2.54 = 190.5 -> 190 (floor-EVEN
            %   ==> DOWN, CPython half-to-even), NOT native MATLAB round()'s 191;
            %   ==> y_density=25 -> 63.5 -> 64 (floor 63 ODD, both agree). A silent
            %   ==> revert to round() would give 191 on horz and go RED. <==
            % (audit section 3 / validate section 3b, frozen s0086 rows.)
            here = fileparts(mfilename('fullpath'));
            for r = 1:size(testCase.APP0_ORACLE, 1)
                tag   = testCase.APP0_ORACLE{r, 1};
                fn    = testCase.APP0_ORACLE{r, 2};
                wantH = testCase.APP0_ORACLE{r, 3};
                wantV = testCase.APP0_ORACLE{r, 4};
                img = mat2doc.image.Image.from_blob(readBlobBin(fullfile(here, 'data', char(fn))));
                testCase.verifyEqual(double(img.horz_dpi), wantH, sprintf('%s horz_dpi', tag));
                testCase.verifyEqual(double(img.vert_dpi), wantV, sprintf('%s vert_dpi', tag));
            end

            % ★ The decisive tie, called out loudly: 190.5 -> 190 (NOT 191).
            tie = mat2doc.image.Image.from_blob( ...
                readBlobBin(fullfile(here, 'data', 's0086_jfif_cm_tie.jpg')));
            testCase.verifyEqual(double(tie.horz_dpi), 190, ...
                'jfif_cm_tie horz_dpi must be 190 (CPython half-to-even 190.5->190; a 191 here = silent revert to native round())');
            testCase.verifyEqual(double(tie.vert_dpi), 64, ...
                'jfif_cm_tie vert_dpi must be 64 (63.5->64, odd floor)');

            % ★ aspect-only unconditional 72 (CLASS-J revert), called out.
            asp = mat2doc.image.Image.from_blob( ...
                readBlobBin(fullfile(here, 'data', 's0086_jfif_aspect.jpg')));
            testCase.verifyEqual(double(asp.horz_dpi), 72, ...
                'jfif_aspect units=0 -> 72 UNCONDITIONALLY (no Exif fallback; CLASS-J reverted)');
            testCase.verifyEqual(double(asp.vert_dpi), 72, 'jfif_aspect vert_dpi 72');
        end

        % =============================================================== %
        % ★ 4. V-A dpi -> EMU pin (the wp:extent value P7-4 will place)    %
        % =============================================================== %

        function test_dpi_to_emu_pin(testCase)
            % ★ V-A (Gate-2 suggestion, proven at the Image level -- the full
            % wp:extent package is a P7-4 item): the Exif dpi threads through
            % Inches(px/dpi) -> Emu identically to python-docx. exif_xres300
            % (px 120x90, dpi 300/300) -> width_emu 365760 (=120/300*914400),
            % height_emu 274320 (=90/300*914400). This is the exact value P7-4's
            % add_picture will place into wp:extent. (validate section 4.)
            here = fileparts(mfilename('fullpath'));
            img = mat2doc.image.Image.from_blob( ...
                readBlobBin(fullfile(here, 'data', 's0086_exif_xres300.jpg')));
            testCase.verifyEqual(double(img.width),  365760, ...
                'exif_xres300 width_emu must be 365760 (=120/300*914400) -- the P7-4 wp:extent value');
            testCase.verifyEqual(double(img.height), 274320, ...
                'exif_xres300 height_emu must be 274320 (=90/300*914400)');
        end

        % =============================================================== %
        % 5. Marker-walk + SOF fidelity (H1/H2 byte-address; H9 laziness)  %
        % =============================================================== %

        function test_marker_walk_and_sof(testCase)
            % Regression: the per-marker (code, offset, segment_length) trace to SOS
            % on the 2 real JPEGs, driven through the port's own MarkerParser_/
            % MarkerIterator_ cursor, hard-pinned row-for-row to the frozen oracle.
            %   * 300-dpi.jpg: 12 markers SOI->SOS, including TWO APP1s (E1 at idx 2
            %     off 22 / idx 4 off 7684).
            %   * exif-420-dpi.jpg: 6 markers SOI->SOS.
            % Plus SOF px dims (via JfifMarkers_.sof). H9: both stop at SOS (the
            % entropy-coded scan data after SOS is never parsed as markers).
            % (validate section 5a/5b, frozen s0086 MARKWALK/SOF rows.)
            here = fileparts(mfilename('fullpath'));
            walks = { "300-dpi.jpg", testCase.MARKWALK_R300; ...
                      "exif-420-dpi.jpg", testCase.MARKWALK_REXIF };
            SOS = mat2doc.image.JPEG_MARKER_CODE.SOS;
            for w = 1:size(walks, 1)
                fn     = walks{w, 1};
                oracle = walks{w, 2};
                blob   = readBlobBin(fullfile(here, 'data', char(fn)));
                mp = mat2doc.image.MarkerParser_.from_stream(mat2doc.image.BytesIO(blob));
                it = mp.iter_markers();
                idx = 0;
                while it.has_more()
                    marker = it.next_marker();
                    row = oracle(idx + 1, :);
                    testCase.assertEqual(row{1}, idx, sprintf('%s marker index bookkeeping', fn));
                    testCase.verifyEqual(double(marker.marker_code), row{2}, ...
                        sprintf('%s marker %d code', fn, idx));
                    testCase.verifyEqual(double(marker.offset), row{3}, ...
                        sprintf('%s marker %d offset (byte-address arithmetic H1)', fn, idx));
                    testCase.verifyEqual(double(marker.segment_length), row{4}, ...
                        sprintf('%s marker %d segment_length', fn, idx));
                    idx = idx + 1;
                    if marker.marker_code == SOS, break; end
                end
                testCase.verifyEqual(idx, size(oracle, 1), ...
                    sprintf('%s must yield exactly %d markers SOI->SOS', fn, size(oracle, 1)));
            end

            % SOF px dims (JfifMarkers_.sof).
            for s = 1:size(testCase.SOF_ORACLE, 1)
                fn     = testCase.SOF_ORACLE{s, 1};
                wantW  = testCase.SOF_ORACLE{s, 2};
                wantH  = testCase.SOF_ORACLE{s, 3};
                markers = mat2doc.image.JfifMarkers_.from_stream( ...
                    mat2doc.image.BytesIO(readBlobBin(fullfile(here, 'data', char(fn)))));
                testCase.verifyEqual(double(markers.sof.px_width),  wantW, sprintf('%s SOF px_width', fn));
                testCase.verifyEqual(double(markers.sof.px_height), wantH, sprintf('%s SOF px_height', fn));
            end
        end

        % =============================================================== %
        % 6. Error paths -- KeyError (absent APP0/APP1) + truncated EOF    %
        % =============================================================== %

        function test_error_paths(testCase)
            % Error path (VERIFY-2/-3): absent-marker KeyError + truncated stream.
            %   * python-icon.jpeg is a JFIF (no APP1) -> JfifMarkers_.app1 raises
            %     mat2doc:KeyError "no APP1 marker in image".
            %   * exif-420-dpi.jpg is an Exif (no APP0) -> JfifMarkers_.app0 raises
            %     mat2doc:KeyError "no APP0 marker in image".
            %   * a truncated stream (FF D8 FF) -> the marker walk hits EOF ->
            %     mat2doc:Exception "unexpected end of file" (port-authored id,
            %     verbatim CPython message).
            % Bare message compared (validate section 1a: Python str(KeyError) adds
            % quotes; the semantically-raised argument is the bare message).
            % (validate section 5c, frozen s0086 KEYERR/TRUNC rows.)
            here = fileparts(mfilename('fullpath'));

            mkJfif = mat2doc.image.JfifMarkers_.from_stream( ...
                mat2doc.image.BytesIO(readBlobBin(fullfile(here, 'data', 'python-icon.jpeg'))));
            [id1, msg1] = catchErr(@() mkJfif.app1);
            testCase.verifyEqual(id1, 'mat2doc:KeyError', 'JFIF with no APP1 -> mat2doc:KeyError');
            testCase.verifyEqual(msg1, 'no APP1 marker in image', 'no-APP1 message verbatim');

            mkExif = mat2doc.image.JfifMarkers_.from_stream( ...
                mat2doc.image.BytesIO(readBlobBin(fullfile(here, 'data', 'exif-420-dpi.jpg'))));
            [id2, msg2] = catchErr(@() mkExif.app0);
            testCase.verifyEqual(id2, 'mat2doc:KeyError', 'Exif with no APP0 -> mat2doc:KeyError');
            testCase.verifyEqual(msg2, 'no APP0 marker in image', 'no-APP0 message verbatim');

            [id3, msg3] = catchErr(@() mat2doc.image.JfifMarkers_.from_stream( ...
                mat2doc.image.BytesIO(uint8([255 216 255]))));
            testCase.verifyEqual(id3, 'mat2doc:Exception', 'truncated stream -> mat2doc:Exception (VERIFY-2)');
            testCase.verifyEqual(msg3, 'unexpected end of file', 'truncated-stream message verbatim');
        end

        % =============================================================== %
        % 7. Factory dispatch -> REAL Jfif/Exif; ALL 6 parsers LIVE        %
        % =============================================================== %

        function test_factory_dispatch_all_6_live(testCase)
            % Edge: the ImageHeaderFactory_ SIGNATURES table dispatches a JFIF blob
            % to the REAL mat2doc.image.Jfif parser and an Exif blob to the REAL
            % mat2doc.image.Exif parser (class-exact on the header object). AND --
            % the image-parser tier being COMPLETE -- every one of the 6 format
            % parsers is LIVE: NO from_stream raises mat2doc:notYetPorted anymore.
            % A regression that re-stubs any parser goes RED here.
            here = fileparts(mfilename('fullpath'));

            % Class-exact dispatch on well-formed crafted blobs.
            jHdr = mat2doc.image.ImageHeaderFactory_( ...
                mat2doc.image.BytesIO(readBlobBin(fullfile(here, 'data', 's0086_jfif_inch.jpg'))));
            testCase.verifyClass(jHdr, 'mat2doc.image.Jfif', 'JFIF signature must dispatch to the real Jfif parser');
            eHdr = mat2doc.image.ImageHeaderFactory_( ...
                mat2doc.image.BytesIO(readBlobBin(fullfile(here, 'data', 's0086_exif_xres300.jpg'))));
            testCase.verifyClass(eHdr, 'mat2doc.image.Exif', 'Exif signature must dispatch to the real Exif parser');

            % ★ ALL-6-LIVE: no format parser is a notYetPorted stub. Each is called
            % on a 4-byte garbage stream; whatever it raises, it must NOT be
            % mat2doc:notYetPorted (image-parser tier COMPLETE).
            stream = @() mat2doc.image.BytesIO(uint8([0 0 0 0]));
            parsers = { ...
                'Png',  @() mat2doc.image.Png.from_stream(stream()); ...
                'Jfif', @() mat2doc.image.Jfif.from_stream(stream()); ...
                'Exif', @() mat2doc.image.Exif.from_stream(stream()); ...
                'Gif',  @() mat2doc.image.Gif.from_stream(stream()); ...
                'Tiff', @() mat2doc.image.Tiff.from_stream(stream()); ...
                'Bmp',  @() mat2doc.image.Bmp.from_stream(stream())};
            for k = 1:size(parsers, 1)
                lbl = parsers{k, 1};
                [id, ~] = catchErr(parsers{k, 2});
                testCase.verifyNotEqual(id, 'mat2doc:notYetPorted', ...
                    sprintf('%s parser must be LIVE (no notYetPorted stub -- image-parser tier COMPLETE)', lbl));
            end
        end

        % =============================================================== %
        % ★ 8. Full frozen-battery equivalence replay (self-contained)     %
        % =============================================================== %

        function test_replay_all_170_records(testCase)
            % ★ Equivalence (headline): replay the ENTIRE frozen Gate-3 battery
            % (the exact 170-probe sequence of s0086_..._probe.m, rebuilt inline
            % through mat2doc.image.* from the self-contained data\ copies) and
            % assert the port reproduces every record's tagged canon value-identical
            % to the frozen python-docx 1.2.0 oracle. Fixture is
            % references\s0086\probe.json copied verbatim into
            % tests\image\data\s0086_probe.json. Same tagged-canon scheme
            % (OK|int|, OK|str|, OK|hex|, err|Name|msg) the Gate-3 comparator used.
            here = fileparts(mfilename('fullpath'));
            refPath = fullfile(here, 'data', 's0086_probe.json');
            testCase.assertTrue(isfile(refPath), ...
                sprintf('frozen probe fixture missing: %s', refPath));
            ref = jsondecode(readTextUtf8(refPath));

            obs = buildObservedBattery(here);

            keys = fieldnames(ref);
            testCase.verifyEqual(numel(keys), 170, ...
                'frozen s0086 battery must hold 170 records');
            testCase.verifyEqual(double(obs.Count), 170, ...   % Map.Count is uint64
                'rebuilt battery must reproduce exactly 170 records');
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
        % ★ 9. M1 byte-pin (P7-2b is M1-neutral)                          %
        % =============================================================== %

        function test_m1_neutral_styles_document_bytes(testCase)
            % ★ Regression: mat2doc.Document().save() emits word/styles.xml and
            % word/document.xml byte-identical to the frozen s0001 oracle (SHA-256
            % == byte-identity, L1). JPEG parsing touches no oxml registry row, no
            % PartFactory registration, nothing on the save path -- so M1 must stay
            % green. RED here would mean the P7-2b port perturbed the save path.
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

function b = readBin(here, name)
    % Read a raw APP1 segment (data\s0086_<name>.bin).
    b = readBlobBin(fullfile(here, 'data', "s0086_" + string(name) + ".bin"));
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

% ------------------------------------------------------------ M1 zip helpers
function [blobs, names] = zipEntryList(zipBytes)
    % Enumerate an in-memory zip in stream (write) order -> ordered {names, blobs}.
    % (Copied from Test_p7_2a_tiff.m so the pin is independent of the reader under
    % test.)
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
% s0086_p7_2b_jpeg_probe.m (same probe order, same section map: IMG / APP1 /
% MARKWALK / SOF / KEYERR / TRUNC; same canon scheme), rebuilt through
% mat2doc.image.* into a containers.Map keyed by probe name -> tagged canon
% string. The 4 real JPEGs + 9 crafted blobs are read from tests\image\data\
% (self-contained; the scenario read the reals from the READ-ONLY python-docx
% clone and re-crafted the blobs in-memory -- both byte-identical, sha1-verified).

function obs = buildObservedBattery(here)
    obs = containers.Map('KeyType', 'char', 'ValueType', 'char');
    dataDir = fullfile(here, 'data');

    % ===================================================== IMG  end-to-end surface
    real = {"REAL300", "300-dpi.jpg"; "REALIGU", "jfif-iguana.jpg"; ...
            "REALICO", "python-icon.jpeg"; "REALEXIF", "exif-420-dpi.jpg"};
    for r = 1:size(real, 1)
        img = mat2doc.image.Image.from_file(fullfile(dataDir, char(real{r, 2})));
        putImgSurface(obs, "IMG_" + real{r, 1}, img, true);
    end
    crafted = {"jfif_inch", "jfif_cm", "jfif_cm_tie", "jfif_aspect", ...
        "exif_xres300", "exif_cm", "exif_asym"};
    for i = 1:numel(crafted)
        img = mat2doc.image.Image.from_blob(readCraftedJpeg(dataDir, crafted{i}));
        putImgSurface(obs, "IMG_" + crafted{i}, img, false);
    end

    % ===================================================== APP1  App1-via-Tiff guard
    APP1 = mat2doc.image.JPEG_MARKER_CODE.APP1;
    rdr = mat2doc.image.StreamReader( ...
        mat2doc.image.BytesIO(readCraftedBin(dataDir, "nonexif_app1")), ...
        mat2doc.image.StreamReader.BIG_ENDIAN);
    m = mat2doc.image.App1Marker_.from_stream(rdr, APP1, 0);
    put(obs, "APP1_nonexif_horz_dpi", P(@() m.horz_dpi));
    put(obs, "APP1_nonexif_vert_dpi", P(@() m.vert_dpi));
    rdr = mat2doc.image.StreamReader( ...
        mat2doc.image.BytesIO(readCraftedBin(dataDir, "exif_app1_x300")), ...
        mat2doc.image.StreamReader.BIG_ENDIAN);
    m = mat2doc.image.App1Marker_.from_stream(rdr, APP1, 0);
    put(obs, "APP1_exif_horz_dpi", P(@() m.horz_dpi));
    put(obs, "APP1_exif_vert_dpi", P(@() m.vert_dpi));

    % ===================================================== MARKWALK  marker fidelity
    walk = {"R300", "300-dpi.jpg"; "REXIF", "exif-420-dpi.jpg"};
    SOS = mat2doc.image.JPEG_MARKER_CODE.SOS;
    for r = 1:size(walk, 1)
        tag  = walk{r, 1};
        blob = readBlobBin(fullfile(dataDir, char(walk{r, 2})));
        mp = mat2doc.image.MarkerParser_.from_stream(mat2doc.image.BytesIO(blob));
        it = mp.iter_markers();
        idx = 0;
        while it.has_more()
            marker = it.next_marker();
            k = sprintf("MARKWALK_%s_%d", tag, idx);
            put(obs, k + "_code", P(@() marker.marker_code));
            put(obs, k + "_off",  P(@() marker.offset));
            put(obs, k + "_seg",  P(@() marker.segment_length));
            idx = idx + 1;
            if marker.marker_code == SOS, break; end
        end
        put(obs, "MARKWALK_" + tag + "_count", P(@() idx));
    end

    % ===================================================== SOF  px dims (real files)
    for r = 1:size(walk, 1)
        tag = walk{r, 1};
        markers = mat2doc.image.JfifMarkers_.from_stream( ...
            mat2doc.image.BytesIO(readBlobBin(fullfile(dataDir, char(walk{r, 2})))));
        put(obs, "SOF_" + tag + "_px_width",  P(@() markers.sof.px_width));
        put(obs, "SOF_" + tag + "_px_height", P(@() markers.sof.px_height));
    end

    % ===================================================== KEYERR  absent APP0/APP1
    mkJfif = mat2doc.image.JfifMarkers_.from_stream( ...
        mat2doc.image.BytesIO(readBlobBin(fullfile(dataDir, "python-icon.jpeg"))));
    put(obs, "KEYERR_no_app1", P(@() mkJfif.app1));
    mkExif = mat2doc.image.JfifMarkers_.from_stream( ...
        mat2doc.image.BytesIO(readBlobBin(fullfile(dataDir, "exif-420-dpi.jpg"))));
    put(obs, "KEYERR_no_app0", P(@() mkExif.app0));

    % ===================================================== TRUNC  unexpected EOF
    put(obs, "TRUNC_eof", P(@() mat2doc.image.JfifMarkers_.from_stream( ...
        mat2doc.image.BytesIO(uint8([255 216 255])))));
end

function blob = readCraftedJpeg(dataDir, name)
    % Read a crafted s0086 JPEG (data\s0086_<name>.jpg). The scenario keyed the
    % corpus by bare name; the self-contained fixtures carry the s0086_ prefix.
    blob = readBlobBin(fullfile(dataDir, "s0086_" + string(name) + ".jpg"));
end

function blob = readCraftedBin(dataDir, name)
    % Read a crafted s0086 raw APP1 segment (data\s0086_<name>.bin).
    blob = readBlobBin(fullfile(dataDir, "s0086_" + string(name) + ".bin"));
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
% (verbatim from the Gate-3 scenario s0086_..._probe.m so the canon is identical)

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

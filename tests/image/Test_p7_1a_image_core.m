classdef Test_p7_1a_image_core < matlab.unittest.TestCase
% TEST_P7_1A_IMAGE_CORE  Gate-4 permanent unit tests for Mat2Doc P7-1a
%   (the image-parsing CORE re-port).
%
%   Surface under test (re-homed from Mat2Ppt +image / ported from python-docx
%   v1.2.0 src/docx/image/*):
%     - mat2doc.image.Image        -- the value object: content_type / px_* /
%       *_dpi delegation, width/height Length math (Inches->EMU), ext
%       (splitext_ext, the F-1 fix), filename basename, sha1 END-TO-END, and
%       scaled_dimensions (pyRound banker's rounding). (image.py::Image)
%     - mat2doc.image.StreamReader -- fixed-width BE/LE integer + UTF-8 string
%       reads with EOF detection. (helpers.py::StreamReader)
%     - mat2doc.image.BytesIO      -- the seekable io.BytesIO analogue.
%     - mat2doc.opc.sha1_hexdigest -- hashlib.sha1(...).hexdigest() replica.
%     - mat2doc.image.{MIME_TYPE,PNG_CHUNK_TYPE,TIFF_FLD,TIFF_TAG,
%       JPEG_MARKER_CODE} -- the five constants tables. (constants.py)
%     - mat2doc.image.ImageHeaderFactory_ -- the SIGNATURES dispatch + the
%       UnrecognizedImageError terminus (image.py::_ImageHeaderFactory,
%       __init__.py SIGNATURES); the 6 format parsers (Png/Jfif/Exif/Gif/Tiff/
%       Bmp) are clean notYetPorted stubs owned by P7-1b/P7-2.
%
%   Provenance (Gate-1..3, all 2026-08-01):
%     * Audit    : validation\mat2doc\audit_P7-1a_image_core.md  (Porter Gate-1 +
%                  Fable Gate-2 REVISE->F-1 fixed; the F-1 fix is Image.ext ->
%                  private splitext_ext replicating CPython os.path.splitext).
%     * Validate : validation\mat2doc\validate_P7-1a_image_core.md  (Gate-3 PASS,
%                  195/195 value-identical vs python-docx 1.2.0, 0 new D-numbers,
%                  NO re-pin list -- M1-neutral).
%     * Scenario : validation\mat2doc\scenarios\s0083_p7_1a_image_core_probe.{py,m}
%                  (+ the test-double header s0083_DoubleHeader.m, co-located in
%                  this folder as the self-contained helper; the frozen oracle
%                  references\s0083\probe.json).
%     * Frozen refs (python-docx 1.2.0 oracle, frozen ONCE):
%         references\s0083\probe.json -- the 195-value image-core oracle, copied
%           verbatim into tests\image\data\probe.json (co-located `probe.json
%           text eol=lf` .gitattributes) so this suite is self-contained; the two
%           real test images 300-dpi.png / 300-dpi.jpg copied into
%           tests\image\data\ (co-located `* binary` .gitattributes -- they are
%           the sha1 END-TO-END equivalence blobs).
%     * M1 byte-pin values reuse the frozen s0001 manifest (word/styles.xml
%       02d71a68..., word/document.xml 0e4dd503...) already owned by
%       Test_p1_8_skeleton_m1; image parsing is M1-neutral (touches no oxml
%       registry row, no PartFactory registration, nothing on the save path).
%
%   Coverage taxonomy
%   -----------------
%   * ★ Equivalence (headline) -- test_replay_all_195_records replays the ENTIRE
%     frozen Gate-3 battery (the exact probe sequence of s0083_..._probe.m,
%     rebuilt inline through mat2doc.image.* / mat2doc.opc.*) and asserts the
%     port reproduces every one of the 195 records' tagged canon
%     value-identical. Self-contained: the images + oracle are copied into
%     tests\image\data\.
%   * ★ Image-core equivalence (regression) -- 300-dpi.png / .jpg content_type /
%     px_* / *_dpi (seeded) + width/height EMU + ext + filename + sha1 END-TO-END
%     hard-pinned to the frozen s0083 oracle (sha1 ca11f589... PNG / 4040957e...
%     JPG; width/height EMU 2621280/1764792 and 4584192/5900928).
%   * ★ splitext_ext (F-1) regression pin (LOUD) -- the 23 ext vectors through
%     the LIVE Image.ext surface; the F-1 regression guard (a future re-port that
%     re-breaks the CPython splitext dotfile / basename-only / case-kept rule goes
%     RED here).
%   * StreamReader (nominal + edge + error path) -- BE/LE read_byte/short/long,
%     read_str ASCII + UTF-8, base_offset shift, unknown byte-order -> BE
%     fallback, and read-past-EOF -> mat2doc:UnexpectedEndOfFileError (identifier
%     pinned).
%   * BytesIO (edge) -- read(n) / read-all / seek+read / tell / read-past-EOF ->
%     empty (offset-indexing, 0-based cursor).
%   * sha1_hexdigest (regression) -- empty / "abc" / bytes 0..255 (the Java
%     signed-byte >127 boundary) / non-ASCII "héllo" utf-8 == hashlib.
%   * Constants (regression) -- MIME_TYPE / PNG_CHUNK_TYPE (pHYs mixed case) /
%     TIFF_FLD (+ field_type_name + KeyError) / TIFF_TAG (+ tag_name + KeyError) /
%     JPEG_MARKER_CODE (all 49 marker-code constants + STANDALONE (11) / SOF (13)
%     tuples + is_standalone + marker_name + KeyError); spot + count pins.
%   * scaled_dimensions + pyRound (edge) -- native / both / one-given, and the
%     half-to-EVEN banker's tie battery (0.5->0, 1.5->2, 2.5->2, 3.5->4,
%     -1.5->-2) driven through the real scaled_dimensions API.
%   * ★ WMF/EMF-exclusion (regression trap guard, LOUD) -- a WMF placeable
%     signature (D7 CD C6 9A) and an EMF signature (01 00 00 00) ->
%     mat2doc:UnrecognizedImageError. docx has NO WMF/EMF parser (unlike the
%     Mat2Ppt PIL seam); a future sloppy re-port that re-adds WMF goes RED here.
%   * Factory dispatch (edge) -- the 8 SIGNATURES rows dispatch the PNG/JPEG/GIF/
%     BMP/TIFF signatures to their (stubbed) parser, each raising
%     mat2doc:notYetPorted (P7-1b/P7-2 boundary); unrecognized ->
%     mat2doc:UnrecognizedImageError.
%   * NIE abstract members (error path) -- BaseImageHeader.content_type /
%     default_ext on a plain base -> mat2doc:NotImplementedError, message
%     byte-verbatim (cross-faithful upstream text).
%   * ★ M1 byte-pin (regression) -- mat2doc.Document().save() -> word/styles.xml
%     sha256 02d71a68... + word/document.xml sha256 0e4dd503... (image core is
%     M1-neutral; RED here would mean the re-port perturbed the save path).
%
%   Deviations exercised (adopt-only, ZERO new D-numbers -- Gate-3): the raised
%   error identifiers map under the pre-adopted D-003/D-004 error families
%   (faithful class, port-authored messages), so error paths pin the IDENTIFIER
%   (mat2doc:UnrecognizedImageError / mat2doc:UnexpectedEndOfFileError /
%   mat2doc:notYetPorted / mat2doc:KeyError); the two NotImplementedError messages
%   additionally match verbatim (cross-faithful upstream text). No serialization
%   occurs in the image core, so no L1 byte surface is touched (equivalence is
%   value-level; see validate_P7-1a section 2).
%
%   Determinism: no network, no hard-coded absolute paths -- the frozen oracle and
%   the two test images resolve relative to this file via
%   fileparts(mfilename('fullpath')); the test-double header s0083_DoubleHeader
%   resolves from this folder (runtests cd's here). Every image read is BINARY
%   ('r'/'n' -- no CRLF translation).

    properties (Constant)
        % ★ The frozen s0083 image-core equivalence values (python-docx 1.2.0
        % oracle, references\s0083\probe.json). Cols: tag, content_type, px_width,
        % px_height, horz_dpi, vert_dpi, width_EMU, height_EMU, ext, filename,
        % sha1. sha1 is the END-TO-END hash of the real blob; width/height EMU are
        % the genuine Inches/Length math (Inches(px/dpi) truncated via int()).
        IMG_ORACLE = { ...
            "PNG", "image/png",  860,  579, 300, 300, 2621280, 1764792, "png", "300-dpi.png", "ca11f589af080a26b619b76823897001d34c7e44"; ...
            "JPG", "image/jpeg", 1504, 1936, 300, 300, 4584192, 5900928, "jpg", "300-dpi.jpg", "4040957ee43594f8eaa6abe4cd5c1fe0d6fc6a55"};

        % ★ splitext_ext (F-1) vectors: {filename, expected ext} == CPython
        % os.path.splitext(fn)[1][1:] (case kept, basename-only, both separators).
        SPL_VECTORS = { ...
            ".bashrc", "";        ".png", "";           "..png", ""; ...
            ".myimage", "";       "...", "";            "a.png", "png"; ...
            "IMG.PNG", "PNG";     "a..png", "png";      ".tar.gz", "gz"; ...
            "archive.tar.gz", "gz"; "file.", "";        "trailing.dot.", ""; ...
            "image", "";          "noext", "";          "a.b.c.jpeg", "jpeg"; ...
            "image.J p g", "J p g"; "/x/y.z/a.png", "png"; "/x/y.z/noext", ""; ...
            "300-dpi.png", "png"; "300-dpi.jpg", "jpg"; "C:\dir.x\img.png", "png"; ...
            "C:\dir.x\noext", ""; "", ""};

        % ★ M1 byte-pin (== Test_p1_8_skeleton_m1.M1_MANIFEST rows; python-docx
        % 1.2.0 s0001 oracle). SHA-256 equality IS byte-identity (L1). Image core
        % is M1-neutral, so these must stay green.
        M1_STYLES_SHA   = "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"
        M1_DOCUMENT_SHA = "0e4dd503bc095219cd24e8312b372c588ca56bb834aacac53eafb7bf47836327"
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied from the proven tests\shared\Test_p1_1_shared.m. The
            % test-double header s0083_DoubleHeader.m lives in THIS folder and
            % resolves via the cwd runtests cd's into (proven by
            % tests\shared\Test_p2_1_proxy_tier.m + its s0013_* siblings).
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\image
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % ★ 1. Image-core equivalence (regression, the headline)          %
        % =============================================================== %

        function test_image_core_equivalence_png_jpg(testCase)
            % ★ Regression: Image core props on the two REAL test blobs ==
            % the frozen s0083 oracle. content_type / px_* / *_dpi are the seeded
            % (test-double header) delegation leg; width/height EMU / ext / filename
            % / sha1 are the GENUINE core (sha1 hashes the real blob end-to-end,
            % width/height run the real Inches/Length math). The RAW header parse is
            % re-proven end-to-end at P7-1b (VERIFY V-2) -- here the delegation seed
            % is the oracle's own parse (validate_P7-1a section 2).
            here = fileparts(mfilename('fullpath'));
            for r = 1:size(testCase.IMG_ORACLE, 1)
                tag      = testCase.IMG_ORACLE{r, 1};
                wantCT   = testCase.IMG_ORACLE{r, 2};
                wantPxW  = testCase.IMG_ORACLE{r, 3};
                wantPxH  = testCase.IMG_ORACLE{r, 4};
                wantHDpi = testCase.IMG_ORACLE{r, 5};
                wantVDpi = testCase.IMG_ORACLE{r, 6};
                wantWEmu = testCase.IMG_ORACLE{r, 7};
                wantHEmu = testCase.IMG_ORACLE{r, 8};
                wantExt  = testCase.IMG_ORACLE{r, 9};
                wantFn   = testCase.IMG_ORACLE{r, 10};
                wantSha1 = testCase.IMG_ORACLE{r, 11};

                fn   = char(wantFn);
                blob = readBlobBin(fullfile(here, 'data', fn));
                % Seed the header with the oracle's own px/dpi/content_type parse
                % (the P7-1a format parsers are stubbed; V-2 re-proves the parse).
                hdr  = s0083_DoubleHeader(wantPxW, wantPxH, wantHDpi, wantVDpi, ...
                                          wantCT, wantExt);
                img  = mat2doc.image.Image(blob, wantFn, hdr);

                testCase.verifyEqual(string(img.content_type), string(wantCT), ...
                    sprintf('%s content_type', tag));
                testCase.verifyEqual(double(img.px_width),  wantPxW,  sprintf('%s px_width', tag));
                testCase.verifyEqual(double(img.px_height), wantPxH,  sprintf('%s px_height', tag));
                testCase.verifyEqual(double(img.horz_dpi),  wantHDpi, sprintf('%s horz_dpi', tag));
                testCase.verifyEqual(double(img.vert_dpi),  wantVDpi, sprintf('%s vert_dpi', tag));
                % Genuine Length math: Inches(px/dpi) truncated via int() -> EMU.
                testCase.verifyEqual(double(img.width),  wantWEmu, sprintf('%s width EMU', tag));
                testCase.verifyEqual(double(img.height), wantHEmu, sprintf('%s height EMU', tag));
                testCase.verifyEqual(string(img.ext),      string(wantExt), sprintf('%s ext', tag));
                testCase.verifyEqual(string(img.filename), string(wantFn),  sprintf('%s filename', tag));
                % ★ sha1 END-TO-END (hashes the real blob -- no parser needed).
                testCase.verifyEqual(string(img.sha1), string(wantSha1), ...
                    sprintf('%s sha1 END-TO-END must equal the frozen oracle', tag));
            end
        end

        % =============================================================== %
        % ★ 2. splitext_ext (F-1) regression pin -- THE F-1 GUARD          %
        % =============================================================== %

        function test_splitext_ext_f1_regression_pin(testCase)
            % ★★ F-1 REGRESSION GUARD (LOUD): Image.ext == CPython
            % os.path.splitext(fn)[1][1:]. Gate-2 Fable found MATLAB's fileparts
            % treats ".bashrc" as an extension; the F-1 fix routes Image.ext through
            % the private splitext_ext replicating CPython exactly -- leading-dot
            % dotfiles yield "", the split is basename-only (dir dots ignored), case
            % is PRESERVED (IMG.PNG->PNG), and both / and \ separators are honored.
            % A future re-port that re-breaks ANY of these vectors goes RED HERE.
            % (validate_P7-1a section 3, 23/23 value-identical.)
            for i = 1:size(testCase.SPL_VECTORS, 1)
                fn   = testCase.SPL_VECTORS{i, 1};
                want = testCase.SPL_VECTORS{i, 2};
                img  = mat2doc.image.Image(uint8([]), fn, ...
                    s0083_DoubleHeader(1, 1, 1, 1, "image/png", "png"));
                testCase.verifyEqual(string(img.ext), string(want), ...
                    sprintf('F-1 splitext_ext: ext("%s") must be "%s"', fn, want));
            end
        end

        % =============================================================== %
        % 3. StreamReader (nominal + edge + error path)                   %
        % =============================================================== %

        function test_streamreader_reads_be_le(testCase)
            % Nominal + edge: BE/LE fixed-width reads at offsets, read_str ASCII +
            % UTF-8, base_offset shift, and unknown byte-order token -> BE fallback.
            % pat = 12 34 56 78 9A BC DE F0 00 FF 41 42 43 44 + utf8("é€").
            pat = uint8([18 52 86 120 154 188 222 240 0 255 65 66 67 68 195 169 226 130 172]);
            be  = mat2doc.image.StreamReader(mat2doc.image.BytesIO(pat), ">");
            le  = mat2doc.image.StreamReader(mat2doc.image.BytesIO(pat), "<");
            unk = mat2doc.image.StreamReader(mat2doc.image.BytesIO(pat), "X");   % -> BE
            bo  = mat2doc.image.StreamReader(mat2doc.image.BytesIO(pat), ">", 2);

            testCase.verifyEqual(be.read_byte(0), 18);
            testCase.verifyEqual(be.read_byte(1), 52);
            testCase.verifyEqual(be.read_byte(0, 1), 52);            % base+offset
            testCase.verifyEqual(be.read_short(0), 4660);            % 0x1234 BE
            testCase.verifyEqual(le.read_short(0), 13330);           % 0x3412 LE
            testCase.verifyEqual(be.read_long(0), 305419896);        % 0x12345678 BE
            testCase.verifyEqual(le.read_long(0), 2018915346);       % 0x78563412 LE
            testCase.verifyEqual(be.read_short(8), 255);             % 0x00FF
            testCase.verifyEqual(be.read_str(4, 10), "ABCD");        % ASCII
            testCase.verifyEqual(be.read_str(5, 14), "é€");          % UTF-8 2B+3B
            testCase.verifyEqual(unk.read_short(0), 4660);           % unknown -> BE
            testCase.verifyEqual(bo.read_byte(0), 86);               % base_offset=2 -> byte 2
            bo.seek(0);
            testCase.verifyEqual(bo.tell(), 2);                      % tell reflects base_offset
            % pass-through read(n)
            r = mat2doc.image.StreamReader(mat2doc.image.BytesIO(pat), ">").read(4);
            testCase.verifyEqual(r, uint8([18 52 86 120]));
        end

        function test_streamreader_eof_raises(testCase)
            % Error path (EOF): a fixed-width read past EOF raises
            % mat2doc:UnexpectedEndOfFileError (identifier pinned -- the re-home
            % trap: the class name IS the Python exception name).
            sr = mat2doc.image.StreamReader(mat2doc.image.BytesIO(uint8([1 2])), ">");
            caught = [];
            try
                sr.read_long(0);          % needs 4 bytes, only 2 available
            catch ME
                caught = ME;
            end
            testCase.assertNotEmpty(caught, 'read past EOF must raise');
            testCase.verifyEqual(caught.identifier, 'mat2doc:UnexpectedEndOfFileError', ...
                'EOF must raise mat2doc:UnexpectedEndOfFileError (== Python class)');
        end

        % =============================================================== %
        % 4. BytesIO (edge, offset-indexing)                              %
        % =============================================================== %

        function test_bytesio_offset_indexing(testCase)
            % Edge: read(n) / read-all / seek+read / tell / read-past-EOF -> empty.
            % 0-based cursor, byte results as 1xN uint8 (empty 1x0 at EOF).
            pat = uint8([18 52 86 120 154 188 222 240 0 255 65 66 67 68 195 169 226 130 172]);
            bio = mat2doc.image.BytesIO(pat);
            testCase.verifyEqual(bio.read(4), uint8([18 52 86 120]));       % first 4
            testCase.verifyEqual(bio.read(), pat(5:end));                   % read-all remaining
            bio2 = mat2doc.image.BytesIO(pat);
            bio2.seek(10);
            testCase.verifyEqual(bio2.read(4), uint8([65 66 67 68]));       % seek+read "ABCD"
            testCase.verifyEqual(bio2.tell(), 14);                         % 0-based cursor
            bio3 = mat2doc.image.BytesIO(pat);
            bio3.seek(999);
            testCase.verifyEqual(numel(bio3.read(4)), 0);                  % past EOF -> empty
            testCase.verifyEqual(numel(bio3.read()), 0);                   % read-all at EOF -> empty
            % short read (2 of 10 requested) returns the 2 available bytes
            testCase.verifyEqual(mat2doc.image.BytesIO(uint8([170 187])).read(10), uint8([170 187]));
        end

        % =============================================================== %
        % 5. sha1_hexdigest (regression) -- Java signed-byte boundary     %
        % =============================================================== %

        function test_sha1_hexdigest_known_blobs(testCase)
            % Regression: sha1_hexdigest == hashlib.sha1(...).hexdigest(). The
            % bytes-0..255 vector exercises the Java signed-byte boundary (values
            % >127 would saturate without the audited bytesToJava converter).
            sha = @(b) mat2doc.opc.sha1_hexdigest(b);
            testCase.verifyEqual(string(sha(uint8([]))), ...
                "da39a3ee5e6b4b0d3255bfef95601890afd80709");           % sha1("")
            testCase.verifyEqual(string(sha(uint8('abc'))), ...
                "a9993e364706816aba3e25717850c26c9cd0d89d");           % sha1("abc")
            testCase.verifyEqual(string(sha(uint8(0:255))), ...
                "4916d6bdb7f78e6803698cab32d1586ea457dfc8");           % bytes 0..255 (>127)
            testCase.verifyEqual(string(sha(uint8([104 195 169 108 108 111]))), ...
                "35b5ea45c5e41f78b46a937cc74d41dfea920890");           % "héllo" utf-8
        end

        % =============================================================== %
        % 6. Constants tables (regression) -- spot + count pins           %
        % =============================================================== %

        function test_constants_mime_png_chunk(testCase)
            % Regression: MIME_TYPE (5) + PNG_CHUNK_TYPE (pHYs mixed case H15).
            MIME = mat2doc.image.MIME_TYPE;
            testCase.verifyEqual(string(MIME.BMP),  "image/bmp");
            testCase.verifyEqual(string(MIME.GIF),  "image/gif");
            testCase.verifyEqual(string(MIME.JPEG), "image/jpeg");
            testCase.verifyEqual(string(MIME.PNG),  "image/png");
            testCase.verifyEqual(string(MIME.TIFF), "image/tiff");
            PC = mat2doc.image.PNG_CHUNK_TYPE;
            testCase.verifyEqual(string(PC.IHDR), "IHDR");
            testCase.verifyEqual(string(PC.pHYs), "pHYs");   % H15 mixed case preserved
            testCase.verifyEqual(string(PC.IEND), "IEND");
        end

        function test_constants_tiff_fld_tag(testCase)
            % Regression: TIFF_FLD (5 values + field_type_name 1-5 + KeyError) and
            % TIFF_TAG (5 codes + tag_name incl. 0xC4A5->PrintImageMatching +
            % KeyError). KeyError is the Python-dict-index error identifier.
            FLD = mat2doc.image.TIFF_FLD;
            testCase.verifyEqual(double(FLD.BYTE), 1);
            testCase.verifyEqual(double(FLD.ASCII), 2);
            testCase.verifyEqual(double(FLD.SHORT), 3);
            testCase.verifyEqual(double(FLD.LONG), 4);
            testCase.verifyEqual(double(FLD.RATIONAL), 5);
            testCase.verifyEqual(string(FLD.field_type_name(1)), "BYTE");
            testCase.verifyEqual(string(FLD.field_type_name(2)), "ASCII char");
            testCase.verifyEqual(string(FLD.field_type_name(5)), "RATIONAL");
            verifyKeyError(testCase, @() mat2doc.image.TIFF_FLD.field_type_name(99));

            TAG = mat2doc.image.TIFF_TAG;
            testCase.verifyEqual(double(TAG.IMAGE_WIDTH), 256);
            testCase.verifyEqual(double(TAG.IMAGE_LENGTH), 257);
            testCase.verifyEqual(double(TAG.X_RESOLUTION), 282);
            testCase.verifyEqual(double(TAG.Y_RESOLUTION), 283);
            testCase.verifyEqual(double(TAG.RESOLUTION_UNIT), 296);
            testCase.verifyEqual(string(TAG.tag_name(256)),   "ImageWidth");
            testCase.verifyEqual(string(TAG.tag_name(34665)), "ExifTag");
            testCase.verifyEqual(string(TAG.tag_name(50341)), "PrintImageMatching");  % 0xC4A5
            verifyKeyError(testCase, @() mat2doc.image.TIFF_TAG.tag_name(39321));      % 0x9999
        end

        function test_constants_jpeg_marker_code(testCase)
            % Regression: JPEG_MARKER_CODE -- all 49 single-byte marker-code
            % constants (uint8 -> int value), the STANDALONE (11) / SOF (13)
            % tuples, is_standalone, and marker_name (incl. APP13/APP14 + KeyError).
            % Representative spot checks + full-table + COUNT pin.
            JMC = mat2doc.image.JPEG_MARKER_CODE;
            % Representative marker values (constants.py).
            testCase.verifyEqual(double(JMC.TEM),  1);
            testCase.verifyEqual(double(JMC.SOF0), 192);
            testCase.verifyEqual(double(JMC.SOF2), 194);
            testCase.verifyEqual(double(JMC.DHT),  196);
            testCase.verifyEqual(double(JMC.SOI),  216);
            testCase.verifyEqual(double(JMC.EOI),  217);
            testCase.verifyEqual(double(JMC.SOS),  218);
            testCase.verifyEqual(double(JMC.APP0), 224);
            testCase.verifyEqual(double(JMC.APPD), 237);   % APP13 code
            testCase.verifyEqual(double(JMC.APPE), 238);   % APP14 code
            testCase.verifyEqual(double(JMC.APPF), 239);
            % COUNT pins: the tuples must hold exactly 11 / 13 entries.
            testCase.verifyEqual(numel(JMC.STANDALONE_MARKERS), 11, ...
                'STANDALONE_MARKERS must hold 11 codes');
            testCase.verifyEqual(numel(JMC.SOF_MARKER_CODES), 13, ...
                'SOF_MARKER_CODES must hold 13 codes');
            testCase.verifyEqual(double(JMC.STANDALONE_MARKERS), ...
                [1 216 217 208 209 210 211 212 213 214 215]);
            testCase.verifyEqual(double(JMC.SOF_MARKER_CODES), ...
                [192 193 194 195 197 198 199 201 202 203 205 206 207]);
            % is_standalone
            testCase.verifyTrue(JMC.is_standalone(uint8(216)));   % SOI standalone
            testCase.verifyFalse(JMC.is_standalone(uint8(224)));  % APP0 not standalone
            % marker_name
            testCase.verifyEqual(string(JMC.marker_name(uint8(0))),   "UNKNOWN");
            testCase.verifyEqual(string(JMC.marker_name(uint8(192))), "SOF0");
            testCase.verifyEqual(string(JMC.marker_name(uint8(237))), "APP13");
            testCase.verifyEqual(string(JMC.marker_name(uint8(238))), "APP14");
            verifyKeyError(testCase, @() mat2doc.image.JPEG_MARKER_CODE.marker_name(uint8(119)));

            % ALL 44 marker values, value-by-value (the full constants.py table).
            names = ["TEM" "DHT" "DAC" "JPG" ...
                "SOF0" "SOF1" "SOF2" "SOF3" "SOF5" "SOF6" "SOF7" "SOF9" ...
                "SOFA" "SOFB" "SOFD" "SOFE" "SOFF" ...
                "RST0" "RST1" "RST2" "RST3" "RST4" "RST5" "RST6" "RST7" ...
                "SOI" "EOI" "SOS" "DQT" "DNL" "DRI" "DHP" "EXP" ...
                "APP0" "APP1" "APP2" "APP3" "APP4" "APP5" "APP6" "APP7" ...
                "APP8" "APP9" "APPA" "APPB" "APPC" "APPD" "APPE" "APPF"];
            vals = [1 196 204 200 ...
                192 193 194 195 197 198 199 201 ...
                202 203 205 206 207 ...
                208 209 210 211 212 213 214 215 ...
                216 217 218 219 220 221 222 223 ...
                224 225 226 227 228 229 230 231 ...
                232 233 234 235 236 237 238 239];
            % COUNT pin: JPEG_MARKER_CODE defines 49 single-byte marker-code
            % constants (the full python-docx constants.py table: 4 misc + 13 SOF
            % + 8 RST + 8 misc2 + 16 APP). (validate_P7-1a section 4 rounds this to
            % "the 44 markers" informally; the exact class surface is 49 and every
            % value is checked below.)
            testCase.verifyEqual(numel(names), 49, 'the JPEG marker table has 49 codes');
            for i = 1:numel(names)
                testCase.verifyEqual(double(JMC.(names(i))), vals(i), ...
                    sprintf('JMC.%s value', names(i)));
            end
        end

        % =============================================================== %
        % 7. scaled_dimensions + pyRound (banker's half-to-even)          %
        % =============================================================== %

        function test_scaled_dimensions_native_and_given(testCase)
            % Nominal + edge: native / both-given / one-given scaled_dimensions on a
            % seeded 200x100 px @ 100 dpi Image (native 1828800 x 914400 EMU).
            img = seededImage();
            [w, h] = img.scaled_dimensions();
            testCase.verifyEqual(double(w), 1828800);   % native width
            testCase.verifyEqual(double(h), 914400);    % native height
            [w2, h2] = img.scaled_dimensions(500000, 300000);
            testCase.verifyEqual(double(w2), 500000);   % both given -> as given
            testCase.verifyEqual(double(h2), 300000);
            [w3, h3] = img.scaled_dimensions(457200);   % width-only -> aspect-kept
            testCase.verifyEqual(double(w3), 457200);
            testCase.verifyEqual(double(h3), 228600);
            [w4, h4] = img.scaled_dimensions([], 457200);   % height-only
            testCase.verifyEqual(double(w4), 914400);
            testCase.verifyEqual(double(h4), 457200);
        end

        function test_pyround_banker_half_to_even(testCase)
            % ★ Edge: pyRound is half-to-EVEN (banker's), NOT MATLAB's half-away
            % round. Driven through the real scaled_dimensions API (pyRound is
            % private): a native 200x100 @ 100 dpi image scaled to width w has
            % height = pyRound(w/2), so w=1,3,5,7 -> h=0,2,2,4 and negative widths
            % w=-3,-5 -> h=-2,-2 (all ties resolve to the EVEN integer).
            % round(0.5)=0, round(1.5)=2, round(2.5)=2, round(3.5)=4, round(-1.5)=-2.
            img = seededImage();
            ties = { 1, 0; 3, 2; 5, 2; 7, 4; -3, -2; -5, -2 };  % {width, expected height}
            for i = 1:size(ties, 1)
                w = ties{i, 1};
                wantH = ties{i, 2};
                [~, h] = img.scaled_dimensions(w);
                testCase.verifyEqual(double(h), wantH, ...
                    sprintf('scaled height for width %d must be %d (banker''s round)', w, wantH));
            end
        end

        % =============================================================== %
        % ★ 8. WMF/EMF-exclusion (regression trap guard, LOUD)            %
        % =============================================================== %

        function test_wmf_emf_exclusion_unrecognized(testCase)
            % ★★ RE-PORT TRAP GUARD (LOUD): docx has NO WMF/EMF parser (unlike the
            % Mat2Ppt PIL seam), so a WMF placeable signature (D7 CD C6 9A) and an
            % EMF signature (01 00 00 00) are UNRECOGNIZED ->
            % mat2doc:UnrecognizedImageError. A future sloppy re-port that re-adds a
            % WMF/EMF parser (the pptx behaviour) MUST go RED here.
            % (validate_P7-1a section 6.)
            wmf = uint8([215 205 198 154, zeros(1, 60)]);   % D7 CD C6 9A + pad
            emf = uint8([1 0 0 0, zeros(1, 60)]);           % 01 00 00 00 + pad
            verifyUnrecognized(testCase, @() mat2doc.image.Image.from_blob(wmf), 'WMF');
            verifyUnrecognized(testCase, @() mat2doc.image.Image.from_blob(emf), 'EMF');
            % plus the other unrecognized termini: garbage / empty / short header.
            verifyUnrecognized(testCase, @() mat2doc.image.Image.from_blob( ...
                uint8('not an image at all, just text bytes here 0123456789')), 'garbage');
            verifyUnrecognized(testCase, @() mat2doc.image.Image.from_blob(uint8([])), 'empty');
            verifyUnrecognized(testCase, @() mat2doc.image.Image.from_blob(uint8([137 80])), 'short');
        end

        % =============================================================== %
        % 9. Factory dispatch -> stubbed parsers (P7-1b/P7-2 boundary)    %
        % =============================================================== %

        function test_factory_dispatch_to_stubs(testCase)
            % Edge: the 8 SIGNATURES rows dispatch a recognized signature to its
            % (stubbed) format parser, each raising mat2doc:notYetPorted (the parser
            % is un-ported until P7-1b/P7-2). Order is behavior (first match wins);
            % the signature bytes are laid at each row's offset.
            cases = { ...
                'PNG',   [137 80 78 71 13 10 26 10, zeros(1, 24)]; ...
                'GIF87', [double('GIF87a'), zeros(1, 26)]; ...
                'GIF89', [double('GIF89a'), zeros(1, 26)]; ...
                'BMP',   [double('BM'), zeros(1, 30)]; ...
                'TIFF_MM', [77 77 0 42, zeros(1, 28)]; ...
                'TIFF_II', [73 73 42 0, zeros(1, 28)]; ...
                'JFIF',  [255 216 255 224 0 16, double('JFIF'), 0, zeros(1, 21)]; ...
                'Exif',  [255 216 255 225 0 16, double('Exif'), 0, zeros(1, 21)]};
            for k = 1:size(cases, 1)
                lbl  = cases{k, 1};
                blob = uint8(cases{k, 2});
                caught = [];
                try
                    mat2doc.image.Image.from_blob(blob);
                catch ME
                    caught = ME;
                end
                testCase.assertNotEmpty(caught, ...
                    sprintf('%s signature must dispatch (to a stub)', lbl));
                testCase.verifyEqual(caught.identifier, 'mat2doc:notYetPorted', ...
                    sprintf(['%s signature must dispatch to a notYetPorted parser ' ...
                    'stub (P7-1b/P7-2 boundary)'], lbl));
            end
        end

        function test_parser_stubs_raise_notyetported_directly(testCase)
            % Edge: each of the 6 format-parser stubs raises mat2doc:notYetPorted
            % when invoked directly (the P7-1b/P7-2 owning-WP markers).
            stream = mat2doc.image.BytesIO(uint8([0 0 0 0]));
            stubs = { ...
                @() mat2doc.image.Png.from_stream(stream); ...
                @() mat2doc.image.Jfif.from_stream(stream); ...
                @() mat2doc.image.Exif.from_stream(stream); ...
                @() mat2doc.image.Gif.from_stream(stream); ...
                @() mat2doc.image.Tiff.from_stream(stream); ...
                @() mat2doc.image.Bmp.from_stream(stream)};
            for k = 1:numel(stubs)
                caught = [];
                try
                    stubs{k}();
                catch ME
                    caught = ME;
                end
                testCase.assertNotEmpty(caught, 'parser stub must raise');
                testCase.verifyEqual(caught.identifier, 'mat2doc:notYetPorted', ...
                    'each format-parser stub must raise mat2doc:notYetPorted');
            end
        end

        % =============================================================== %
        % 10. NIE abstract members (error path, verbatim message)         %
        % =============================================================== %

        function test_baseimageheader_abstract_members_raise(testCase)
            % Error path: BaseImageHeader.content_type / default_ext on a plain base
            % -> mat2doc:NotImplementedError with the upstream message VERBATIM
            % (cross-faithful text -- these are faithful, not port-authored).
            base = mat2doc.image.BaseImageHeader(1, 1, 1, 1);
            checkRaises(testCase, @() base.content_type, ...
                'mat2doc:NotImplementedError', ...
                'content_type property must be implemented by all subclasses of BaseImageHeader');
            checkRaises(testCase, @() base.default_ext, ...
                'mat2doc:NotImplementedError', ...
                'default_ext property must be implemented by all subclasses of BaseImageHeader');
        end

        % =============================================================== %
        % ★ 11. Full frozen-battery equivalence replay (self-contained)   %
        % =============================================================== %

        function test_replay_all_195_records(testCase)
            % ★ Equivalence (headline): replay the ENTIRE frozen Gate-3 battery
            % (the exact 195-probe sequence of s0083_..._probe.m, rebuilt inline
            % through mat2doc.image.* / mat2doc.opc.*) and assert the port
            % reproduces every record's tagged canon value-identical to the frozen
            % python-docx 1.2.0 oracle. Fixture is references\s0083\probe.json copied
            % verbatim into tests\image\data\probe.json (self-contained). The tagged
            % canon (OK|int|, OK|str|, OK|hex|, OK|bool|, err|<Name>[|<msg>]) is the
            % same cross-language type-lattice-proof scheme the Gate-3 comparator
            % used (validate_P7-1a section 1).
            here = fileparts(mfilename('fullpath'));
            refPath = fullfile(here, 'data', 'probe.json');
            testCase.assertTrue(isfile(refPath), ...
                sprintf('frozen probe fixture missing: %s', refPath));
            % UTF-8 read (the frozen oracle is UTF-8, ensure_ascii=False -- the
            % SR_be_str_utf8 record carries "é€"; a locale-default fileread on
            % Windows would corrupt it and false-fail the equivalence).
            ref = jsondecode(readTextUtf8(refPath));

            obs = buildObservedBattery(here);

            keys = fieldnames(ref);
            testCase.verifyEqual(numel(keys), 195, ...
                'frozen s0083 battery must hold 195 records');
            testCase.verifyEqual(double(obs.Count), 195, ...   % Map.Count is uint64
                'rebuilt battery must reproduce exactly 195 records');
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
        % ★ 12. M1 byte-pin (image core is M1-neutral)                    %
        % =============================================================== %

        function test_m1_neutral_styles_document_bytes(testCase)
            % ★ Regression: mat2doc.Document().save() emits word/styles.xml and
            % word/document.xml byte-identical to the frozen s0001 oracle (SHA-256
            % == byte-identity, L1). Image parsing touches no oxml registry row, no
            % PartFactory registration, nothing on the save path -- so M1 must stay
            % green. RED here would mean the P7-1a re-port perturbed the save path.
            % (validate_P7-1a section 8.)
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
% Local functions in a classdef file are visible to the class methods above.

function b = readBlobBin(path)
    % Binary read (no CRLF translation): MATLAB 'r' is binary by default.
    fid = fopen(char(path), 'r', 'n');
    assert(fid >= 0, 'could not open for read: %s', char(path));
    c = onCleanup(@() fclose(fid));
    b = fread(fid, inf, '*uint8')';
end

function txt = readTextUtf8(path)
    % Explicit UTF-8 text read (the frozen probe.json is UTF-8, ensure_ascii=False).
    fid = fopen(char(path), 'r', 'n', 'UTF-8');
    assert(fid >= 0, 'could not open for read: %s', char(path));
    c = onCleanup(@() fclose(fid));
    txt = fread(fid, inf, '*char')';
end

function img = seededImage()
    % A 200x100 px @ 100 dpi seeded Image (native 1828800 x 914400 EMU) -- the
    % scaled_dimensions / pyRound driver (validate_P7-1a section 5).
    hdr = s0083_DoubleHeader(200, 100, 100, 100, "image/png", "png");
    img = mat2doc.image.Image(uint8('seedblob'), "seed.png", hdr);
end

function verifyKeyError(testCase, fn)
    % A Python-dict-index miss -> mat2doc:KeyError (identifier pinned).
    caught = [];
    try
        fn();
    catch ME
        caught = ME;
    end
    testCase.assertNotEmpty(caught, 'expected a KeyError but none was raised');
    testCase.verifyEqual(caught.identifier, 'mat2doc:KeyError', ...
        'a dict-index miss must raise mat2doc:KeyError');
end

function verifyUnrecognized(testCase, fn, label)
    % An unrecognized image blob -> mat2doc:UnrecognizedImageError (identifier).
    caught = [];
    try
        fn();
    catch ME
        caught = ME;
    end
    testCase.assertNotEmpty(caught, sprintf('%s must raise', label));
    testCase.verifyEqual(caught.identifier, 'mat2doc:UnrecognizedImageError', ...
        sprintf('%s must raise mat2doc:UnrecognizedImageError', label));
end

function checkRaises(testCase, fn, expId, expMsg)
    % Assert fn() raises; verify the identifier and the byte-exact message.
    caught = [];
    try
        fn();
    catch ME
        caught = ME;
    end
    testCase.assertNotEmpty(caught, 'expected an error but none was raised');
    testCase.verifyEqual(caught.identifier, expId, ...
        sprintf('expected identifier %s, got %s', expId, caught.identifier));
    testCase.verifyEqual(caught.message, char(expMsg), ...
        'error message must be byte-exact');
end

% ------------------------------------------------------------ M1 zip helpers
function [blobs, names] = zipEntryList(zipBytes)
    % Enumerate an in-memory zip in stream (write) order -> ordered {names, blobs}.
    % (Copied from tests\api\Test_p1_8_skeleton_m1.m so the order pin is
    % independent of the reader under test.)
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
% s0083_p7_1a_image_core_probe.m (same probe order, same P/PC/canon canon
% scheme), rebuilt through mat2doc.image.* / mat2doc.opc.* into a
% containers.Map keyed by probe name -> tagged canon string. The two test
% images are read from tests\image\data\ (self-contained; the scenario read
% them from the READ-ONLY python-docx clone -- byte-identical, sha1-verified).

function obs = buildObservedBattery(dataRoot)
    obs = containers.Map('KeyType', 'char', 'ValueType', 'char');

    PNG_PATH = fullfile(dataRoot, 'data', '300-dpi.png');
    JPG_PATH = fullfile(dataRoot, 'data', '300-dpi.jpg');

    % ===================================================== IMG  real-image core
    pngBlob = readBlobBin(PNG_PATH);
    jpgBlob = readBlobBin(JPG_PATH);
    pngHdr = s0083_DoubleHeader(860, 579, 300, 300, "image/png", "png");
    jpgHdr = s0083_DoubleHeader(1504, 1936, 300, 300, "image/jpeg", "jpg");
    pngImg = mat2doc.image.Image(pngBlob, "300-dpi.png", pngHdr);
    jpgImg = mat2doc.image.Image(jpgBlob, "300-dpi.jpg", jpgHdr);
    imgs = {"PNG", pngImg; "JPG", jpgImg};
    for r = 1:size(imgs, 1)
        tag = imgs{r, 1};
        img = imgs{r, 2};
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

    % ===================================================== SHA  sha1 corner cases
    sha = @(b) mat2doc.opc.sha1_hexdigest(b);
    put(obs, "SHA_empty",      P(@() sha(uint8([]))));
    put(obs, "SHA_abc",        P(@() sha(uint8('abc'))));
    put(obs, "SHA_0to255",     P(@() sha(uint8(0:255))));
    put(obs, "SHA_hello_utf8", P(@() sha(uint8([104 195 169 108 108 111]))));

    % ===================================================== SPL  splitext_ext (F-1)
    spl_vectors = [".bashrc", ".png", "..png", ".myimage", "...", "a.png", ...
        "IMG.PNG", "a..png", ".tar.gz", "archive.tar.gz", "file.", ...
        "trailing.dot.", "image", "noext", "a.b.c.jpeg", "image.J p g", ...
        "/x/y.z/a.png", "/x/y.z/noext", "300-dpi.png", "300-dpi.jpg", ...
        "C:\dir.x\img.png", "C:\dir.x\noext", ""];
    for i = 1:numel(spl_vectors)
        put(obs, sprintf("SPL_%02d", i), P(@() extOf(spl_vectors(i))));
    end

    % ===================================================== SR  StreamReader
    pat = uint8([18 52 86 120 154 188 222 240 0 255 65 66 67 68 195 169 226 130 172]);
    be  = mat2doc.image.StreamReader(mat2doc.image.BytesIO(pat), ">");
    le  = mat2doc.image.StreamReader(mat2doc.image.BytesIO(pat), "<");
    unk = mat2doc.image.StreamReader(mat2doc.image.BytesIO(pat), "X");
    bo  = mat2doc.image.StreamReader(mat2doc.image.BytesIO(pat), ">", 2);
    put(obs, "SR_be_byte0",    P(@() be.read_byte(0)));
    put(obs, "SR_be_byte1",    P(@() be.read_byte(1)));
    put(obs, "SR_be_byte_off", P(@() be.read_byte(0, 1)));
    put(obs, "SR_be_short0",   P(@() be.read_short(0)));
    put(obs, "SR_le_short0",   P(@() le.read_short(0)));
    put(obs, "SR_be_long0",    P(@() be.read_long(0)));
    put(obs, "SR_le_long0",    P(@() le.read_long(0)));
    put(obs, "SR_be_short8",   P(@() be.read_short(8)));
    put(obs, "SR_be_str",      P(@() be.read_str(4, 10)));
    put(obs, "SR_be_str_utf8", P(@() be.read_str(5, 14)));
    put(obs, "SR_unk_short0",  P(@() unk.read_short(0)));
    put(obs, "SR_bo_byte0",    P(@() bo.read_byte(0)));
    put(obs, "SR_bo_tell",     P(@() sr_bo_tell(bo)));
    put(obs, "SR_read_pass",   P(@() mat2doc.image.StreamReader(mat2doc.image.BytesIO(pat), ">").read(4)));
    put(obs, "SR_eof",         PC(@() mat2doc.image.StreamReader(mat2doc.image.BytesIO(uint8([1 2])), ">").read_long(0)));

    % ===================================================== BIO  BytesIO analogue
    bio  = mat2doc.image.BytesIO(pat);
    put(obs, "BIO_read4",     P(@() bio.read(4)));
    put(obs, "BIO_read_rest", P(@() bio.read()));
    bio2 = mat2doc.image.BytesIO(pat);
    put(obs, "BIO_seek_read", P(@() bio_seek_read(bio2)));
    put(obs, "BIO_tell",      P(@() bio2.tell()));
    bio3 = mat2doc.image.BytesIO(pat);
    put(obs, "BIO_read_past",      P(@() bio_read_past(bio3)));
    put(obs, "BIO_read_all_empty", P(@() bio_read_all_empty(bio3)));
    put(obs, "BIO_short_read",     P(@() mat2doc.image.BytesIO(uint8([170 187])).read(10)));

    % ===================================================== K  constants tables
    MIME = mat2doc.image.MIME_TYPE;
    put(obs, "K_MIME_BMP",  P(@() MIME.BMP));
    put(obs, "K_MIME_GIF",  P(@() MIME.GIF));
    put(obs, "K_MIME_JPEG", P(@() MIME.JPEG));
    put(obs, "K_MIME_PNG",  P(@() MIME.PNG));
    put(obs, "K_MIME_TIFF", P(@() MIME.TIFF));
    PCH = mat2doc.image.PNG_CHUNK_TYPE;
    put(obs, "K_PNGC_IHDR", P(@() PCH.IHDR));
    put(obs, "K_PNGC_pHYs", P(@() PCH.pHYs));
    put(obs, "K_PNGC_IEND", P(@() PCH.IEND));
    put(obs, "K_TFLD_alias",    P(@() "same"));
    put(obs, "K_TFLD_BYTE",     P(@() mat2doc.image.TIFF_FLD.BYTE));
    put(obs, "K_TFLD_ASCII",    P(@() mat2doc.image.TIFF_FLD.ASCII));
    put(obs, "K_TFLD_SHORT",    P(@() mat2doc.image.TIFF_FLD.SHORT));
    put(obs, "K_TFLD_LONG",     P(@() mat2doc.image.TIFF_FLD.LONG));
    put(obs, "K_TFLD_RATIONAL", P(@() mat2doc.image.TIFF_FLD.RATIONAL));
    for code = 1:5
        put(obs, sprintf("K_TFLDname_%d", code), P(@() mat2doc.image.TIFF_FLD.field_type_name(code)));
    end
    put(obs, "K_TFLDname_bad", PC(@() mat2doc.image.TIFF_FLD.field_type_name(99)));
    TTAG = mat2doc.image.TIFF_TAG;
    put(obs, "K_TTAG_IMAGE_WIDTH",     P(@() TTAG.IMAGE_WIDTH));
    put(obs, "K_TTAG_IMAGE_LENGTH",    P(@() TTAG.IMAGE_LENGTH));
    put(obs, "K_TTAG_X_RESOLUTION",    P(@() TTAG.X_RESOLUTION));
    put(obs, "K_TTAG_Y_RESOLUTION",    P(@() TTAG.Y_RESOLUTION));
    put(obs, "K_TTAG_RESOLUTION_UNIT", P(@() TTAG.RESOLUTION_UNIT));
    for code = [254 256 282 34665 34853 50341]
        put(obs, sprintf("K_TTAGname_%d", code), P(@() mat2doc.image.TIFF_TAG.tag_name(code)));
    end
    put(obs, "K_TTAGname_bad", PC(@() mat2doc.image.TIFF_TAG.tag_name(39321)));
    JMC = mat2doc.image.JPEG_MARKER_CODE;
    jmc_names = ["TEM" "DHT" "DAC" "JPG" ...
        "SOF0" "SOF1" "SOF2" "SOF3" "SOF5" "SOF6" "SOF7" "SOF9" ...
        "SOFA" "SOFB" "SOFD" "SOFE" "SOFF" ...
        "RST0" "RST1" "RST2" "RST3" "RST4" "RST5" "RST6" "RST7" ...
        "SOI" "EOI" "SOS" "DQT" "DNL" "DRI" "DHP" "EXP" ...
        "APP0" "APP1" "APP2" "APP3" "APP4" "APP5" "APP6" "APP7" ...
        "APP8" "APP9" "APPA" "APPB" "APPC" "APPD" "APPE" "APPF"];
    for i = 1:numel(jmc_names)
        nm = jmc_names(i);
        put(obs, "K_JMC_" + nm, P(@() double(JMC.(nm))));
    end
    put(obs, "K_JMC_STANDALONE", P(@() joinInts(JMC.STANDALONE_MARKERS)));
    put(obs, "K_JMC_SOF",        P(@() joinInts(JMC.SOF_MARKER_CODES)));
    put(obs, "K_JMC_is_sa_SOI",  P(@() JMC.is_standalone(uint8(216))));
    put(obs, "K_JMC_is_sa_APP0", P(@() JMC.is_standalone(uint8(224))));
    jmc_codes = [0 192 194 196 216 217 218 219 224 225 226 237 238];
    for i = 1:numel(jmc_codes)
        code = jmc_codes(i);
        put(obs, sprintf("K_JMCname_%d", code), P(@() mat2doc.image.JPEG_MARKER_CODE.marker_name(uint8(code))));
    end
    put(obs, "K_JMCname_bad", PC(@() mat2doc.image.JPEG_MARKER_CODE.marker_name(uint8(119))));

    % ===================================================== SD  scaled_dimensions + pyRound
    seedImg = seededImage();
    put(obs, "SD_native_w", P(@() sd_w(seedImg)));
    put(obs, "SD_native_h", P(@() sd_h(seedImg)));
    put(obs, "SD_both_w",   P(@() sd_bw(seedImg, 500000, 300000)));
    put(obs, "SD_both_h",   P(@() sd_bh(seedImg, 500000, 300000)));
    put(obs, "SD_wonly_w",  P(@() sd_ww(seedImg, 457200)));
    put(obs, "SD_wonly_h",  P(@() sd_wh(seedImg, 457200)));
    put(obs, "SD_honly_w",  P(@() sd_hw(seedImg, 457200)));
    put(obs, "SD_honly_h",  P(@() sd_hh(seedImg, 457200)));
    tie_ws = [1 3 5 7 -3 -5];
    for i = 1:numel(tie_ws)
        w = tie_ws(i);
        if w < 0, lbl = "m" + string(-w); else, lbl = string(w); end
        put(obs, "SD_tie_" + lbl + "_w", P(@() sd_ww(seedImg, w)));
        put(obs, "SD_tie_" + lbl + "_h", P(@() sd_wh(seedImg, w)));
    end

    % ===================================================== FAC  factory / WMF excl.
    put(obs, "FAC_garbage", PC(@() mat2doc.image.Image.from_blob(uint8('not an image at all, just text bytes here 0123456789'))));
    put(obs, "FAC_wmf",     PC(@() mat2doc.image.Image.from_blob(uint8([215 205 198 154, zeros(1,60)]))));
    put(obs, "FAC_emf",     PC(@() mat2doc.image.Image.from_blob(uint8([1 0 0 0, zeros(1,60)]))));
    put(obs, "FAC_empty",   PC(@() mat2doc.image.Image.from_blob(uint8([]))));
    put(obs, "FAC_short",   PC(@() mat2doc.image.Image.from_blob(uint8([137 80]))));

    % ===================================================== NIE  abstract members
    base = mat2doc.image.BaseImageHeader(1, 1, 1, 1);
    put(obs, "NIE_content_type", P(@() base.content_type));
    put(obs, "NIE_default_ext",  P(@() base.default_ext));
end

function put(map, key, canonStr)
    map(char(key)) = char(canonStr);
end

function e = extOf(fn)
    img = mat2doc.image.Image(uint8([]), fn, s0083_DoubleHeader(1, 1, 1, 1, "image/png", "png"));
    e = img.ext;
end

function v = sr_bo_tell(bo),        bo.seek(0); v = bo.tell(); end
function v = bio_seek_read(bio),    bio.seek(10); v = bio.read(4); end
function v = bio_read_past(bio),    bio.seek(999); v = bio.read(4); end
function v = bio_read_all_empty(bio), bio.seek(999); v = bio.read(); end

function v = sd_w(img),  [c, ~] = img.scaled_dimensions();            v = c; end
function v = sd_h(img),  [~, c] = img.scaled_dimensions();            v = c; end
function v = sd_bw(img, w, h), [c, ~] = img.scaled_dimensions(w, h);  v = c; end
function v = sd_bh(img, w, h), [~, c] = img.scaled_dimensions(w, h);  v = c; end
function v = sd_ww(img, w), [c, ~] = img.scaled_dimensions(w);        v = c; end
function v = sd_wh(img, w), [~, c] = img.scaled_dimensions(w);        v = c; end
function v = sd_hw(img, h), [c, ~] = img.scaled_dimensions([], h);    v = c; end
function v = sd_hh(img, h), [~, c] = img.scaled_dimensions([], h);    v = c; end

function s = joinInts(vec)
    parts = arrayfun(@(x) string(double(x)), vec);
    s = strjoin(parts, "|");
end

% ------------------------------------------------------------ probe wrappers
% (verbatim from the Gate-3 scenario s0083_..._probe.m so the canon is identical)

function s = P(fn)
% Full canon; on error "err|<Name>|<message>".
try
    s = "OK|" + canon(fn());
catch e
    s = "err|" + errName(e) + "|" + string(e.message);
end
end

function s = PC(fn)
% Class-only error canon (port-authored-message family D-003/D-004).
try
    s = "OK|" + canon(fn());
catch e
    s = "err|" + errName(e);
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

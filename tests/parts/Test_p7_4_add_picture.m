classdef Test_p7_4_add_picture < matlab.unittest.TestCase
% TEST_P7_4_ADD_PICTURE  Gate-4 permanent unit tests for Mat2Doc P7-4 [N]
%   -- THE PICTURE MILESTONE: the FIRST runtime IMAGE part in Mat2Doc.
%   src/docx/parts/image.py::{ImagePart} -> +mat2doc\+parts\ImagePart;
%   src/docx/image/... reused via +mat2doc\+package\ImageParts (the package.py
%   ImageParts collection); src/docx/opc/package.py::OpcPackage.{image_parts,
%   get_or_add_image_part} -> +mat2doc\+package\Package; src/docx/parts/story.py::
%   BaseStoryPart.{get_or_add_image,new_pic_inline} -> +mat2doc\+parts\StoryPart;
%   src/docx/text/run.py::Run.add_picture; src/docx/document.py::Document.add_picture;
%   the PartFactory IMAGE->ImagePart flip; and (the P5 payoff) add_picture routed
%   into a header StoryPart.
%
%   P7-4 is the FIRST WP to emit a runtime IMAGE part: a picture becomes a
%   word/media/imageN.<ext> binary part (an EXACT byte copy of the source file)
%   with its own rId on the owning story part + a [Content_Types] Default for the
%   extension + a w:drawing/wp:inline (the P7-3 tree) in document.xml. Equivalence
%   is therefore FULL-PACKAGE byte identity: the media blob, document.xml, every
%   new rel, the Content_Types Default, the zip-entry order (media lands AFTER
%   word/numbering.xml, BEFORE docProps/thumbnail.jpeg). This class permanently
%   FREEZES that surface -- byte-identical to python-docx v1.2.0.
%
%   THE HIGHEST-VALUE PERMANENT PINS IN THIS WP (a regression goes RED loudly):
%     * ** THE P7 HEADLINE ** add_picture full-package byte pin
%       (test_add_picture_full_package_byte_pin): Document(); add_picture(<PNG>);
%       save() -> ALL 18 parts byte-identical to the frozen s0090 python-docx
%       oracle. word/media/image1.png == 6111 B / b60f7099 (the FIRST runtime
%       image part -- an EXACT source-file byte copy), word/document.xml == 2335 B
%       / d0559ef6 (the paragraph/run/w:drawing/wp:inline), document.xml.rels ==
%       dc5eaed8 (rId9 -> media/image1.png, Type .../image), [Content_Types].xml ==
%       1a1b7bf7 (Default png added), + the other 14 parts unchanged. The permanent
%       picture regression guard.
%     * ** dedupe ** (test_dedupe_full_package_byte_pin, s0091): the SAME PNG added
%       twice -> EXACTLY ONE word/media/image1.png; both w:drawings share rId9
%       (SHA1 dedupe in ImageParts + the rel dedupe in relate_to). 18/18.
%     * ** two images ** (test_two_images_full_package_byte_pin, s0092): PNG then
%       JPEG -> image1.png + image2.jpeg, two rels (rId9/rId10), two CT Defaults.
%       19/19.
%     * ** JPEG ** (s0093, 18/18, image1.jpeg ad5afb51) and ** width override **
%       (s0094, 18/18, wp:extent cx=1828800 cy=731520 via scaled_dimensions).
%     * ** C6 HEADER image ** (test_header_image_full_package_byte_pin, s0096): a
%       picture added into sections[0].header -> word/header1.xml (w:drawing) +
%       the FIRST-EVER word/_rels/header1.xml.rels (rId1 -> media/image1.png) + the
%       media part. 20/20. The P5-3b header StoryPart payoff.
%     * ** DEFECT-1 regression pin, scenario J ** (test_defect1_scenario_j...,
%       s0095): reopen a docx whose ONLY media part is a NON-numbered
%       /word/media/logo.png (PackURI.idx == []), then add a JPEG -> NO CRASH,
%       media = {logo.png (retained, contributes no number), image1.jpeg}, 19/19
%       byte-identical. Guards the ImageParts.next_image_partname_ idx-None-skip
%       fix; a regression that re-breaks the numbering crashes here (before the fix
%       Mat2Doc threw `used_numbers(end+1)=[]` singleSubscript). Loudly commented.
%     * ** M1 byte-neutrality ** (test_m1_byte_neutrality): Document().save() ->
%       styles.xml 02d71a68 + document.xml 0e4dd503 -- the LIVE _gather_image_parts
%       transit is byte-neutral on the empty default (its thumbnail is a THUMBNAIL
%       reltype, not RT.IMAGE, so the collection is never grown).
%
%   Provenance (all Gate-3 frozen 2026-08-02):
%     * Audit    : validation\mat2doc\audit_P7-4_add_picture.md (Porter Gate-1 +
%                  Fable/mso-auditor Gate-2 REVISE -> DEFECT-1 fixed + re-verified).
%     * Validate : validation\mat2doc\validate_P7-4_add_picture.md (Gate-3 PASS --
%                  ZERO new D-numbers; s0090 18/18 L1+bin; s0091 18/18 dedupe; s0092
%                  19/19; s0093 18/18; s0094 18/18; s0096 20/20; s0097 20/20; s0095
%                  (DEFECT-1 J) 19/19; probe_diff s0098 MATCH; M1 17/17 re-derived).
%     * Scenarios: validation\mat2doc\scenarios\s0090..s0098_p7_4_*.{py,m} (the
%                  IDENTICAL-sequence byte/probe twins; their build bodies are
%                  replayed VERBATIM by the tests below + runS0098Probe()).
%     * Frozen refs (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE at Gate-3):
%         references\s0090..s0097\ -- the full-package manifests (name|size|sha256
%           in frozen zip-entry order) embedded below as the S00xx_MANIFEST
%           constants so this suite is self-contained; SHA-256 equality IS byte-
%           identity (L1). references\s0001\ reused as the M1 byte guard.
%         tests\parts\data\python-powered.png -- the source PNG (byte-exact copy of
%           python-docx tests\test_files\python-powered.png, sha256 b60f7099 /
%           6111 B == the frozen word/media/image1.png). The add_picture INPUT.
%         tests\parts\data\python-icon.jpeg -- the source JPEG (byte-exact copy,
%           sha256 ad5afb51 / 3277 B).
%         tests\parts\data\s0095_input_logo.docx -- the python-generated docx whose
%           sole media part is the NON-numbered /word/media/logo.png (frozen
%           references\s0095\input_logo.docx), the DEFECT-1 scenario-J INPUT.
%         tests\parts\data\s0098\probe.json -- the frozen s0098 probe oracle (value
%           JSON; jsondecode line-ending agnostic -> no binary pin, s0057/s0089
%           precedent), for the Equivalence leg.
%         All frozen BINARY fixtures carry a co-located `.gitattributes` `* binary`
%           pin (the Gate-4 byte-fixture lesson) so a master checkout does not
%           line-ending-mangle them.
%
%   Coverage taxonomy
%   -----------------
%   * Nominal      -- the add_picture PNG G-scenario (the milestone); ImagePart
%                     from_image (sha1/filename/content_type/default_cx-cy/lazy
%                     image); ImageParts get_or_add_image_part/contains/to_array/
%                     len_; StoryPart get_or_add_image/new_pic_inline; Run/Document
%                     add_picture -> InlineShape; InlineShapes.part.
%   * Edge         -- dedupe (same image -> one media part, one rId); two distinct
%                     images (partname numbering image1/image2, two rels/CT Defaults);
%                     JPEG (a different extension); width override (scaled_dimensions
%                     wp:extent); the C6 header StoryPart route (first-ever header
%                     .rels); the DEFECT-1 non-numbered-partname reopen (scenario J);
%                     ImagePart.load generic filename ("image.png"); the error path
%                     (a missing file -> mat2doc:FileNotFoundError identifier).
%   * Equivalence  -- test_equivalence_s0098_full_probe replays the ENTIRE s0098
%                     probe (ImagePart/ImageParts/StoryPart/Run/Document/InlineShapes)
%                     and compares every leaf to the frozen python-docx 1.2.0 oracle
%                     copied into data\s0098\probe.json (Gate-3 probe_diff MATCH).
%   * Regression   -- hard-coded full-package SHA-256 (+ size) byte pins for
%                     s0090..s0097 (in frozen zip-entry order) + the M1 SHAs.
%   * Upstream     -- the media blob being an EXACT source-file byte copy, the
%                     partname numbering (image1/image2, non-numbered-skip), the
%                     image rels/rIds (rId9 body / rId1 header / rId10 second image,
%                     insertion order H11 no-sort), the [Content_Types] Defaults, the
%                     zip-entry order, the wp:extent EMU, and the header .rels ARE the
%                     python-docx package.py / parts/image.py / story.py contract; the
%                     frozen oracle IS lxml/python-docx's expected output.
%
%   Byte-level (L1) note: every full-package assertion is a SHA-256 (+ size) pin of
%   the raw shipping bytes of each zip part, in frozen zip-entry order. SHA-256
%   equality == byte identity (L1) -- for the binary media blobs this is bit-exact
%   image bytes; for the XML parts it is byte-identical serialization. NO D-number
%   granted any L2 relaxation in this WP (Gate-3: ZERO new), so every byte pin is
%   L1. The equivalence leaf-key floor is the only looser-than-byte check and is
%   commented at its site.
%
%   Determinism: no network, no absolute paths. The source images, the scenario-J
%   input docx and the probe oracle resolve relative to this file via
%   fileparts(mfilename('fullpath')); saves go to tempname .docx deleted via
%   onCleanup; every file read is binary ('r','n'). The +mat2doc package resolves
%   via the MANDATORY PathFixture(worktree-root) in TestClassSetup (WP9-F4 lesson).

    properties (Constant)
        % ==== s0090 -- ** THE P7 HEADLINE ** add PNG, 18 parts (frozen zip order) ====
        S0090_MANIFEST = [ ...
            "[Content_Types].xml",               "1788",      "1a1b7bf7aeab6ab754ba491b21600be5de0b06b910a5e1034eed09dfba6f5f1e"; ...
            "_rels/.rels",                       "734",       "a9ae57efe9186f07d48303bcac5d54c7e359bf3f503939c03ad2954e5c59a5c4"; ...
            "docProps/core.xml",                 "721",       "d14be8284e406d14dc2a576b19f15e7f637b955f9e27a3d9ed871ec99606f7ba"; ...
            "docProps/app.xml",                  "1132",      "be664981c3141cddfc59362beb287ebf20d0773660e2dd6faac5968a5930a081"; ...
            "word/document.xml",                 "2335",      "d0559ef6211745b27da82073e02337997b4c8bd9ed35c9aa42efae162de43eaf"; ...
            "word/_rels/document.xml.rels",      "1359",      "dc5eaed833ad9f3e84a8a809036e4d4cd217373c559111f2460655c102596c4d"; ...
            "word/styles.xml",                   "349458",    "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"; ...
            "word/stylesWithEffects.xml",        "438131",    "463ae0928cf0d84775dbf8cf18d6c3029f6707c81bf590f6d6dd8757a5e93f15"; ...
            "word/settings.xml",                 "2535",      "51a0d348fe85965c66e4748a03c3c0d055d78455514f03cb121c334af7d73689"; ...
            "word/webSettings.xml",              "438",       "349d36de7434d09f86987ff671d8814964a0588c1e630c06e562cda7e75e9f95"; ...
            "word/fontTable.xml",                "2811",      "79385fb7f60247507ecaffc292e9ebd52ea0657b8634f629ba6fccc54011d6bb"; ...
            "word/theme/theme1.xml",             "10939",     "e3a8ab7db9ca7afca56f5f2820a56e8b660016c647773555b060b0a02ac76941"; ...
            "customXml/item1.xml",               "262",       "a86086ffc5d8e83ebd6c71a55d1d2efaa31b137977f5f3a752366e1023612144"; ...
            "customXml/_rels/item1.xml.rels",    "295",       "1ca6c9a64edcebe24ee703a54403611b322d96da33371779e742d2d3f7ed7a6c"; ...
            "customXml/itemProps1.xml",          "354",       "c542307b13ec29a8b546217bb37936ab4822e044b265d2952985ec3d6afed24e"; ...
            "word/numbering.xml",                "5513",      "70976f19cbcd896e51890859fe6ecb3467a5a7ad0c040160fedc5a1993cb09ce"; ...
            "word/media/image1.png",             "6111",      "b60f709965eb54e1400cba53617ae2c02f8028f5399cafa32d8ef18bfc803915"; ...
            "docProps/thumbnail.jpeg",           "8324",      "96367138dc44ce09bf2c8f0f8e49348a1478d2c5c0af69bbc2bbc38b63cdcead"];

        % ==== s0091 -- dedupe: same PNG x2 -> ONE media part, 18 parts ====
        S0091_MANIFEST = [ ...
            "[Content_Types].xml",               "1788",      "1a1b7bf7aeab6ab754ba491b21600be5de0b06b910a5e1034eed09dfba6f5f1e"; ...
            "_rels/.rels",                       "734",       "a9ae57efe9186f07d48303bcac5d54c7e359bf3f503939c03ad2954e5c59a5c4"; ...
            "docProps/core.xml",                 "721",       "d14be8284e406d14dc2a576b19f15e7f637b955f9e27a3d9ed871ec99606f7ba"; ...
            "docProps/app.xml",                  "1132",      "be664981c3141cddfc59362beb287ebf20d0773660e2dd6faac5968a5930a081"; ...
            "word/document.xml",                 "3122",      "62968269b3d2860a04b122a0f628ce78ca8fa4f82acc34e33438c78d84f07063"; ...
            "word/_rels/document.xml.rels",      "1359",      "dc5eaed833ad9f3e84a8a809036e4d4cd217373c559111f2460655c102596c4d"; ...
            "word/styles.xml",                   "349458",    "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"; ...
            "word/stylesWithEffects.xml",        "438131",    "463ae0928cf0d84775dbf8cf18d6c3029f6707c81bf590f6d6dd8757a5e93f15"; ...
            "word/settings.xml",                 "2535",      "51a0d348fe85965c66e4748a03c3c0d055d78455514f03cb121c334af7d73689"; ...
            "word/webSettings.xml",              "438",       "349d36de7434d09f86987ff671d8814964a0588c1e630c06e562cda7e75e9f95"; ...
            "word/fontTable.xml",                "2811",      "79385fb7f60247507ecaffc292e9ebd52ea0657b8634f629ba6fccc54011d6bb"; ...
            "word/theme/theme1.xml",             "10939",     "e3a8ab7db9ca7afca56f5f2820a56e8b660016c647773555b060b0a02ac76941"; ...
            "customXml/item1.xml",               "262",       "a86086ffc5d8e83ebd6c71a55d1d2efaa31b137977f5f3a752366e1023612144"; ...
            "customXml/_rels/item1.xml.rels",    "295",       "1ca6c9a64edcebe24ee703a54403611b322d96da33371779e742d2d3f7ed7a6c"; ...
            "customXml/itemProps1.xml",          "354",       "c542307b13ec29a8b546217bb37936ab4822e044b265d2952985ec3d6afed24e"; ...
            "word/numbering.xml",                "5513",      "70976f19cbcd896e51890859fe6ecb3467a5a7ad0c040160fedc5a1993cb09ce"; ...
            "word/media/image1.png",             "6111",      "b60f709965eb54e1400cba53617ae2c02f8028f5399cafa32d8ef18bfc803915"; ...
            "docProps/thumbnail.jpeg",           "8324",      "96367138dc44ce09bf2c8f0f8e49348a1478d2c5c0af69bbc2bbc38b63cdcead"];

        % ==== s0092 -- two DISTINCT images (PNG then JPEG), 19 parts ====
        S0092_MANIFEST = [ ...
            "[Content_Types].xml",               "1788",      "1a1b7bf7aeab6ab754ba491b21600be5de0b06b910a5e1034eed09dfba6f5f1e"; ...
            "_rels/.rels",                       "734",       "a9ae57efe9186f07d48303bcac5d54c7e359bf3f503939c03ad2954e5c59a5c4"; ...
            "docProps/core.xml",                 "721",       "d14be8284e406d14dc2a576b19f15e7f637b955f9e27a3d9ed871ec99606f7ba"; ...
            "docProps/app.xml",                  "1132",      "be664981c3141cddfc59362beb287ebf20d0773660e2dd6faac5968a5930a081"; ...
            "word/document.xml",                 "3123",      "3b6fd03df1798ba9b0740bf527dac3bde814b873bc4b41fb058fa0795a84f480"; ...
            "word/_rels/document.xml.rels",      "1493",      "ae8501016223eb77759334aa67983d7b57259781b3e73005e8ab74d60cd75c78"; ...
            "word/styles.xml",                   "349458",    "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"; ...
            "word/stylesWithEffects.xml",        "438131",    "463ae0928cf0d84775dbf8cf18d6c3029f6707c81bf590f6d6dd8757a5e93f15"; ...
            "word/settings.xml",                 "2535",      "51a0d348fe85965c66e4748a03c3c0d055d78455514f03cb121c334af7d73689"; ...
            "word/webSettings.xml",              "438",       "349d36de7434d09f86987ff671d8814964a0588c1e630c06e562cda7e75e9f95"; ...
            "word/fontTable.xml",                "2811",      "79385fb7f60247507ecaffc292e9ebd52ea0657b8634f629ba6fccc54011d6bb"; ...
            "word/theme/theme1.xml",             "10939",     "e3a8ab7db9ca7afca56f5f2820a56e8b660016c647773555b060b0a02ac76941"; ...
            "customXml/item1.xml",               "262",       "a86086ffc5d8e83ebd6c71a55d1d2efaa31b137977f5f3a752366e1023612144"; ...
            "customXml/_rels/item1.xml.rels",    "295",       "1ca6c9a64edcebe24ee703a54403611b322d96da33371779e742d2d3f7ed7a6c"; ...
            "customXml/itemProps1.xml",          "354",       "c542307b13ec29a8b546217bb37936ab4822e044b265d2952985ec3d6afed24e"; ...
            "word/numbering.xml",                "5513",      "70976f19cbcd896e51890859fe6ecb3467a5a7ad0c040160fedc5a1993cb09ce"; ...
            "word/media/image1.png",             "6111",      "b60f709965eb54e1400cba53617ae2c02f8028f5399cafa32d8ef18bfc803915"; ...
            "word/media/image2.jpeg",            "3277",      "ad5afb51d11ed1b0bcd80103519ad198129eb9489637b00b9547f208408548f0"; ...
            "docProps/thumbnail.jpeg",           "8324",      "96367138dc44ce09bf2c8f0f8e49348a1478d2c5c0af69bbc2bbc38b63cdcead"];

        % ==== s0093 -- single JPEG -> image1.jpeg, 18 parts ====
        S0093_MANIFEST = [ ...
            "[Content_Types].xml",               "1738",      "66c84fb7a6aa3c4ead49f895e4a7044df1fb57de1ed76d09b2686e91f5bed5b4"; ...
            "_rels/.rels",                       "734",       "a9ae57efe9186f07d48303bcac5d54c7e359bf3f503939c03ad2954e5c59a5c4"; ...
            "docProps/core.xml",                 "721",       "d14be8284e406d14dc2a576b19f15e7f637b955f9e27a3d9ed871ec99606f7ba"; ...
            "docProps/app.xml",                  "1132",      "be664981c3141cddfc59362beb287ebf20d0773660e2dd6faac5968a5930a081"; ...
            "word/document.xml",                 "2335",      "dad33ee50304bf5761ea0a09595d31675c6f643670bd213e623e4759294de4ad"; ...
            "word/_rels/document.xml.rels",      "1360",      "5d71f2705b0e71cabfc71524137d4219be246265f658515fb401e728941a04fc"; ...
            "word/styles.xml",                   "349458",    "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"; ...
            "word/stylesWithEffects.xml",        "438131",    "463ae0928cf0d84775dbf8cf18d6c3029f6707c81bf590f6d6dd8757a5e93f15"; ...
            "word/settings.xml",                 "2535",      "51a0d348fe85965c66e4748a03c3c0d055d78455514f03cb121c334af7d73689"; ...
            "word/webSettings.xml",              "438",       "349d36de7434d09f86987ff671d8814964a0588c1e630c06e562cda7e75e9f95"; ...
            "word/fontTable.xml",                "2811",      "79385fb7f60247507ecaffc292e9ebd52ea0657b8634f629ba6fccc54011d6bb"; ...
            "word/theme/theme1.xml",             "10939",     "e3a8ab7db9ca7afca56f5f2820a56e8b660016c647773555b060b0a02ac76941"; ...
            "customXml/item1.xml",               "262",       "a86086ffc5d8e83ebd6c71a55d1d2efaa31b137977f5f3a752366e1023612144"; ...
            "customXml/_rels/item1.xml.rels",    "295",       "1ca6c9a64edcebe24ee703a54403611b322d96da33371779e742d2d3f7ed7a6c"; ...
            "customXml/itemProps1.xml",          "354",       "c542307b13ec29a8b546217bb37936ab4822e044b265d2952985ec3d6afed24e"; ...
            "word/numbering.xml",                "5513",      "70976f19cbcd896e51890859fe6ecb3467a5a7ad0c040160fedc5a1993cb09ce"; ...
            "word/media/image1.jpeg",            "3277",      "ad5afb51d11ed1b0bcd80103519ad198129eb9489637b00b9547f208408548f0"; ...
            "docProps/thumbnail.jpeg",           "8324",      "96367138dc44ce09bf2c8f0f8e49348a1478d2c5c0af69bbc2bbc38b63cdcead"];

        % ==== s0094 -- width override Inches(2) -> wp:extent cx=1828800 cy=731520 ====
        S0094_MANIFEST = [ ...
            "[Content_Types].xml",               "1788",      "1a1b7bf7aeab6ab754ba491b21600be5de0b06b910a5e1034eed09dfba6f5f1e"; ...
            "_rels/.rels",                       "734",       "a9ae57efe9186f07d48303bcac5d54c7e359bf3f503939c03ad2954e5c59a5c4"; ...
            "docProps/core.xml",                 "721",       "d14be8284e406d14dc2a576b19f15e7f637b955f9e27a3d9ed871ec99606f7ba"; ...
            "docProps/app.xml",                  "1132",      "be664981c3141cddfc59362beb287ebf20d0773660e2dd6faac5968a5930a081"; ...
            "word/document.xml",                 "2335",      "3919505068354865100f645901db1b64d60fe9217e1e32f70c4e8453b3df6e6c"; ...
            "word/_rels/document.xml.rels",      "1359",      "dc5eaed833ad9f3e84a8a809036e4d4cd217373c559111f2460655c102596c4d"; ...
            "word/styles.xml",                   "349458",    "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"; ...
            "word/stylesWithEffects.xml",        "438131",    "463ae0928cf0d84775dbf8cf18d6c3029f6707c81bf590f6d6dd8757a5e93f15"; ...
            "word/settings.xml",                 "2535",      "51a0d348fe85965c66e4748a03c3c0d055d78455514f03cb121c334af7d73689"; ...
            "word/webSettings.xml",              "438",       "349d36de7434d09f86987ff671d8814964a0588c1e630c06e562cda7e75e9f95"; ...
            "word/fontTable.xml",                "2811",      "79385fb7f60247507ecaffc292e9ebd52ea0657b8634f629ba6fccc54011d6bb"; ...
            "word/theme/theme1.xml",             "10939",     "e3a8ab7db9ca7afca56f5f2820a56e8b660016c647773555b060b0a02ac76941"; ...
            "customXml/item1.xml",               "262",       "a86086ffc5d8e83ebd6c71a55d1d2efaa31b137977f5f3a752366e1023612144"; ...
            "customXml/_rels/item1.xml.rels",    "295",       "1ca6c9a64edcebe24ee703a54403611b322d96da33371779e742d2d3f7ed7a6c"; ...
            "customXml/itemProps1.xml",          "354",       "c542307b13ec29a8b546217bb37936ab4822e044b265d2952985ec3d6afed24e"; ...
            "word/numbering.xml",                "5513",      "70976f19cbcd896e51890859fe6ecb3467a5a7ad0c040160fedc5a1993cb09ce"; ...
            "word/media/image1.png",             "6111",      "b60f709965eb54e1400cba53617ae2c02f8028f5399cafa32d8ef18bfc803915"; ...
            "docProps/thumbnail.jpeg",           "8324",      "96367138dc44ce09bf2c8f0f8e49348a1478d2c5c0af69bbc2bbc38b63cdcead"];

        % ==== s0095 -- DEFECT-1 scenario J: reopen non-numbered logo.png + JPEG, 19 parts ====
        S0095_MANIFEST = [ ...
            "[Content_Types].xml",               "1788",      "1a1b7bf7aeab6ab754ba491b21600be5de0b06b910a5e1034eed09dfba6f5f1e"; ...
            "_rels/.rels",                       "734",       "a9ae57efe9186f07d48303bcac5d54c7e359bf3f503939c03ad2954e5c59a5c4"; ...
            "docProps/core.xml",                 "721",       "d14be8284e406d14dc2a576b19f15e7f637b955f9e27a3d9ed871ec99606f7ba"; ...
            "docProps/app.xml",                  "1132",      "be664981c3141cddfc59362beb287ebf20d0773660e2dd6faac5968a5930a081"; ...
            "word/document.xml",                 "3123",      "3b6fd03df1798ba9b0740bf527dac3bde814b873bc4b41fb058fa0795a84f480"; ...
            "word/_rels/document.xml.rels",      "1491",      "8dee3ac274d8cea3b469a5c687427248ea5cb6c371905df1a9e962ec3f90202f"; ...
            "word/styles.xml",                   "349458",    "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"; ...
            "word/stylesWithEffects.xml",        "438131",    "463ae0928cf0d84775dbf8cf18d6c3029f6707c81bf590f6d6dd8757a5e93f15"; ...
            "word/settings.xml",                 "2535",      "51a0d348fe85965c66e4748a03c3c0d055d78455514f03cb121c334af7d73689"; ...
            "word/webSettings.xml",              "438",       "349d36de7434d09f86987ff671d8814964a0588c1e630c06e562cda7e75e9f95"; ...
            "word/fontTable.xml",                "2811",      "79385fb7f60247507ecaffc292e9ebd52ea0657b8634f629ba6fccc54011d6bb"; ...
            "word/theme/theme1.xml",             "10939",     "e3a8ab7db9ca7afca56f5f2820a56e8b660016c647773555b060b0a02ac76941"; ...
            "customXml/item1.xml",               "262",       "a86086ffc5d8e83ebd6c71a55d1d2efaa31b137977f5f3a752366e1023612144"; ...
            "customXml/_rels/item1.xml.rels",    "295",       "1ca6c9a64edcebe24ee703a54403611b322d96da33371779e742d2d3f7ed7a6c"; ...
            "customXml/itemProps1.xml",          "354",       "c542307b13ec29a8b546217bb37936ab4822e044b265d2952985ec3d6afed24e"; ...
            "word/numbering.xml",                "5513",      "70976f19cbcd896e51890859fe6ecb3467a5a7ad0c040160fedc5a1993cb09ce"; ...
            "word/media/logo.png",               "6111",      "b60f709965eb54e1400cba53617ae2c02f8028f5399cafa32d8ef18bfc803915"; ...
            "word/media/image1.jpeg",            "3277",      "ad5afb51d11ed1b0bcd80103519ad198129eb9489637b00b9547f208408548f0"; ...
            "docProps/thumbnail.jpeg",           "8324",      "96367138dc44ce09bf2c8f0f8e49348a1478d2c5c0af69bbc2bbc38b63cdcead"];

        % ==== s0096 -- ** C6 HEADER image **, 20 parts (header1.xml + first-ever .rels) ====
        S0096_MANIFEST = [ ...
            "[Content_Types].xml",               "1916",      "930facc5fe21587ce1966e88e6b3e21002e62c409d2ad8645bcc662713f891d9"; ...
            "_rels/.rels",                       "734",       "a9ae57efe9186f07d48303bcac5d54c7e359bf3f503939c03ad2954e5c59a5c4"; ...
            "docProps/core.xml",                 "721",       "d14be8284e406d14dc2a576b19f15e7f637b955f9e27a3d9ed871ec99606f7ba"; ...
            "docProps/app.xml",                  "1132",      "be664981c3141cddfc59362beb287ebf20d0773660e2dd6faac5968a5930a081"; ...
            "word/document.xml",                 "1597",      "e3b5e38706b3bf9de625e5f684a2c01ae4330727c2c91518c39b72acbd4cfc44"; ...
            "word/_rels/document.xml.rels",      "1355",      "29b4d0f49e573a7cc8901b470fe1de2d453ebd7edbb67cc53b806120c1f978c2"; ...
            "word/styles.xml",                   "349458",    "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"; ...
            "word/stylesWithEffects.xml",        "438131",    "463ae0928cf0d84775dbf8cf18d6c3029f6707c81bf590f6d6dd8757a5e93f15"; ...
            "word/settings.xml",                 "2535",      "51a0d348fe85965c66e4748a03c3c0d055d78455514f03cb121c334af7d73689"; ...
            "word/webSettings.xml",              "438",       "349d36de7434d09f86987ff671d8814964a0588c1e630c06e562cda7e75e9f95"; ...
            "word/fontTable.xml",                "2811",      "79385fb7f60247507ecaffc292e9ebd52ea0657b8634f629ba6fccc54011d6bb"; ...
            "word/theme/theme1.xml",             "10939",     "e3a8ab7db9ca7afca56f5f2820a56e8b660016c647773555b060b0a02ac76941"; ...
            "customXml/item1.xml",               "262",       "a86086ffc5d8e83ebd6c71a55d1d2efaa31b137977f5f3a752366e1023612144"; ...
            "customXml/_rels/item1.xml.rels",    "295",       "1ca6c9a64edcebe24ee703a54403611b322d96da33371779e742d2d3f7ed7a6c"; ...
            "customXml/itemProps1.xml",          "354",       "c542307b13ec29a8b546217bb37936ab4822e044b265d2952985ec3d6afed24e"; ...
            "word/numbering.xml",                "5513",      "70976f19cbcd896e51890859fe6ecb3467a5a7ad0c040160fedc5a1993cb09ce"; ...
            "word/header1.xml",                  "2069",      "a4064db029092038573c0e3b237447a3f566025f51152f84d1d5a46551ddcd05"; ...
            "word/_rels/header1.xml.rels",       "288",       "a134ffb58bad658becad1362f28a7fcb92904a928e5a686fbae9bed93092e9d8"; ...
            "word/media/image1.png",             "6111",      "b60f709965eb54e1400cba53617ae2c02f8028f5399cafa32d8ef18bfc803915"; ...
            "docProps/thumbnail.jpeg",           "8324",      "96367138dc44ce09bf2c8f0f8e49348a1478d2c5c0af69bbc2bbc38b63cdcead"];

        % ==== s0097 -- body+header SHARE one media part, 20 parts ====
        S0097_MANIFEST = [ ...
            "[Content_Types].xml",               "1916",      "930facc5fe21587ce1966e88e6b3e21002e62c409d2ad8645bcc662713f891d9"; ...
            "_rels/.rels",                       "734",       "a9ae57efe9186f07d48303bcac5d54c7e359bf3f503939c03ad2954e5c59a5c4"; ...
            "docProps/core.xml",                 "721",       "d14be8284e406d14dc2a576b19f15e7f637b955f9e27a3d9ed871ec99606f7ba"; ...
            "docProps/app.xml",                  "1132",      "be664981c3141cddfc59362beb287ebf20d0773660e2dd6faac5968a5930a081"; ...
            "word/document.xml",                 "2385",      "50b9885c5ac399e0f30f846a4ae33ee02ab727549d2560d73953b75051b3fffa"; ...
            "word/_rels/document.xml.rels",      "1488",      "3b631cdc8447a32f8746740b4a5df5161d44a056bc5c4c8af1c9096b3613d363"; ...
            "word/styles.xml",                   "349458",    "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"; ...
            "word/stylesWithEffects.xml",        "438131",    "463ae0928cf0d84775dbf8cf18d6c3029f6707c81bf590f6d6dd8757a5e93f15"; ...
            "word/settings.xml",                 "2535",      "51a0d348fe85965c66e4748a03c3c0d055d78455514f03cb121c334af7d73689"; ...
            "word/webSettings.xml",              "438",       "349d36de7434d09f86987ff671d8814964a0588c1e630c06e562cda7e75e9f95"; ...
            "word/fontTable.xml",                "2811",      "79385fb7f60247507ecaffc292e9ebd52ea0657b8634f629ba6fccc54011d6bb"; ...
            "word/theme/theme1.xml",             "10939",     "e3a8ab7db9ca7afca56f5f2820a56e8b660016c647773555b060b0a02ac76941"; ...
            "customXml/item1.xml",               "262",       "a86086ffc5d8e83ebd6c71a55d1d2efaa31b137977f5f3a752366e1023612144"; ...
            "customXml/_rels/item1.xml.rels",    "295",       "1ca6c9a64edcebe24ee703a54403611b322d96da33371779e742d2d3f7ed7a6c"; ...
            "customXml/itemProps1.xml",          "354",       "c542307b13ec29a8b546217bb37936ab4822e044b265d2952985ec3d6afed24e"; ...
            "word/numbering.xml",                "5513",      "70976f19cbcd896e51890859fe6ecb3467a5a7ad0c040160fedc5a1993cb09ce"; ...
            "word/media/image1.png",             "6111",      "b60f709965eb54e1400cba53617ae2c02f8028f5399cafa32d8ef18bfc803915"; ...
            "word/header1.xml",                  "2069",      "a4064db029092038573c0e3b237447a3f566025f51152f84d1d5a46551ddcd05"; ...
            "word/_rels/header1.xml.rels",       "288",       "a134ffb58bad658becad1362f28a7fcb92904a928e5a686fbae9bed93092e9d8"; ...
            "docProps/thumbnail.jpeg",           "8324",      "96367138dc44ce09bf2c8f0f8e49348a1478d2c5c0af69bbc2bbc38b63cdcead"];

        % ---- M1 default-save byte guards (references\s0001; _gather_image_parts neutral) ----
        STYLES_SIZE_M1 = 349458
        STYLES_SHA_M1  = "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"
        DOC_SIZE_M1    = 1548
        DOC_SHA_M1     = "0e4dd503bc095219cd24e8312b372c588ca56bb834aacac53eafb7bf47836327"

        % ---- the source-image SHAs / EMU pins (headline witnesses) ----
        PNG_SHA  = "b60f709965eb54e1400cba53617ae2c02f8028f5399cafa32d8ef18bfc803915"
        JPEG_SHA = "ad5afb51d11ed1b0bcd80103519ad198129eb9489637b00b9547f208408548f0"
        WIDTH_OVERRIDE_CX = 1828800   % Inches(2) -> wp:extent cx
        WIDTH_OVERRIDE_CY = 731520    % scaled_dimensions aspect-preserving cy
        NATIVE_CX = 1778000           % default_cx (native PNG width)
        NATIVE_CY = 711200            % default_cy (native PNG height, horz_dpi quirk)
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied from tests\parts\Test_p2_2_storypart_parts.m. here is
            % tests\parts; the worktree root is two levels up.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\parts
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. ** THE P7 HEADLINE ** add_picture full-package byte pin (s0090)%
        % =============================================================== %

        function test_add_picture_full_package_byte_pin(testCase)
            % Regression (L1+bin, THE headline P7 pin): the exact Gate-3 G-scenario --
            %   Document(); add_picture(<PNG>); save()
            % emits ALL 18 parts byte-identical (size + SHA-256, in frozen zip-entry
            % order) to the frozen s0090 python-docx oracle. The FIRST runtime image
            % part word/media/image1.png (6111 B / b60f7099) is an EXACT byte copy of
            % the source PNG; word/document.xml (2335 B / d0559ef6) carries the
            % paragraph/run/w:drawing/wp:inline; document.xml.rels adds rId9 ->
            % media/image1.png (Type .../image); [Content_Types].xml adds
            % <Default Extension="png">. The NEW media part lands AFTER
            % word/numbering.xml, BEFORE docProps/thumbnail.jpeg (H11 zip-order pin).
            % RED on ANY single-byte / rId / Default / zip-order / image-blob drift.
            d = mat2doc.Document();
            d.add_picture(testCase.dataFile('python-powered.png'));   % Python: add_picture(IMG)
            zipBytes = testCase.saveZip(d);
            [blobs, names] = zipEntryList(zipBytes);
            testCase.assertPackage(zipBytes, testCase.S0090_MANIFEST, ...
                's0090 add PNG (** THE P7 HEADLINE **)');

            % --- loud headline witnesses (backing the SHA pin, human-readable) ---
            img = entryBlob(blobs, names, "word/media/image1.png");
            testCase.verifyEqual(numel(img), 6111, ...
                'headline: word/media/image1.png is exactly 6111 B (exact source copy)');
            testCase.verifyEqual(sha256hex(img), testCase.PNG_SHA, ...
                'headline: word/media/image1.png is a BYTE-EXACT copy of the source PNG (b60f7099)');
            rels = char(entryBlob(blobs, names, "word/_rels/document.xml.rels"));
            testCase.verifyTrue(contains(rels, 'Id="rId9"'), ...
                'headline: the image rel is rId9 (insertion order, no rId sort -- H11)');
            testCase.verifyTrue(contains(rels, 'Target="media/image1.png"'), ...
                'headline: rId9 targets media/image1.png');
            testCase.verifyTrue(contains(rels, ...
                '/image"'), 'headline: rId9 Type is the .../relationships/image reltype');
            ct = char(entryBlob(blobs, names, "[Content_Types].xml"));
            testCase.verifyTrue(contains(ct, '<Default Extension="png" ContentType="image/png"/>'), ...
                'headline: [Content_Types].xml adds <Default Extension="png" ContentType="image/png"/>');
            doc = char(entryBlob(blobs, names, "word/document.xml"));
            testCase.verifyTrue(contains(doc, '<w:drawing>') && contains(doc, 'r:embed="rId9"'), ...
                'headline: document.xml carries a w:drawing whose blip embeds rId9');
        end

        % =============================================================== %
        % 2. ** dedupe ** -- same PNG x2 -> ONE media part (s0091)          %
        % =============================================================== %

        function test_dedupe_full_package_byte_pin(testCase)
            % Regression (L1, dedupe): the SAME PNG added twice collapses (SHA1 dedupe
            % in ImageParts.get_or_add_image_part + the rel dedupe in relate_to) to
            % EXACTLY ONE word/media/image1.png, and BOTH w:drawings reuse rId9 ->
            % 18/18 byte-identical to the frozen s0091 oracle. document.xml grows the
            % second drawing (3122 B / 62968269) but the media inventory does NOT.
            d = mat2doc.Document();
            d.add_picture(testCase.dataFile('python-powered.png'));   % Python: add_picture(IMG)
            d.add_picture(testCase.dataFile('python-powered.png'));   % Python: add_picture(IMG) -- dedupe
            zipBytes = testCase.saveZip(d);
            [blobs, names] = zipEntryList(zipBytes);
            testCase.assertPackage(zipBytes, testCase.S0091_MANIFEST, ...
                's0091 dedupe (same PNG x2 -> ONE media part)');
            % readable witness: exactly ONE media/image*.png part; both drawings rId9.
            mediaCount = sum(startsWith(names, "word/media/image"));
            testCase.verifyEqual(mediaCount, 1, ...
                'dedupe: exactly ONE word/media/image* part after adding the same PNG twice');
            doc = char(entryBlob(blobs, names, "word/document.xml"));
            testCase.verifyEqual(numel(strfind(doc, 'r:embed="rId9"')), 2, ...
                'dedupe: BOTH w:drawings share the SAME rId9 (rel dedupe)');
        end

        % =============================================================== %
        % 3. ** two images ** -- PNG + JPEG -> image1.png + image2.jpeg     %
        % =============================================================== %

        function test_two_images_full_package_byte_pin(testCase)
            % Regression (L1, partname numbering): two DISTINCT images (PNG then JPEG)
            % -> word/media/image1.png + word/media/image2.jpeg (the number is unique
            % WITHOUT regard to extension), two rels (rId9/rId10, insertion order), two
            % [Content_Types] Defaults (png+jpeg) -> 19/19 byte-identical to the frozen
            % s0092 oracle.
            d = mat2doc.Document();
            d.add_picture(testCase.dataFile('python-powered.png'));   % Python: add_picture(PNG)
            d.add_picture(testCase.dataFile('python-icon.jpeg'));     % Python: add_picture(JPEG)
            zipBytes = testCase.saveZip(d);
            [blobs, names] = zipEntryList(zipBytes);
            testCase.assertPackage(zipBytes, testCase.S0092_MANIFEST, ...
                's0092 two DISTINCT images (image1.png + image2.jpeg)');
            % readable witness: two media parts, rId9 before rId10 (insertion order).
            testCase.verifyTrue(any(names == "word/media/image1.png") && ...
                any(names == "word/media/image2.jpeg"), ...
                'two-images: media = {image1.png, image2.jpeg} (number unique across extensions)');
            rels = char(entryBlob(blobs, names, "word/_rels/document.xml.rels"));
            i9 = strfind(rels, 'Id="rId9"'); i10 = strfind(rels, 'Id="rId10"');
            testCase.verifyNotEmpty(i9,  'rId9 (PNG rel) present');
            testCase.verifyNotEmpty(i10, 'rId10 (JPEG rel) present');
            testCase.verifyLessThan(i9(1), i10(1), ...
                'two-images: rId9 (PNG) precedes rId10 (JPEG) -- insertion order, no sort (H11)');
        end

        % =============================================================== %
        % 4. ** JPEG ** + ** width override **                              %
        % =============================================================== %

        function test_jpeg_full_package_byte_pin(testCase)
            % Regression (L1): a single JPEG -> word/media/image1.jpeg (3277 B /
            % ad5afb51), [Content_Types] Default jpeg -> 18/18 byte-identical to the
            % frozen s0093 oracle. The default.docx already carries a jpeg Default (the
            % thumbnail) so [Content_Types].xml stays at 1738 B / 66c84fb7.
            d = mat2doc.Document();
            d.add_picture(testCase.dataFile('python-icon.jpeg'));     % Python: add_picture(JPEG)
            zipBytes = testCase.saveZip(d);
            testCase.assertPackage(zipBytes, testCase.S0093_MANIFEST, ...
                's0093 single JPEG (image1.jpeg)');
        end

        function test_width_override_full_package_byte_pin(testCase)
            % Regression (L1): add_picture(PNG, width=Inches(2)) -> scaled_dimensions
            % derives the aspect-preserving height; the wp:inline carries
            % wp:extent cx=1828800 cy=731520 -> 18/18 byte-identical to the frozen
            % s0094 oracle (document.xml 39195050). The media blob is unchanged (the
            % override affects only the drawing EMU, not the stored image bytes).
            d = mat2doc.Document();
            d.add_picture(testCase.dataFile('python-powered.png'), mat2doc.shared.Inches(2)); % Python: width=Inches(2)
            zipBytes = testCase.saveZip(d);
            [blobs, names] = zipEntryList(zipBytes);
            testCase.assertPackage(zipBytes, testCase.S0094_MANIFEST, ...
                's0094 width override Inches(2)');
            % readable witness: the exact scaled wp:extent EMU in document.xml.
            doc = char(entryBlob(blobs, names, "word/document.xml"));
            testCase.verifyTrue(contains(doc, ...
                sprintf('<wp:extent cx="%d" cy="%d"/>', testCase.WIDTH_OVERRIDE_CX, testCase.WIDTH_OVERRIDE_CY)), ...
                'width override: wp:extent cx=1828800 cy=731520 (scaled_dimensions)');
        end

        % =============================================================== %
        % 5. ** C6 HEADER image ** (s0096) + body+header share (s0097)       %
        % =============================================================== %

        function test_header_image_full_package_byte_pin(testCase)
            % Regression (L1, the P5-3b payoff): add_picture routed into a HEADER
            % StoryPart. Accessing sections[0].header.paragraphs[0] on the first
            % section materializes word/header1.xml; the run relates the media part
            % into the header part with its OWN rId1 -> the FIRST-EVER
            % word/_rels/header1.xml.rels in Mat2Doc. 20/20 byte-identical to the
            % frozen s0096 oracle (header1.xml 2069 B / a4064db0; header1.xml.rels
            % 288 B / a134ffb5; the shared media/image1.png).
            d = mat2doc.Document();
            hdr = d.sections.getitem_(0).header;     % Python: sections[0].header
            ps = hdr.paragraphs;                     % Python: hdr.paragraphs[0]
            r = ps(1).add_run();                     % Python: r = p.add_run()
            r.add_picture(testCase.dataFile('python-powered.png'));   % Python: r.add_picture(IMG)
            zipBytes = testCase.saveZip(d);
            [blobs, names] = zipEntryList(zipBytes);
            testCase.assertPackage(zipBytes, testCase.S0096_MANIFEST, ...
                's0096 ** C6 HEADER image ** (first-ever header .rels)');
            % readable witness: header1.xml.rels rId1 -> media/image1.png.
            hrels = char(entryBlob(blobs, names, "word/_rels/header1.xml.rels"));
            testCase.verifyTrue(contains(hrels, 'Id="rId1"') && ...
                contains(hrels, 'Target="media/image1.png"'), ...
                'C6 header: word/_rels/header1.xml.rels rId1 -> media/image1.png');
        end

        function test_body_header_share_full_package_byte_pin(testCase)
            % Regression (L1, package-level media sharing): the SAME PNG in the body
            % AND the header shares ONE package-level word/media/image1.png (image_parts
            % is package-level); the body's DocumentPart relates it rId9, the header's
            % HeaderPart relates it rId1 (relationships are per story part) -> 20/20
            % byte-identical to the frozen s0097 oracle, with a SINGLE media part.
            d = mat2doc.Document();
            d.add_picture(testCase.dataFile('python-powered.png'));   % Python: add_picture(IMG) (body)
            hdr = d.sections.getitem_(0).header;     % Python: sections[0].header
            ps = hdr.paragraphs;
            r = ps(1).add_run();
            r.add_picture(testCase.dataFile('python-powered.png'));   % Python: r.add_picture(IMG) (header)
            zipBytes = testCase.saveZip(d);
            [blobs, names] = zipEntryList(zipBytes);
            testCase.assertPackage(zipBytes, testCase.S0097_MANIFEST, ...
                's0097 body+header SHARE one media part');
            mediaCount = sum(startsWith(names, "word/media/image"));
            testCase.verifyEqual(mediaCount, 1, ...
                'body+header share: exactly ONE package-level media part (shared)');
        end

        % =============================================================== %
        % 6. ** DEFECT-1 regression pin -- scenario J ** (s0095)             %
        % =============================================================== %

        function test_defect1_scenario_j_reopen_logo_add_jpeg(testCase)
            % ** DEFECT-1 REGRESSION GUARD ** (L1, Edge/error-path-forward): open a
            % docx whose ONLY media part carries a NON-numbered partname
            % (/word/media/logo.png -> PackURI.idx == []), then add a JPEG.
            %
            % Before the Gate-2 DEFECT-1 fix, Mat2Doc CRASHED here: appending the
            % None (idx==[]) to the used_numbers vector ran `used_numbers(end+1)=[]`
            % which DELETES an element (singleSubscript error). python-docx succeeds
            % (a None entry never equals a candidate integer). The fixed
            % ImageParts.next_image_partname_ SKIPS the [] idx and reaches the
            % IDENTICAL next-partname result: image1.jpeg (logo.png contributes NO
            % number). A regression that re-breaks the numbering CRASHES here.
            %
            % Result: NO CRASH, media = {logo.png (retained), image1.jpeg}, 19/19
            % byte-identical to the frozen s0095 oracle. The input bytes are the
            % frozen python-generated s0095_input_logo.docx.
            src = testCase.dataFile('s0095_input_logo.docx');
            d = mat2doc.Document(src);               % Python: Document(input_logo)
            d.add_picture(testCase.dataFile('python-icon.jpeg'));   % Python: add_picture(JPEG)
            zipBytes = testCase.saveZip(d);
            [~, names] = zipEntryList(zipBytes);
            testCase.assertPackage(zipBytes, testCase.S0095_MANIFEST, ...
                's0095 DEFECT-1 scenario J (non-numbered logo.png + JPEG, NO CRASH)');
            % readable witness: logo.png retained, the new part numbered image1.jpeg.
            testCase.verifyTrue(any(names == "word/media/logo.png"), ...
                'DEFECT-1: the non-numbered logo.png is RETAINED');
            testCase.verifyTrue(any(names == "word/media/image1.jpeg"), ...
                'DEFECT-1: the new JPEG is numbered image1.jpeg (idx-None skipped, next number = 1)');
        end

        % =============================================================== %
        % 7. ImagePart -- from_image / lazy image / default_cx-cy / load     %
        % =============================================================== %

        function test_imagepart_from_image_and_lazy(testCase)
            % Nominal + Edge (parts/image.py): ImagePart.from_image builds a part from
            % an Image + a PackURI; sha1/filename/content_type read through; default_cx
            % (Inches, native width) = 1778000 and default_cy (Emu, native height via
            % the horz_dpi-for-height quirk) = 711200; image() lazily decodes
            % Image.from_blob on first access (the manual None-sentinel cache).
            img = mat2doc.image.Image.from_file(testCase.dataFile('python-powered.png'));
            ip = mat2doc.parts.ImagePart.from_image(img, ...
                mat2doc.opc.PackURI("/word/media/image1.png"));
            testCase.verifyClass(ip, 'mat2doc.parts.ImagePart', 'from_image -> an ImagePart');
            testCase.verifyEqual(string(ip.sha1()), ...
                "b0a1e6cf904691e6fa42bd9e72acc2b05280dc86", 'ImagePart.sha1 (blob SHA1)');
            testCase.verifyEqual(string(ip.filename()), "python-powered.png", ...
                'ImagePart.filename from the source Image');
            testCase.verifyEqual(string(ip.content_type()), "image/png", 'ImagePart.content_type');
            testCase.verifyEqual(double(ip.default_cx()), testCase.NATIVE_CX, ...
                'ImagePart.default_cx = 1778000 (native width, Inches)');
            testCase.verifyEqual(double(ip.default_cy()), testCase.NATIVE_CY, ...
                'ImagePart.default_cy = 711200 (native height via horz_dpi quirk, Emu)');
            % image() -> the |Image| (lazy on a from_blob part; here it is the stored one)
            testCase.verifyClass(ip.image(), 'mat2doc.image.Image', 'ImagePart.image -> an Image');

            % lazy decode: a part built WITHOUT an image (via load) decodes on first
            % image() access (Image.from_blob), and the sha1 identity is preserved.
            ipLazy = mat2doc.parts.ImagePart.load( ...
                mat2doc.opc.PackURI("/word/media/image9.png"), ...
                mat2doc.opc.CONTENT_TYPE.PNG, img.blob(), []);
            testCase.verifyClass(ipLazy.image(), 'mat2doc.image.Image', ...
                'ImagePart.image lazily decodes Image.from_blob on first access');
            testCase.verifyEqual(string(ipLazy.sha1()), string(ip.sha1()), ...
                'the lazily-decoded part has the SAME blob sha1 (identity preserved)');
        end

        function test_imagepart_load_generic_filename(testCase)
            % Edge (parts/image.py load / filename): ImagePart.load (the PartFactory
            % entry) builds a part with NO source Image, so filename falls back to the
            % generic "image.<ext>" derived from the partname extension. sha1 still
            % matches the blob's.
            img = mat2doc.image.Image.from_file(testCase.dataFile('python-powered.png'));
            ip2 = mat2doc.parts.ImagePart.load( ...
                mat2doc.opc.PackURI("/word/media/image7.png"), ...
                mat2doc.opc.CONTENT_TYPE.PNG, img.blob(), []);
            testCase.verifyEqual(string(ip2.filename()), "image.png", ...
                'ImagePart.load filename falls back to generic "image.png" (no source Image)');
            testCase.verifyEqual(string(ip2.sha1()), ...
                "b0a1e6cf904691e6fa42bd9e72acc2b05280dc86", ...
                'ImagePart.load sha1 matches the blob (identity)');
        end

        % =============================================================== %
        % 8. ImageParts -- dedupe / membership / numbering                  %
        % =============================================================== %

        function test_imageparts_dedupe_membership_numbering(testCase)
            % Nominal + Edge (package.py ImageParts): get_or_add_image_part dedupes by
            % SHA1 (the same source PNG twice -> the SAME handle, len_ stays 1); a
            % distinct JPEG grows to len_ 2 with partname numbering image1.png /
            % image2.jpeg (the number is unique WITHOUT regard to extension);
            % contains() is membership by handle identity; to_array() materializes the
            % ImagePart vector; append() adds directly.
            PNG  = testCase.dataFile('python-powered.png');
            JPEG = testCase.dataFile('python-icon.jpeg');
            ips = mat2doc.package.ImageParts();
            testCase.verifyEqual(ips.len_(), 0, 'a fresh ImageParts is empty');

            ip_a = ips.get_or_add_image_part(PNG);
            ip_b = ips.get_or_add_image_part(PNG);       % SAME sha1 -> dedupe
            testCase.verifyTrue(ip_a == ip_b, ...
                'dedupe: the same PNG returns the SAME ImagePart handle (SHA1 match)');
            testCase.verifyEqual(ips.len_(), 1, 'dedupe: len_ stays 1 after two adds of the same image');
            testCase.verifyEqual(string(ip_a.partname().string()), "/word/media/image1.png", ...
                'first image partname /word/media/image1.png');

            ip_c = ips.get_or_add_image_part(JPEG);      % distinct -> grows
            testCase.verifyTrue(ip_a ~= ip_c, 'a distinct JPEG is a distinct handle');
            testCase.verifyEqual(ips.len_(), 2, 'len_ 2 after a distinct image');
            testCase.verifyEqual(string(ip_c.partname().string()), "/word/media/image2.jpeg", ...
                'second image partname /word/media/image2.jpeg (number unique across extensions)');

            % contains (membership by handle identity) + to_array
            testCase.verifyTrue(ips.contains(ip_a), 'contains(ip_a) true');
            testCase.verifyTrue(ips.contains(ip_c), 'contains(ip_c) true');
            arr = ips.to_array();
            testCase.verifyEqual(numel(arr), 2, 'to_array materializes the 2 ImageParts');
            testCase.verifyClass(arr, 'mat2doc.parts.ImagePart', 'to_array is an ImagePart vector');

            % a fresh empty collection: contains -> false (empty-collection edge)
            empty = mat2doc.package.ImageParts();
            testCase.verifyFalse(empty.contains(ip_a), 'empty ImageParts contains -> false');
        end

        % =============================================================== %
        % 9. StoryPart / Run / Document -- the add_picture wiring            %
        % =============================================================== %

        function test_storypart_get_or_add_image_and_new_pic_inline(testCase)
            % Nominal (story.py): StoryPart.get_or_add_image returns [rId, Image] --
            % the body DocumentPart relates the media part as rId9 and returns the
            % decoded Image; new_pic_inline returns the CT_Inline (the P7-3 wp:inline
            % tree) that add_drawing embeds.
            PNG = testCase.dataFile('python-powered.png');
            d = mat2doc.Document();
            story = d.part();                            % the DocumentPart (a StoryPart)
            [rId, image] = story.get_or_add_image(PNG);
            testCase.verifyEqual(string(rId), "rId9", 'get_or_add_image rId is rId9');
            testCase.verifyClass(image, 'mat2doc.image.Image', 'get_or_add_image returns an Image');
            inline = story.new_pic_inline(PNG);
            testCase.verifyClass(inline, 'mat2doc.oxml.shape.CT_Inline', ...
                'new_pic_inline returns a CT_Inline (the P7-3 wp:inline tree)');
        end

        function test_run_and_document_add_picture_inline_shape(testCase)
            % Nominal (run.py / document.py): Run.add_picture and Document.add_picture
            % both return an InlineShape whose width/height are the native EMU
            % (1778000 x 711200) and whose type is PICTURE. InlineShapes over the saved
            % body reports len_ 1 and .part -> the DocumentPart (VERIFY-3).
            PNG = testCase.dataFile('python-powered.png');

            % Document.add_picture -> InlineShape
            d = mat2doc.Document();
            sh = d.add_picture(PNG);
            testCase.verifyClass(sh, 'mat2doc.shape.InlineShape', 'Document.add_picture -> InlineShape');
            testCase.verifyEqual(double(sh.width),  testCase.NATIVE_CX, 'InlineShape.width native 1778000');
            testCase.verifyEqual(double(sh.height), testCase.NATIVE_CY, 'InlineShape.height native 711200');
            testCase.verifyEqual(string(sh.type), "PICTURE", 'InlineShape.type PICTURE');

            % Run.add_picture -> InlineShape (over a live paragraph run)
            d2 = mat2doc.Document();
            p = d2.add_paragraph();
            r = p.add_run();
            sh2 = r.add_picture(PNG);
            testCase.verifyClass(sh2, 'mat2doc.shape.InlineShape', 'Run.add_picture -> InlineShape');

            % InlineShapes.part (VERIFY-3): parent story part is the DocumentPart.
            body = d.part().element().xpath("//w:body");
            shapes = mat2doc.shape.InlineShapes(body(1), d.part());
            testCase.verifyEqual(shapes.len_(), 1, 'InlineShapes.len_ = 1');
            testCase.verifyClass(shapes.part(), 'mat2doc.parts.DocumentPart', ...
                'InlineShapes.part -> the DocumentPart (VERIFY-3)');
        end

        function test_add_picture_missing_file_raises(testCase)
            % Edge (error path, IDENTIFIER-pinned): add_picture on a NON-existent path
            % fails inside Image.from_file with the faithful mat2doc:FileNotFoundError
            % (image.py 35-50) -- the python-docx FileNotFoundError analogue. Pinned by
            % IDENTIFIER (not merely "it throws"), per the Gate-4 error-path rule.
            d = mat2doc.Document();
            testCase.verifyError(@() d.add_picture(testCase.dataFile('does_not_exist.png')), ...
                'mat2doc:FileNotFoundError', ...
                'add_picture on a missing file raises mat2doc:FileNotFoundError');
        end

        % =============================================================== %
        % 10. ** M1 byte-neutrality ** (the _gather_image_parts transit)     %
        % =============================================================== %

        function test_m1_byte_neutrality(testCase)
            % Regression (byte-neutrality, L1): _gather_image_parts is LIVE on EVERY
            % open at P7-4, but default.docx has no RT.IMAGE relationship (its
            % thumbnail is a THUMBNAIL reltype, short-circuited before any append), so
            % a bare mat2doc.Document().save() STILL emits word/styles.xml at 349458 B
            % / 02d71a68 AND word/document.xml at 1548 B / 0e4dd503 (== the frozen
            % s0001 python-docx oracle). SHA equality is L1. One save emits both parts.
            d = mat2doc.Document();
            zipBytes = testCase.saveZip(d);
            [blobs, names] = zipEntryList(zipBytes);
            sty = entryBlob(blobs, names, "word/styles.xml");
            doc = entryBlob(blobs, names, "word/document.xml");
            testCase.verifyEqual(numel(sty), testCase.STYLES_SIZE_M1, ...
                'M1 styles.xml is exactly 349458 B (_gather_image_parts byte-neutral)');
            testCase.verifyEqual(sha256hex(sty), testCase.STYLES_SHA_M1, ...
                'M1 styles.xml SHA == frozen s0001 oracle (L1)');
            testCase.verifyEqual(numel(doc), testCase.DOC_SIZE_M1, ...
                'M1 document.xml is exactly 1548 B (_gather_image_parts byte-neutral)');
            testCase.verifyEqual(sha256hex(doc), testCase.DOC_SHA_M1, ...
                'M1 document.xml SHA == frozen s0001 oracle (L1)');
        end

        % =============================================================== %
        % 11. EQUIVALENCE -- the full s0098 probe vs the frozen oracle       %
        % =============================================================== %

        function test_equivalence_s0098_full_probe(testCase)
            % Equivalence: replay the ENTIRE frozen s0098 probe (runS0098Probe -- the
            % .m twin's body VERBATIM: ImagePart from_image/load; ImageParts dedupe/
            % len/partnames/contains; StoryPart get_or_add_image/new_pic_inline;
            % Run/Document add_picture -> InlineShape; InlineShapes.part) and compare
            % EVERY tagged leaf to the frozen python-docx 1.2.0 oracle copied into
            % data\s0098\probe.json. Gate-3 probe_diff found ZERO divergences (MATCH),
            % so every leaf must be byte/value-identical.
            here = fileparts(mfilename('fullpath'));
            port   = runS0098Probe(fullfile(here, 'data'));
            oracle = loadJson(fullfile(here, 'data', 's0098', 'probe.json'));

            pKeys = sort(fieldnames(port));
            oKeys = sort(fieldnames(oracle));
            testCase.verifyEqual(pKeys, oKeys, ...
                'the replayed s0098 probe and the frozen oracle must have identical keys');
            % Non-trivial floor guarding a silent-empty replay. The only looser-than-
            % byte assertion in this class; justified as a leaf-count floor.
            testCase.verifyGreaterThanOrEqual(numel(oKeys), 20, ...
                'the s0098 oracle must expose the full add_picture API probe surface');
            for i = 1:numel(oKeys)
                k = oKeys{i};
                testCase.verifyTrue(isfield(port, k), sprintf('port is missing leaf %s', k));
                testCase.verifyEqual(string(port.(k)), string(oracle.(k)), ...
                    sprintf('probe %s must be byte/value-identical to the frozen s0098 oracle', k));
            end
        end

    end

    % ===================== instance helpers ============================ %
    methods (Access = private)

        function zipBytes = saveZip(~, d)
            % d.save to a BINARY-mode temp .docx (the writer opens "wb"), read the
            % whole-zip bytes back, delete. Returns the uint8 zip vector.
            tmp = [tempname '.docx'];
            cleanup = onCleanup(@() deleteIfExists(tmp)); %#ok<NASGU>
            d.save(tmp);
            zipBytes = readBytes(tmp);
        end

        function assertPackage(testCase, zipBytes, MANIFEST, label)
            % Full-package byte pin: enumerate the saved zip in stream (write) order
            % and assert (a) the part inventory + EXACT frozen zip-entry order ==
            % MANIFEST col 1, and (b) every part's size + SHA-256 == MANIFEST cols
            % 2/3 (SHA-256 equality == byte-identity, L1 -- bit-exact for the binary
            % media blobs). Copied from Test_p5_3b_hdrftr_api.m.
            [blobs, names] = zipEntryList(zipBytes);
            want = MANIFEST(:, 1)';
            testCase.verifyEqual(numel(names), size(MANIFEST, 1), ...
                sprintf('%s: must emit exactly %d parts', label, size(MANIFEST, 1)));
            testCase.verifyEqual(names, want, ...
                sprintf('%s: zip-entry sequence must equal the frozen order', label));
            for k = 1:size(MANIFEST, 1)
                nm       = MANIFEST(k, 1);
                wantSize = str2double(MANIFEST(k, 2));
                wantSha  = MANIFEST(k, 3);
                got = entryBlob(blobs, names, nm);
                testCase.verifyEqual(numel(got), wantSize, ...
                    sprintf('%s: part %s must be exactly %d B', label, nm, wantSize));
                testCase.verifyEqual(sha256hex(got), wantSha, ...
                    sprintf('%s: part %s SHA-256 must equal the frozen oracle (byte-identical L1)', label, nm));
            end
        end

        function p = dataFile(~, name)
            here = fileparts(mfilename('fullpath'));   % tests\parts
            p = char(fullfile(here, 'data', name));
        end
    end
end

% ===================== file-local helpers ============================== %

function P = runS0098Probe(dataDir)
    % Replay the s0098 probe sequence (the .m twin's body, VERBATIM tags/inputs/
    % order) and return the struct of tagged canonical strings. Embedded here so the
    % Equivalence leg is self-contained (the validation-folder scenario is NOT on the
    % toolbox path and must not be a dependency). Mirrors
    % validation\mat2doc\scenarios\s0098_p7_4_api_probe.m. The absolute source-image
    % paths in the .m twin are replaced by the co-located data\ copies (byte-exact).
    %
    % NB: NO `import mat2doc...` here -- a specific (non-.*) import is resolved at
    % suite-CREATION PARSE time, BEFORE TestClassSetup's PathFixture puts +mat2doc on
    % the path (the "specific import fails in test class" lesson). Fully qualified
    % everywhere, which resolves at RUN time (after PathFixture).
    PNG  = char(fullfile(dataDir, 'python-powered.png'));
    JPEG = char(fullfile(dataDir, 'python-icon.jpeg'));

    P = struct();

    % --- 1. ImagePart.from_image + lazy image / default_cx-cy ---
    img = mat2doc.image.Image.from_file(PNG);
    ip = mat2doc.parts.ImagePart.from_image( ...
        img, mat2doc.opc.PackURI("/word/media/image1.png"));
    P.ip_sha1 = cs(ip.sha1());
    P.ip_filename = cs(ip.filename());
    P.ip_content_type = cs(ip.content_type());
    P.ip_default_cx = ci(ip.default_cx());
    P.ip_default_cy = ci(ip.default_cy());
    P.ip_image_class = cs(basename_(ip.image()));

    % --- 2. ImagePart.load -> generic filename, sha1 identity ---
    ip2 = mat2doc.parts.ImagePart.load( ...
        mat2doc.opc.PackURI("/word/media/image7.png"), ...
        mat2doc.opc.CONTENT_TYPE.PNG, img.blob(), []);
    P.ipload_filename = cs(ip2.filename());
    P.ipload_sha1_eq = cb(ip2.sha1() == ip.sha1());

    % --- 3. Package.get_or_add_image_part + ImageParts dedupe ---
    d = mat2doc.Document();
    pkg = d.part().package();
    ip_a = pkg.get_or_add_image_part(PNG);
    ip_b = pkg.get_or_add_image_part(PNG);
    P.dedupe_same_handle = cb(ip_a == ip_b);
    P.dedupe_len_after2 = ci(pkg.image_parts().len_());
    ip_c = pkg.get_or_add_image_part(JPEG);
    P.dedupe_len_after3 = ci(pkg.image_parts().len_());
    P.dedupe_distinct = cb(ip_a ~= ip_c);
    P.p1_partname = cs(ip_a.partname().string());
    P.p3_partname = cs(ip_c.partname().string());
    P.ip_contains_a = cb(pkg.image_parts().contains(ip_a));

    % --- 4. StoryPart.get_or_add_image / new_pic_inline ---
    d2 = mat2doc.Document();
    story = d2.part();
    [rId, image] = story.get_or_add_image(PNG);
    P.story_rid = cs(rId);
    P.story_image_class = cs(basename_(image));
    inline = story.new_pic_inline(PNG);
    P.inline_class = cs(basename_(inline));

    % --- 5. Document.add_picture -> InlineShape ---
    d3 = mat2doc.Document();
    sh = d3.add_picture(PNG);
    P.shape_class = cs(basename_(sh));
    P.shape_width = ci(sh.width);
    P.shape_height = ci(sh.height);
    P.shape_type_name = cs(string(sh.type));

    % --- 6. InlineShapes.part (VERIFY-3) ---
    bodies = d3.part().element().xpath("//w:body");
    body = bodies(1);
    shapes = mat2doc.shape.InlineShapes(body, d3.part());
    P.shapes_len = ci(shapes.len_());
    P.shapes_part_class = cs(basename_(shapes.part()));
end

function s = basename_(x)
    c = split(string(class(x)), ".");
    s = c(end);
end

function s = ci(v)
    % canonical int tag "int|<decimal>" (matches the s0098 oracle format)
    s = "int|" + string(int64(double(v)));
end

function s = cs(v)
    s = "str|" + string(v);
end

function s = cb(v)
    s = "bool|" + string(logical(v));   % "true"/"false"
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

function [blobs, names] = zipEntryList(zipBytes)
    % Enumerate an in-memory zip in stream (write) order -> ordered {names, blobs}.
    % java.util.zip.ZipInputStream reads local file headers in physical order, so
    % `names` is the true zip-entry write sequence. Kept file-local so the order pin
    % is independent of the reader under test. (Copied from Test_p5_3b_hdrftr_api.m.)
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

function blob = entryBlob(blobs, names, membername)
    i = find(names == string(membername), 1);
    assert(~isempty(i), 'zip entry not found: %s', membername);
    blob = blobs{i};
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

function h = sha256hex(bytes)
    % Lowercase hex SHA-256 of a uint8 vector (base MATLAB / Java MessageDigest).
    md  = java.security.MessageDigest.getInstance('SHA-256');
    dig = md.digest(typecast(uint8(bytes(:)'), 'int8'));
    b   = typecast(dig, 'uint8');
    h   = lower(string(sprintf('%02x', b)));
end

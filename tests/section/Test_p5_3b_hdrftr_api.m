classdef Test_p5_3b_hdrftr_api < matlab.unittest.TestCase
% TEST_P5_3B_HDRFTR_API  Gate-4 permanent unit tests for Mat2Doc P5-3b
%   (the _Header/_Footer proxy tier + the hdr/ftr SEPARATE-PART wiring --
%   risk-register #4, the cross-part hazard; the FIRST runtime-ADDED parts in
%   Mat2Doc). src/docx/section.py::{_BaseHeaderFooter,_Header,_Footer} ->
%   +mat2doc\+section\{BaseHeaderFooter_,Header_,Footer_}; the six Section hdr/ftr
%   members; src/docx/parts/hdrftr.py::{HeaderPart,FooterPart} ->
%   +mat2doc\+parts\{HeaderPart,FooterPart}; the five DocumentPart un-stubs
%   (add_header_part/add_footer_part/header_part/footer_part/drop_header_part);
%   the PartFactory WML_HEADER/WML_FOOTER flip; the C3 BlockItemContainer element
%   seam; and the byte-identical header/footer templates.
%
%   P5-3b is the FIRST WP that CREATES parts at runtime (a header/footer is a
%   separate word/headerN.xml / word/footerN.xml part with its own rId +
%   [Content_Types] Override + document.xml headerReference). Equivalence is
%   therefore FULL-PACKAGE byte identity (not just document.xml): every added
%   part, every new rel, every new Override, the zip-entry order, AND the drop
%   path leaving zero residue. This class permanently FREEZES that surface --
%   byte-identical to python-docx v1.2.0.
%
%   THE HIGHEST-VALUE PERMANENT PINS IN THIS WP (a regression goes RED loudly):
%     * ** THE P5 HEADLINE ** header full-package byte pin
%       (test_header_full_package_byte_pin): Document(); sections[0].header
%       unlinked; paragraphs[0].text="Hello, Header!"; save() -> ALL 18 parts
%       byte-identical to the frozen s0052 python-docx oracle. header1.xml ==
%       1329 B / 3c75200f, [Content_Types].xml == 4a3612de (Override added),
%       document.xml.rels == 29b4d0f4 (rId9 header rel), document.xml == e3b5e387
%       (headerReference), + the other 14 parts unchanged.
%     * Variant full-package byte pins: footer-only (s0053, footer1.xml
%       d4a3334d); BOTH header+footer (s0054, 19 parts, rId9 header + rId10
%       footer); even+first-page headers + titlePg (s0055, 20 parts, header1/2/3,
%       document.xml 79022ccd carrying w:type even/first); the add->drop TOGGLE
%       (is_linked_to_previous False->True -> package BACK to the pristine 17-part
%       default == s0001, zero residue).
%     * Reload byte pin (test_reload_roundtrip_byte_pin): open the frozen
%       python-generated header+footer s0058 _source.docx in Mat2Doc -> save() ->
%       19/19 byte-identical to python-docx's OWN open/save (py_roundtrip). This
%       is the ONLY leg that exercises the PartFactory WML_HEADER->HeaderPart.load
%       / WML_FOOTER->FooterPart.load flip on LOAD (the inherited-static-trap
%       path, untouched by every add-side scenario) -- header1 237dfed9
%       "Reloaded Header", footer1 9e92124a "Reloaded Footer".
%     * C3 seam byte-neutrality pins: M1 Document().save() document.xml == 1548 B
%       / 0e4dd503; M2 add_heading x3 + add_paragraph document.xml == 1865 B /
%       a71e5502 (== frozen s0033). These guard the live base-class refactor
%       (property->protected element_() method + element_store_ rename).
%     * Section proxies + WD_HEADER_FOOTER wiring: all six members ->
%       Header_/Footer_ (header/footer @lazyproperty CACHED same-handle;
%       even/first FRESH each access); is_linked_to_previous get/set (add + drop);
%       the 3-section inherit-walk resolution ([T,T,T]->[T,F,T]->[F,F,T]; sec2/sec1
%       share header1 via case-2 recursion; sec0 gets header2 via case-3 add).
%     * HeaderPart/FooterPart: new -> header1.xml/footer1.xml + WML_HEADER/
%       WML_FOOTER; the inherited-static load dispatches to the subclass.
%     * DocumentPart un-stubs: add_header_part/add_footer_part -> [part, rId];
%       header_part(rId)/footer_part(rId) round-trip; drop_header_part; RT.HEADER/
%       RT.FOOTER constants.
%
%   Provenance (Gate-1..3, all 2026-07-31):
%     * Audit    : validation\mat2doc\audit_P5-3b_hdrftr_api.md (Porter Gate-1 +
%                  Fable/mso-auditor Gate-2 APPROVE -- 9 full-package byte scenarios).
%     * Validate : validation\mat2doc\validate_P5-3b_hdrftr_api.md (Gate-3 PASS --
%                  ZERO new D-numbers; header G-scenario 18/18 L1 re-derived; footer
%                  18/18; both 19/19 (rId9/rId10); even+first 20/20 (79022ccd);
%                  add->drop toggle 17/17 == pristine s0001; probe_diff s0057 MATCH
%                  over the full Section hdr/ftr surface + 3-section inherit-walk;
%                  RELOAD 19/19 L1 + probe MATCH (HeaderPart.load/FooterPart.load
%                  byte-proven); C3 M1 17/17 (0e4dd503) + M2 17/17 (a71e5502)
%                  re-derived byte-neutral).
%     * Scenarios: validation\mat2doc\scenarios\s0052..s0058_p5_3b_*.{py,m} (the
%                  IDENTICAL-sequence byte/probe twins; their build bodies are
%                  replayed VERBATIM by the emit helpers + runS0057Probe() below).
%     * Frozen refs (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE at Gate-3):
%         references\s0052..s0055\, s0058\ -- the full-package manifests
%           (name|size|sha256 in frozen zip-entry order) are embedded below as the
%           S00xx_MANIFEST constants so this suite is self-contained; SHA-256
%           equality IS byte-identity (L1). references\s0001\ (M1 + toggle oracle)
%           and references\s0033\ (M2 oracle) reused as the C3 guards.
%         tests\section\data\s0058_source.docx -- the python-generated header+footer
%           package (references\s0058\_source.docx), copied byte-for-byte (co-located
%           `*.docx binary` .gitattributes) as the INPUT for the reload leg.
%         tests\section\data\s0057_hdrftr_probe_oracle.json -- the frozen s0057
%           probe oracle (value JSON; jsondecode line-ending agnostic -> no binary
%           pin, s0045 precedent), for the Equivalence leg.
%     * Template : +mat2doc\templates\default.docx + default-header.xml /
%                  default-footer.xml -- ship in the toolbox, so the suite is
%                  self-contained relative to the worktree.
%
%   Coverage taxonomy
%   -----------------
%   * Nominal      -- the header/footer G-scenario (add + set text + save); the
%                     six Section members -> correct Header_/Footer_; is_linked
%                     get/set/get round-trip; HeaderPart/FooterPart.new; the
%                     DocumentPart add/get/drop un-stubs.
%   * Edge         -- footer-only + both + even/first variants; the add->drop
%                     TOGGLE (zero residue); the RELOAD path (LOAD-side flip); the
%                     3-section inherit-walk (case-2 recursion vs case-3 add);
%                     lazy caching (header cached, even/first fresh); non-ASCII is
%                     covered by the equivalence probe replay; drop_header_part
%                     rel removal (header_part(rId) then raises mat2doc:KeyError).
%   * Equivalence  -- test_equivalence_full_probe_vs_frozen_oracle replays the
%                     ENTIRE s0057 probe (kinds / default_linked / roundtrip /
%                     paragraphs / header_partname / the inherit-walk) and
%                     flatten-compares every leaf to the frozen python-docx 1.2.0
%                     oracle (Gate-3 probe_diff found ZERO divergences).
%   * Regression   -- hard-coded full-package SHA-256 (+ size) byte pins for
%                     s0052/s0053/s0054/s0055/s0058, the s0056 toggle == s0001, and
%                     the C3 M1/M2 document.xml SHAs.
%   * Upstream     -- the header/footer template bytes, the rId numbering
%                     (rId9/rId10/rId11 in insertion order, H11 no-sort), the
%                     partname numbering (headerN/footerN via next_partname), the
%                     Override insertion, the headerReference w:type wiring, the
%                     inherit-walk case-2/case-3 resolution, and the round-trip
%                     re-serialization ARE the python-docx section.py /
%                     parts/hdrftr.py contract; the frozen oracle IS lxml's
%                     expected output for these sequences.
%
%   Byte-level (L1) note: every full-package assertion is a SHA-256 (+ size) pin of
%   the raw shipping bytes of each zip part, in frozen zip-entry order. SHA-256
%   equality == byte identity (L1). NO D-number granted any L2 relaxation in this
%   WP (Gate-3: ZERO new), so every byte pin is L1. The equivalence leaf-key-count
%   floor is the only looser-than-byte check and is commented at its site.
%
%   Determinism: no network, no absolute paths. The frozen source docx + oracle
%   json resolve relative to this file via fileparts(mfilename('fullpath')); the
%   template ships in the toolbox; saves go to tempname .docx deleted via
%   onCleanup; every file read is binary ('r','n'). The +mat2doc package resolves
%   via the MANDATORY PathFixture(worktree-root) in TestClassSetup (WP9-F4 lesson).

    properties (Constant)
        HEADER_CLS = 'mat2doc.section.Header_'
        FOOTER_CLS = 'mat2doc.section.Footer_'

        % ---- shared unchanging parts (byte-identical across s0052..s0058 and to
        %      the M1 default) -- kept once so the per-scenario manifests read as
        %      the header/footer DELTA over the frozen default. ----

        % ==== s0052 -- header G-scenario, 18 parts (frozen zip-entry order) ====
        % ** THE P5 HEADLINE BYTE ORACLE ** (validate_P5-3b section 1).
        S0052_MANIFEST = [ ...
            "[Content_Types].xml",             "1866",   "4a3612de6729d64f3a8cec6a0b61764cb396b8a48fc1b607f905437ce55b1e4a"; ...
            "_rels/.rels",                     "734",    "a9ae57efe9186f07d48303bcac5d54c7e359bf3f503939c03ad2954e5c59a5c4"; ...
            "docProps/core.xml",               "721",    "d14be8284e406d14dc2a576b19f15e7f637b955f9e27a3d9ed871ec99606f7ba"; ...
            "docProps/app.xml",                "1132",   "be664981c3141cddfc59362beb287ebf20d0773660e2dd6faac5968a5930a081"; ...
            "word/document.xml",               "1597",   "e3b5e38706b3bf9de625e5f684a2c01ae4330727c2c91518c39b72acbd4cfc44"; ...
            "word/_rels/document.xml.rels",    "1355",   "29b4d0f49e573a7cc8901b470fe1de2d453ebd7edbb67cc53b806120c1f978c2"; ...
            "word/styles.xml",                 "349458", "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"; ...
            "word/stylesWithEffects.xml",      "438131", "463ae0928cf0d84775dbf8cf18d6c3029f6707c81bf590f6d6dd8757a5e93f15"; ...
            "word/settings.xml",               "2535",   "51a0d348fe85965c66e4748a03c3c0d055d78455514f03cb121c334af7d73689"; ...
            "word/webSettings.xml",            "438",    "349d36de7434d09f86987ff671d8814964a0588c1e630c06e562cda7e75e9f95"; ...
            "word/fontTable.xml",              "2811",   "79385fb7f60247507ecaffc292e9ebd52ea0657b8634f629ba6fccc54011d6bb"; ...
            "word/theme/theme1.xml",           "10939",  "e3a8ab7db9ca7afca56f5f2820a56e8b660016c647773555b060b0a02ac76941"; ...
            "customXml/item1.xml",             "262",    "a86086ffc5d8e83ebd6c71a55d1d2efaa31b137977f5f3a752366e1023612144"; ...
            "customXml/_rels/item1.xml.rels",  "295",    "1ca6c9a64edcebe24ee703a54403611b322d96da33371779e742d2d3f7ed7a6c"; ...
            "customXml/itemProps1.xml",        "354",    "c542307b13ec29a8b546217bb37936ab4822e044b265d2952985ec3d6afed24e"; ...
            "word/numbering.xml",              "5513",   "70976f19cbcd896e51890859fe6ecb3467a5a7ad0c040160fedc5a1993cb09ce"; ...
            "word/header1.xml",                "1329",   "3c75200fabed242609e4d3a2032d62e24a74c6c428a6ffb78187e1075ed48365"; ...
            "docProps/thumbnail.jpeg",         "8324",   "96367138dc44ce09bf2c8f0f8e49348a1478d2c5c0af69bbc2bbc38b63cdcead"];

        % ==== s0053 -- footer G-scenario, 18 parts ====
        S0053_MANIFEST = [ ...
            "[Content_Types].xml",             "1866",   "98da1c3b53db2295e8bf7570d35f0c59f98f650930a9b61be218fb8b9a04a101"; ...
            "_rels/.rels",                     "734",    "a9ae57efe9186f07d48303bcac5d54c7e359bf3f503939c03ad2954e5c59a5c4"; ...
            "docProps/core.xml",               "721",    "d14be8284e406d14dc2a576b19f15e7f637b955f9e27a3d9ed871ec99606f7ba"; ...
            "docProps/app.xml",                "1132",   "be664981c3141cddfc59362beb287ebf20d0773660e2dd6faac5968a5930a081"; ...
            "word/document.xml",               "1597",   "d8fac257e3af9530717ec94f57395b0d95962744a56355a84fada4d008c4927d"; ...
            "word/_rels/document.xml.rels",    "1355",   "125a941674a26403972a869dd60fcdb82da089ff35ac3b297be07240db44afd2"; ...
            "word/styles.xml",                 "349458", "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"; ...
            "word/stylesWithEffects.xml",      "438131", "463ae0928cf0d84775dbf8cf18d6c3029f6707c81bf590f6d6dd8757a5e93f15"; ...
            "word/settings.xml",               "2535",   "51a0d348fe85965c66e4748a03c3c0d055d78455514f03cb121c334af7d73689"; ...
            "word/webSettings.xml",            "438",    "349d36de7434d09f86987ff671d8814964a0588c1e630c06e562cda7e75e9f95"; ...
            "word/fontTable.xml",              "2811",   "79385fb7f60247507ecaffc292e9ebd52ea0657b8634f629ba6fccc54011d6bb"; ...
            "word/theme/theme1.xml",           "10939",  "e3a8ab7db9ca7afca56f5f2820a56e8b660016c647773555b060b0a02ac76941"; ...
            "customXml/item1.xml",             "262",    "a86086ffc5d8e83ebd6c71a55d1d2efaa31b137977f5f3a752366e1023612144"; ...
            "customXml/_rels/item1.xml.rels",  "295",    "1ca6c9a64edcebe24ee703a54403611b322d96da33371779e742d2d3f7ed7a6c"; ...
            "customXml/itemProps1.xml",        "354",    "c542307b13ec29a8b546217bb37936ab4822e044b265d2952985ec3d6afed24e"; ...
            "word/numbering.xml",              "5513",   "70976f19cbcd896e51890859fe6ecb3467a5a7ad0c040160fedc5a1993cb09ce"; ...
            "word/footer1.xml",                "1329",   "d4a3334dcdb463196a1a156205bd1ebbaa81bef0ae088bdbb4a327668bbd5832"; ...
            "docProps/thumbnail.jpeg",         "8324",   "96367138dc44ce09bf2c8f0f8e49348a1478d2c5c0af69bbc2bbc38b63cdcead"];

        % ==== s0054 -- BOTH header+footer, 19 parts (two runtime parts) ====
        S0054_MANIFEST = [ ...
            "[Content_Types].xml",             "1994",   "479ad1a2258771db93540fbde6ad43e861d8d39dd2559e3f6db17286b11cc6ab"; ...
            "_rels/.rels",                     "734",    "a9ae57efe9186f07d48303bcac5d54c7e359bf3f503939c03ad2954e5c59a5c4"; ...
            "docProps/core.xml",               "721",    "d14be8284e406d14dc2a576b19f15e7f637b955f9e27a3d9ed871ec99606f7ba"; ...
            "docProps/app.xml",                "1132",   "be664981c3141cddfc59362beb287ebf20d0773660e2dd6faac5968a5930a081"; ...
            "word/document.xml",               "1647",   "edd4618049dde9f4fbe42a870d537572c481093809d33360872b221cf800ea19"; ...
            "word/_rels/document.xml.rels",    "1484",   "157ed86827865b095c9c6b39e7c3451315a4b0d6127daa8636026684d8b9df89"; ...
            "word/styles.xml",                 "349458", "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"; ...
            "word/stylesWithEffects.xml",      "438131", "463ae0928cf0d84775dbf8cf18d6c3029f6707c81bf590f6d6dd8757a5e93f15"; ...
            "word/settings.xml",               "2535",   "51a0d348fe85965c66e4748a03c3c0d055d78455514f03cb121c334af7d73689"; ...
            "word/webSettings.xml",            "438",    "349d36de7434d09f86987ff671d8814964a0588c1e630c06e562cda7e75e9f95"; ...
            "word/fontTable.xml",              "2811",   "79385fb7f60247507ecaffc292e9ebd52ea0657b8634f629ba6fccc54011d6bb"; ...
            "word/theme/theme1.xml",           "10939",  "e3a8ab7db9ca7afca56f5f2820a56e8b660016c647773555b060b0a02ac76941"; ...
            "customXml/item1.xml",             "262",    "a86086ffc5d8e83ebd6c71a55d1d2efaa31b137977f5f3a752366e1023612144"; ...
            "customXml/_rels/item1.xml.rels",  "295",    "1ca6c9a64edcebe24ee703a54403611b322d96da33371779e742d2d3f7ed7a6c"; ...
            "customXml/itemProps1.xml",        "354",    "c542307b13ec29a8b546217bb37936ab4822e044b265d2952985ec3d6afed24e"; ...
            "word/numbering.xml",              "5513",   "70976f19cbcd896e51890859fe6ecb3467a5a7ad0c040160fedc5a1993cb09ce"; ...
            "word/header1.xml",                "1316",   "f43c0d348130ecb487ec76b92f10218bb2921df6c2908ed69e386aeaf7f51178"; ...
            "word/footer1.xml",                "1316",   "423cfd57d37ebccc51f7a46c2e48a7d97ba038735ef31727941ad0ee31104041"; ...
            "docProps/thumbnail.jpeg",         "8324",   "96367138dc44ce09bf2c8f0f8e49348a1478d2c5c0af69bbc2bbc38b63cdcead"];

        % ==== s0055 -- even+first-page headers + titlePg, 20 parts (3 header parts) ====
        S0055_MANIFEST = [ ...
            "[Content_Types].xml",             "2122",   "b2d6874d007d0dc4499d89e15b6bacd090424ed1531b812a16d4bd55dfd2f48d"; ...
            "_rels/.rels",                     "734",    "a9ae57efe9186f07d48303bcac5d54c7e359bf3f503939c03ad2954e5c59a5c4"; ...
            "docProps/core.xml",               "721",    "d14be8284e406d14dc2a576b19f15e7f637b955f9e27a3d9ed871ec99606f7ba"; ...
            "docProps/app.xml",                "1132",   "be664981c3141cddfc59362beb287ebf20d0773660e2dd6faac5968a5930a081"; ...
            "word/document.xml",               "1704",   "79022ccded9d84605ae2d53e15528c96e292000a68a4fe994bcf11fb79d23ac5"; ...
            "word/_rels/document.xml.rels",    "1613",   "2ce2029186079705509f8658dfa8937a1ba63d9940219815905de30518040c1f"; ...
            "word/styles.xml",                 "349458", "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"; ...
            "word/stylesWithEffects.xml",      "438131", "463ae0928cf0d84775dbf8cf18d6c3029f6707c81bf590f6d6dd8757a5e93f15"; ...
            "word/settings.xml",               "2535",   "51a0d348fe85965c66e4748a03c3c0d055d78455514f03cb121c334af7d73689"; ...
            "word/webSettings.xml",            "438",    "349d36de7434d09f86987ff671d8814964a0588c1e630c06e562cda7e75e9f95"; ...
            "word/fontTable.xml",              "2811",   "79385fb7f60247507ecaffc292e9ebd52ea0657b8634f629ba6fccc54011d6bb"; ...
            "word/theme/theme1.xml",           "10939",  "e3a8ab7db9ca7afca56f5f2820a56e8b660016c647773555b060b0a02ac76941"; ...
            "customXml/item1.xml",             "262",    "a86086ffc5d8e83ebd6c71a55d1d2efaa31b137977f5f3a752366e1023612144"; ...
            "customXml/_rels/item1.xml.rels",  "295",    "1ca6c9a64edcebe24ee703a54403611b322d96da33371779e742d2d3f7ed7a6c"; ...
            "customXml/itemProps1.xml",        "354",    "c542307b13ec29a8b546217bb37936ab4822e044b265d2952985ec3d6afed24e"; ...
            "word/numbering.xml",              "5513",   "70976f19cbcd896e51890859fe6ecb3467a5a7ad0c040160fedc5a1993cb09ce"; ...
            "word/header1.xml",                "1322",   "a18e175adc361682b0ee3f5754814fcd8fcc8ded1a89d7b630db356008bbcd4e"; ...
            "word/header2.xml",                "1319",   "15034863aee287100f979d6a11d3cc4c1f98dbb4ebc721ee005c46a1440d72aa"; ...
            "word/header3.xml",                "1320",   "a7edde6aa3161db2814a002e6528bb3127843a5659f07249b3d0840f8ec90ff9"; ...
            "docProps/thumbnail.jpeg",         "8324",   "96367138dc44ce09bf2c8f0f8e49348a1478d2c5c0af69bbc2bbc38b63cdcead"];

        % ==== s0058 -- RELOAD round-trip oracle (py_roundtrip), 19 parts ====
        % The frozen py_roundtrip.docx == python-docx's OWN Document(_source).save().
        % Shares [Content_Types]/document.xml/rels with s0054 (same wiring); the
        % header1/footer1 bytes differ (Reloaded Header/Footer text).
        S0058_MANIFEST = [ ...
            "[Content_Types].xml",             "1994",   "479ad1a2258771db93540fbde6ad43e861d8d39dd2559e3f6db17286b11cc6ab"; ...
            "_rels/.rels",                     "734",    "a9ae57efe9186f07d48303bcac5d54c7e359bf3f503939c03ad2954e5c59a5c4"; ...
            "docProps/core.xml",               "721",    "d14be8284e406d14dc2a576b19f15e7f637b955f9e27a3d9ed871ec99606f7ba"; ...
            "docProps/app.xml",                "1132",   "be664981c3141cddfc59362beb287ebf20d0773660e2dd6faac5968a5930a081"; ...
            "word/document.xml",               "1647",   "edd4618049dde9f4fbe42a870d537572c481093809d33360872b221cf800ea19"; ...
            "word/_rels/document.xml.rels",    "1484",   "157ed86827865b095c9c6b39e7c3451315a4b0d6127daa8636026684d8b9df89"; ...
            "word/styles.xml",                 "349458", "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"; ...
            "word/stylesWithEffects.xml",      "438131", "463ae0928cf0d84775dbf8cf18d6c3029f6707c81bf590f6d6dd8757a5e93f15"; ...
            "word/settings.xml",               "2535",   "51a0d348fe85965c66e4748a03c3c0d055d78455514f03cb121c334af7d73689"; ...
            "word/webSettings.xml",            "438",    "349d36de7434d09f86987ff671d8814964a0588c1e630c06e562cda7e75e9f95"; ...
            "word/fontTable.xml",              "2811",   "79385fb7f60247507ecaffc292e9ebd52ea0657b8634f629ba6fccc54011d6bb"; ...
            "word/theme/theme1.xml",           "10939",  "e3a8ab7db9ca7afca56f5f2820a56e8b660016c647773555b060b0a02ac76941"; ...
            "customXml/item1.xml",             "262",    "a86086ffc5d8e83ebd6c71a55d1d2efaa31b137977f5f3a752366e1023612144"; ...
            "customXml/_rels/item1.xml.rels",  "295",    "1ca6c9a64edcebe24ee703a54403611b322d96da33371779e742d2d3f7ed7a6c"; ...
            "customXml/itemProps1.xml",        "354",    "c542307b13ec29a8b546217bb37936ab4822e044b265d2952985ec3d6afed24e"; ...
            "word/numbering.xml",              "5513",   "70976f19cbcd896e51890859fe6ecb3467a5a7ad0c040160fedc5a1993cb09ce"; ...
            "word/header1.xml",                "1330",   "237dfed9488df8e6b47f3b355dfe0f525b10c31828a91b8347d6608a6ebe2104"; ...
            "word/footer1.xml",                "1330",   "9e92124a0c3e34ed638b1d722e17d8b5d0f220881309707c4b8df78732067a37"; ...
            "docProps/thumbnail.jpeg",         "8324",   "96367138dc44ce09bf2c8f0f8e49348a1478d2c5c0af69bbc2bbc38b63cdcead"];

        % ==== s0001 -- the pristine 17-part M1 default (toggle target + C3 M1 guard) ====
        % Identical to Test_p1_8_skeleton_m1.M1_MANIFEST (same frozen oracle).
        S0001_MANIFEST = [ ...
            "[Content_Types].xml",             "1738",   "66c84fb7a6aa3c4ead49f895e4a7044df1fb57de1ed76d09b2686e91f5bed5b4"; ...
            "_rels/.rels",                     "734",    "a9ae57efe9186f07d48303bcac5d54c7e359bf3f503939c03ad2954e5c59a5c4"; ...
            "docProps/core.xml",               "721",    "d14be8284e406d14dc2a576b19f15e7f637b955f9e27a3d9ed871ec99606f7ba"; ...
            "docProps/app.xml",                "1132",   "be664981c3141cddfc59362beb287ebf20d0773660e2dd6faac5968a5930a081"; ...
            "word/document.xml",               "1548",   "0e4dd503bc095219cd24e8312b372c588ca56bb834aacac53eafb7bf47836327"; ...
            "word/_rels/document.xml.rels",    "1227",   "1e7f0eb144a98e199249314f61ff32a8de2a27e56d8e9ee6b524e1c6b235d377"; ...
            "word/styles.xml",                 "349458", "02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384"; ...
            "word/stylesWithEffects.xml",      "438131", "463ae0928cf0d84775dbf8cf18d6c3029f6707c81bf590f6d6dd8757a5e93f15"; ...
            "word/settings.xml",               "2535",   "51a0d348fe85965c66e4748a03c3c0d055d78455514f03cb121c334af7d73689"; ...
            "word/webSettings.xml",            "438",    "349d36de7434d09f86987ff671d8814964a0588c1e630c06e562cda7e75e9f95"; ...
            "word/fontTable.xml",              "2811",   "79385fb7f60247507ecaffc292e9ebd52ea0657b8634f629ba6fccc54011d6bb"; ...
            "word/theme/theme1.xml",           "10939",  "e3a8ab7db9ca7afca56f5f2820a56e8b660016c647773555b060b0a02ac76941"; ...
            "customXml/item1.xml",             "262",    "a86086ffc5d8e83ebd6c71a55d1d2efaa31b137977f5f3a752366e1023612144"; ...
            "customXml/_rels/item1.xml.rels",  "295",    "1ca6c9a64edcebe24ee703a54403611b322d96da33371779e742d2d3f7ed7a6c"; ...
            "customXml/itemProps1.xml",        "354",    "c542307b13ec29a8b546217bb37936ab4822e044b265d2952985ec3d6afed24e"; ...
            "word/numbering.xml",              "5513",   "70976f19cbcd896e51890859fe6ecb3467a5a7ad0c040160fedc5a1993cb09ce"; ...
            "docProps/thumbnail.jpeg",         "8324",   "96367138dc44ce09bf2c8f0f8e49348a1478d2c5c0af69bbc2bbc38b63cdcead"];

        % ---- C3 element-seam byte-neutrality single-part guards ----
        DOC_SIZE_M1 = 1548
        DOC_SHA_M1  = "0e4dd503bc095219cd24e8312b372c588ca56bb834aacac53eafb7bf47836327"
        DOC_SIZE_M2 = 1865
        DOC_SHA_M2  = "a71e550253b8c6c9f472c740b13f6e184b29bf7ac7e5694dcdfba00ecbef7c2c"
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied from tests\section\Test_p5_3a_sections_api.m. here is
            % tests\section; the worktree root is two levels up.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\section
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. ** THE P5 HEADLINE ** header full-package byte pin (s0052)     %
        % =============================================================== %

        function test_header_full_package_byte_pin(testCase)
            % Regression (L1, THE headline P5 pin): the exact Gate-3 G-scenario --
            %   Document(); h = sections[0].header; h.is_linked_to_previous = False;
            %   h.paragraphs[0].text = "Hello, Header!"; save()
            % emits ALL 18 parts byte-identical (size + SHA-256, in frozen
            % zip-entry order) to the frozen s0052 python-docx oracle. The four
            % load-bearing parts -- header1.xml (NEW, 3c75200f), [Content_Types].xml
            % (Override added, 4a3612de), document.xml.rels (rId9 header rel,
            % 29b4d0f4), document.xml (headerReference, e3b5e387) -- plus the 14
            % unchanged parts. NEW word/header1.xml lands AFTER word/numbering.xml,
            % BEFORE docProps/thumbnail.jpeg (H11 zip-order pin). RED on ANY
            % single-byte / rId / Override / zip-order / header-reference drift.
            d = mat2doc.Document();
            h = d.sections.getitem_(0).header;             % Python: sections[0].header
            h.is_linked_to_previous = false;               % Python: = False
            ps = h.paragraphs;                             % Python: h.paragraphs[0]
            ps(1).text = "Hello, Header!";                 % Python: ...text = "Hello, Header!"
            zipBytes = testCase.saveZip(d);
            testCase.assertPackage(zipBytes, testCase.S0052_MANIFEST, ...
                's0052 header G-scenario (** THE P5 HEADLINE **)');
        end

        % =============================================================== %
        % 2. Variant full-package byte pins (footer / both / even+first)   %
        % =============================================================== %

        function test_footer_full_package_byte_pin(testCase)
            % Regression (L1): footer-only G-scenario -> 18/18 byte-identical to the
            % frozen s0053 oracle. footer1.xml == 1329 B / d4a3334d, footer Override,
            % rId9 footer rel, footerReference in document.xml.
            d = mat2doc.Document();
            f = d.sections.getitem_(0).footer;             % Python: sections[0].footer
            f.is_linked_to_previous = false;
            ps = f.paragraphs;
            ps(1).text = "Hello, Footer!";
            zipBytes = testCase.saveZip(d);
            testCase.assertPackage(zipBytes, testCase.S0053_MANIFEST, ...
                's0053 footer G-scenario');
        end

        function test_header_footer_both_full_package_byte_pin(testCase)
            % Regression (L1, cross-part): ONE section grows TWO runtime parts --
            % header1.xml (rId9) then footer1.xml (rId10) -> 19/19 byte-identical to
            % the frozen s0054 oracle. Two [Content_Types] Overrides, two rels in
            % INSERTION order (rId9 header before rId10 footer, H11 no-sort). The
            % readable rel-order witness backs up the SHA pin.
            d = mat2doc.Document();
            s = d.sections.getitem_(0);
            h = s.header;
            h.is_linked_to_previous = false;
            hp = h.paragraphs; hp(1).text = "H";
            f = s.footer;
            f.is_linked_to_previous = false;
            fp = f.paragraphs; fp(1).text = "F";
            zipBytes = testCase.saveZip(d);
            [blobs, names] = zipEntryList(zipBytes);
            testCase.assertPackage(zipBytes, testCase.S0054_MANIFEST, ...
                's0054 BOTH header+footer (two runtime parts)');
            % readable cross-part witness: rId9 -> header1.xml, rId10 -> footer1.xml,
            % rId9 BEFORE rId10 (insertion order, no rId sort -- H11).
            rels = char(entryBlob(blobs, names, "word/_rels/document.xml.rels"));
            i9 = strfind(rels, 'Id="rId9"');   i10 = strfind(rels, 'Id="rId10"');
            testCase.verifyNotEmpty(i9,  'rId9 (header rel) must be present');
            testCase.verifyNotEmpty(i10, 'rId10 (footer rel) must be present');
            testCase.verifyLessThan(i9(1), i10(1), ...
                'rId9 (header) must precede rId10 (footer) -- insertion order, no sort');
            testCase.verifyTrue(contains(rels, 'Target="header1.xml"'), ...
                'rId9 must target header1.xml');
            testCase.verifyTrue(contains(rels, 'Target="footer1.xml"'), ...
                'rId10 must target footer1.xml');
        end

        function test_even_first_headers_full_package_byte_pin(testCase)
            % Regression (L1): one section with different_first_page_header_footer
            % grows THREE distinct header parts -- header1 (PRIMARY, rId9), header2
            % (EVEN_PAGE, rId10, w:type="even"), header3 (FIRST_PAGE, rId11,
            % w:type="first") -> 20/20 byte-identical to the frozen s0055 oracle.
            % document.xml == 1704 B / 79022ccd. A readable w:type witness confirms
            % the WD_HEADER_FOOTER index -> headerReference w:type wiring at the
            % byte level.
            d = mat2doc.Document();
            s = d.sections.getitem_(0);
            s.different_first_page_header_footer = true;
            h = s.header;
            h.is_linked_to_previous = false;
            hp = h.paragraphs; hp(1).text = "Primary";
            eh = s.even_page_header;
            eh.is_linked_to_previous = false;
            ehp = eh.paragraphs; ehp(1).text = "Even";
            fh = s.first_page_header;
            fh.is_linked_to_previous = false;
            fhp = fh.paragraphs; fhp(1).text = "First";
            zipBytes = testCase.saveZip(d);
            [blobs, names] = zipEntryList(zipBytes);
            testCase.assertPackage(zipBytes, testCase.S0055_MANIFEST, ...
                's0055 even+first-page headers + titlePg (3 header parts)');
            % readable w:type witness: document.xml carries even + first references.
            doc = char(entryBlob(blobs, names, "word/document.xml"));
            testCase.verifyTrue(contains(doc, 'w:type="even"'), ...
                'document.xml must carry the even-page headerReference (w:type="even")');
            testCase.verifyTrue(contains(doc, 'w:type="first"'), ...
                'document.xml must carry the first-page headerReference (w:type="first")');
        end

        function test_add_drop_toggle_returns_to_pristine(testCase)
            % Edge/Regression (L1, zero-residue drop path): add a header + content,
            % then is_linked_to_previous = True DROPS it -- the package must return
            % to the pristine 17-part default byte-for-byte (== frozen s0001). The
            % drop removes the part, the rel (refcount), the Override AND the
            % headerReference; ANY residue (an orphan Override, a stale rel, a left
            % headerReference) makes a part diverge and this RED. document.xml back
            % to 1548 B / 0e4dd503.
            d = mat2doc.Document();
            h = d.sections.getitem_(0).header;
            h.is_linked_to_previous = false;
            hp = h.paragraphs; hp(1).text = "X";
            h.is_linked_to_previous = true;                % drop everything
            zipBytes = testCase.saveZip(d);
            testCase.assertPackage(zipBytes, testCase.S0001_MANIFEST, ...
                's0056 add->drop toggle == pristine s0001 (zero residue)');
        end

        % =============================================================== %
        % 3. ** RELOAD ** round-trip byte pin (the LOAD-side flip, s0058)   %
        % =============================================================== %

        function test_reload_roundtrip_byte_pin(testCase)
            % Regression (L1) + Edge (LOAD-side dispatch): open the frozen
            % python-generated header+footer package s0058 _source.docx in Mat2Doc
            % -- exercising the PartFactory WML_HEADER -> HeaderPart.load and
            % WML_FOOTER -> FooterPart.load flip on LOAD (the inherited-static-trap
            % path, reachable ONLY through reload) -- then save() UNCHANGED. The
            % result is 19/19 byte-identical to python-docx's OWN Document(_source)
            % .save() (frozen py_roundtrip / s0058 manifest). header1 == 237dfed9
            % "Reloaded Header", footer1 == 9e92124a "Reloaded Footer". Also assert
            % the read-back through the reloaded proxies.
            src = testCase.dataFile('s0058_source.docx');
            d = mat2doc.Document(src);                      % LOAD -> HeaderPart/FooterPart.load
            s = d.sections.getitem_(0);

            % read-back through the reloaded proxies (probe MATCH, validate_P5-3b s4)
            testCase.verifyFalse(s.header.is_linked_to_previous, ...
                'reloaded header has an explicit definition -> not linked');
            hp = s.header.paragraphs;
            testCase.verifyEqual(string(hp(1).text), "Reloaded Header", ...
                'reloaded header text read back through HeaderPart.load');
            testCase.verifyEqual(string(s.header.part.partname), "/word/header1.xml", ...
                'reloaded header part partname');
            testCase.verifyClass(s.header.part, 'mat2doc.parts.HeaderPart', ...
                'the reloaded header part is a HeaderPart (WML_HEADER load flip)');
            testCase.verifyFalse(s.footer.is_linked_to_previous, ...
                'reloaded footer has an explicit definition -> not linked');
            fp = s.footer.paragraphs;
            testCase.verifyEqual(string(fp(1).text), "Reloaded Footer", ...
                'reloaded footer text read back through FooterPart.load');
            testCase.verifyEqual(string(s.footer.part.partname), "/word/footer1.xml", ...
                'reloaded footer part partname');
            testCase.verifyClass(s.footer.part, 'mat2doc.parts.FooterPart', ...
                'the reloaded footer part is a FooterPart (WML_FOOTER load flip)');

            % save UNCHANGED -> 19/19 byte-identical to py_roundtrip
            zipBytes = testCase.saveZip(d);
            testCase.assertPackage(zipBytes, testCase.S0058_MANIFEST, ...
                's0058 reload round-trip == python-docx own open/save');
        end

        % =============================================================== %
        % 4. C3 element-seam byte-neutrality (the live-refactor guard)      %
        % =============================================================== %

        function test_c3_m1_byte_neutrality(testCase)
            % Regression (byte-neutrality, L1): the BlockItemContainer element-seam
            % refactor (property -> protected element_() method + element_store_
            % rename) perturbs ZERO default-save bytes -- a bare Document().save()
            % STILL emits word/document.xml at EXACTLY 1548 B / 0e4dd503 (== M1).
            d = mat2doc.Document();
            doc = testCase.saveExtract(d, "word/document.xml");
            testCase.verifyEqual(numel(doc), testCase.DOC_SIZE_M1, ...
                sprintf('M1 document.xml must be exactly %d B (C3 seam neutrality)', testCase.DOC_SIZE_M1));
            testCase.verifyEqual(sha256hex(doc), testCase.DOC_SHA_M1, ...
                'M1 document.xml SHA unchanged after the C3 seam refactor (byte-identical L1)');
        end

        function test_c3_m2_byte_neutrality(testCase)
            % Regression (byte-neutrality, L1): the content-bearing M2 path
            % (add_heading x3 + add_paragraph -- Body_.add_paragraph / add_heading
            % flow THROUGH the seam) is byte-neutral too -- word/document.xml ==
            % EXACTLY 1865 B / a71e5502 (== frozen s0033). Guards the seam on the
            % live content path, not just the empty default.
            d = mat2doc.Document();
            d.add_heading("Document Title", 0);
            d.add_heading("First Section", 1);
            d.add_heading("A Subsection", 2);
            d.add_paragraph("Body paragraph text.");
            doc = testCase.saveExtract(d, "word/document.xml");
            testCase.verifyEqual(numel(doc), testCase.DOC_SIZE_M2, ...
                sprintf('M2 document.xml must be exactly %d B (C3 seam neutrality)', testCase.DOC_SIZE_M2));
            testCase.verifyEqual(sha256hex(doc), testCase.DOC_SHA_M2, ...
                'M2 document.xml SHA == frozen s0033 (C3 seam neutral on the content path, L1)');
        end

        % =============================================================== %
        % 5. Section proxies -- kinds, WD_HEADER_FOOTER wiring, caching     %
        % =============================================================== %

        function test_section_six_members_resolve_to_proxies(testCase)
            % Nominal (section.py 61-142): all SIX Section header/footer members
            % RESOLVE to the correct Header_/Footer_ proxy (the P5-3a stub battery
            % is now live -- see Test_p5_3a re-pin). header/footer -> PRIMARY;
            % even_page_* -> EVEN_PAGE; first_page_* -> FIRST_PAGE. The
            % WD_HEADER_FOOTER index is not public; the class (Header_ vs Footer_)
            % is pinned here, the index -> w:type byte wiring is proven by s0055.
            sec = mat2doc.Document().sections.getitem_(0);
            testCase.verifyClass(sec.header,             testCase.HEADER_CLS, 'header -> Header_');
            testCase.verifyClass(sec.footer,             testCase.FOOTER_CLS, 'footer -> Footer_');
            testCase.verifyClass(sec.even_page_header,   testCase.HEADER_CLS, 'even_page_header -> Header_');
            testCase.verifyClass(sec.even_page_footer,   testCase.FOOTER_CLS, 'even_page_footer -> Footer_');
            testCase.verifyClass(sec.first_page_header,  testCase.HEADER_CLS, 'first_page_header -> Header_');
            testCase.verifyClass(sec.first_page_footer,  testCase.FOOTER_CLS, 'first_page_footer -> Footer_');
        end

        function test_header_footer_lazy_caching(testCase)
            % Edge (section.py 97/135 @lazyproperty vs 61-95 plain @property):
            % header/footer CACHE their proxy -- repeated reads return the SAME
            % handle; even_page_*/first_page_* mint a FRESH proxy each access. The
            % proxies are plain handle objects (BlockItemContainer < StoryChild <
            % handle -- NOT an ElementProxy), so `==` is instance identity with NO
            % element side effect (a mistaken element-identity `==` would lazily
            % create a part just to compare -- pinned against here).
            sec = mat2doc.Document().sections.getitem_(0);
            testCase.verifyTrue(sec.header == sec.header, ...
                'header is @lazyproperty-cached -> SAME handle on repeat (H5/H9)');
            testCase.verifyTrue(sec.footer == sec.footer, ...
                'footer is @lazyproperty-cached -> SAME handle on repeat (H5/H9)');
            testCase.verifyFalse(sec.even_page_header == sec.even_page_header, ...
                'even_page_header is a plain @property -> FRESH proxy each access');
            testCase.verifyFalse(sec.first_page_footer == sec.first_page_footer, ...
                'first_page_footer is a plain @property -> FRESH proxy each access');
        end

        function test_is_linked_get_set_add_and_drop(testCase)
            % Nominal + Edge (section.py 302-325): is_linked_to_previous get/set.
            % Default True (no explicit definition). Setting False ADDS an empty
            % definition (get -> False). Setting True again DROPS it (get -> True).
            % Siblings are INDEPENDENT (distinct references): setting header does
            % NOT move footer / even_page_header off their default True.
            sec = mat2doc.Document().sections.getitem_(0);
            h = sec.header;
            testCase.verifyTrue(h.is_linked_to_previous, 'default header is_linked_to_previous True');
            h.is_linked_to_previous = false;               % add empty definition
            testCase.verifyFalse(h.is_linked_to_previous, 'set False -> has explicit definition');
            testCase.verifyTrue(sec.footer.is_linked_to_previous, ...
                'footer independent (still linked after header add)');
            testCase.verifyTrue(sec.even_page_header.is_linked_to_previous, ...
                'even_page_header independent (still linked after header add)');
            h.is_linked_to_previous = true;                % drop the definition
            testCase.verifyTrue(h.is_linked_to_previous, 'set True again -> definition dropped, linked');
        end

        function test_inherit_walk_three_sections(testCase)
            % Edge (section.py 356-374 _get_or_add_definition, the inherit-walk):
            % a 3-section document. Initially all headers linked ([T,T,T]). Define
            % section-1's header -> [T,F,T]. Then:
            %   * sec2.header.part == sec1.header.part == /word/header1.xml  (case-2:
            %     sec2 has no definition, recurses via preceding_sectPr to sec1).
            %   * sec0.header.part == /word/header2.xml  (case-3: sec0 is the FIRST
            %     section, preceding_sectPr is None, so it ADDS a new part; header1
            %     already exists -> next_partname gives header2).
            %   * final linked == [F,F,T]  (sec0 defined via header2, sec1 defined
            %     via header1, sec2 still linked -- it resolved by inheritance
            %     WITHOUT creating its own definition).
            % This is the exact frozen s0057 inherit oracle, step-for-step.
            d = mat2doc.Document();
            d.add_section(mat2doc.enum.section.WD_SECTION.NEW_PAGE);
            d.add_section(mat2doc.enum.section.WD_SECTION.NEW_PAGE);
            secs = d.sections;
            linked0 = arrayfun(@(i) secs.getitem_(i).header.is_linked_to_previous, 0:2);
            testCase.verifyEqual(logical(linked0), [true true true], ...
                'linked0 [T,T,T] -- all three headers initially linked');

            h1 = secs.getitem_(1).header;
            h1.is_linked_to_previous = false;              % define section-1
            h1p = h1.paragraphs; h1p(1).text = "S1";
            linked1 = arrayfun(@(i) secs.getitem_(i).header.is_linked_to_previous, 0:2);
            testCase.verifyEqual(logical(linked1), [true false true], ...
                'linked1 [T,F,T] -- only section-1 has an explicit definition');

            % ORDER MATTERS: p2 (case-2, no add), p1 (case-1), p0 (case-3, adds header2).
            p2 = string(secs.getitem_(2).header.part.partname);
            p1 = string(secs.getitem_(1).header.part.partname);
            testCase.verifyEqual(p2, "/word/header1.xml", ...
                'sec2 inherits sec1 header1.xml (case-2 recursion via preceding_sectPr)');
            testCase.verifyEqual(p1, "/word/header1.xml", 'sec1 owns header1.xml (case-1)');
            testCase.verifyEqual(p2, p1, 'sec2 and sec1 share header1.xml');
            p0 = string(secs.getitem_(0).header.part.partname);
            testCase.verifyEqual(p0, "/word/header2.xml", ...
                'sec0 (first section) ADDS header2.xml (case-3; header1 already taken)');

            linked2 = arrayfun(@(i) secs.getitem_(i).header.is_linked_to_previous, 0:2);
            testCase.verifyEqual(logical(linked2), [false false true], ...
                'linked2 [F,F,T] -- sec0/sec1 defined, sec2 still inherits (no own def)');
        end

        % =============================================================== %
        % 6. HeaderPart / FooterPart new + inherited-static load            %
        % =============================================================== %

        function test_headerpart_footerpart_new_and_load(testCase)
            % Nominal + Edge (parts/hdrftr.py): HeaderPart.new / FooterPart.new
            % build the right partname (/word/header1.xml, /word/footer1.xml via
            % next_partname) and content_type (WML_HEADER, WML_FOOTER). The
            % inherited-static `load` OVERRIDE dispatches to the SUBCLASS (the
            % inherited-static trap: without the own static, load would silently
            % build a base XmlPart) -- HeaderPart.load -> a HeaderPart, FooterPart
            % .load -> a FooterPart, re-parsing the blob byte-round-trippable.
            CT = mat2doc.opc.CONTENT_TYPE;
            pkg = mat2doc.package.Package.open(char(testCase.templatePath()));

            hp = mat2doc.parts.HeaderPart.new(pkg);
            testCase.verifyClass(hp, 'mat2doc.parts.HeaderPart', 'HeaderPart.new -> a HeaderPart');
            testCase.verifyEqual(string(hp.partname), "/word/header1.xml", ...
                'HeaderPart.new partname /word/header1.xml (next_partname)');
            testCase.verifyEqual(string(hp.content_type), string(CT.WML_HEADER), ...
                'HeaderPart.new content_type WML_HEADER');

            fp = mat2doc.parts.FooterPart.new(pkg);
            testCase.verifyClass(fp, 'mat2doc.parts.FooterPart', 'FooterPart.new -> a FooterPart');
            testCase.verifyEqual(string(fp.partname), "/word/footer1.xml", ...
                'FooterPart.new partname /word/footer1.xml (next_partname)');
            testCase.verifyEqual(string(fp.content_type), string(CT.WML_FOOTER), ...
                'FooterPart.new content_type WML_FOOTER');

            % inherited-static load dispatches to the subclass (NOT base XmlPart)
            hp2 = mat2doc.parts.HeaderPart.load(hp.partname, CT.WML_HEADER, hp.blob, pkg);
            testCase.verifyClass(hp2, 'mat2doc.parts.HeaderPart', ...
                'HeaderPart.load builds a HeaderPart (inherited-static trap fix)');
            fp2 = mat2doc.parts.FooterPart.load(fp.partname, CT.WML_FOOTER, fp.blob, pkg);
            testCase.verifyClass(fp2, 'mat2doc.parts.FooterPart', ...
                'FooterPart.load builds a FooterPart (inherited-static trap fix)');
        end

        % =============================================================== %
        % 7. DocumentPart un-stubs -- add / get / drop + RT constants       %
        % =============================================================== %

        function test_documentpart_add_get_drop_header_footer(testCase)
            % Nominal + Edge (parts/document.py un-stubs, section 6 re-pin of
            % Test_p2_2's stub battery): add_header_part / add_footer_part return
            % [part, rId] (a HeaderPart/FooterPart + a string rId); header_part(rId)
            % / footer_part(rId) round-trip to the SAME related handle;
            % drop_header_part (-> drop_rel) removes the REL entry (rels().contains
            % goes false). NOTE it does NOT invalidate related_parts: drop_rel is
            % the inherited dict.__delitem__, which prunes only the rels dict and
            % LEAVES the parallel related_parts map STALE (the faithful python-docx
            % quirk pinned in Test_p2_2), so header_part(rId) would still resolve the
            % stale part -- the rel-removal is therefore the correct drop pin, not a
            % KeyError.
            RT = mat2doc.opc.RELATIONSHIP_TYPE;
            testCase.verifyEqual(string(RT.HEADER), ...
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header", ...
                'RT.HEADER constant');
            testCase.verifyEqual(string(RT.FOOTER), ...
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer", ...
                'RT.FOOTER constant');

            pkg = mat2doc.package.Package.open(char(testCase.templatePath()));
            dp = pkg.main_document_part();

            [hp, hRid] = dp.add_header_part();
            testCase.verifyClass(hp, 'mat2doc.parts.HeaderPart', 'add_header_part -> HeaderPart');
            testCase.verifyClass(hRid, 'string', 'add_header_part -> a string rId');
            testCase.verifyTrue(dp.header_part(hRid) == hp, ...
                'header_part(rId) round-trips to the SAME added HeaderPart handle');

            [fp, fRid] = dp.add_footer_part();
            testCase.verifyClass(fp, 'mat2doc.parts.FooterPart', 'add_footer_part -> FooterPart');
            testCase.verifyTrue(dp.footer_part(fRid) == fp, ...
                'footer_part(rId) round-trips to the SAME added FooterPart handle');

            % drop_header_part -> drop_rel removes the REL (refcount 0 < 2). The
            % header part just added has no headerReference (added directly, not via
            % the section proxy), so refcount is 0 and the rel is dropped.
            testCase.verifyTrue(dp.rels().contains(hRid), 'precondition: header rel present');
            dp.drop_header_part(hRid);
            testCase.verifyFalse(dp.rels().contains(hRid), ...
                'after drop_header_part the header rel is removed (rels().contains -> false)');
        end

        % =============================================================== %
        % 8. EQUIVALENCE -- full s0057 probe vs the frozen oracle           %
        % =============================================================== %

        function test_equivalence_full_probe_vs_frozen_oracle(testCase)
            % Equivalence: replay the ENTIRE frozen s0057 probe (runS0057Probe --
            % the .m twin's body VERBATIM: kinds / default_linked / roundtrip /
            % paragraphs / header_partname / the 3-section inherit-walk) and
            % flatten-compare EVERY leaf to the frozen python-docx 1.2.0 oracle
            % copied into data\s0057_hdrftr_probe_oracle.json. Gate-3 probe_diff
            % found ZERO divergences (MATCH exit 0), so every leaf must be
            % byte/value-identical.
            port   = runS0057Probe();
            oracle = testCase.loadProbeOracle();
            pMap = containers.Map('KeyType', 'char', 'ValueType', 'char');
            oMap = containers.Map('KeyType', 'char', 'ValueType', 'char');
            flattenLeaves(port,   '', pMap);
            flattenLeaves(oracle, '', oMap);
            pKeys = sort(pMap.keys());
            oKeys = sort(oMap.keys());
            testCase.verifyEqual(pKeys, oKeys, ...
                'the replayed s0057 probe and the frozen oracle must have identical leaf keys');
            % Non-trivial floor guarding a silent-empty replay. The only looser-
            % than-byte assertion in this class; justified as a leaf-count floor.
            testCase.verifyGreaterThan(numel(oKeys), 20, ...
                'the flattened s0057 oracle must expose the full hdr/ftr probe surface');
            for i = 1:numel(oKeys)
                k = oKeys{i};
                testCase.verifyTrue(isKey(pMap, k), sprintf('port is missing leaf %s', k));
                testCase.verifyEqual(pMap(k), oMap(k), ...
                    sprintf('leaf %s must be byte/value-identical to the frozen s0057 oracle', k));
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

        function bytes = saveExtract(testCase, d, partname)
            % Save d and return the raw bytes of a single zip part (frozen order).
            zipBytes = testCase.saveZip(d);
            [blobs, names] = zipEntryList(zipBytes);
            bytes = entryBlob(blobs, names, partname);
        end

        function assertPackage(testCase, zipBytes, MANIFEST, label)
            % Full-package byte pin: enumerate the saved zip in stream (write)
            % order and assert (a) the part inventory + EXACT frozen zip-entry
            % order == MANIFEST col 1, and (b) every part's size + SHA-256 ==
            % MANIFEST cols 2/3 (SHA-256 equality == byte-identity, L1).
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
            here = fileparts(mfilename('fullpath'));   % tests\section
            p = char(fullfile(here, 'data', name));
        end

        function p = templatePath(~)
            here = fileparts(mfilename('fullpath'));   % tests\section
            root = fileparts(fileparts(here));         % worktree root
            p = fullfile(root, '+mat2doc', 'templates', 'default.docx');
        end

        function o = loadProbeOracle(testCase)
            % Read the co-located frozen s0057 oracle in BINARY mode (no CRLF
            % translation) and decode UTF-8 -> struct. jsondecode is line-ending
            % agnostic, so no `* binary` pin is needed for this value fixture
            % (s0045 precedent).
            p = testCase.dataFile('s0057_hdrftr_probe_oracle.json');
            fid = fopen(p, 'r', 'n');
            assert(fid >= 0, 'cannot open frozen s0057 oracle %s', p);
            raw = fread(fid, Inf, '*uint8')';
            fclose(fid);
            o = jsondecode(native2unicode(raw, 'UTF-8'));
        end
    end
end

% ===================== file-local helpers ============================== %

function P = runS0057Probe()
    % Replay the s0057 probe sequence (the .m twin's body, VERBATIM tags/inputs/
    % order) and return the nested struct of tagged values. Embedded so the
    % Equivalence leg is self-contained (the validation-folder scenario is NOT on
    % the toolbox path and must not be a dependency). Mirrors
    % validation\mat2doc\scenarios\s0057_p5_3b_hdrftr_probe.m lines 13-80.
    P = struct();

    % -- 1. proxy kinds --
    d = mat2doc.Document();
    s = d.sections.getitem_(0);
    K = struct();
    K.header            = probeKind(s.header);
    K.footer            = probeKind(s.footer);
    K.even_page_header  = probeKind(s.even_page_header);
    K.even_page_footer  = probeKind(s.even_page_footer);
    K.first_page_header = probeKind(s.first_page_header);
    K.first_page_footer = probeKind(s.first_page_footer);
    P.kinds = K;

    % -- 2. default is_linked_to_previous --
    L = struct();
    L.header            = s.header.is_linked_to_previous;
    L.footer            = s.footer.is_linked_to_previous;
    L.even_page_header  = s.even_page_header.is_linked_to_previous;
    L.first_page_header = s.first_page_header.is_linked_to_previous;
    P.default_linked = L;

    % -- 3. get/set/get round-trip + independence --
    h = s.header;
    before = h.is_linked_to_previous;
    h.is_linked_to_previous = false;
    after = h.is_linked_to_previous;
    R = struct();
    R.before = before;
    R.after = after;
    R.footer_independent = s.footer.is_linked_to_previous;
    R.even_independent = s.even_page_header.is_linked_to_previous;
    P.roundtrip = R;

    % -- 4. paragraphs count + add_paragraph --
    count0 = numel(h.paragraphs);
    h.add_paragraph("second");
    count1 = numel(h.paragraphs);
    PP = struct(); PP.count0 = count0; PP.count1 = count1;
    P.paragraphs = PP;

    % -- 5. header.part partname (case-3 add on first section) --
    P.header_partname = string(h.part.partname);

    % -- 6. multi-section inherit-walk --
    d2 = mat2doc.Document();
    d2.add_section(mat2doc.enum.section.WD_SECTION.NEW_PAGE);
    d2.add_section(mat2doc.enum.section.WD_SECTION.NEW_PAGE);
    secs = d2.sections;
    linked0 = arrayfun(@(i) secs.getitem_(i).header.is_linked_to_previous, 0:2);
    h1 = secs.getitem_(1).header;
    h1.is_linked_to_previous = false;
    h1p = h1.paragraphs; h1p(1).text = "S1";
    linked1 = arrayfun(@(i) secs.getitem_(i).header.is_linked_to_previous, 0:2);
    p2 = string(secs.getitem_(2).header.part.partname);
    p1 = string(secs.getitem_(1).header.part.partname);
    same_2_1 = (p2 == p1);
    p0 = string(secs.getitem_(0).header.part.partname);
    linked2 = arrayfun(@(i) secs.getitem_(i).header.is_linked_to_previous, 0:2);
    I = struct();
    I.linked0 = logical(linked0);
    I.linked1 = logical(linked1);
    I.sec2_partname = p2;
    I.sec1_partname = p1;
    I.sec2_eq_sec1 = logical(same_2_1);
    I.sec0_partname = p0;
    I.linked2 = logical(linked2);
    P.inherit = I;
end

function k = probeKind(x)
    if isa(x, 'mat2doc.section.Header_')
        k = "header";
    else
        k = "footer";
    end
end

function [blobs, names] = zipEntryList(zipBytes)
    % Enumerate an in-memory zip in stream (write) order -> ordered {names, blobs}.
    % java.util.zip.ZipInputStream reads local file headers in physical order, so
    % `names` is the true zip-entry write sequence. Kept file-local so the order
    % pin is independent of the reader under test. (Copied from
    % Test_p1_8_skeleton_m1.m.)
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

function flattenLeaves(s, prefix, map)
    % Recursively flatten a (possibly nested) struct / cell / array into
    % map(dotted.path) -> canonical string, so a MATLAB-built probe struct and the
    % jsondecode'd oracle compare leaf-by-leaf regardless of container types.
    % (Copied from Test_p5_3a_sections_api.m.)
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
    elseif (islogical(s) || isnumeric(s)) && ~isscalar(s)
        if isempty(s)
            map(prefix) = ''; %#ok<NASGU>
        else
            parts = strings(1, numel(s));
            for i = 1:numel(s)
                parts(i) = canonScalar(s(i));
            end
            map(prefix) = char(join(parts, "|")); %#ok<NASGU>
        end
    else
        map(prefix) = char(canonScalar(s)); %#ok<NASGU>
    end
end

function s = canonScalar(v)
    % Coerce logical -> double FIRST so a boolean is "1"/"0" whether jsondecode
    % returned it as a logical or a double (jsondecode's scalar-bool vs bool-array
    % typing is not something the oracle and the MATLAB replay must agree on --
    % this normalization removes that ambiguity, so both sides canonicalize
    % identically). Strings pass through unchanged.
    if islogical(v)
        v = double(v);
    end
    if isnumeric(v)
        s = string(num2str(v));
    else
        s = string(v);
    end
end

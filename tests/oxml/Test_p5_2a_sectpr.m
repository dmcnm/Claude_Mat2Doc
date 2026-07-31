classdef Test_p5_2a_sectpr < matlab.unittest.TestCase
% TEST_P5_2A_SECTPR  Gate-4 permanent unit tests for Mat2Doc P5-2a
%   (section-properties oxml core + page geometry): src/docx/oxml/section.py ->
%   +mat2doc\+oxml\+section\ CT_HdrFtrRef / CT_PageMar / CT_PageSz / CT_SectType /
%   CT_SectPr, and the SEVEN section-block registry rows in +mat2doc\+oxml\
%   registry.m (w:titlePg->CT_OnOff [P5-1 deferral closed]; w:footerReference /
%   w:headerReference->CT_HdrFtrRef; w:pgMar->CT_PageMar; w:pgSz->CT_PageSz;
%   w:sectPr->CT_SectPr; w:type->CT_SectType).
%
%   P5-2a is a REGISTRY-ADDING, M1-CENTRAL WP: default.docx's word/document.xml
%   carries a <w:sectPr> (with pgSz + pgMar + cols + docGrid), so registering
%   w:sectPr/pgSz/pgMar/type/titlePg flips the PARSE CLASS of that subtree from
%   generic XmlElement to the new CT classes on EVERY load/save. That parse->CT->
%   serialize transit is byte-neutral (registering a CT changes only the parsed
%   node CLASS, never content/order -- P4-6/P5-1 precedent), so word/document.xml
%   stays 1548 B, SHA-256 0e4dd503... This class permanently freezes the
%   guarantees the prior gates established:
%     * Gate-1 Porter  : audit_P5-2a_sectpr.md (self-probe).
%     * Gate-2 Auditor (Fable): audit_P5-2a_sectpr.md GATE-2 -- APPROVE; H11
%       successor-slice mapping proven; H3/H4/H6/H10 accessor semantics proven.
%     * Gate-3 Validator: validate_P5-2a_sectpr.md -- PASS, ZERO new D-numbers.
%       probe_diff s0039 MATCH exit 0 (full CT_HdrFtrRef/CT_PageMar/CT_PageSz/
%       CT_SectType/CT_SectPr surface); M1 17/17 (document.xml 1548 B / 0e4dd503
%       byte-identical through the new CT_SectPr parse path); parse-path round-trip
%       byte-identical; the geometry-write (1568 B / 698367cd) and LANDSCAPE
%       (1569 B / d165a628, NO w/h swap) novel paths 17/17 byte-identical; 3
%       extended scenarios byte-exact; H11 successor-slice adversarial insertion
%       byte-identical.
%
%   THE HIGHEST-VALUE PERMANENT PINS IN THIS WP (a regression goes RED loudly):
%     * (M) M1 document.xml byte-pin (test_m1_document_xml_byte_identical): the
%       registry-adding parse-path guard -- mat2doc.Document().save() ->
%       word/document.xml == 1548 B, SHA-256 0e4dd503... (frozen s0001 oracle).
%       The CENTRAL part that P5-2a could break; SHA equality is an L1 assertion.
%     * (G) Geometry-write byte-pin (test_geometry_write_byte_identical): oxml-level
%       page_width/page_height/orientation/top_margin on the body sectPr -> save ->
%       word/document.xml == 1568 B, SHA-256 698367cd... (frozen s0040). The novel
%       emit path.  <w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/> (orient
%       APPENDED after w,h -- no swap) + <w:pgMar w:top="360" .../>.
%     * (L) LANDSCAPE byte-pin (test_landscape_byte_identical): orientation=LANDSCAPE
%       ONLY -> word/document.xml == 1569 B, SHA-256 d165a628... (frozen s0041), AND
%       w:w/w:h stay the DEFAULT PORTRAIT dims (12240/15840) -- the orientation-
%       semantics / no-w-h-swap guard. A port that swapped dims on landscape diverges.
%     * (H11) SUCCESSOR-SLICE pins (test_h11_successor_slice_pins): the ZeroOrOne/
%       ZeroOrMore successor slices map Python _tag_seq[N:] -> TAG_SEQ(N+1:end).
%         - pgMar lands EXACTLY between pgSz(@4) and paperSrc(@6)  [_tag_seq[5:]->TAG_SEQ(6:end)]
%         - pgSz  lands EXACTLY between type(@3) and pgMar(@5)     [_tag_seq[4:]->TAG_SEQ(5:end)]
%         - a 2nd headerReference lands AFTER an intervening footerReference (refs-
%           first-then-type: neither ref tag is a _tag_seq member, so their
%           successors span the WHOLE sequence).
%       serhex byte pins vs the frozen s0039 oracle (rooted cases) + hard-coded L1
%       strings (loose w-only cases). A wrong slice mis-positions the child -> Word
%       repair -> the pin goes RED.
%     * (T) titlePg H3 tri-state BREADTH (test_ct_sectpr_titlepg_tristate): Python
%       `value in [None, False]` uses ==-MEMBERSHIP, so None / False / 0 / 0.0 ALL
%       REMOVE the child (isequal(value,false); isequal(0,false) true), while
%       True / 1 SET it -> the EMPTY <w:titlePg/> (D-delta-1, CT_OnOff.val==default
%       removes @val). The numeric-0 breadth is pinned distinct from CT_Settings'
%       IDENTITY semantics.
%     * (S) margin SIGN asymmetry (test_ct_pagemar_all_margins): top/bottom use
%       ST_SignedTwipsMeasure (accept negatives), the other five ST_TwipsMeasure --
%       pinned by writing negative top/bottom (-720/-360 twips) that round-trip to
%       signed EMU, alongside the serhex oracle.
%
%   Provenance (all 2026-07-31):
%     * Audit    : validation\mat2doc\audit_P5-2a_sectpr.md
%     * Validate : validation\mat2doc\validate_P5-2a_sectpr.md
%     * Scenarios: validation\mat2doc\scenarios\s0039_p5_2a_sectpr_probe.{py,m}
%                  (its probe body is replayed VERBATIM by runProbes() below);
%                  s0040_..._geometry_write_gscenario.{py,m} (geo byte oracle);
%                  s0041_..._landscape_gscenario.{py,m} (landscape byte oracle).
%     * Frozen refs (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE at Gate-3):
%         references\s0039\probe.json -- copied verbatim (self-contained) into
%           tests\oxml\data\s0039_probe_oracle.json (the Equivalence replay set;
%           value/serhex JSON, jsondecode is line-ending agnostic -> no `* binary`
%           .gitattributes pin needed, per the s0030/s0036 precedent).
%         references\s0040\ (geo 1568 / 698367cd) and references\s0041\ (landscape
%           1569 / d165a628) -- pinned here by SHA-256 of the MATLAB-emitted part
%           (Document().save() itself), so no byte fixture is copied.
%         references\s0001\parts\word\document.xml -- the M1 byte reference
%           (1548 / 0e4dd503); NOT copied (SHA of what Document().save() emits).
%
%   Coverage taxonomy
%   -----------------
%   * Nominal      -- CT_PageSz w/h/orient; CT_PageMar 7 margins; CT_SectType val +
%                     each WD_SECTION_START member; CT_SectPr page geometry / margins
%                     / orientation / start_type / titlePg accessors; hdr/ftr ref
%                     add/get/remove/list; registry resolution.
%   * Edge         -- None ([]) round-trips (H3); orient PORTRAIT (== NON-None
%                     default) removes @w:orient; ORIENT_DEFAULT (absent->PORTRAIT);
%                     titlePg numeric-0 / 0.0 / None / False breadth (all remove);
%                     start_type NEW_PAGE + None identity-remove; signed negative
%                     top/bottom margins; H5 get-is-add identity; H11 adversarial
%                     insertion positions; RequiredAttribute-absent error path.
%   * Equivalence  -- test_equivalence_full_battery_vs_frozen_oracle replays the
%                     ENTIRE s0039 battery (runProbes, the .m twin's body verbatim)
%                     and flatten-compares every leaf to the frozen oracle (0 diffs).
%   * Regression   -- hard-coded expected serialized-XML strings (ASCII == byte-
%                     identical L1) + UPPERCASE serhex vs the frozen oracle + SHA-256
%                     of the M1 / geometry / landscape / extended document.xml parts.
%   * Upstream     -- the H11 successor-slice ordering + geometry/orientation
%                     semantics ARE the python-docx section.py surface; the frozen
%                     oracle IS lxml's expected output for these sequences.
%
%   Byte-level (L1) note: every serialized-XML comparison is either the FULL
%   serialize_part_xml output decoded as an ASCII string (string-equality ==
%   byte-equality L1), or its UPPERCASE hex (serhex) vs the frozen oracle, or the
%   SHA-256 of an emitted document.xml part. NO D-number granted any L2 relaxation
%   in this WP (Gate-3: zero new, none at L2), so every pin here is L1. The
%   equivalence leaf-key-count guard is the only looser-than-byte check and is
%   commented at its site.  (The pre-existing SIGNED D-nsprefix-rewrite governs
%   only LOOSE multi-ref sectPr r:-prefix numbering -- dead-on-generation; this
%   class pins the REACHABLE r:-ROOTED multi-ref path as byte-identical, and does
%   NOT pin the loose divergence as correct -- see test_rooted_multiref_byte_identical.)
%
%   Determinism: no network, no absolute paths. The worktree root and the
%   co-located oracle resolve relative to this file via
%   fileparts(mfilename('fullpath')); saves go to tempname .docx deleted via
%   onCleanup; every file read is binary ('r','n'). The +mat2doc package resolves
%   via the MANDATORY PathFixture(worktree-root) in TestClassSetup (WP9-F4 lesson).

    properties (Constant)
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        % --- registered class names (P5-2a, +oxml\registry.m 235-241) ---
        CT_SECTPR    = 'mat2doc.oxml.section.CT_SectPr'
        CT_PAGESZ    = 'mat2doc.oxml.section.CT_PageSz'
        CT_PAGEMAR   = 'mat2doc.oxml.section.CT_PageMar'
        CT_SECTTYPE  = 'mat2doc.oxml.section.CT_SectType'
        CT_HDRFTRREF = 'mat2doc.oxml.section.CT_HdrFtrRef'
        CT_ONOFF     = 'mat2doc.oxml.shared.CT_OnOff'

        % --- frozen s0001 M1 word/document.xml byte reference (the P5-2a
        %     registry-adding parse-path risk on the CENTRAL part) ---
        DOC_SIZE_M1 = 1548
        DOC_SHA_M1  = "0e4dd503bc095219cd24e8312b372c588ca56bb834aacac53eafb7bf47836327"

        % --- frozen s0040 GEOMETRY-write oracle (page_width/height/orientation/
        %     top_margin on the body sectPr) ---
        DOC_SIZE_GEO = 1568
        DOC_SHA_GEO  = "698367cd72fde4b50706ed17100b5c21421470db39e21ba3197cc8c69cd5f231"

        % --- frozen s0041 LANDSCAPE oracle (orientation only; NO w/h swap) ---
        DOC_SIZE_LANDSCAPE = 1569
        DOC_SHA_LANDSCAPE  = "d165a628ae3db01a0aeb96463af0e804ce0d735055858d7977a28e90f7e6d662"

        % --- extended scenario oracles (validate report §4; byte-verified) ---
        DOC_SIZE_MARGNONE = 1535
        DOC_SHA_MARGNONE  = "b45b2c6b"   % 8-hex prefix pinned (report's frozen prefix)
        DOC_SIZE_STARTDOC = 1574
        DOC_SHA_STARTDOC  = "23f89f6b"
        DOC_SIZE_TITLEPG  = 1560
        DOC_SHA_TITLEPG   = "7b41c7a5"
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied verbatim from tests\oxml\Test_p5_1_settings.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\oxml
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. registry rows (the 7 P5-2a flips)                             %
        % =============================================================== %

        function test_registry_resolves_seven_rows(testCase)
            % Nominal / Regression (registry.m 235-241): the 7 section-block rows.
            % REGISTRATION is what flips the parse class of the sectPr subtree.
            pairs = { ...
                "w:titlePg",         testCase.CT_ONOFF; ...
                "w:footerReference", testCase.CT_HDRFTRREF; ...
                "w:headerReference", testCase.CT_HDRFTRREF; ...
                "w:pgMar",           testCase.CT_PAGEMAR; ...
                "w:pgSz",            testCase.CT_PAGESZ; ...
                "w:sectPr",          testCase.CT_SECTPR; ...
                "w:type",            testCase.CT_SECTTYPE };
            for i = 1:size(pairs, 1)
                tag = pairs{i, 1}; cls = pairs{i, 2};
                r = mat2doc.oxml.registry(mat2doc.oxml.qn(tag));
                testCase.verifyEqual(char(r), cls, ...
                    sprintf('registry must resolve %s -> %s (P5-2a row)', tag, cls));
                % a real OxmlElement(tag) is the exact-class CT
                testCase.verifyEqual(class(mat2doc.oxml.OxmlElement(tag)), cls, ...
                    sprintf('OxmlElement(%s) must be a %s', tag, cls));
            end
        end

        % =============================================================== %
        % 2. CT_PageSz -- w/h/orient (H3 tri-state + H10 orient default)    %
        % =============================================================== %

        function test_ct_pagesz(testCase)
            % Nominal + Edge + Regression (s0039 pagesz): w/h Length round-trip;
            % orient default PORTRAIT when @w:orient ABSENT (ORIENT_DEFAULT); set
            % LANDSCAPE writes; set PORTRAIT (== NON-None default) REMOVES @w:orient;
            % set None ([]) removes. serhex vs the frozen oracle.
            o = loadOracle();

            % bare: w/h None (H3), orient PORTRAIT (H10 NON-None default, @ absent)
            b = mat2doc.oxml.OxmlElement("w:pgSz");
            testCase.verifyEqual(class(b), testCase.CT_PAGESZ);
            testCase.verifyTrue(isequal(b.w, []), 'bare w -> [] (None)');
            testCase.verifyTrue(isequal(b.h, []), 'bare h -> [] (None)');
            testCase.verifyEqual(b.orient, mat2doc.enum.section.WD_ORIENTATION.PORTRAIT, ...
                'bare orient -> PORTRAIT (ORIENT_DEFAULT, @w:orient absent)');

            % set w/h (Twips) -> Length EMU (12240tw=7772400EMU, 15840tw=10058400EMU)
            e = mat2doc.oxml.OxmlElement("w:pgSz");
            e.w = mat2doc.shared.Twips(12240);
            e.h = mat2doc.shared.Twips(15840);
            testCase.verifyEqual(double(e.w), 7772400, 'w EMU exact');
            testCase.verifyEqual(double(e.h), 10058400, 'h EMU exact');
            testCase.verifyEqual(hx_e(e), string(o.pagesz.set_wh.serhex), ...
                'pgSz set w/h serhex (L1) vs frozen oracle');

            % set LANDSCAPE -> @w:orient="landscape"
            e = mat2doc.oxml.OxmlElement("w:pgSz");
            e.orient = mat2doc.enum.section.WD_ORIENTATION.LANDSCAPE;
            testCase.verifyEqual(e.orient, mat2doc.enum.section.WD_ORIENTATION.LANDSCAPE);
            testCase.verifyEqual(hx_e(e), string(o.pagesz.set_landscape.serhex), ...
                'pgSz set LANDSCAPE serhex (L1)');

            % set LANDSCAPE then PORTRAIT -> @w:orient REMOVED (== NON-None default)
            e = mat2doc.oxml.OxmlElement("w:pgSz");
            e.orient = mat2doc.enum.section.WD_ORIENTATION.LANDSCAPE;
            e.orient = mat2doc.enum.section.WD_ORIENTATION.PORTRAIT;
            testCase.verifyEqual(e.orient, mat2doc.enum.section.WD_ORIENTATION.PORTRAIT, ...
                'orient reads PORTRAIT after removal (default)');
            testCase.verifyEqual(hx_e(e), string(o.pagesz.set_portrait_removes.serhex), ...
                'set PORTRAIT (== NON-None default) REMOVES @w:orient (L1)');
            testCase.verifyEmpty(e.attrib_names(), ...
                'no attributes remain after setting orient to the default');

            % set None ([]) -> removes @w:orient too (byte-identical to above)
            e = mat2doc.oxml.OxmlElement("w:pgSz");
            e.orient = mat2doc.enum.section.WD_ORIENTATION.LANDSCAPE;
            e.orient = [];
            testCase.verifyEqual(hx_e(e), string(o.pagesz.set_orient_none.serhex), ...
                'set orient None ([]) removes @w:orient (L1)');
        end

        % =============================================================== %
        % 3. CT_PageMar -- 7 margins + top/bottom SIGNED asymmetry          %
        % =============================================================== %

        function test_ct_pagemar_all_margins(testCase)
            % Nominal + Edge + Regression (s0039 pagemar): all 7 margins bare->None;
            % all set incl. SIGNED NEGATIVE top/bottom (the ST_SignedTwipsMeasure vs
            % ST_TwipsMeasure split); set-None removes @w:top. serhex vs oracle.
            o = loadOracle();

            b = mat2doc.oxml.OxmlElement("w:pgMar");
            testCase.verifyEqual(class(b), testCase.CT_PAGEMAR);
            for f = ["top" "right" "bottom" "left" "header" "footer" "gutter"]
                testCase.verifyTrue(isequal(b.(f), []), ...
                    sprintf('bare pgMar %s -> [] (None, H3)', f));
            end

            % all set: top/bottom NEGATIVE (signed), others incl. gutter=0
            e = mat2doc.oxml.OxmlElement("w:pgMar");
            e.top = mat2doc.shared.Twips(-720);    % SIGNED -> -457200 EMU
            e.right = mat2doc.shared.Twips(1800);
            e.bottom = mat2doc.shared.Twips(-360); % SIGNED -> -228600 EMU
            e.left = mat2doc.shared.Twips(1440);
            e.header = mat2doc.shared.Twips(720);
            e.footer = mat2doc.shared.Twips(720);
            e.gutter = mat2doc.shared.Twips(0);
            % SIGN asymmetry pin: negatives round-trip to signed EMU on top/bottom
            testCase.verifyEqual(double(e.top), -457200, 'top signed EMU (ST_SignedTwipsMeasure)');
            testCase.verifyEqual(double(e.bottom), -228600, 'bottom signed EMU (ST_SignedTwipsMeasure)');
            testCase.verifyEqual(double(e.right), 1143000);
            testCase.verifyEqual(double(e.left), 914400);
            testCase.verifyEqual(double(e.header), 457200);
            testCase.verifyEqual(double(e.footer), 457200);
            testCase.verifyEqual(double(e.gutter), 0, 'gutter 0 present (not removed)');
            testCase.verifyEqual(hx_e(e), string(o.pagemar.all_set.serhex), ...
                'pgMar all-set (signed top/bottom) serhex (L1) vs frozen oracle');

            % set top None -> @w:top removed, rest intact
            e.top = [];
            testCase.verifyTrue(isequal(e.top, []), 'top -> [] after None');
            testCase.verifyEqual(hx_e(e), string(o.pagemar.top_none_removes.serhex), ...
                'set top None ([]) removes @w:top only (L1)');
        end

        % =============================================================== %
        % 4. CT_SectType -- val + every WD_SECTION_START member             %
        % =============================================================== %

        function test_ct_secttype(testCase)
            % Nominal + Edge + Regression (s0039 secttype): bare val None; each
            % WD_SECTION_START member writes <w:type w:val="..."/>; set None removes.
            o = loadOracle();

            b = mat2doc.oxml.OxmlElement("w:type");
            testCase.verifyEqual(class(b), testCase.CT_SECTTYPE);
            testCase.verifyTrue(isequal(b.val, []), 'bare w:type val -> [] (None, H3)');

            members = enumeration('mat2doc.enum.section.WD_SECTION_START');
            for k = 1:numel(members)
                name = char(members(k));
                e = mat2doc.oxml.OxmlElement("w:type");
                e.val = members(k);
                testCase.verifyEqual(e.val, members(k), ...
                    sprintf('w:type val round-trips %s', name));
                testCase.verifyEqual(hx_e(e), string(o.secttype.members.(name).serhex), ...
                    sprintf('w:type %s serhex (L1) vs frozen oracle', name));
            end

            % set NEW_PAGE then None -> @w:val removed -> bare <w:type/>
            e = mat2doc.oxml.OxmlElement("w:type");
            e.val = mat2doc.enum.section.WD_SECTION_START.NEW_PAGE;
            e.val = [];
            testCase.verifyTrue(isequal(e.val, []), 'val -> [] after None');
            testCase.verifyEqual(hx_e(e), string(o.secttype.none_removes.serhex), ...
                'set val None removes @w:val -> bare <w:type/> (L1)');
        end

        % =============================================================== %
        % 5. CT_SectPr -- geometry / margin / orientation / start_type      %
        % =============================================================== %

        function test_ct_sectpr_geometry_accessors(testCase)
            % Nominal + Regression (s0039 sectpr_defaults / geometry / margins):
            % page_width/height/orientation/start_type/titlePg_val/margin defaults;
            % the geometry write serhex; the WRAP-vs-DIRECT margin-setter asymmetry.
            o = loadOracle();

            % defaults on a bare sectPr (H3/H4)
            s = mat2doc.oxml.OxmlElement("w:sectPr");
            testCase.verifyEqual(class(s), testCase.CT_SECTPR);
            testCase.verifyTrue(isequal(s.page_width, []), 'default page_width None');
            testCase.verifyTrue(isequal(s.page_height, []), 'default page_height None');
            testCase.verifyEqual(s.orientation, mat2doc.enum.section.WD_ORIENTATION.PORTRAIT, ...
                'default orientation PORTRAIT (pgSz absent)');
            testCase.verifyEqual(s.start_type, mat2doc.enum.section.WD_SECTION_START.NEW_PAGE, ...
                'default start_type NEW_PAGE (type absent)');
            testCase.verifyFalse(s.titlePg_val, 'default titlePg_val False');
            for f = ["top_margin" "bottom_margin" "left_margin" "right_margin" "header" "footer" "gutter"]
                testCase.verifyTrue(isequal(s.(f), []), ...
                    sprintf('default %s None (pgMar absent)', f));
            end

            % geometry write (element-level mirror of the geo G-scenario)
            s = mat2doc.oxml.OxmlElement("w:sectPr");
            s.page_width = mat2doc.shared.Twips(15840);
            s.page_height = mat2doc.shared.Twips(12240);
            s.orientation = mat2doc.enum.section.WD_ORIENTATION.LANDSCAPE;
            s.top_margin = mat2doc.shared.Twips(360);
            testCase.verifyEqual(double(s.page_width), 10058400);
            testCase.verifyEqual(double(s.page_height), 7772400);
            testCase.verifyEqual(s.orientation, mat2doc.enum.section.WD_ORIENTATION.LANDSCAPE);
            testCase.verifyEqual(double(s.top_margin), 228600);
            testCase.verifyEqual(hx_e(s), string(o.sectpr_geometry.serhex), ...
                'sectPr geometry write serhex (L1) vs frozen oracle');

            % WRAP (raw int -> Length) vs DIRECT (Length) margin setters (faithful
            % asymmetry): bottom_margin = raw 914400 wraps; top_margin = Twips(-720)
            % direct signed. Both read back exact EMU; serhex identical.
            s = mat2doc.oxml.OxmlElement("w:sectPr");
            s.bottom_margin = 914400;                    % raw int -> Length(914400)
            s.top_margin = mat2doc.shared.Twips(-720);   % direct signed Length
            testCase.verifyEqual(double(s.bottom_margin), 914400, 'bottom_margin wrap path exact');
            testCase.verifyEqual(double(s.top_margin), -457200, 'top_margin direct signed exact');
            testCase.verifyEqual(hx_e(s), string(o.sectpr_margins_wrapdirect.serhex), ...
                'sectPr wrap/direct margin serhex (L1)');
        end

        function test_ct_sectpr_start_type(testCase)
            % Regression (s0039 sectpr_start_type): each WD_SECTION_START member
            % writes <w:type>; NEW_PAGE and None each REMOVE the type child (identity
            % semantics); the getter returns NEW_PAGE when the child is absent.
            o = loadOracle();
            members = enumeration('mat2doc.enum.section.WD_SECTION_START');
            for k = 1:numel(members)
                name = char(members(k));
                s = mat2doc.oxml.OxmlElement("w:sectPr");
                s.start_type = members(k);
                testCase.verifyEqual(s.start_type, members(k), ...
                    sprintf('start_type round-trips %s', name));
                testCase.verifyEqual(hx_e(s), string(o.sectpr_start_type.(name).serhex), ...
                    sprintf('start_type %s serhex (L1)', name));
            end
            % NEW_PAGE is the identity-remove case: no w:type child written
            sNP = mat2doc.oxml.OxmlElement("w:sectPr");
            sNP.start_type = mat2doc.enum.section.WD_SECTION_START.NEW_PAGE;
            testCase.verifyTrue(isequal(sNP.type, []), ...
                'start_type NEW_PAGE removes/omits the w:type child (identity)');

            % set EVEN_PAGE then None -> type child REMOVED -> start_type back to NEW_PAGE
            s = mat2doc.oxml.OxmlElement("w:sectPr");
            s.start_type = mat2doc.enum.section.WD_SECTION_START.EVEN_PAGE;
            s.start_type = [];
            testCase.verifyEqual(s.start_type, mat2doc.enum.section.WD_SECTION_START.NEW_PAGE, ...
                'set start_type None removes w:type -> getter NEW_PAGE');
            testCase.verifyEqual(hx_e(s), string(o.sectpr_start_type.none_removes.serhex), ...
                'start_type None removes w:type (L1)');
        end

        function test_ct_sectpr_titlepg_tristate(testCase)
            % (T) Regression (H3 breadth, s0039 sectpr_titlepg): Python
            % `value in [None, False]` is ==-MEMBERSHIP, so None / False / 0 / 0.0
            % ALL REMOVE the child, while True / 1 SET it -> the EMPTY <w:titlePg/>
            % (D-delta-1). The numeric-0 breadth distinguishes this from CT_Settings'
            % IDENTITY `is None or is False`.
            o = loadOracle();

            % set True -> <w:titlePg/> (empty; D-delta-1)
            s = mat2doc.oxml.OxmlElement("w:sectPr");
            s.titlePg_val = true;
            testCase.verifyTrue(s.titlePg_val, 'read-back True');
            testCase.verifyEqual(childLocalnames(s), "titlePg", 'titlePg child present after True');
            testCase.verifyEqual(hx_e(s), string(o.sectpr_titlepg.set_true.serhex), ...
                'titlePg set True -> empty <w:titlePg/> (D-delta-1, L1)');

            % set 1 (numeric) -> also sets (True branch), byte-identical to True
            s = mat2doc.oxml.OxmlElement("w:sectPr");
            s.titlePg_val = 1;
            testCase.verifyTrue(s.titlePg_val, 'numeric 1 sets titlePg -> True');
            testCase.verifyEqual(hx_e(s), string(o.sectpr_titlepg.set_1.serhex), ...
                'titlePg set 1 -> empty <w:titlePg/> (L1)');

            % --- the REMOVE breadth: False / 0 / 0.0 / None ALL remove ---
            removeInputs = {false, 0, 0.0, []};
            oracleKeys   = {'set_false_removes', 'set_0_removes', 'set_0_removes', 'set_none_removes'};
            labels       = {'False', '0 (double)', '0.0', 'None ([])'};
            for i = 1:numel(removeInputs)
                s = mat2doc.oxml.OxmlElement("w:sectPr");
                s.titlePg_val = removeInputs{i};   % starts absent -> stays absent
                testCase.verifyFalse(s.titlePg_val, ...
                    sprintf('titlePg_val %s -> False', labels{i}));
                testCase.verifyEqual(childLocalnames(s), strings(1,0), ...
                    sprintf('titlePg %s removes/omits the child', labels{i}));
                testCase.verifyEqual(hx_e(s), string(o.sectpr_titlepg.(oracleKeys{i}).serhex), ...
                    sprintf('titlePg %s -> bare <w:sectPr/> (L1)', labels{i}));
            end

            % breadth from a PRESENT child: set True then 0 must REMOVE (not keep)
            s = mat2doc.oxml.OxmlElement("w:sectPr");
            s.titlePg_val = true;
            s.titlePg_val = 0;   % numeric 0 == false -> remove
            testCase.verifyEqual(childLocalnames(s), strings(1,0), ...
                'set True then 0 REMOVES the titlePg child (isequal(0,false) breadth)');
            testCase.verifyFalse(s.titlePg_val);
        end

        function test_ct_sectpr_orientation_semantics(testCase)
            % (H4) Regression (s0039 sectpr_orientation): None (falsy) -> PORTRAIT
            % -> @w:orient removed; PORTRAIT (== default) -> removed; LANDSCAPE ->
            % written. All create the pgSz child (get_or_add_pgSz).
            o = loadOracle();

            % None -> get_or_add_pgSz then orient=PORTRAIT -> @w:orient removed (bare pgSz)
            s = mat2doc.oxml.OxmlElement("w:sectPr");
            s.orientation = [];
            testCase.verifyEqual(s.orientation, mat2doc.enum.section.WD_ORIENTATION.PORTRAIT, ...
                'set None -> orientation reads PORTRAIT');
            testCase.verifyEqual(childLocalnames(s), "pgSz", 'None still creates the pgSz child (H4 falsy)');
            testCase.verifyEqual(hx_e(s), string(o.sectpr_orientation.set_none_portrait.serhex), ...
                'orientation None -> <w:pgSz/> (no @orient) (L1)');

            % PORTRAIT -> same (== default removes @orient)
            s = mat2doc.oxml.OxmlElement("w:sectPr");
            s.orientation = mat2doc.enum.section.WD_ORIENTATION.PORTRAIT;
            testCase.verifyEqual(hx_e(s), string(o.sectpr_orientation.set_portrait_removes.serhex), ...
                'orientation PORTRAIT (== default) removes @orient (L1)');

            % LANDSCAPE -> @w:orient="landscape" written
            s = mat2doc.oxml.OxmlElement("w:sectPr");
            s.orientation = mat2doc.enum.section.WD_ORIENTATION.LANDSCAPE;
            testCase.verifyEqual(s.orientation, mat2doc.enum.section.WD_ORIENTATION.LANDSCAPE);
            testCase.verifyEqual(hx_e(s), string(o.sectpr_orientation.set_landscape.serhex), ...
                'orientation LANDSCAPE -> <w:pgSz w:orient="landscape"/> (L1)');
        end

        % =============================================================== %
        % 6. CT_SectPr hdr/ftr references -- add/get/remove/list + H5        %
        % =============================================================== %

        function test_ct_sectpr_hdrftr_refs(testCase)
            % Nominal + Edge + Regression (s0039 hdrftr_surface): add hdr/ftr refs
            % (CT_HdrFtrRef type_+rId), list order + rIds, get returns the SAME live
            % handle as add (H5 identity ==), get-absent -> None, remove returns the
            % rId and detaches, remove-absent raises mat2doc:ValueError verbatim.
            % Built r:-ROOTED (real-package faithful).
            o = loadOracle();

            % ---- header ----
            s = parse("<w:sectPr " + nsWR() + "></w:sectPr>");
            h1 = s.add_headerReference(mat2doc.enum.section.WD_HEADER_FOOTER.PRIMARY, "rId3");
            s.add_headerReference(mat2doc.enum.section.WD_HEADER_FOOTER.FIRST_PAGE, "rId4");
            testCase.verifyEqual(class(h1), testCase.CT_HDRFTRREF, 'add returns a CT_HdrFtrRef');
            testCase.verifyEqual(h1.type_, mat2doc.enum.section.WD_HEADER_FOOTER.PRIMARY);
            testCase.verifyEqual(char(h1.rId), 'rId3');
            lst = s.headerReference_lst;
            testCase.verifyEqual(numel(lst), 2, 'header lst length 2');
            testCase.verifyEqual(rids(lst), {'rId3','rId4'}, 'header lst rIds in doc order');
            getp = s.get_headerReference(mat2doc.enum.section.WD_HEADER_FOOTER.PRIMARY);
            testCase.verifyEqual(char(getp.rId), 'rId3', 'get PRIMARY -> rId3');
            % H5 identity: get returns the SAME element as add
            testCase.verifyTrue(getp == h1, 'H5: get_headerReference == the added handle');
            testCase.verifyTrue(isequal(s.get_headerReference(mat2doc.enum.section.WD_HEADER_FOOTER.EVEN_PAGE), []), ...
                'get EVEN_PAGE (absent) -> [] (None)');
            testCase.verifyEqual(hx_e(s), string(o.hdrftr_surface.header.serhex_after_adds), ...
                'header refs serhex (rooted, L1) vs frozen oracle');
            removed = s.remove_headerReference(mat2doc.enum.section.WD_HEADER_FOOTER.PRIMARY);
            testCase.verifyEqual(char(removed), 'rId3', 'remove returns the rId');
            testCase.verifyEqual(rids(s.headerReference_lst), {'rId4'}, 'header lst after remove');
            % remove-absent -> mat2doc:ValueError verbatim
            bareH = mat2doc.oxml.OxmlElement("w:sectPr");
            testCase.verifyError(@() bareH.remove_headerReference(mat2doc.enum.section.WD_HEADER_FOOTER.PRIMARY), ...
                'mat2doc:ValueError', 'remove-absent header raises mat2doc:ValueError');

            % ---- footer ----
            s = parse("<w:sectPr " + nsWR() + "></w:sectPr>");
            f1 = s.add_footerReference(mat2doc.enum.section.WD_HEADER_FOOTER.PRIMARY, "rId5");
            s.add_footerReference(mat2doc.enum.section.WD_HEADER_FOOTER.EVEN_PAGE, "rId6");
            lst = s.footerReference_lst;
            testCase.verifyEqual(rids(lst), {'rId5','rId6'}, 'footer lst rIds in doc order');
            getp = s.get_footerReference(mat2doc.enum.section.WD_HEADER_FOOTER.PRIMARY);
            testCase.verifyTrue(getp == f1, 'H5: get_footerReference == the added handle');
            testCase.verifyTrue(isequal(s.get_footerReference(mat2doc.enum.section.WD_HEADER_FOOTER.FIRST_PAGE), []), ...
                'get FIRST_PAGE (absent) -> [] (None)');
            testCase.verifyEqual(hx_e(s), string(o.hdrftr_surface.footer.serhex_after_adds), ...
                'footer refs serhex (rooted, L1) vs frozen oracle');
            removed = s.remove_footerReference(mat2doc.enum.section.WD_HEADER_FOOTER.PRIMARY);
            testCase.verifyEqual(char(removed), 'rId5', 'remove returns the rId');
            testCase.verifyEqual(rids(s.footerReference_lst), {'rId6'}, 'footer lst after remove');
            bareF = mat2doc.oxml.OxmlElement("w:sectPr");
            testCase.verifyError(@() bareF.remove_footerReference(mat2doc.enum.section.WD_HEADER_FOOTER.PRIMARY), ...
                'mat2doc:ValueError', 'remove-absent footer raises mat2doc:ValueError');
        end

        function test_hdrftrref_required_attr_errors(testCase)
            % Edge (error path, s0039 hdrftrref): both attributes are REQUIRED --
            % reading @w:type or @r:id on a bare CT_HdrFtrRef raises the IDENTIFIER
            % mat2doc:InvalidXmlError with the verbatim message. Built value pinned too.
            o = loadOracle();

            e = mat2doc.oxml.OxmlElement("w:headerReference");
            e.type_ = mat2doc.enum.section.WD_HEADER_FOOTER.PRIMARY;
            e.rId = "rId7";
            testCase.verifyEqual(hx_e(e), string(o.hdrftrref.built.serhex), ...
                'built headerReference serhex (loose single-ref, ns0; L1) vs oracle');

            bare = mat2doc.oxml.OxmlElement("w:headerReference");
            testCase.verifyError(@() bare.type_, 'mat2doc:InvalidXmlError', ...
                'absent @w:type raises mat2doc:InvalidXmlError');
            testCase.verifyError(@() bare.rId, 'mat2doc:InvalidXmlError', ...
                'absent @r:id raises mat2doc:InvalidXmlError');
            % message VERBATIM (matches the python-docx wording in the oracle)
            testCase.verifyEqual(errmsg(@() bare.type_), string(o.hdrftrref.err_type_absent), ...
                'absent @w:type message verbatim');
            testCase.verifyEqual(errmsg(@() bare.rId), string(o.hdrftrref.err_rId_absent), ...
                'absent @r:id message verbatim');
        end

        % =============================================================== %
        % 7. (H11) successor-slice insertion pins (THE crux)                %
        % =============================================================== %

        function test_h11_successor_slice_pins(testCase)
            % (H11) Regression (the highest-value ordering pins, s0039 h11): the
            % ZeroOrOne/ZeroOrMore successor slices _tag_seq[N:] -> TAG_SEQ(N+1:end).
            % A wrong slice mis-positions the child -> Word repair -> these pins RED.
            o = loadOracle();

            % -- pgMar between pgSz(@4) and paperSrc(@6)  [_tag_seq[5:]->TAG_SEQ(6:end)] --
            e = parse("<w:sectPr " + nsW() + "><w:pgSz w:w=""12240""/><w:paperSrc/></w:sectPr>");
            e.get_or_add_pgMar();
            testCase.verifyEqual(childLocalnames(e), ["pgSz" "pgMar" "paperSrc"], ...
                'H11: pgMar lands EXACTLY between pgSz and paperSrc');
            testCase.verifyEqual(hx_e(e), string(o.h11.pgMar_neighbors.serhex), ...
                'H11 pgMar_neighbors serhex (L1) vs frozen oracle');

            % -- pgSz between type(@3) and pgMar(@5)  [_tag_seq[4:]->TAG_SEQ(5:end)] --
            %    a geometry descriptor inserted among a before + after neighbor;
            %    loose w-only element -> the serialized string is a hard-coded L1 pin.
            e = parse("<w:sectPr " + nsW() + "><w:type w:val=""nextPage""/><w:pgMar/></w:sectPr>");
            e.get_or_add_pgSz();
            testCase.verifyEqual(childLocalnames(e), ["type" "pgSz" "pgMar"], ...
                'H11: pgSz lands EXACTLY between type and pgMar');
            testCase.verifyEqual(ser(e), decl() + newline + ...
                "<w:sectPr xmlns:w=""" + testCase.W + """>" + ...
                "<w:type w:val=""nextPage""/><w:pgSz/><w:pgMar/></w:sectPr>", ...
                'H11 pgSz-between-neighbors serialized bytes (L1) hard-coded');

            % -- 2nd headerReference AFTER an intervening footerReference (refs-first-
            %    then-type: neither ref tag is a _tag_seq member, so their successors
            %    span the WHOLE sequence). r:-ROOTED (real-package faithful). --
            s = parse("<w:sectPr " + nsWR() + "></w:sectPr>");
            s.add_headerReference(mat2doc.enum.section.WD_HEADER_FOOTER.PRIMARY, "rId1");
            s.add_footerReference(mat2doc.enum.section.WD_HEADER_FOOTER.PRIMARY, "rId2");
            s.add_headerReference(mat2doc.enum.section.WD_HEADER_FOOTER.FIRST_PAGE, "rId3");
            testCase.verifyEqual(childLocalnames(s), ...
                ["headerReference" "footerReference" "headerReference"], ...
                'H11: 2nd headerReference lands AFTER the intervening footerReference');
            testCase.verifyEqual(hx_e(s), string(o.h11.second_href_after_footer.serhex), ...
                'H11 second_href_after_footer serhex (rooted, L1) vs frozen oracle');
        end

        % =============================================================== %
        % 8. (M) M1 document.xml byte-pin (registry-adding parse-path guard) %
        % =============================================================== %

        function test_m1_document_xml_byte_identical(testCase)
            % (M) Regression (byte-neutrality, L1): mat2doc.Document().save() emits
            % word/document.xml at EXACTLY 1548 B with the frozen s0001 SHA-256 --
            % byte-identical DESPITE the <w:sectPr> subtree now parsing to CT_SectPr /
            % CT_PageSz / CT_PageMar on the save path (P5-2a's CENTRAL parse-path
            % risk). SHA-256 equality is an L1 assertion.
            bytes = emitDocPart('document.xml');
            testCase.verifyEqual(numel(bytes), testCase.DOC_SIZE_M1, ...
                sprintf('word/document.xml must be exactly %d B after the section registry rows', ...
                    testCase.DOC_SIZE_M1));
            testCase.verifyEqual(sha256hex(bytes), testCase.DOC_SHA_M1, ...
                'word/document.xml SHA-256 must equal the frozen s0001 oracle (byte-identical L1)');
        end

        function test_document_xml_parse_serialize_roundtrip_L1(testCase)
            % Regression (byte-neutrality, L1): parse the emitted document.xml back
            % through mat2doc.oxml.parse_xml (instantiating CT_Document at the root
            % and CT_SectPr/CT_PageSz/CT_PageMar in the sectPr subtree) and
            % re-serialize -- must be byte-identical to the input.
            inBytes  = emitDocPart('document.xml');
            root     = mat2doc.oxml.parse_xml(inBytes);
            outBytes = mat2doc.oxml.serialize_part_xml(root);
            testCase.verifyEqual(uint8(outBytes(:)'), uint8(inBytes(:)'), ...
                'document.xml parse->serialize must be byte-identical (CT_SectPr path byte-neutral)');
            % the body sectPr parses to CT_SectPr with the real default geometry
            sectPrs = root.xpath("./w:body/w:sectPr");
            testCase.verifyEqual(numel(sectPrs), 1, 'exactly one body sectPr');
            sp = sectPrs(1);
            testCase.verifyEqual(class(sp), testCase.CT_SECTPR, 'body sectPr parses as CT_SectPr');
            testCase.verifyEqual(double(sp.page_width), 7772400, 'parsed page_width 12240tw');
            testCase.verifyEqual(double(sp.page_height), 10058400, 'parsed page_height 15840tw');
            testCase.verifyEqual(sp.orientation, mat2doc.enum.section.WD_ORIENTATION.PORTRAIT, ...
                'parsed default orientation PORTRAIT');
        end

        % =============================================================== %
        % 9. (G) Geometry-write + (L) LANDSCAPE + extended byte-pins         %
        % =============================================================== %

        function test_geometry_write_byte_identical(testCase)
            % (G) Regression (novel WRITE path, L1): page_width/height/orientation/
            % top_margin on the body sectPr -> save -> word/document.xml == 1568 B,
            % SHA-256 698367cd... (frozen s0040). The mutated subtree is
            % <w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/> +
            % <w:pgMar w:top="360" .../>; cols/docGrid untouched.
            bytes = emitDocPartMutated(@geoMutate);
            testCase.verifyEqual(numel(bytes), testCase.DOC_SIZE_GEO, ...
                sprintf('geometry-write word/document.xml must be exactly %d B', testCase.DOC_SIZE_GEO));
            testCase.verifyEqual(sha256hex(bytes), testCase.DOC_SHA_GEO, ...
                'geometry-write document.xml SHA-256 must equal the frozen s0040 oracle (L1)');
            % structural corroboration: orient APPENDED after w,h (no swap)
            xml = string(native2unicode(bytes, 'UTF-8'));
            testCase.verifyTrue(contains(xml, '<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>'), ...
                'geo pgSz: w,h,orient in schema order');
            testCase.verifyTrue(contains(xml, '<w:pgMar w:top="360"'), 'geo pgMar top=360');
        end

        function test_landscape_byte_identical(testCase)
            % (L) Regression (orientation-semantics / NO w/h swap guard, L1): setting
            % ONLY orientation=LANDSCAPE on the body sectPr -> word/document.xml ==
            % 1569 B, SHA-256 d165a628... (frozen s0041). CRUCIALLY w:w/w:h stay the
            % DEFAULT PORTRAIT dims (12240/15840) -- a port that swapped dims on
            % landscape would diverge here.
            bytes = emitDocPartMutated(@landscapeMutate);
            testCase.verifyEqual(numel(bytes), testCase.DOC_SIZE_LANDSCAPE, ...
                sprintf('landscape word/document.xml must be exactly %d B', testCase.DOC_SIZE_LANDSCAPE));
            testCase.verifyEqual(sha256hex(bytes), testCase.DOC_SHA_LANDSCAPE, ...
                'landscape document.xml SHA-256 must equal the frozen s0041 oracle (L1)');
            % *** NO-SWAP GUARD: w/h are the UNCHANGED default portrait dims ***
            xml = string(native2unicode(bytes, 'UTF-8'));
            testCase.verifyTrue(contains(xml, '<w:pgSz w:w="12240" w:h="15840" w:orient="landscape"/>'), ...
                'LANDSCAPE: w:w/w:h UNCHANGED (12240/15840); only @w:orient appended (NO swap)');
        end

        function test_extended_scenarios_byte_identical(testCase)
            % Regression (extended novel paths, L1; validate report §4): three
            % single-mutation saves pinned by size + SHA-256 prefix.
            %   * margin removed : top_margin = None  -> 1535 B / b45b2c6b...
            %   * start_type     : start_type = EVEN_PAGE -> 1574 B / 23f89f6b...
            %   * titlePg        : titlePg_val = True  -> 1560 B / 7b41c7a5...
            % SHA prefixes are pinned (the report froze prefixes); size is exact.
            cases = { ...
                @margnoneMutate, testCase.DOC_SIZE_MARGNONE, testCase.DOC_SHA_MARGNONE, 'margin-removed'; ...
                @startdocMutate, testCase.DOC_SIZE_STARTDOC, testCase.DOC_SHA_STARTDOC, 'start_type=EVEN_PAGE'; ...
                @titlepgMutate,  testCase.DOC_SIZE_TITLEPG,  testCase.DOC_SHA_TITLEPG,  'titlePg=True' };
            for i = 1:size(cases, 1)
                bytes = emitDocPartMutated(cases{i, 1});
                testCase.verifyEqual(numel(bytes), cases{i, 2}, ...
                    sprintf('%s document.xml size', cases{i, 4}));
                testCase.verifyEqual(extractBefore(sha256hex(bytes), 9), cases{i, 3}, ...
                    sprintf('%s document.xml SHA-256 prefix (L1)', cases{i, 4}));
            end
        end

        % =============================================================== %
        % 10. D-nsprefix-rewrite disposition -- rooted multi-ref is L1       %
        % =============================================================== %

        function test_rooted_multiref_byte_identical(testCase)
            % Regression (D-nsprefix-rewrite disposition, L1): the pre-existing SIGNED
            % L2 deviation governs ONLY loose (un-rooted) multi-ref sectPr r:-prefix
            % numbering (lxml ns0/ns1/ns2 vs the port's reused ns0) -- dead-on-
            % generation. Every REACHABLE (r:-ROOTED, real-document) multi-ref path
            % is BYTE-IDENTICAL: the refs render with the in-scope r: prefix
            % (r:id="rId1"...), and NO ns0/nsN ever appears. This pins the clean
            % reachable path; it does NOT pin the loose divergence as correct.
            s = parse("<w:sectPr " + nsWR() + "></w:sectPr>");
            s.add_headerReference(mat2doc.enum.section.WD_HEADER_FOOTER.PRIMARY, "rId1");
            s.add_footerReference(mat2doc.enum.section.WD_HEADER_FOOTER.PRIMARY, "rId2");
            s.add_headerReference(mat2doc.enum.section.WD_HEADER_FOOTER.FIRST_PAGE, "rId3");
            xml = ser(s);
            testCase.verifyEqual(xml, decl() + newline + ...
                "<w:sectPr xmlns:w=""" + testCase.W + """ " + ...
                "xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">" + ...
                "<w:headerReference w:type=""default"" r:id=""rId1""/>" + ...
                "<w:footerReference w:type=""default"" r:id=""rId2""/>" + ...
                "<w:headerReference w:type=""first"" r:id=""rId3""/></w:sectPr>", ...
                'rooted multi-ref sectPr is byte-identical (r: prefix, rId1..rId3; NO ns0/nsN)');
            testCase.verifyFalse(contains(xml, "ns0") || contains(xml, "ns1"), ...
                'no auto-numbered ns prefix appears in the reachable rooted path');
        end

        % =============================================================== %
        % 11. EQUIVALENCE -- full s0039 battery vs the frozen oracle         %
        % =============================================================== %

        function test_equivalence_full_battery_vs_frozen_oracle(testCase)
            % Equivalence: replay the ENTIRE s0039 battery (runProbes -- the .m twin's
            % body VERBATIM: pagesz/pagemar/secttype/hdrftrref/sectpr_* / hdrftr_surface
            % / h11) and flatten-compare EVERY leaf to the frozen python-docx 1.2.0
            % oracle copied into data\s0039_probe_oracle.json. Gate-3 found ZERO
            % divergences (probe_diff exit 0), so every leaf must be byte/value-
            % identical. Ties the suite to the Gate-3 output.
            port   = runProbes();
            oracle = loadOracle();
            pMap = containers.Map('KeyType','char','ValueType','char');
            oMap = containers.Map('KeyType','char','ValueType','char');
            flattenLeaves(port,   '', pMap);
            flattenLeaves(oracle, '', oMap);

            pKeys = sort(pMap.keys());
            oKeys = sort(oMap.keys());
            testCase.verifyEqual(pKeys, oKeys, ...
                'the replayed battery and the frozen oracle must have identical leaf keys');
            % Non-trivial size guard (guards a silent-empty replay). The only looser-
            % than-byte assertion in this class; justified as a floor on leaf count.
            testCase.verifyGreaterThan(numel(oKeys), 60, ...
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

function s = decl()
    s = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>";
end

function s = W_()
    s = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
end

function s = nsW()
    s = "xmlns:w=""" + W_() + """";
end

function s = nsWR()
    s = "xmlns:w=""" + W_() + """ " + ...
        "xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""";
end

function e = parse(xml)
    e = mat2doc.oxml.parse_xml(uint8(unicode2native(char(xml), "UTF-8")));
end

function s = ser(e)
    % Decode serialize_part_xml bytes to a string (ASCII content -> string-equality
    % is a byte-identical L1 assertion; non-ASCII round-trips via UTF-8).
    s = string(native2unicode(mat2doc.oxml.serialize_part_xml(e), "UTF-8"));
end

function h = hx(raw)
    % UPPERCASE hex of raw UTF-8 bytes (matches Python bytes.hex().upper()).
    h = string(sprintf('%02X', uint8(raw)));
end

function h = hx_e(e)
    h = hx(mat2doc.oxml.serialize_part_xml(e));
end

function names = childLocalnames(e)
    kids = e.xpath("./*");
    names = strings(1, numel(kids));
    for k = 1:numel(kids)
        names(k) = string(kids(k).local_part);
    end
end

function C = lnsCell(e)
    % Ordered child localnames as a 1xN cell row of char ({} when no children).
    kids = e.xpath("./*");
    if isempty(kids)
        C = cell(1, 0);
        return
    end
    C = cell(1, numel(kids));
    for k = 1:numel(kids)
        C{k} = char(kids(k).local_part);
    end
end

function C = rids(lst)
    % Cell row of raw rId strings (mirrors python [x.rId for x in lst]).
    if isempty(lst)
        C = cell(1, 0);
        return
    end
    C = cell(1, numel(lst));
    for k = 1:numel(lst)
        C{k} = char(lst(k).rId);
    end
end

function s = rv(x)
    % Uniform accessor repr mirroring the s0039 rv(): None->"None",
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

function o = loadOracle()
    % Read the co-located frozen oracle in BINARY mode (no CRLF translation) and
    % decode UTF-8 -> struct. jsondecode is whitespace/line-ending agnostic, so no
    % `* binary` .gitattributes pin is needed (value-based fixture, s0030/s0036
    % precedent).
    here = fileparts(mfilename('fullpath'));
    p = fullfile(here, 'data', 's0039_probe_oracle.json');
    fid = fopen(p, 'r', 'n');
    if fid < 0
        error('mat2doc:test:openFail', 'cannot open frozen oracle %s', p);
    end
    raw = fread(fid, Inf, '*uint8')';
    fclose(fid);
    o = jsondecode(native2unicode(raw, 'UTF-8'));
end

% ---- geometry / mutation closures for the document.xml byte-pins ----

function sp = bodySectPr(d)
    sectPrs = d.element.xpath("./w:body/w:sectPr");
    sp = sectPrs(1);
end

function geoMutate(d)
    sp = bodySectPr(d);
    sp.page_width  = mat2doc.shared.Twips(15840);
    sp.page_height = mat2doc.shared.Twips(12240);
    sp.orientation = mat2doc.enum.section.WD_ORIENTATION.LANDSCAPE;
    sp.top_margin  = mat2doc.shared.Twips(360);
end

function landscapeMutate(d)
    sp = bodySectPr(d);
    sp.orientation = mat2doc.enum.section.WD_ORIENTATION.LANDSCAPE;
end

function margnoneMutate(d)
    sp = bodySectPr(d);
    sp.top_margin = [];          % remove @w:top from the default pgMar
end

function startdocMutate(d)
    sp = bodySectPr(d);
    sp.start_type = mat2doc.enum.section.WD_SECTION_START.EVEN_PAGE;
end

function titlepgMutate(d)
    sp = bodySectPr(d);
    sp.titlePg_val = true;       % insert <w:titlePg/>
end

function bytes = emitDocPart(partLeaf)
    % Document().save() to a temp .docx, unzip, return word/<partLeaf> raw bytes.
    d = mat2doc.Document();
    bytes = saveAndExtract(d, partLeaf);
end

function bytes = emitDocPartMutated(mutate)
    % Document().save() after applying a mutation closure to the body sectPr,
    % then extract word/document.xml raw bytes (the geometry novel paths).
    d = mat2doc.Document();
    mutate(d);
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

function P = runProbes()
    % Replay the s0039 probe sequence (the .m twin's P-struct body, VERBATIM
    % tags/inputs/order) and return the nested struct of tagged canonical values.
    % Embedded here so the Equivalence leg is self-contained (the validation-folder
    % scenario is NOT on the toolbox path and must not be a dependency). Mirrors
    % validation\mat2doc\scenarios\s0039_p5_2a_sectpr_probe.m lines 22-202.
    WDO  = @(n) mat2doc.enum.section.WD_ORIENTATION.(n);
    WDS  = @(n) mat2doc.enum.section.WD_SECTION_START.(n);
    WDHF = @(n) mat2doc.enum.section.WD_HEADER_FOOTER.(n);
    TW   = @(v) mat2doc.shared.Twips(v);

    P = struct();

    % ---- pagesz ----
    ps = struct();
    b = nw("w:pgSz");
    ps.bare = struct("w", rv(b.w), "h", rv(b.h), "orient", rv(b.orient));
    e = nw("w:pgSz"); e.w = TW(12240); e.h = TW(15840);
    ps.set_wh = struct("serhex", hx_e(e), "w", rv(e.w), "h", rv(e.h));
    e = nw("w:pgSz"); e.orient = WDO("LANDSCAPE");
    ps.set_landscape = struct("serhex", hx_e(e), "orient", rv(e.orient));
    e = nw("w:pgSz"); e.orient = WDO("LANDSCAPE"); e.orient = WDO("PORTRAIT");
    ps.set_portrait_removes = struct("serhex", hx_e(e), "orient", rv(e.orient));
    e = nw("w:pgSz"); e.orient = WDO("LANDSCAPE"); e.orient = [];
    ps.set_orient_none = struct("serhex", hx_e(e), "orient", rv(e.orient));
    P.pagesz = ps;

    % ---- pagemar ----
    pm = struct();
    b = nw("w:pgMar");
    pm.bare = struct("top", rv(b.top), "right", rv(b.right), "bottom", rv(b.bottom), ...
        "left", rv(b.left), "header", rv(b.header), "footer", rv(b.footer), ...
        "gutter", rv(b.gutter));
    e = nw("w:pgMar");
    e.top = TW(-720); e.right = TW(1800); e.bottom = TW(-360); e.left = TW(1440);
    e.header = TW(720); e.footer = TW(720); e.gutter = TW(0);
    pm.all_set = struct("serhex", hx_e(e), "top", rv(e.top), "right", rv(e.right), ...
        "bottom", rv(e.bottom), "left", rv(e.left), "header", rv(e.header), ...
        "footer", rv(e.footer), "gutter", rv(e.gutter));
    e.top = [];
    pm.top_none_removes = struct("serhex", hx_e(e), "top", rv(e.top));
    P.pagemar = pm;

    % ---- secttype ----
    st = struct();
    b = nw("w:type");
    st.bare = struct("val", rv(b.val));
    members = struct();
    marr = enumeration('mat2doc.enum.section.WD_SECTION_START');
    for k = 1:numel(marr)
        e = nw("w:type"); e.val = marr(k);
        members.(char(marr(k))) = struct("serhex", hx_e(e), "val", rv(e.val));
    end
    st.members = members;
    e = nw("w:type"); e.val = WDS("NEW_PAGE"); e.val = [];
    st.none_removes = struct("serhex", hx_e(e), "val", rv(e.val));
    P.secttype = st;

    % ---- hdrftrref ----
    hr = struct();
    e = nw("w:headerReference"); e.type_ = WDHF("PRIMARY"); e.rId = "rId7";
    hr.built = struct("serhex", hx_e(e), "type_", rv(e.type_), "rId", rv(e.rId));
    bare = nw("w:headerReference");
    hr.err_type_absent = errmsg(@() bare.type_);
    hr.err_rId_absent = errmsg(@() bare.rId);
    P.hdrftrref = hr;

    % ---- sectpr_defaults ----
    s = nw("w:sectPr");
    P.sectpr_defaults = struct( ...
        "page_width", rv(s.page_width), "page_height", rv(s.page_height), ...
        "orientation", rv(s.orientation), "start_type", rv(s.start_type), ...
        "titlePg_val", rv(s.titlePg_val), "top_margin", rv(s.top_margin), ...
        "bottom_margin", rv(s.bottom_margin), "left_margin", rv(s.left_margin), ...
        "right_margin", rv(s.right_margin), "header", rv(s.header), ...
        "footer", rv(s.footer), "gutter", rv(s.gutter));

    % ---- sectpr_geometry ----
    s = nw("w:sectPr");
    s.page_width = TW(15840); s.page_height = TW(12240);
    s.orientation = WDO("LANDSCAPE"); s.top_margin = TW(360);
    P.sectpr_geometry = struct("serhex", hx_e(s), "page_width", rv(s.page_width), ...
        "page_height", rv(s.page_height), "orientation", rv(s.orientation), ...
        "top_margin", rv(s.top_margin));

    % ---- sectpr_margins_wrapdirect ----
    s = nw("w:sectPr");
    s.bottom_margin = 914400;      % raw int -> WRAP path Length(914400)
    s.top_margin = TW(-720);       % direct SIGNED Length
    P.sectpr_margins_wrapdirect = struct("serhex", hx_e(s), ...
        "bottom_margin", rv(s.bottom_margin), "top_margin", rv(s.top_margin));

    % ---- sectpr_start_type ----
    stp = struct();
    for k = 1:numel(marr)
        s = nw("w:sectPr"); s.start_type = marr(k);
        stp.(char(marr(k))) = struct("serhex", hx_e(s), "start_type", rv(s.start_type));
    end
    s = nw("w:sectPr"); s.start_type = WDS("EVEN_PAGE"); s.start_type = [];
    stp.none_removes = struct("serhex", hx_e(s), "start_type", rv(s.start_type));
    P.sectpr_start_type = stp;

    % ---- sectpr_titlepg ----
    tp = struct();
    s = nw("w:sectPr"); s.titlePg_val = true;
    tp.set_true = struct("serhex", hx_e(s), "val", rv(s.titlePg_val));
    s = nw("w:sectPr"); s.titlePg_val = false;
    tp.set_false_removes = struct("serhex", hx_e(s), "val", rv(s.titlePg_val));
    s = nw("w:sectPr"); s.titlePg_val = 0;
    tp.set_0_removes = struct("serhex", hx_e(s), "val", rv(s.titlePg_val));
    s = nw("w:sectPr"); s.titlePg_val = [];
    tp.set_none_removes = struct("serhex", hx_e(s), "val", rv(s.titlePg_val));
    s = nw("w:sectPr"); s.titlePg_val = 1;
    tp.set_1 = struct("serhex", hx_e(s), "val", rv(s.titlePg_val));
    P.sectpr_titlepg = tp;

    % ---- sectpr_orientation ----
    orx = struct();
    s = nw("w:sectPr"); s.orientation = [];
    orx.set_none_portrait = struct("serhex", hx_e(s), "orientation", rv(s.orientation));
    s = nw("w:sectPr"); s.orientation = WDO("PORTRAIT");
    orx.set_portrait_removes = struct("serhex", hx_e(s), "orientation", rv(s.orientation));
    s = nw("w:sectPr"); s.orientation = WDO("LANDSCAPE");
    orx.set_landscape = struct("serhex", hx_e(s), "orientation", rv(s.orientation));
    P.sectpr_orientation = orx;

    % ---- hdrftr_surface (r:-ROOTED) ----
    surf = struct();
    s = parse("<w:sectPr " + nsWR() + "></w:sectPr>");
    h1 = s.add_headerReference(WDHF("PRIMARY"), "rId3");
    s.add_headerReference(WDHF("FIRST_PAGE"), "rId4");
    lst = s.headerReference_lst;
    getp = s.get_headerReference(WDHF("PRIMARY"));
    hh = struct();
    hh.lst_len = rv(numel(lst));
    hh.lst_rids = rids(lst);
    hh.get_primary_rid = rv(getp.rId);
    hh.get_even_absent = rv(s.get_headerReference(WDHF("EVEN_PAGE")));
    hh.eq_get_is_add = rv(getp == h1);
    hh.serhex_after_adds = hx_e(s);
    removed = s.remove_headerReference(WDHF("PRIMARY"));
    lst2 = s.headerReference_lst;
    hh.removed_rid = rv(removed);
    hh.lst_len_after = rv(numel(lst2));
    hh.lst_rids_after = rids(lst2);
    hh.err_remove_absent = errmsg(@() nw("w:sectPr").remove_headerReference(WDHF("PRIMARY")));
    surf.header = hh;
    s = parse("<w:sectPr " + nsWR() + "></w:sectPr>");
    f1 = s.add_footerReference(WDHF("PRIMARY"), "rId5");
    s.add_footerReference(WDHF("EVEN_PAGE"), "rId6");
    lst = s.footerReference_lst;
    getp = s.get_footerReference(WDHF("PRIMARY"));
    ff = struct();
    ff.lst_len = rv(numel(lst));
    ff.lst_rids = rids(lst);
    ff.get_primary_rid = rv(getp.rId);
    ff.get_first_absent = rv(s.get_footerReference(WDHF("FIRST_PAGE")));
    ff.eq_get_is_add = rv(getp == f1);
    ff.serhex_after_adds = hx_e(s);
    removed = s.remove_footerReference(WDHF("PRIMARY"));
    lst2 = s.footerReference_lst;
    ff.removed_rid = rv(removed);
    ff.lst_len_after = rv(numel(lst2));
    ff.lst_rids_after = rids(lst2);
    ff.err_remove_absent = errmsg(@() nw("w:sectPr").remove_footerReference(WDHF("PRIMARY")));
    surf.footer = ff;
    P.hdrftr_surface = surf;

    % ---- h11 ----
    h11 = struct();
    e = parse("<w:sectPr " + nsW() + "><w:pgSz w:w=""12240""/><w:paperSrc/></w:sectPr>");
    e.get_or_add_pgMar();
    h11.pgMar_neighbors = struct("localnames", {lnsCell(e)}, "serhex", hx_e(e));
    s = parse("<w:sectPr " + nsWR() + "></w:sectPr>");
    s.add_headerReference(WDHF("PRIMARY"), "rId1");
    s.add_footerReference(WDHF("PRIMARY"), "rId2");
    s.add_headerReference(WDHF("FIRST_PAGE"), "rId3");
    h11.second_href_after_footer = struct("localnames", {lnsCell(s)}, "serhex", hx_e(s));
    P.h11 = h11;
end

function e = nw(tag)
    e = mat2doc.oxml.OxmlElement(tag);
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

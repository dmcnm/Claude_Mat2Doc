classdef Test_p8_2_comments < matlab.unittest.TestCase
% TEST_P8_2_COMMENTS  Gate-4 permanent unit tests for Mat2Doc P8-2 [N] -- the
%   COMMENTS tier: oxml/comments + the comments API + CommentsPart on-demand +
%   the WML_COMMENTS PartFactory flip. The LAST new part type in the port.
%   src/docx/oxml/comments.py -> +mat2doc\+oxml\+comments\{CT_Comments, CT_Comment};
%   src/docx/comments.py -> +mat2doc\+comments\{Comments, Comment};
%   src/docx/parts/comments.py -> +mat2doc\+parts\CommentsPart;
%   src/docx/document.py::Document.{add_comment, comments} (un-stubbed);
%   src/docx/parts/document.py::DocumentPart.{comments, _comments_part}
%   (un-stubbed); + the 2 comments registry rows (w:comments -> CT_Comments,
%   w:comment -> CT_Comment) + the WML_COMMENTS -> CommentsPart PartFactory flip.
%
%   P8-2 is an ON-DEMAND-PART, byte-critical WP: default.docx ships NO comments
%   part, so the FIRST comment access materializes word/comments.xml, adds the
%   [Content_Types].xml <Override>, inserts the document.xml.rels COMMENTS rel,
%   and writes the comment-range markers (w:commentRangeStart/End +
%   w:commentReference) into document.xml. Equivalence is therefore FULL-PACKAGE
%   byte identity (18 parts = 17 default + word/comments.xml). This class freezes
%   that surface byte-identical to python-docx v1.2.0.
%
%   This class permanently freezes what the prior gates established:
%     * Gate-1 Porter  : audit_P8-2_comments.md (self-probe 13/13; headline
%                        comments.xml caee0a20...; M1 17/17; ST_DateTime discharge
%                        with ZERO new D-numbers, residual under D-002).
%     * Gate-2 Auditor (Fable): APPROVE -- 18/18 parts x 7 packages byte-identical;
%                        ST_DateTime zero-new-D CONFIRMED; VERIFY-COMMENTTEXT keep
%                        p_lst(end) CONFIRMED; C3 seam PASS; no fixes.
%     * Gate-3 Validator: validate_P8-2_comments.md -- PASS, 7/7 byte scenarios
%                        full-package L1, probe MATCH (s0107), M1 17/17 (s0108),
%                        ZERO new D-number. Froze references\s0101..s0108.
%
%   THE HIGHEST-VALUE PERMANENT PINS IN THIS WP (a regression goes RED loudly):
%     * ★ (s0101 HEADLINE) test_s0101_one_comment_full_package -- the on-demand
%       comments part end-to-end: Document(); add_paragraph; add_run; add_comment;
%       (date pinned tz-aware UTC 2021-06-15T09:30:00Z via the PUBLIC accessor);
%       save() -> ALL 18 parts byte-identical (size + SHA-256, in frozen zip-entry
%       order) to the frozen s0101 python-docx oracle. word/comments.xml == 990 B
%       / 6a927c1e...; the comment-range markers in document.xml; the on-demand
%       [Content_Types] Override + document.xml.rels COMMENTS rel (rId9). RED on
%       ANY single-byte / attr-order / rel / Override / zip-order drift.
%     * ★ (H3 ABSENT) test_s0102_empty_no_initials_full_package + the CT-level
%       test_ct_comment_initials_absent_when_none -- initials passed [] (None)
%       OMITS @w:initials; the attr order stays w:id, w:author, w:date. Byte-pinned
%       (comments.xml 940 B) AND serialized-substring-pinned.
%     * ★ (VERIFY-COMMENTTEXT) test_s0103_multi_para_full_package + the CT-level
%       test_verify_commenttext_append_path -- the newline-split multi-paragraph
%       path: each appended <w:p> gets the "CommentText" pStyle set DIRECTLY on the
%       element, read back as element_().p_lst(end). Byte-pinned (comments.xml
%       1162 B) AND p_lst(end).style pinned.
%     * ★ (H2 UTF-8) test_s0104_nonascii_full_package -- é / café / CJK 第二行 /
%       José / Jé all emitted as raw UTF-8 bytes in @w:author, @w:initials, <w:t>.
%       Byte-pinned (comments.xml 1081 B / caee0a20... == the porter's §5 headline).
%     * ★ (id auto-increment) test_s0105_two_comments_full_package + the CT-level
%       test_ct_comments_add_comment_autoincrement -- two comments -> ids 0,1; two
%       range-marker pairs. Byte-pinned + CT-level id arithmetic.
%     * ★ (REOPEN / factory flip) test_s0106_roundtrip_full_package +
%       test_reopen_loads_commentspart -- open a python-docx-built one-comment
%       .docx (WML_COMMENTS dispatches to CommentsPart) and save UNCHANGED ->
%       18/18 byte-identical to Python's OWN round-trip (SAME 6a927c1e...).
%     * ★ (ST_DateTime) test_st_datetime_roundtrip_utc -- the FIRST reachable
%       ST_DateTime consumer: set date = datetime(2021,6,15,9,30,0,UTC) -> the raw
%       @w:date and the .date getter both round-trip 2021-06-15T09:30:00Z (whole
%       seconds + literal Z). Pinned tz-aware UTC (NEVER naive/date-only -- those
%       are host-tz-dependent, Gate-2/Gate-3 binding note).
%     * ★ (M1) test_s0108_m1_17_of_17 -- Document().save() emits all 17 default
%       parts byte-identical (NO word/comments.xml): the two comments registry rows
%       + the WML_COMMENTS->CommentsPart flip re-class ZERO M1 parts.
%     * (s0107 EQUIVALENCE) test_s0107_api_probe_vs_frozen_oracle -- the full
%       22-fact Comments/Comment/CommentsPart API probe value-identical to the
%       frozen probe.json (Gate-3 probe_diff MATCH, exit 0).
%
%   Provenance (all Gate-3 frozen 2026-08-02):
%     * Audit    : validation\mat2doc\audit_P8-2_comments.md
%     * Validate : validation\mat2doc\validate_P8-2_comments.md
%     * Scenarios: validation\mat2doc\scenarios\s0101..s0108_p8_2_*.{py,m} (the
%                  IDENTICAL-sequence byte/probe twins; their build bodies are
%                  replayed VERBATIM by the runS01xx helpers below).
%     * Frozen refs (python-docx 1.2.0 / lxml 5.3.0 oracle, frozen ONCE at Gate-3):
%         references\s0101..s0106,s0108\manifest.json (name|size|sha256 in frozen
%           zip-entry order) copied into tests\parts\data\s01NN\; references\s0106\
%           input.docx (the frozen reopen INPUT) copied to data\s0106\input.docx;
%           references\s0107\probe.json copied to data\s0107\probe.json. All are
%           covered by the co-located tests\parts\data\.gitattributes `* binary`
%           pin (the Gate-4 byte-fixture lesson -- no line-ending mangle on a
%           master checkout; jsondecode is line-ending agnostic regardless).
%
%   Coverage taxonomy
%   -----------------
%   * Nominal      -- CT_Comments.add_comment id auto-increment + get_comment_by_id;
%                     CT_Comment attrs (id/author/initials/date); Comments
%                     iter/len_/add_comment/get; Comment C3 seam (paragraphs/
%                     add_paragraph/text/add_run); CommentsPart.default (empty);
%                     the s0101 headline full-package.
%   * Edge         -- H3 initials-absent-when-None (attr OMITTED); empty text ->
%                     single empty paragraph (no run); get(99) -> [] (None);
%                     get_comment_by_id(99) -> []; H2 non-ASCII author/initials/
%                     text; multi-paragraph CommentText append path; the reopen
%                     path (a doc WITH a comments part loads CommentsPart).
%   * Equivalence  -- the s0101..s0106 full-package byte batteries (18/18 L1 each)
%                     + s0108 M1 17/17 + s0107 22-fact API probe vs the frozen
%                     oracle (Gate-3 probe_diff exit 0).
%   * Regression   -- hard-coded full-package SHA-256 (+ size) byte pins (from the
%                     frozen manifests, in frozen zip-entry order) + the tz-aware
%                     UTC ST_DateTime round-trip string + serialized-XML substrings.
%   * Upstream     -- the on-demand comments.xml, the attr order (w:id, w:author,
%                     [w:initials,] w:date), the annotationRef/CommentReference run
%                     scaffold, the comment-range markers, the [Content_Types]
%                     Override + document.xml.rels COMMENTS rel, and the zip-entry
%                     order ARE the python-docx comments.py / parts/comments.py
%                     contract; the frozen oracle IS lxml/python-docx's output.
%
%   Byte-level (L1) note: every full-package assertion is a SHA-256 (+ size) pin of
%   the raw shipping bytes of each zip part, in frozen zip-entry order. SHA-256
%   equality == byte identity (L1). NO D-number granted any L2 relaxation in this
%   WP (Gate-3: ZERO new; the ST_DateTime read-path under-accept residual is under
%   the pre-existing signed D-002, not a P8-2 deviation), so every byte pin here is
%   L1. The only looser-than-byte check is the s0107 API-value probe (tagged
%   strings, not bytes) -- justified because Gate-3 probe_diff already proved value
%   equivalence (exit 0). The mat2doc exception identifiers/messages equal the
%   signed exception-model mapping (design.md section 2 -- non-byte, NOT a D-number).
%
%   Determinism: no network, no absolute paths. The frozen manifests, the reopen
%   input docx, and the probe oracle resolve relative to this file via
%   fileparts(mfilename('fullpath')); saves go to tempname .docx deleted via
%   onCleanup; every file read is binary ('r','n'); the date is PINNED tz-aware UTC
%   (host-stable). The +mat2doc package resolves via the MANDATORY
%   PathFixture(worktree-root) in TestClassSetup (WP9-F4 lesson).

    properties (Constant)
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        % --- the tz-aware UTC instant pinned on every frozen w:date (Gate-2/3) ---
        PINNED_DATE_STR = "2021-06-15T09:30:00Z"

        % --- registered class names (the 2 P8-2 rows) ---
        CT_COMMENTS = 'mat2doc.oxml.comments.CT_Comments'
        CT_COMMENT  = 'mat2doc.oxml.comments.CT_Comment'

        % --- headline witness SHAs (backing the manifest pins, human-readable) ---
        S0101_COMMENTS_SHA = "6a927c1ef5f8f3111e00c250a6fc01f10ac2adf2926c6f28cf8d1b07d761e742"
        S0104_COMMENTS_SHA = "caee0a20c04eb6a1e1a1796bfffa35a44eeea0d608247bb66c6232064649104f"
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied verbatim from tests\parts\Test_p7_4_add_picture.m. here
            % is tests\parts; the worktree root is two levels up.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\parts
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. ★ s0101 HEADLINE -- one comment, full-package byte pin          %
        % =============================================================== %

        function test_s0101_one_comment_full_package(testCase)
            % ★ Regression (L1, THE P8-2 headline): the exact Gate-3 s0101 scenario --
            %   Document(); p=add_paragraph(); r=p.add_run("anchored text");
            %   add_comment(r, "Reviewer note", "Ann", "A"); date pinned tz-aware UTC;
            %   save()
            % emits ALL 18 parts byte-identical (size + SHA-256, in frozen zip-entry
            % order) to the frozen s0101 python-docx oracle. word/comments.xml lands
            % AFTER word/numbering.xml, BEFORE docProps/thumbnail.jpeg (H11 zip-order
            % pin -- the on-demand-part insertion point). RED on ANY single-byte /
            % attr-order / rel / Override / zip-order drift.
            d = runS0101();
            zipBytes = saveZip(d);
            [blobs, names] = zipEntryList(zipBytes);
            testCase.assertPackage(zipBytes, 's0101', 's0101 one comment (** HEADLINE **)');

            % --- loud headline witnesses (backing the SHA pin, human-readable) ---
            cx = entryBlob(blobs, names, "word/comments.xml");
            testCase.verifyEqual(numel(cx), 990, ...
                'headline: word/comments.xml is exactly 990 B');
            testCase.verifyEqual(sha256hex(cx), testCase.S0101_COMMENTS_SHA, ...
                'headline: word/comments.xml SHA == frozen s0101 oracle (6a927c1e)');
            cxs = char(native2unicode(cx, 'UTF-8'));
            testCase.verifyTrue(contains(cxs, 'w:author="Ann"'), ...
                'headline: comments.xml carries @w:author="Ann"');
            testCase.verifyTrue(contains(cxs, ['w:date="' char(testCase.PINNED_DATE_STR) '"']), ...
                'headline: comments.xml @w:date pinned tz-aware UTC 2021-06-15T09:30:00Z');
            doc = char(native2unicode(entryBlob(blobs, names, "word/document.xml"), 'UTF-8'));
            testCase.verifyTrue(contains(doc, '<w:commentRangeStart') && ...
                contains(doc, '<w:commentRangeEnd') && contains(doc, '<w:commentReference'), ...
                'headline: document.xml carries the comment-range markers + reference run');
            rels = char(native2unicode(entryBlob(blobs, names, "word/_rels/document.xml.rels"), 'UTF-8'));
            testCase.verifyTrue(contains(rels, 'Target="comments.xml"'), ...
                'headline: the on-demand document.xml.rels COMMENTS rel targets comments.xml');
            ct = char(native2unicode(entryBlob(blobs, names, "[Content_Types].xml"), 'UTF-8'));
            testCase.verifyTrue(contains(ct, '/word/comments.xml'), ...
                'headline: [Content_Types].xml adds the comments.xml Override');
        end

        % =============================================================== %
        % 2. ★ s0102 H3 -- empty text / initials None -> @w:initials ABSENT  %
        % =============================================================== %

        function test_s0102_empty_no_initials_full_package(testCase)
            % ★ Regression (L1, H3 tri-state): initials passed [] (None) OMITS
            % @w:initials from <w:comment>; text="" yields a single empty paragraph
            % (no run). Attr order stays w:id, w:author, w:date (initials skipped).
            % 18/18 byte-identical to the frozen s0102 oracle (comments.xml 940 B).
            d = runS0102();
            zipBytes = saveZip(d);
            [blobs, names] = zipEntryList(zipBytes);
            testCase.assertPackage(zipBytes, 's0102', 's0102 empty text / initials None (H3 absent)');
            cxs = char(native2unicode(entryBlob(blobs, names, "word/comments.xml"), 'UTF-8'));
            testCase.verifyFalse(contains(cxs, 'w:initials'), ...
                'H3: @w:initials is ABSENT when initials passed [] (None)');
            testCase.verifyTrue(contains(cxs, 'w:author="Bob"'), ...
                's0102: @w:author="Bob" present (required attr, "" would also be present)');
        end

        % =============================================================== %
        % 3. ★ s0103 VERIFY-COMMENTTEXT -- multi-paragraph CommentText path   %
        % =============================================================== %

        function test_s0103_multi_para_full_package(testCase)
            % ★ Regression (L1, VERIFY-COMMENTTEXT): a "\n"-split text -> first
            % paragraph gets the seed run; each appended paragraph gets the
            % "CommentText" pStyle set DIRECTLY on the element (read back as
            % element_().p_lst(end)). 18/18 byte-identical to the frozen s0103 oracle
            % (comments.xml 1162 B); 3 <w:p>.
            d = runS0103();
            zipBytes = saveZip(d);
            [blobs, names] = zipEntryList(zipBytes);
            testCase.assertPackage(zipBytes, 's0103', 's0103 multi-paragraph (CommentText append path)');
            cxs = char(native2unicode(entryBlob(blobs, names, "word/comments.xml"), 'UTF-8'));
            testCase.verifyEqual(numel(strfind(cxs, '<w:p>')) + numel(strfind(cxs, '<w:p ')), 3, ...
                's0103: exactly 3 <w:p> in the comment (seed + 2 appended)');
            testCase.verifyGreaterThanOrEqual(numel(strfind(cxs, 'w:val="CommentText"')), 3, ...
                's0103: each paragraph carries the CommentText pStyle');
        end

        % =============================================================== %
        % 4. ★ s0104 H2 -- non-ASCII author/initials/text (UTF-8 bytes)       %
        % =============================================================== %

        function test_s0104_nonascii_full_package(testCase)
            % ★ Regression (L1, H2 UTF-8): accented Latin (Héllo café / José / Jé) +
            % CJK (第二行) emitted as raw UTF-8 bytes (never numeric entities) in
            % @w:author / @w:initials and <w:t> text, across the newline-split path.
            % 18/18 byte-identical to the frozen s0104 oracle (comments.xml 1081 B /
            % caee0a20 == the porter's §5 headline).
            d = runS0104();
            zipBytes = saveZip(d);
            [blobs, names] = zipEntryList(zipBytes);
            testCase.assertPackage(zipBytes, 's0104', 's0104 non-ASCII author/initials/text (H2 UTF-8)');
            cx = entryBlob(blobs, names, "word/comments.xml");
            testCase.verifyEqual(sha256hex(cx), testCase.S0104_COMMENTS_SHA, ...
                'H2: word/comments.xml SHA == frozen s0104 oracle (caee0a20)');
            % the raw UTF-8 bytes for "café" (é == 0xC3 0xA9) appear verbatim (not "&#233;")
            testCase.verifyTrue(any(cx == 195) && any(cx == 169), ...
                'H2: the UTF-8 bytes of é (0xC3 0xA9) are present raw (not an entity)');
            testCase.verifyFalse(contains(char(native2unicode(cx, 'UTF-8')), '&#'), ...
                'H2: no numeric character entities in comments.xml (raw UTF-8)');
        end

        % =============================================================== %
        % 5. ★ s0105 -- two comments, id auto-increment 0,1                  %
        % =============================================================== %

        function test_s0105_two_comments_full_package(testCase)
            % ★ Regression (L1, id auto-increment): two comments anchored to two runs
            % -> ids 0 then 1, two <w:comment> in comments.xml (document order), two
            % independent comment-range marker pairs in document.xml. 18/18
            % byte-identical to the frozen s0105 oracle (comments.xml 1255 B).
            d = runS0105();
            zipBytes = saveZip(d);
            [blobs, names] = zipEntryList(zipBytes);
            testCase.assertPackage(zipBytes, 's0105', 's0105 two comments (id auto-increment)');
            cxs = char(native2unicode(entryBlob(blobs, names, "word/comments.xml"), 'UTF-8'));
            testCase.verifyTrue(contains(cxs, 'w:id="0"') && contains(cxs, 'w:id="1"'), ...
                's0105: comments.xml carries @w:id 0 and 1 (auto-increment)');
        end

        % =============================================================== %
        % 6. ★ s0106 REOPEN -- open a comment doc + save unchanged           %
        % =============================================================== %

        function test_s0106_roundtrip_full_package(testCase)
            % ★ Regression (L1, PartFactory WML_COMMENTS->CommentsPart open path):
            % open a python-docx-built one-comment .docx (dispatched to CommentsPart
            % by the P8-2 factory flip) and save UNCHANGED -> 18/18 byte-identical to
            % python-docx's OWN open/save output of the SAME input (comments.xml
            % re-serializes to the SAME 6a927c1e as s0101). Proves the flip is
            % byte-neutral on the reopen path.
            input = testCase.dataFile('s0106/input.docx');
            d = mat2doc.Document(input);   % Python: docx.Document(input_docx)
            zipBytes = saveZip(d);
            [blobs, names] = zipEntryList(zipBytes);
            testCase.assertPackage(zipBytes, 's0106', 's0106 reopen-and-save (CommentsPart open path)');
            testCase.verifyEqual(sha256hex(entryBlob(blobs, names, "word/comments.xml")), ...
                testCase.S0101_COMMENTS_SHA, ...
                's0106: the reopened comments.xml re-serializes to the SAME s0101 SHA (byte-neutral)');
        end

        function test_reopen_loads_commentspart(testCase)
            % Nominal (reopen / factory flip, readable witness): opening a doc WITH a
            % comments part dispatches WML_COMMENTS -> CommentsPart (not the base
            % XmlPart); the loaded Comments has len_ 1 and the first comment reads
            % back its author (Ann) / id (0) / text (Reviewer note) -- the s0101 input.
            input = testCase.dataFile('s0106/input.docx');
            d = mat2doc.Document(input);
            cp = d.part.part_related_by(mat2doc.opc.RELATIONSHIP_TYPE.COMMENTS);
            testCase.verifyClass(cp, 'mat2doc.parts.CommentsPart', ...
                'reopen: WML_COMMENTS dispatches to CommentsPart (the P8-2 flip)');
            cs = d.comments;
            testCase.verifyEqual(cs.len_(), 1, 'reopen: the loaded document has 1 comment');
            c0 = cs.get(0);
            testCase.verifyEqual(c0.comment_id, 0, 'reopen: first comment id 0');
            testCase.verifyEqual(string(c0.author), "Ann", 'reopen: first comment author Ann');
            testCase.verifyEqual(string(c0.text), "Reviewer note", 'reopen: first comment text');
        end

        % =============================================================== %
        % 7. ★ s0108 -- M1 17/17 byte-neutrality                             %
        % =============================================================== %

        function test_s0108_m1_17_of_17(testCase)
            % ★ Regression (byte-neutrality, L1): a bare Document().save() emits ALL
            % 17 default parts byte-identical (NO word/comments.xml) to the frozen
            % s0108 oracle -- the two comments registry rows + the WML_COMMENTS ->
            % CommentsPart PartFactory flip re-class ZERO M1 parts (default.docx
            % ships no comments.xml). The on-demand-part neutrality guard.
            d = mat2doc.Document();   % Python: docx.Document()
            zipBytes = saveZip(d);
            [~, names] = zipEntryList(zipBytes);
            testCase.assertPackage(zipBytes, 's0108', 's0108 M1 default (17/17 byte-neutral)');
            testCase.verifyFalse(any(names == "word/comments.xml"), ...
                'M1: default.docx save ships NO word/comments.xml');
        end

        % =============================================================== %
        % 8. CT_Comments -- add_comment id auto-increment + get_comment_by_id%
        % =============================================================== %

        function test_ct_comments_add_comment_autoincrement(testCase)
            % Nominal (comments.py CT_Comments.add_comment): the explicit adder
            % allocates a unique w:id (max+1; empty -> 0), appends the minimum-valid
            % <w:comment> (id, empty author, a CommentText paragraph with a
            % CommentReference/annotationRef run), and returns a CT_Comment.
            cs = mat2doc.oxml.OxmlElement("w:comments");
            testCase.verifyClass(cs, testCase.CT_COMMENTS, 'w:comments -> CT_Comments');
            c0 = cs.add_comment();
            c1 = cs.add_comment();
            c2 = cs.add_comment();
            testCase.verifyClass(c0, testCase.CT_COMMENT, 'add_comment returns a CT_Comment');
            testCase.verifyEqual(c0.id, 0, 'first add_comment -> @w:id 0');
            testCase.verifyEqual(c1.id, 1, 'second add_comment -> @w:id 1');
            testCase.verifyEqual(c2.id, 2, 'third add_comment -> @w:id 2');
            testCase.verifyEqual(numel(cs.comment_lst), 3, '3 <w:comment> children (append)');
            % the minimum-valid scaffold: a CommentReference run + annotationRef
            xml = ser(c0);
            testCase.verifyTrue(contains(xml, 'w:val="CommentReference"') && ...
                contains(xml, '<w:annotationRef/>'), ...
                'add_comment scaffold has the CommentReference run + annotationRef');
        end

        function test_ct_comments_get_comment_by_id_hit_and_miss(testCase)
            % Nominal + Edge (H3): get_comment_by_id returns the matching <w:comment>
            % as the SAME handle in comment_lst (H5), and [] (None) on a miss (NOT an
            % empty typed array).
            cs = mat2doc.oxml.OxmlElement("w:comments");
            cs.add_comment();          % id 0
            cs.add_comment();          % id 1
            hit = cs.get_comment_by_id(1);
            testCase.verifyEqual(hit.id, 1, 'get_comment_by_id(1) -> the id==1 <w:comment>');
            lst = cs.comment_lst;
            testCase.verifyTrue(hit == lst(2), ...
                'H5: get_comment_by_id returns the SAME handle as comment_lst(2)');
            testCase.verifyTrue(isequal(cs.get_comment_by_id(99), []), ...
                'H3: a miss returns [] (None), not an empty typed array');
        end

        % =============================================================== %
        % 9. CT_Comment -- attrs + H3 initials-absent-when-None              %
        % =============================================================== %

        function test_ct_comment_attrs_roundtrip(testCase)
            % Nominal (comments.py CT_Comment attrs): id (ST_DecimalNumber), author
            % (ST_String), initials (ST_String), date (ST_DateTime) round-trip via
            % the get/set descriptors. Serialized attr order is w:id, w:author,
            % w:initials, w:date (parse-then-set insertion order == lxml).
            cs = mat2doc.oxml.OxmlElement("w:comments");
            c = cs.add_comment();
            c.author = "Amy";
            c.initials = "AJ";
            c.date = datetime(2021,6,15,9,30,0,"TimeZone","UTC");
            testCase.verifyEqual(c.id, 0, 'id 0');
            testCase.verifyEqual(string(c.author), "Amy", 'author round-trip');
            testCase.verifyEqual(string(c.initials), "AJ", 'initials round-trip');
            % raw @w:date is the whole-seconds + literal-Z ST_DateTime form
            testCase.verifyEqual(string(c.get(mat2doc.oxml.qn("w:date"))), ...
                testCase.PINNED_DATE_STR, 'raw @w:date == 2021-06-15T09:30:00Z (ST_DateTime to_xml)');
        end

        function test_ct_comment_initials_absent_when_none(testCase)
            % ★ Edge (H3 tri-state): assigning initials = [] (None) REMOVES @w:initials
            % (OptionalAttribute absent-branch); assigning "" keeps it PRESENT-empty.
            % The raw element get(qn) is [] when absent, "" when present-empty.
            cs = mat2doc.oxml.OxmlElement("w:comments");
            c = cs.add_comment();
            c.author = "Bob";
            c.initials = [];                      % Python: initials = None
            testCase.verifyTrue(isequal(c.initials, []), ...
                'H3: initials getter returns [] (None) when @w:initials absent');
            testCase.verifyTrue(isequal(c.get(mat2doc.oxml.qn("w:initials")), []), ...
                'H3: the raw element has NO @w:initials attribute');
            testCase.verifyFalse(contains(ser(c), 'w:initials'), ...
                'H3: @w:initials absent from the serialized <w:comment>');
            % present-empty: assigning "" ADDS the attribute (distinct from None)
            c.initials = "";
            testCase.verifyEqual(string(c.initials), "", ...
                'H3: initials "" is present-empty (distinct from None)');
            testCase.verifyTrue(contains(ser(c), 'w:initials=""'), ...
                'H3: @w:initials="" present when assigned the empty string');
        end

        % =============================================================== %
        % 10. ★ ST_DateTime -- tz-aware UTC round-trip (first reachable use) %
        % =============================================================== %

        function test_st_datetime_roundtrip_utc(testCase)
            % ★ Nominal (comments.py CT_Comment.date -- the FIRST reachable ST_DateTime
            % consumer): set date = datetime(2021,6,15,9,30,0,"TimeZone","UTC") -> the
            % raw @w:date serializes whole-seconds + literal Z (2021-06-15T09:30:00Z),
            % and the .date getter round-trips the SAME instant. Pinned tz-aware UTC
            % (NEVER naive/date-only -- those are host-tz-dependent in BOTH engines,
            % Gate-2/Gate-3 binding note; ZERO new D-number, residual under D-002).
            cs = mat2doc.oxml.OxmlElement("w:comments");
            c = cs.add_comment();
            c.date = datetime(2021,6,15,9,30,0,"TimeZone","UTC");
            % raw attribute (to_xml)
            testCase.verifyEqual(string(c.get(mat2doc.oxml.qn("w:date"))), ...
                testCase.PINNED_DATE_STR, 'ST_DateTime to_xml -> whole-seconds + literal Z');
            % getter (from_xml): a tz-aware datetime pointing at the SAME instant
            dt = c.date;
            testCase.verifyClass(dt, 'datetime', 'date getter -> a datetime');
            dt.TimeZone = "UTC";
            testCase.verifyEqual(string(dt, "yyyy-MM-dd'T'HH:mm:ss'Z'"), testCase.PINNED_DATE_STR, ...
                'ST_DateTime from_xml round-trips the SAME UTC instant');
            % H3: date absent -> [] (None)
            c2 = cs.add_comment();
            testCase.verifyTrue(isequal(c2.date, []), 'H3: date getter [] when @w:date absent');
        end

        % =============================================================== %
        % 11. Comments API -- iter / len_ / add_comment / get                %
        % =============================================================== %

        function test_comments_api_len_add_get(testCase)
            % Nominal (comments.py Comments): a bare Document().comments materializes
            % the on-demand part (len_ 0); add_comment mints a Comment (len_ grows);
            % get(id) hits / get(miss) -> []; to_array materializes a Comment per
            % <w:comment>. H5: fresh Comment views each access.
            d = mat2doc.Document();
            cs = d.comments;                          % materializes the part on demand
            testCase.verifyEqual(cs.len_(), 0, 'a fresh document has 0 comments');
            c0 = cs.add_comment("Hello", "Amy", "AJ");
            testCase.verifyClass(c0, 'mat2doc.comments.Comment', 'add_comment -> a Comment');
            testCase.verifyEqual(cs.len_(), 1, 'len_ 1 after one add_comment');
            testCase.verifyEqual(c0.comment_id, 0, 'first comment id 0');
            testCase.verifyEqual(string(c0.author), "Amy", 'author Amy');
            testCase.verifyEqual(string(c0.initials), "AJ", 'initials AJ');
            testCase.verifyEqual(string(c0.text), "Hello", 'text Hello');
            testCase.verifyEqual(numel(c0.paragraphs), 1, 'single paragraph');

            c1 = cs.add_comment("A" + newline + "B", "Bob", []);   % Python: "A\nB", None
            testCase.verifyEqual(cs.len_(), 2, 'len_ 2 after a second add_comment');
            testCase.verifyEqual(c1.comment_id, 1, 'second comment id 1 (auto-increment)');
            testCase.verifyTrue(isequal(c1.initials, []), 'H3: initials [] when None passed');
            testCase.verifyEqual(string(c1.text), "A" + newline + "B", 'multi-line text joined by \n');
            testCase.verifyEqual(numel(c1.paragraphs), 2, 'two paragraphs for one "\n"');

            % to_array (__iter__) + get hit/miss
            arr = cs.to_array();
            testCase.verifyEqual(numel(arr), 2, 'to_array materializes 2 Comments');
            testCase.verifyClass(arr, 'mat2doc.comments.Comment', 'to_array is a Comment vector');
            testCase.verifyEqual(string(cs.get(0).author), "Amy", 'get(0) hit -> author Amy');
            testCase.verifyTrue(isequal(cs.get(99), []), 'H3: get(99) miss -> [] (None)');
        end

        % =============================================================== %
        % 12. Comment C3 seam -- paragraphs / add_paragraph / add_run / text %
        % =============================================================== %

        function test_comment_c3_seam(testCase)
            % Nominal (comments.py Comment(BlockItemContainer) C3 seam): the inherited
            % BlockItemContainer surface operates directly on the <w:comment>.
            % add_run through the comment's first paragraph, add_paragraph appends a
            % block, and text reflects the joined paragraph content.
            d = mat2doc.Document();
            cs = d.comments;
            c = cs.add_comment("", "Amy", "AJ");     % empty text -> single empty paragraph
            ps = c.paragraphs();
            testCase.verifyEqual(numel(ps), 1, 'empty-text comment starts with one empty paragraph');
            ps(1).add_run("through the comment");    % C3 seam: add_run on the comment paragraph
            testCase.verifyEqual(string(c.text), "through the comment", ...
                'C3: add_run through the comment paragraph updates text');
            p2 = c.add_paragraph("second block");    % C3 seam: add_paragraph appends
            testCase.verifyClass(p2, 'mat2doc.text.Paragraph', 'add_paragraph -> a Paragraph');
            testCase.verifyEqual(numel(c.paragraphs), 2, 'two paragraphs after add_paragraph');
            testCase.verifyEqual(string(c.text), "through the comment" + newline + "second block", ...
                'C3: text joins the two paragraphs with \n (H16: no strip)');
        end

        function test_verify_commenttext_append_path(testCase)
            % ★ Nominal (VERIFY-COMMENTTEXT, comments.py Comment.add_paragraph): when
            % style is [] (None) the appended <w:p> gets the "CommentText" paragraph
            % style set DIRECTLY on the element (paragraph.style would RAISE when the
            % style is absent from the styles part). MATLAB reads the just-appended
            % element back as element_().p_lst(end); Gate-2 CONFIRMED this is
            % handle-identical to paragraph._p under the append invariant (successors
            % = () -> the new <w:p> is always the last child). Probed here at the CT
            % layer via the CommentsPart's comment element.
            d = mat2doc.Document();
            p = d.add_paragraph();
            r = p.add_run("anchor");
            d.add_comment(r, "Line one" + newline + "Line two", "Cara", "C");   % "\n"-split
            cp = d.part.part_related_by(mat2doc.opc.RELATIONSHIP_TYPE.COMMENTS);
            ce = cp.element().comment_lst(1);
            testCase.verifyEqual(numel(ce.p_lst), 2, 'two <w:p> (seed + one appended)');
            new_p = ce.p_lst(end);                   % the just-appended paragraph's element
            testCase.verifyEqual(string(new_p.style), "CommentText", ...
                'VERIFY-COMMENTTEXT: p_lst(end).style == "CommentText" (element-level set)');
            testCase.verifyTrue(contains(ser(new_p), '<w:pStyle w:val="CommentText"/>'), ...
                'VERIFY-COMMENTTEXT: the appended <w:p> serializes with the CommentText pStyle');
        end

        % =============================================================== %
        % 13. CommentsPart.default / _default_comments_xml on-demand         %
        % =============================================================== %

        function test_commentspart_default_empty(testCase)
            % Nominal (parts/comments.py CommentsPart.default + _default_comments_xml):
            % default(package) builds a CommentsPart whose <w:comments> root is empty
            % (0 comments) -- the on-demand template. comments proxy is minted fresh.
            d = mat2doc.Document();
            pkg = d.part.package;
            cpd = mat2doc.parts.CommentsPart.default(pkg);
            testCase.verifyClass(cpd, 'mat2doc.parts.CommentsPart', 'default -> a CommentsPart');
            testCase.verifyEqual(cpd.comments.len_(), 0, ...
                '_default_comments_xml is an EMPTY <w:comments> (0 comments)');
        end

        % =============================================================== %
        % 14. EQUIVALENCE -- s0107 API-value probe vs the frozen oracle       %
        % =============================================================== %

        function test_s0107_api_probe_vs_frozen_oracle(testCase)
            % Equivalence (values, not bytes): replay the ENTIRE s0107 API probe
            % (runS0107 -- the .m twin's body VERBATIM: len_ 0->1->2->3, id
            % auto-increment, author/initials(present/empty/None)/text/paragraphs,
            % to_array authors, get hit/miss, the tz-aware UTC date write+read
            % round-trip, CommentsPart.default) and compare EACH tagged fact to the
            % frozen python-docx 1.2.0 oracle in data\s0107\probe.json. Gate-3
            % probe_diff was exit 0, so every fact must be value-identical. Looser-
            % than-byte (tagged strings); justified by the Gate-3 probe_diff proof.
            here   = fileparts(mfilename('fullpath'));
            port   = runS0107();
            oracle = loadJson(fullfile(here, 'data', 's0107', 'probe.json'));

            keys = fieldnames(oracle);
            % Non-triviality floor guarding a silent-empty replay (the frozen
            % probe.json exposes 22 tagged facts; Gate-3 probe_diff was exit 0).
            testCase.verifyGreaterThanOrEqual(numel(keys), 22, ...
                'the s0107 oracle must expose all 22 tagged facts');
            testCase.verifyEqual(sort(fieldnames(port)), sort(keys), ...
                'the replayed probe and the frozen oracle expose the same keys');
            for i = 1:numel(keys)
                k = keys{i};
                testCase.verifyEqual(string(port.(k)), string(oracle.(k)), ...
                    sprintf('s0107 fact "%s" must be value-identical to the frozen oracle', k));
            end
        end

    end

    % ===================== instance helpers ============================ %
    methods (Access = private)

        function assertPackage(testCase, zipBytes, scn, label)
            % Full-package byte pin from a frozen manifest.json: read
            % data\<scn>\manifest.json (parts: name/size/sha256, in frozen zip-entry
            % order), enumerate the saved zip in stream (write) order, and assert (a)
            % the part inventory + EXACT frozen zip-entry order == the manifest
            % parts sequence, and (b) every part's size + SHA-256 == the manifest
            % (SHA-256 equality == byte-identity, L1).
            here = fileparts(mfilename('fullpath'));
            man  = loadJson(fullfile(here, 'data', scn, 'manifest.json'));
            parts = man.parts;
            wantNames = strings(1, numel(parts));
            for k = 1:numel(parts)
                wantNames(k) = string(parts(k).name);
            end
            [blobs, names] = zipEntryList(zipBytes);
            testCase.verifyEqual(numel(names), numel(parts), ...
                sprintf('%s: must emit exactly %d parts', label, numel(parts)));
            testCase.verifyEqual(names, wantNames, ...
                sprintf('%s: zip-entry sequence must equal the frozen manifest order', label));
            for k = 1:numel(parts)
                nm       = string(parts(k).name);
                wantSize = double(parts(k).size);
                wantSha  = string(parts(k).sha256);
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
% NB: NO `import mat2doc...` at class or function scope -- a specific (non-.*)
% import is resolved at suite-CREATION PARSE time, BEFORE TestClassSetup's
% PathFixture puts +mat2doc on the path (the "specific import fails in test class"
% lesson). Everything below is fully qualified, which resolves at RUN time.

% ---- scenario build-body replays (the .m twins' bodies, VERBATIM) ----

function d = runS0101()
    % s0101 headline one comment -- the .m twin body VERBATIM (date pinned UTC).
    d = mat2doc.Document();
    p = d.add_paragraph();
    r = p.add_run("anchored text");
    d.add_comment(r, "Reviewer note", "Ann", "A");
    cp = d.part.part_related_by(mat2doc.opc.RELATIONSHIP_TYPE.COMMENTS);
    cp.element().comment_lst(1).date = datetime(2021,6,15,9,30,0,"TimeZone","UTC");
end

function d = runS0102()
    % s0102 empty text / initials None (H3 absent) -- the .m twin body VERBATIM.
    d = mat2doc.Document();
    p = d.add_paragraph();
    r = p.add_run("anchored text");
    d.add_comment(r, "", "Bob", []);   % Python: add_comment(r, "", "Bob", None)
    cp = d.part.part_related_by(mat2doc.opc.RELATIONSHIP_TYPE.COMMENTS);
    cp.element().comment_lst(1).date = datetime(2021,6,15,9,30,0,"TimeZone","UTC");
end

function d = runS0103()
    % s0103 multi-paragraph CommentText path -- the .m twin body VERBATIM.
    d = mat2doc.Document();
    p = d.add_paragraph();
    r = p.add_run("anchored text");
    d.add_comment(r, "Line one" + newline + "Line two" + newline + "Line three", "Cara", "C");
    cp = d.part.part_related_by(mat2doc.opc.RELATIONSHIP_TYPE.COMMENTS);
    cp.element().comment_lst(1).date = datetime(2021,6,15,9,30,0,"TimeZone","UTC");
end

function d = runS0104()
    % s0104 non-ASCII author/initials/text (H2 UTF-8) -- the .m twin body VERBATIM.
    d = mat2doc.Document();
    p = d.add_paragraph();
    r = p.add_run("anchored text");
    d.add_comment(r, "Héllo café" + newline + "第二行", "José", "Jé");
    cp = d.part.part_related_by(mat2doc.opc.RELATIONSHIP_TYPE.COMMENTS);
    cp.element().comment_lst(1).date = datetime(2021,6,15,9,30,0,"TimeZone","UTC");
end

function d = runS0105()
    % s0105 two comments (id auto-increment) -- the .m twin body VERBATIM.
    d = mat2doc.Document();
    p1 = d.add_paragraph(); r1 = p1.add_run("first anchor");
    p2 = d.add_paragraph(); r2 = p2.add_run("second anchor");
    d.add_comment(r1, "First comment", "Amy", "A");
    d.add_comment(r2, "Second comment", "Ben", "B");
    cp = d.part.part_related_by(mat2doc.opc.RELATIONSHIP_TYPE.COMMENTS);
    lst = cp.element().comment_lst;
    for k = 1:numel(lst)
        lst(k).date = datetime(2021,6,15,9,30,0,"TimeZone","UTC");
    end
end

% ---- s0107 API-value probe replay (the .m twin body, VERBATIM) ----

function P = runS0107()
    % Replay the s0107 Comments/Comment/CommentsPart API-value probe and return a
    % struct of tagged canonical strings (int|/str|/bool|/null). Embedded here so
    % the Equivalence leg is self-contained. Mirrors
    % validation\mat2doc\scenarios\s0107_p8_2_api_probe.m VERBATIM. Fully-qualified
    % inline (NO specific import -- parse-time hazard; see memory).
    P = struct();
    d = mat2doc.Document();
    cs = d.comments;                          % materializes the comments part on demand
    P.len0 = ci(cs.len_());                   % 0

    c0 = cs.add_comment("Hello", "Amy", "AJ");
    P.len1 = ci(cs.len_());                   % 1
    P.c0_id = ci(c0.comment_id);              % 0
    P.c0_author = cs_(c0.author);             % "Amy"
    P.c0_initials = cs_(c0.initials);         % "AJ"
    P.c0_text = cs_(c0.text);                 % "Hello"
    P.c0_npara = ci(numel(c0.paragraphs));    % 1

    c1 = cs.add_comment("A" + newline + "B", "Bob", []);   % Python: "A\nB", initials None
    P.len2 = ci(cs.len_());                   % 2
    P.c1_id = ci(c1.comment_id);              % 1 (auto-increment)
    P.c1_initials = cs_(c1.initials);         % [] -> null
    P.c1_text = cs_(c1.text);                 % "A\nB"
    P.c1_npara = ci(numel(c1.paragraphs));    % 2

    c2 = cs.add_comment("x", "Cara", "");     % initials present-empty
    P.len3 = ci(cs.len_());                   % 3
    P.c2_initials = cs_(c2.initials);         % "" -> str|

    % -- __iter__ (to_array) length + authors --
    arr = cs.to_array();
    auth = strings(1, numel(arr));
    for k = 1:numel(arr), auth(k) = arr(k).author; end
    P.iter_len = ci(numel(arr));              % 3
    P.iter_authors = "str|" + join(auth, "|");  % Amy|Bob|Cara

    % -- get hit + miss --
    P.get0_author = cs_(cs.get(0).author);    % "Amy"
    P.get99_isnone = cb(isequal(cs.get(99), []));  % true

    % -- date write+read round-trip (pin tz-aware UTC on the element) --
    cp = d.part.part_related_by(mat2doc.opc.RELATIONSHIP_TYPE.COMMENTS);
    ce0 = cp.element().comment_lst(1);
    ce0.date = datetime(2021,6,15,9,30,0,"TimeZone","UTC");
    P.c0_date_attr = cs_(ce0.get(mat2doc.oxml.qn("w:date")));   % raw attr
    ts = cs.get(0).timestamp; ts.TimeZone = "UTC";
    P.c0_timestamp = "str|" + string(ts, "yyyy-MM-dd'T'HH:mm:ss'Z'");  % getter round-trip

    % -- CommentsPart.default(package): class + empty comments --
    pkg = d.part.package;
    cpd = mat2doc.parts.CommentsPart.default(pkg);
    cparts = split(string(class(cpd)), ".");
    P.default_class = "str|" + cparts(end);   % CommentsPart
    P.default_len = ci(cpd.comments.len_());  % 0
end

function s = ci(v)
    s = "int|" + mat2doc.shared.pyStr(double(v), "int");
end
function s = cs_(v)
    if isequal(v, []), s = "null"; else, s = "str|" + string(v); end
end
function s = cb(v)
    if v, s = "bool|true"; else, s = "bool|false"; end
end

% ---- serialization / io / hashing helpers ----

function s = ser(e)
    % Decode serialize_part_xml bytes to a string (ASCII content -> string-equality
    % is a byte-identical L1 assertion; non-ASCII round-trips via UTF-8).
    s = string(native2unicode(mat2doc.oxml.serialize_part_xml(e), "UTF-8"));
end

function zipBytes = saveZip(d)
    % d.save to a BINARY temp .docx (the writer opens "wb"), read the whole-zip
    % bytes back, delete. Returns the uint8 zip vector.
    tmp = [tempname '.docx'];
    cleanup = onCleanup(@() deleteIfExists(tmp)); %#ok<NASGU>
    d.save(tmp);
    zipBytes = readBytes(tmp);
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
    % is independent of the reader under test. (Copied from Test_p7_4_add_picture.m.)
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
    % Lowercase hex SHA-256 of a uint8 vector (base MATLAB / Java MessageDigest;
    % matches the python hashlib manifest SHAs).
    md  = java.security.MessageDigest.getInstance('SHA-256');
    dig = md.digest(typecast(uint8(bytes(:)'), 'int8'));
    b   = typecast(dig, 'uint8');
    h   = lower(string(sprintf('%02x', b)));
end

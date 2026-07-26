classdef Test_p1_2_oxml < matlab.unittest.TestCase
% TEST_P1_2_OXML  Gate-4 permanent unit tests for Mat2Doc P1-2 (oxml core).
%
%   The FIRST Mat2Doc tests\oxml\ class. Pins the byte-fidelity foundation of
%   M1: the observable output of the oxml tree/parser/serializer core is
%   serialized XML bytes, produced by
%       mat2doc.oxml.serialize_part_xml(mat2doc.oxml.parse_xml(bytes)).
%
%   Surface under test (ported from python-docx v1.2.0 src/docx/oxml/parser.py
%   + src/docx/opc/oxml.py, realised in +mat2doc\+oxml\): parse_xml,
%   serialize_part_xml, XmlParser, XmlElement, and the nsmap/qn/OxmlElement
%   supporting layer they exercise.
%
%   This class ENCODES THE FROZEN GATE-3 27-CASE BATTERY verbatim -- the
%   Validator's oracle (lxml 5.3.0 / libxml2 2.13.9 under python-docx 1.2.0's
%   exact parser+serializer config) copied byte-for-byte into
%   tests\oxml\data\p1_2\ so the suite is self-contained. Provenance:
%   validation\mat2doc\references\p1_2\ (manifest mso-p1_2-oracle/1, frozen
%   2026-07-25); Gate-3 report validation\mat2doc\reports\p1_2_validation.md.
%
%   Coverage taxonomy:
%     * Equivalence / Regression -- every case compares the port's serialized
%       bytes against the frozen lxml .oracle.xml. Because the oracle IS
%       lxml's own bytes, byte-identical here == byte-identical to
%       python-docx: these are hard-coded expected-XML regression pins AND the
%       cross-implementation equivalence check in one.
%     * Nominal        -- the RT battery: the four real default.docx parts
%                         (document/styles/settings/numbering) parse+serialize
%                         byte-identical (M1 pin).
%     * Edge           -- nsdecl adversarials (redundant/unused/shadowed/
%                         default-ns/xml:space/char-refs/CDATA decls), escaping,
%                         non-ASCII (e / CJK / emoji), empty-element tri-state,
%                         ws-only-tail drop, and the error/reject paths.
%     * Upstream shape -- the a05 default-ns Types case mirrors the
%                         [Content_Types].xml surface; a06 mirrors Word's
%                         <w:t xml:space="preserve"> emission.
%
%   *** a03 known-deviation pin (L2 D-nsprefix-rewrite) ***
%   a03 (same URI bound to a second prefix) is the sole L2 case: lxml keeps the
%   as-written second prefix (<q:r q:val="1"/>); the port re-renders through the
%   first same-URI binding (<w:r w:val="1"/>). test_a03_known_deviation_bytes
%   pins that the port's bytes DIFFER from lxml (verifyNotEqual) AND
%   test_a03_canonical_equivalence pins that the two are canonically equal
%   (identical expanded-name trees). Any future change to the a03 behaviour --
%   in EITHER direction -- surfaces as a deliberate red. Mirrors the Mat2Ppt
%   D-nsprefix-rewrite fixture pattern. D-nsprefix-rewrite was SIGNED as a
%   permanent L2 class 2026-07-18; NO new D-number.
%
%   Adopted-D reject cases exercised (decision_2026-07-25_mat2doc_deviation_
%   preadoption.md): D-006 (DOCTYPE rejected -> mat2doc:XMLSyntaxError; lxml
%   accepts/ignores) and D-002 (ASCII char-ref digit grammar -> reject; lxml
%   also rejects). Asserted by error IDENTIFIER (mat2doc:XMLSyntaxError), which
%   is the faithful class for these paths.
%
%   Explicitly EXCLUDED (per Gate-3 report, recorded WHY): the V-BLANK
%   accepted-unreachable shapes (char-ref/CDATA-derived inter-element
%   whitespace) -- unreachable in real OOXML, so not pinned.
%
%   Determinism: no network, no absolute paths -- every fixture is resolved
%   relative to this file via fileparts(mfilename('fullpath')).

    properties (TestParameter)
        % --- RT battery: real default.docx parts, parse+serialize == lxml ----
        %     base name; on-disk files are <base>.xml.{in,oracle}.xml.
        rtCase = { ...
            'word_document', 'word_styles', 'word_settings', 'word_numbering'};

        % --- nsdecl adversarials that are L1 byte-identical (the H8 crux) ----
        nsL1Case = { ...
            'a01_nested_redundant', 'a02_unused_decl', 'a04_inner_rebind_shadow', ...
            'a05_default_ns_types', 'a06_xml_space_preserve', 'a07_char_refs', ...
            'a08_cdata', 'a11_three_deep_redundant', 'a13_ancestor_declared_rid'};

        % --- FIX-1 regression pins: ws-only trailing tail dropped / kept -----
        fixCase = { ...
            'a09d_leadws_trailws_tail', 'a09e_trailws_tail', ...
            'a09ctl_nonblank_text_keeps_tail'};

        % --- escaping / UTF-8 / empty-element tri-state ----------------------
        escCase = { ...
            'esc_text_amp_lt_gt_cr', 'esc_attr_quote_cr_lf_tab', ...
            'utf8_raw_accent_cjk_emoji', 'empty_selfclose_text_none'};

        % --- reject-as-designed (D-006 / D-002) ------------------------------
        rejectCase = {'d006_doctype', 'd002_hex_zz', 'd002_dec_1a'};
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into
            % the test folder, so without the worktree root on the path a COLD
            % run cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Copy of the proven idiom from tests\shared\Test_p1_1_shared.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\oxml
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));
        end
    end

    methods (Test)

        % ============================================ RT battery (M1 pin) =====
        function test_roundtrip_byte_identical(testCase, rtCase)
            % Nominal + M1: parse+serialize the raw default.docx part blob and
            % assert byte-identical to lxml's own reserialization (== the
            % python-docx open->save form). L1; byte-level assertion (the ladder
            % demanded L1 for the whole RT battery). Fixtures rt\<base>.xml.*.
            [inb, orb] = testCase.loadCase('rt', rtCase + ".xml");
            out = roundtrip(inb);
            verifyByteIdentical(testCase, out, orb, rtCase);
        end

        % ==================================== nsdecl adversarials (L1) ========
        function test_nsdecl_L1(testCase, nsL1Case)
            % Edge (H8 verbatim-until-moved): parse the frozen .in.xml and
            % assert the serialized bytes are byte-identical to lxml. Redundant/
            % unused/shadowed/default-ns/xml:space/char-ref/CDATA declarations
            % all reproduce lxml exactly. L1 byte-level.
            [inb, orb] = testCase.loadCase('nsdecl', nsL1Case);
            out = roundtrip(inb);
            verifyByteIdentical(testCase, out, orb, nsL1Case);
        end

        % ==================================== FIX-1 ws-tail-drop pins =========
        function test_fix1_wstail(testCase, fixCase)
            % Regression (Gate-2 FIX-1): a09d/a09e drop the ws-only trailing
            % tail lxml drops; a09ctl (non-blank TEXT) KEEPS its ws tail -- the
            % control proving FIX-1 did not over-drop. L1 byte-level.
            [inb, orb] = testCase.loadCase('nsdecl', fixCase);
            out = roundtrip(inb);
            verifyByteIdentical(testCase, out, orb, fixCase);
        end

        % ==================================== escaping / UTF-8 / tri-state ====
        function test_escaping_utf8_tristate(testCase, escCase)
            % Edge: text escaping (& < > CR), attr escaping (" CR LF TAB), raw
            % UTF-8 non-ASCII (e / CJK / emoji astral -> 4-byte), and the
            % text==[] (None) self-close case. L1 byte-level.
            [inb, orb] = testCase.loadCase('esc', escCase);
            out = roundtrip(inb);
            verifyByteIdentical(testCase, out, orb, escCase);
        end

        % ==================================== reject-as-designed (D-006/002) ==
        function test_reject_raises(testCase, rejectCase)
            % Edge / error path: D-006 (DOCTYPE) and D-002 (invalid char-ref
            % digit grammar) inputs must RAISE. The identifier is asserted --
            % mat2doc:XMLSyntaxError is the faithful class for these reject
            % paths (Gate-3 report). Provenance reject\reject_cases.json.
            inb = testCase.loadInput('reject', rejectCase);
            testCase.verifyError(@() roundtrip(inb), 'mat2doc:XMLSyntaxError', ...
                sprintf('%s must reject with mat2doc:XMLSyntaxError', rejectCase));
        end

        % ==================================== a03 known-deviation pin =========
        function test_a03_known_deviation_bytes(testCase)
            % *** L2 D-nsprefix-rewrite (SIGNED 2026-07-18) ***
            % The port's serialized bytes DIFFER from lxml's: lxml keeps the
            % as-written second prefix (<q:r q:val="1"/>), the port re-renders
            % through the first same-URI binding (<w:r w:val="1"/>). Pinned as a
            % verifyNotEqual so this deviation staying-put is asserted; if the
            % port ever matches lxml byte-for-byte (or drifts otherwise) this
            % goes red and forces a deliberate review.
            [inb, orb] = testCase.loadCase('nsdecl', 'a03_same_uri_second_prefix');
            out = roundtrip(inb);
            testCase.verifyNotEqual(out, orb, ...
                'a03 is the L2 D-nsprefix-rewrite case: port bytes must DIFFER from lxml');
            % Belt-and-braces: the divergence is exactly the prefix q->w on the
            % child (first same-URI binding). Confirm the port emits w:r, lxml q:r.
            testCase.verifyTrue(contains(bytes2str(out), "<w:r w:val=""1""/>"), ...
                'port must render the child through the first (w) binding');
            testCase.verifyTrue(contains(bytes2str(orb), "<q:r q:val=""1""/>"), ...
                'oracle (lxml) keeps the as-written (q) second prefix');
        end

        function test_a03_canonical_equivalence(testCase)
            % *** L2 D-nsprefix-rewrite: canonical equality proof ***
            % Although the bytes differ, both bind the child element and its
            % attribute to the SAME URI -- canonically equal, Office-safe,
            % non-corrupting (this is WHY the deviation is a PASS). Compare the
            % expanded-name (Clark) trees of the port output and the lxml oracle.
            [inb, orb] = testCase.loadCase('nsdecl', 'a03_same_uri_second_prefix');
            out = roundtrip(inb);
            portCanon = canonForm(mat2doc.oxml.parse_xml(strip_decl(out)));
            lxmlCanon = canonForm(mat2doc.oxml.parse_xml(strip_decl(orb)));
            testCase.verifyEqual(portCanon, lxmlCanon, ...
                'a03: port and lxml must be canonically (expanded-name) equal');
        end

        % ==================================== generation cases (b04/b06/gen) ==
        function test_gen_b04_append_suppression(testCase)
            % L1 generation: parse a parent + a fragment, append the fragment;
            % the fragment's redundant xmlns:w is suppressed at move time
            % (moveNodeToDocument analogue). Mirrors the oracle API op sequence
            % (harness\mat2doc\validate_p1_2.m). gen\b04_*.oracle.xml.
            Uw = wURI();
            par  = mat2doc.oxml.parse_xml("<w:body xmlns:w=""" + Uw + """/>");
            frag = mat2doc.oxml.parse_xml("<w:p xmlns:w=""" + Uw + """><w:r/></w:p>");
            par.append(frag);
            out = mat2doc.oxml.serialize_part_xml(par);
            orb = testCase.loadOracle('gen', 'b04_append_suppression');
            verifyByteIdentical(testCase, out, orb, 'b04_append_suppression');
        end

        function test_gen_b06_move_subtree_reconcile(testCase)
            % L1 generation: move a subtree carrying a nested redundant xmlns:w
            % under a new root; both redundant decls suppressed at move time.
            Uw = wURI();
            newroot = mat2doc.oxml.parse_xml("<w:document xmlns:w=""" + Uw + """/>");
            sub = mat2doc.oxml.parse_xml("<w:p xmlns:w=""" + Uw + """><w:r xmlns:w=""" ...
                + Uw + """><w:t/></w:r></w:p>");
            newroot.append(sub);
            out = mat2doc.oxml.serialize_part_xml(newroot);
            orb = testCase.loadOracle('gen', 'b06_move_subtree_reconcile');
            verifyByteIdentical(testCase, out, orb, 'b06_move_subtree_reconcile');
        end

        function test_gen_empty_open_text_emptystring(testCase)
            % L1 generation (H3 tri-state): text set to "" (empty string, not
            % None) serializes as <w:t></w:t>, NOT the self-closing <w:t/>.
            Uw = wURI();
            t = mat2doc.oxml.parse_xml("<w:t xmlns:w=""" + Uw + """/>");
            t.text = "";
            out = mat2doc.oxml.serialize_part_xml(t);
            orb = testCase.loadOracle('gen', 'empty_open_text_emptystring');
            verifyByteIdentical(testCase, out, orb, 'empty_open_text_emptystring');
        end

    end

    % ============================= instance fixture loaders ================
    methods (Access = private)
        function [inb, orb] = loadCase(testCase, sub, base)
            inb = testCase.loadInput(sub, base);
            orb = testCase.loadOracle(sub, base);
        end

        function b = loadInput(testCase, sub, base)
            b = readFixture(dataPath(sub, base + ".in.xml"));
            testCase.assertNotEmpty(b, ...
                sprintf('input fixture missing/empty: %s/%s.in.xml', sub, base));
        end

        function b = loadOracle(testCase, sub, base)
            b = readFixture(dataPath(sub, base + ".oracle.xml"));
            testCase.assertNotEmpty(b, ...
                sprintf('oracle fixture missing/empty: %s/%s.oracle.xml', sub, base));
        end
    end
end

% ============================= file-local helpers ==========================

function p = dataPath(sub, name)
    here = fileparts(mfilename('fullpath'));   % tests\oxml
    p = fullfile(here, 'data', 'p1_2', char(sub), char(name));
end

function b = readFixture(p)
    f = fopen(p, 'r');
    if f < 0; b = uint8([]); return; end
    b = fread(f, Inf, '*uint8')';
    fclose(f);
end

function out = roundtrip(inbytes)
    % parse_xml wants a uint8 ROW vector of on-disk bytes.
    out = mat2doc.oxml.serialize_part_xml(mat2doc.oxml.parse_xml(uint8(inbytes(:)')));
    out = uint8(out(:)');
end

function verifyByteIdentical(testCase, got, want, label)
    % Byte-level assertion (L1). On mismatch, report sizes and the first
    % differing offset so a regression is diagnosable.
    got = uint8(got(:)');
    want = uint8(want(:)');
    if ~isequal(got, want)
        n = min(numel(got), numel(want));
        d = find(got(1:n) ~= want(1:n), 1);
        if isempty(d); d = n + 1; end
        diag = sprintf(['%s: bytes differ (got %d B, want %d B); first diff @%d ' ...
            '(got 0x%02X, want 0x%02X)'], char(label), numel(got), numel(want), ...
            d, byteAt(got, d), byteAt(want, d));
    else
        diag = char(label);
    end
    testCase.verifyEqual(got, want, diag);
end

function v = byteAt(b, i)
    if i >= 1 && i <= numel(b); v = double(b(i)); else; v = 0; end
end

function s = bytes2str(b)
    s = string(native2unicode(uint8(b(:)'), "UTF-8"));
end

function s = strip_decl(b)
    % Drop the XML declaration line so parse_xml sees a bare root element.
    s = bytes2str(b);
    s = regexprep(s, "^<\?xml[^>]*\?>\s*", "");
end

function c = canonForm(elm)
    % Canonical (expanded-name) serialization of a parsed tree: tag as its
    % Clark {uri}local name, attributes sorted by Clark name, text, then the
    % children canonicalized in order. Prefix-independent -- two trees that bind
    % the same elements/attributes to the same URIs canonicalize identically.
    tag = elm.tag;
    names = sort(elm.attrib_names());
    attrs = "";
    for k = 1:numel(names)
        attrs = attrs + "|" + names(k) + "=" + string(elm.get(names(k)));
    end
    txt = elm.text;
    if isequal(txt, []); txt = ""; end
    kids = elm.to_array();
    inner = "";
    for k = 1:numel(kids)
        inner = inner + canonForm(kids(k));
    end
    c = "(" + tag + attrs + "#" + string(txt) + inner + ")";
end

function u = wURI()
    u = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
end

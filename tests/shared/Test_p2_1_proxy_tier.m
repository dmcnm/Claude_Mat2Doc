classdef Test_p2_1_proxy_tier < matlab.unittest.TestCase
% TEST_P2_1_PROXY_TIER  Gate-4 permanent unit tests for Mat2Doc P2-1
%   (the shared proxy base tier + exceptions + types + the Document
%   ElementProxy retrofit).
%
%   Freezes the P2-1 guarantees every later API proxy (P3-P6: Paragraph, Run,
%   Table, Section, ...) depends on:
%     - mat2doc.shared.ElementProxy  -- H5 ELEMENT-IDENTITY eq/ne, the `element`
%       accessor, and the `.part` None-guard + parent delegation (shared.py
%       277-316).
%     - mat2doc.shared.Parented / mat2doc.shared.StoryChild -- parent-only
%       delegators that hold NO element and (deliberately) DO NOT override
%       eq/ne (shared.py 319-353).
%     - mat2doc.shared.lazyproperty -- the non-callable idiom guide + the
%       first-access-caches / read-only semantics of the realised idiom.
%     - mat2doc.exc.PythonDocxError / InvalidSpanError / InvalidXmlError -- the
%       raisers (mat2doc:<Name> flat identifiers, verbatim messages).
%     - mat2doc.document.Document retrofitted onto ElementProxy
%       (VERIFY-M1-DOC-BASE) -- isa + one-part byte spot-check (the full 17/17
%       M1 sweep is owned by tests\api\Test_p1_8_skeleton_m1, not duplicated).
%
%   Provenance (Gate-1..3, all 2026-07-26):
%     * Audit    : validation\mat2doc\audit_P2-1_proxy_tier.md  (Porter Gate-1 +
%                  Opus Gate-2 adversarial APPROVE; VERIFY-1 RATIFIED,
%                  VERIFY-2/3/4 ratified; three Low doc-comment fixes applied).
%     * Validate : validation\mat2doc\validate_P2-1_proxy_tier.md  (Gate-3 PASS,
%                  6 bars MATCH, regression 361/361, 0 new D-numbers).
%     * Scenario : validation\mat2doc\scenarios\s0013_proxy_tier_probes.{py,m}
%                  (+ helpers s0013_PartStub.m / s0013_LazyHost.m; the frozen
%                  python-docx oracle references\s0013\probe.json).
%     * Frozen refs (python-docx 1.2.0 oracle): the word/document.xml byte spot
%       (Bar 4 below) reuses the frozen s0001 manifest value already pinned by
%       Test_p1_8_skeleton_m1.M1_MANIFEST (SHA-256 equality == byte-identity L1).
%
%   Coverage taxonomy
%   -----------------
%   * H5 crux (the identity trap the whole proxy layer rests on) -- same element
%     handle -> EQUAL; distinct BYTE-IDENTICAL elements -> NOT equal (identity,
%     not content); self -> equal; vs non-proxy -> false; vs [] (None) -> false;
%     both operand orders. (validate_P2-1 Bar 1.)
%   * VERIFY-1 Document == (ratified) -- dp.document == dp.document -> TRUE
%     (fresh proxies, same element); two independent mat2doc.Document() -> not
%     equal; and the stronger NON-CACHE pin (equal-but-DISTINCT handles: two
%     dp.document proxies are ==-equal yet one can be deleted while the other
%     stays valid). (validate_P2-1 Bar 2.)
%   * .part None-guard + delegation -- ElementProxy with parent=[] -> raises
%     mat2doc:ValueError with the verbatim message; with a stub parent it
%     delegates to parent.part; Parented / StoryChild delegate too (and are NOT
%     ElementProxy). (validate_P2-1 Bar 3.)
%   * VERIFY-M1-DOC-BASE byte-neutrality (M1 guard reinforcement) -- Document
%     isa mat2doc.shared.ElementProxy, and a save spot-check pins word/document.xml
%     byte-identical to the frozen s0001 oracle. Full 17/17 sweep NOT duplicated
%     (owned by Test_p1_8_skeleton_m1). (validate_P2-1 Bar 4.)
%   * lazyproperty -- first access computes once (count 1), second returns cache
%     (count still 1), assignment raises; and the guide itself is non-callable.
%     (validate_P2-1 Bar 5.)
%   * Exceptions -- the three raisers emit mat2doc:<Name> with verbatim messages
%     including a %-literal ("span 100% invalid"); the PythonDocxError-family
%     hierarchy INTENT is documented (flat identifiers under the shared mat2doc:
%     prefix -- VERIFY-2). (validate_P2-1 Bar 6.)
%
%   Deviations exercised: 0 new D-numbers (Gate-3). This tier emits no serialized
%   output of its own; the Bar-4 spot-check rides the standing adopt-only set
%   (D-001 own OOXML parser/attr-order, byte-proven by the frozen s0001 value).
%   VERIFY-2 (flat identifiers, no identifier-level is-a) is a ratified design
%   choice, inert at runtime (no src catch site), and needs no D.
%
%   Determinism: no network, no absolute paths. The default template resolves
%   package-relative via the MANDATORY PathFixture(worktree-root); every file
%   write is BINARY ("w"/'n', never 'wt') and every read is binary ('r'/'n').

    properties (Constant)
        % word/document.xml frozen byte spot (== Test_p1_8_skeleton_m1.M1_MANIFEST
        % row for word/document.xml, the python-docx 1.2.0 s0001 oracle). SHA-256
        % equality IS byte-identity (L1). One-part reinforcement of the M1 guard
        % after the ElementProxy retrofit; the FULL 17/17 sweep is owned by
        % Test_p1_8_skeleton_m1 (this class deliberately does not duplicate it).
        SPOT_PART     = "word/document.xml"
        SPOT_SIZE     = 1548
        SPOT_SHA      = "0e4dd503bc095219cd24e8312b372c588ca56bb834aacac53eafb7bf47836327"
    end

    properties (Access = private)
        % One default-template document + its saved zip, built ONCE in setup
        % (styles 349 KB + stylesWithEffects 438 KB make re-serialization
        % non-trivial). Shared read-only across the VERIFY-1 and Bar-4 cases.
        doc_
        candNames_
        candBlobs_
    end

    methods (TestClassSetup)
        function addWorktreeToPath(testCase)
            % MANDATORY PathFixture (WP9-F4 lesson): R2024b runtests cd's into the
            % test folder, so without the worktree root on the path a COLD run
            % cannot resolve the +mat2doc package (MATLAB:undefinedVarOrClass).
            % Idiom copied from the proven tests\shared\Test_p1_1_shared.m.
            import matlab.unittest.fixtures.PathFixture
            here = fileparts(mfilename('fullpath'));   % tests\shared
            root = fileparts(fileparts(here));         % worktree root
            testCase.applyFixture(PathFixture(root));

            % Build the shared default-template document + its saved zip ONCE,
            % AFTER the PathFixture is active (so +mat2doc resolves).
            testCase.doc_ = mat2doc.Document();        % no arg -> default template
            zipBytes = testCase.saveToTemp(testCase.doc_);
            [b, n] = zipEntryList(zipBytes);
            testCase.candBlobs_ = b;
            testCase.candNames_ = n;
        end
    end

    methods (Test)

        % =============================================================== %
        % 1. H5 ELEMENT-IDENTITY eq/ne -- THE CRUX PIN                     %
        %    (every P3-P6 proxy's equality contract bottoms out here)      %
        % =============================================================== %

        function test_h5_same_element_handle_is_equal(testCase)
            % H5: two DIFFERENT proxy instances wrapping the SAME oxml element
            % compare EQUAL (identity of the wrapped element, not of the proxy).
            % (validate_P2-1 Bar 1: same-handle == -> true / != -> false.)
            e = mat2doc.oxml.XmlElement("w:p");
            a     = mat2doc.shared.ElementProxy(e);
            aSame = mat2doc.shared.ElementProxy(e);     % distinct proxy, SAME element
            testCase.verifyTrue(a == aSame, ...
                'proxies over the SAME element must be ElementProxy-equal (H5)');
            testCase.verifyFalse(a ~= aSame, ...
                '~= must be the complement for same-element proxies');
        end

        function test_h5_distinct_byte_identical_elements_not_equal(testCase)
            % H5 -- THE TRAP: two SEPARATELY-CONSTRUCTED but byte-identical
            % <w:p/> elements are DIFFERENT handles, so proxies over them are
            % NOT equal. Equality is element IDENTITY, never content. A proxy
            % that ever compared these equal would silently alias distinct nodes.
            % (validate_P2-1 Bar 1: distinct-byte-identical == -> false / != -> true.)
            e1 = mat2doc.oxml.XmlElement("w:p");
            e2 = mat2doc.oxml.XmlElement("w:p");        % byte-identical, DISTINCT handle
            a = mat2doc.shared.ElementProxy(e1);
            b = mat2doc.shared.ElementProxy(e2);
            testCase.verifyFalse(a == b, ...
                'byte-identical BUT distinct elements must NOT be equal (identity, not content)');
            testCase.verifyTrue(a ~= b, ...
                '~= must be true for distinct byte-identical elements');
        end

        function test_h5_self_is_equal(testCase)
            % H5: a proxy equals itself. (validate_P2-1 Bar 1: self == -> true.)
            e = mat2doc.oxml.XmlElement("w:p");
            a = mat2doc.shared.ElementProxy(e);
            testCase.verifyTrue(a == a, 'a proxy must equal itself');
            testCase.verifyFalse(a ~= a, 'a proxy must not be unequal to itself');
        end

        function test_h5_vs_nonproxy_is_false(testCase)
            % H5: comparison against a non-proxy is FALSE for == (mirrors Python's
            % `isinstance(other, ElementProxy)` guard) and TRUE for ~=.
            % (validate_P2-1 Bar 1: vs non-proxy.)
            e = mat2doc.oxml.XmlElement("w:p");
            a = mat2doc.shared.ElementProxy(e);
            notpx = "not a proxy";
            testCase.verifyFalse(a == notpx, 'proxy == non-proxy must be false');
            testCase.verifyTrue(a ~= notpx, 'proxy ~= non-proxy must be true');
        end

        function test_h5_nonproxy_lhs_reflected(testCase)
            % H5 (both operand orders): with the NON-proxy on the LHS, MATLAB
            % still dispatches ElementProxy.eq/ne (a user classdef dominates the
            % built-in string), so the result is symmetric -- false / true.
            % (validate_P2-1 Bar 1: non-proxy-LHS reflected.)
            e = mat2doc.oxml.XmlElement("w:p");
            a = mat2doc.shared.ElementProxy(e);
            notpx = "not a proxy";
            testCase.verifyFalse(notpx == a, 'non-proxy == proxy (reflected) must be false');
            testCase.verifyTrue(notpx ~= a, 'non-proxy ~= proxy (reflected) must be true');
        end

        function test_h5_vs_empty_none_is_false(testCase)
            % H5 / H3: comparison against [] (the None sentinel) is FALSE for ==
            % and TRUE for ~=. (validate_P2-1 Bar 1: vs None/[].)
            e = mat2doc.oxml.XmlElement("w:p");
            a = mat2doc.shared.ElementProxy(e);
            testCase.verifyFalse(a == [], 'proxy == [] (None) must be false');
            testCase.verifyTrue(a ~= [], 'proxy ~= [] (None) must be true');
        end

        % =============================================================== %
        % 2. VERIFY-1 Document == (ratified)                              %
        % =============================================================== %

        function test_verify1_document_eq_document_true(testCase)
            % VERIFY-1 (ratified): dp.document builds a FRESH Document proxy per
            % access (a plain @property, docx/parts/document.py 58-59), but both
            % wrap the SAME w:document element, so they compare EQUAL -- the
            % Python-faithful result (`part.document == part.document` is True).
            % (validate_P2-1 Bar 2.)
            dp = testCase.doc_.part();
            docX = dp.document();
            docY = dp.document();
            testCase.verifyTrue(docX == docY, ...
                'two fresh dp.document proxies wrap the SAME element -> equal (VERIFY-1)');
            testCase.verifyFalse(docX ~= docY, ...
                '~= must be the complement for two dp.document proxies');
        end

        function test_verify1_two_independent_documents_not_equal(testCase)
            % VERIFY-1: two INDEPENDENT mat2doc.Document() open distinct
            % w:document elements, so they are NOT equal (element identity).
            % (validate_P2-1 Bar 2.)
            d1 = testCase.doc_;
            d2 = mat2doc.Document();                     % independent document
            testCase.verifyFalse(d1 == d2, ...
                'two independent documents wrap distinct elements -> not equal');
            testCase.verifyTrue(d1 ~= d2, ...
                'two independent documents must be ~= true');
        end

        function test_verify1_fresh_proxies_are_distinct_handles(testCase)
            % VERIFY-1 NON-CACHE PIN (Gate-2 recommended, the stronger one):
            % the two dp.document proxies are ==-equal YET DISTINCT handle
            % instances -- non-caching that `==` alone can NEVER show (it is now
            % element-identity, exactly as Python's `is`-only distinctness). Prove
            % it structurally: deleting one leaves the other valid.
            % (validate_P2-1 Bar 2: fresh proxies distinct instances.)
            dp = testCase.doc_.part();
            d3 = dp.document();
            d4 = dp.document();
            testCase.verifyTrue(d3 == d4, ...
                'the two fresh proxies must be ==-equal (same element)');
            delete(d3);
            testCase.verifyTrue(isvalid(d4), ...
                ['deleting one fresh proxy must leave the other VALID -> they are ' ...
                 'distinct handle instances (non-caching, invisible to element-identity ==)']);
        end

        % =============================================================== %
        % 3. .part None-guard + Parented/StoryChild delegation            %
        % =============================================================== %

        function test_part_none_guard_valueerror_verbatim(testCase)
            % ElementProxy with parent defaulting to [] (None): .part raises
            % mat2doc:ValueError with the byte-verbatim message. Identifier
            % local-part == the Python exception class name; message == Python
            % str(ValueError). (validate_P2-1 Bar 3.)
            e = mat2doc.oxml.XmlElement("w:p");
            epNone = mat2doc.shared.ElementProxy(e);     % parent -> [] (None)
            caught = [];
            try
                epNone.part();
            catch ME
                caught = ME;
            end
            testCase.assertNotEmpty(caught, 'parentless .part must raise (not return)');
            testCase.verifyEqual(caught.identifier, 'mat2doc:ValueError', ...
                'None-guard must raise mat2doc:ValueError (== Python ValueError)');
            testCase.verifyEqual(caught.message, ...
                'part is not accessible from this element', ...
                'None-guard message must be byte-verbatim to python-docx');
        end

        function test_part_elementproxy_delegates_to_parent(testCase)
            % ElementProxy WITH a parent delegates .part -> parent.part.
            % (validate_P2-1 Bar 3.)
            e = mat2doc.oxml.XmlElement("w:p");
            stub = PartStub_p2_1();
            ep = mat2doc.shared.ElementProxy(e, stub);
            testCase.verifyEqual(ep.part(), "STUBPART", ...
                'ElementProxy.part must delegate to parent.part');
        end

        function test_part_parented_delegates_and_is_not_elementproxy(testCase)
            % Parented holds only a parent and delegates .part; it is a SEPARATE
            % class, NOT a subclass of ElementProxy (docx folds parent into
            % ElementProxy but keeps Parented standalone -- shared.py 319-333).
            % (validate_P2-1 Bar 3.)
            stub = PartStub_p2_1();
            par = mat2doc.shared.Parented(stub);
            testCase.verifyEqual(par.part(), "STUBPART", ...
                'Parented.part must delegate to parent.part');
            testCase.verifyFalse(isa(par, 'mat2doc.shared.ElementProxy'), ...
                'Parented must NOT be an ElementProxy (standalone class)');
        end

        function test_part_storychild_delegates_and_is_not_elementproxy(testCase)
            % StoryChild is structurally identical to Parented but a DISTINCT
            % class (its parent is a ProvidesStoryPart); block-item subclasses
            % derive from it, not from Parented (shared.py 336-353).
            % (validate_P2-1 Bar 3.)
            stub = PartStub_p2_1();
            sc = mat2doc.shared.StoryChild(stub);
            testCase.verifyEqual(sc.part(), "STUBPART", ...
                'StoryChild.part must delegate to parent.part');
            testCase.verifyFalse(isa(sc, 'mat2doc.shared.ElementProxy'), ...
                'StoryChild must NOT be an ElementProxy (standalone class)');
        end

        % =============================================================== %
        % 4. VERIFY-M1-DOC-BASE byte-neutrality (M1 guard reinforcement)  %
        % =============================================================== %

        function test_document_isa_elementproxy(testCase)
            % The Document retrofit: mat2doc.document.Document now derives
            % mat2doc.shared.ElementProxy (document.py 28), gaining the H5
            % identity eq/ne + element accessor. RED if a future change reparents
            % Document off the shared base. (validate_P2-1 Bar 4.)
            testCase.verifyTrue(isa(testCase.doc_, 'mat2doc.shared.ElementProxy'), ...
                'mat2doc.document.Document must be an ElementProxy after the P2-1 retrofit');
        end

        function test_document_save_spot_check_document_xml_byte_identical(testCase)
            % Byte-neutrality spot-check (L1): after the ElementProxy retrofit,
            % mat2doc.Document().save still emits word/document.xml byte-identical
            % to the frozen s0001 python-docx oracle (size + SHA-256). The FULL
            % 17/17 M1 sweep is owned by tests\api\Test_p1_8_skeleton_m1; this is
            % the lighter one-part reinforcement the P2-1 Gate-4 brief specifies.
            % (validate_P2-1 Bar 4.)
            got = testCase.entryBlob(testCase.SPOT_PART);
            testCase.verifyEqual(numel(got), testCase.SPOT_SIZE, ...
                sprintf('%s must be exactly %d B', testCase.SPOT_PART, testCase.SPOT_SIZE));
            testCase.verifyEqual(sha256hex(got), testCase.SPOT_SHA, ...
                sprintf(['%s SHA-256 must equal the frozen s0001 oracle ' ...
                '(byte-identical L1 -- retrofit changed no bytes)'], testCase.SPOT_PART));
        end

        % =============================================================== %
        % 5. lazyproperty (idiom guide + realised semantics)              %
        % =============================================================== %

        function test_lazyproperty_computes_once_and_caches(testCase)
            % First access computes (value 42, count 1); second access returns the
            % CACHE (count still 1 -- getter body NOT re-run). (validate_P2-1 Bar 5.)
            h = LazyHost_p2_1();
            testCase.verifyEqual(h.val, 42, 'first access must compute the value');
            testCase.verifyEqual(h.compute_count, 1, 'first access must run the getter body once');
            testCase.verifyEqual(h.val, 42, 'second access must return the cached value');
            testCase.verifyEqual(h.compute_count, 1, ...
                'second access must NOT re-run the getter body (cached)');
        end

        function test_lazyproperty_set_raises(testCase)
            % Read-only: assignment to the lazy property raises (MATLAB read-only
            % Dependent == Python lazyproperty.__set__ -> AttributeError).
            % (validate_P2-1 Bar 5.)
            h = LazyHost_p2_1();
            raised = false;
            try
                h.val = 999; %#ok<NASGU>
            catch
                raised = true;
            end
            testCase.verifyTrue(raised, 'assignment to the lazy property must raise (read-only)');
        end

        function test_lazyproperty_guide_is_not_callable(testCase)
            % mat2doc.shared.lazyproperty is a NON-callable idiom guide (there are
            % no decorators in MATLAB); calling it raises a dedicated identifier.
            caught = [];
            try
                mat2doc.shared.lazyproperty();
            catch ME
                caught = ME;
            end
            testCase.assertNotEmpty(caught, 'the lazyproperty guide must not be callable');
            testCase.verifyEqual(caught.identifier, 'mat2doc:lazyproperty:notCallable', ...
                'calling the guide must raise mat2doc:lazyproperty:notCallable');
        end

        % =============================================================== %
        % 6. Exceptions (flat identifiers, verbatim messages)             %
        % =============================================================== %

        function test_exc_pythondocxerror_identifier_and_message(testCase)
            % PythonDocxError raiser -> mat2doc:PythonDocxError, verbatim message.
            % (validate_P2-1 Bar 6.)
            [id, msg] = raiseAndCapture(@() mat2doc.exc.PythonDocxError("generic boom"));
            testCase.verifyEqual(id, 'mat2doc:PythonDocxError', ...
                'PythonDocxError must raise the mat2doc:PythonDocxError identifier');
            testCase.verifyEqual(msg, 'generic boom', 'message must be verbatim');
        end

        function test_exc_invalidspanerror_percent_literal_message(testCase)
            % InvalidSpanError raiser -> mat2doc:InvalidSpanError, with a %-LITERAL
            % message. Pins the single-string "%s" rule: a bare `%` in the message
            % must survive verbatim (never consumed as a format spec).
            % (validate_P2-1 Bar 6.)
            [id, msg] = raiseAndCapture(@() mat2doc.exc.InvalidSpanError("span 100% invalid"));
            testCase.verifyEqual(id, 'mat2doc:InvalidSpanError', ...
                'InvalidSpanError must raise the mat2doc:InvalidSpanError identifier');
            testCase.verifyEqual(msg, 'span 100% invalid', ...
                'the %-literal message must survive verbatim (single-string "%s" rule)');
        end

        function test_exc_invalidxmlerror_identifier_and_message(testCase)
            % InvalidXmlError raiser -> mat2doc:InvalidXmlError, verbatim message.
            % (validate_P2-1 Bar 6.)
            [id, msg] = raiseAndCapture(@() mat2doc.exc.InvalidXmlError("missing <w:sectPr>"));
            testCase.verifyEqual(id, 'mat2doc:InvalidXmlError', ...
                'InvalidXmlError must raise the mat2doc:InvalidXmlError identifier');
            testCase.verifyEqual(msg, 'missing <w:sectPr>', 'message must be verbatim');
        end

        function test_exc_hierarchy_intent_flat_identifiers(testCase)
            % Hierarchy INTENT documented (VERIFY-2 ratified): in python-docx the
            % family is InvalidSpanError / InvalidXmlError issubclass
            % PythonDocxError issubclass Exception (exceptions.py 7-18). MATLAB
            % identifiers are FLAT -- there is no identifier-level is-a -- so all
            % three map to DISTINCT identifiers under the SHARED "mat2doc:" prefix,
            % which is the (broader) analogue of a "catch any python-docx error".
            % This is inert at runtime: the clone src has ZERO `except` sites for
            % any of the three (raise-only). Pin the flat-identifier model here.
            ids = strings(1, 3);
            [ids(1), ~] = raiseAndCapture(@() mat2doc.exc.PythonDocxError("x"));
            [ids(2), ~] = raiseAndCapture(@() mat2doc.exc.InvalidSpanError("x"));
            [ids(3), ~] = raiseAndCapture(@() mat2doc.exc.InvalidXmlError("x"));
            testCase.verifyTrue(all(startsWith(ids, "mat2doc:")), ...
                'all three raisers must live under the shared mat2doc: prefix');
            testCase.verifyEqual(numel(unique(ids)), 3, ...
                'the three exceptions must map to three DISTINCT flat identifiers');
        end

    end

    % ===================== instance helpers ============================ %
    methods (Access = private)

        function bytes = saveToTemp(~, d)
            % d.save to a BINARY-mode temp .docx (the writer opens "wb"), read the
            % bytes back, delete. Returns the whole-zip uint8 vector.
            tmp = [tempname '.docx'];
            cleanup = onCleanup(@() deleteIfExists(tmp)); %#ok<NASGU>
            d.save(tmp);
            bytes = readBytes(tmp);
        end

        function blob = entryBlob(testCase, membername)
            % The bytes of one member of the cached save (built in setup).
            i = find(testCase.candNames_ == string(membername), 1);
            assert(~isempty(i), 'zip entry not found: %s', membername);
            blob = testCase.candBlobs_{i};
        end
    end
end

% ===================== file-local helpers ============================== %

function [id, msg] = raiseAndCapture(fn)
    % Run fn() expecting it to raise; return (identifier, message) as char, or
    % ('','') if it did not raise.
    id = ''; msg = '';
    try
        fn();
    catch ME
        id = ME.identifier;
        msg = ME.message;
    end
end

function [blobs, names] = zipEntryList(zipBytes)
    % Enumerate an in-memory zip in stream (write) order -> ordered {names, blobs}.
    % java.util.zip.ZipInputStream reads local file headers in physical order, so
    % `names` is the true zip-entry write sequence. (Copied from
    % tests\api\Test_p1_8_skeleton_m1.m so the order pin is reader-independent.)
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

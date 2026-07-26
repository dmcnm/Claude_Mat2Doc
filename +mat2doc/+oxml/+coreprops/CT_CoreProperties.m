classdef CT_CoreProperties < mat2doc.oxml.BaseOxmlElement
% CT_COREPROPERTIES Custom element class for the <cp:coreProperties> element.
%
%   The root element of the Core Properties part stored as /docProps/core.xml.
%   Implements many of the Dublin Core document metadata elements. String
%   elements resolve to an empty string ("") if the element is not present in
%   the XML. String elements are limited in length to 255 unicode characters.
%
%   PRECEDENT (near-direct re-port): Mat2Ppt already ported this exact lineage
%   from python-pptx (Mat2Ppt/+mat2ppt/+oxml/+coreprops/CT_CoreProperties.m).
%   python-docx v1.2.0 coreprops.py is STRUCTURALLY IDENTICAL to python-pptx
%   v1.0.2 (both inherit the shared OPC core-properties lineage) -- every symbol
%   verified line-by-line against docx v1.2.0 (see audit_P1-7_coreprops.md). The
%   Mat2Ppt solutions to the hard parts (W3CDTF round-trip, the xsi nsdecl hoist,
%   the revision int grammar) carry over verbatim; docx deltas noted below.
%
%   DESCRIPTORS: 15 ZeroOrOne child descriptors, all with successors=() (each
%   child appends at end of the element, matching Python -- coreprops.py 27-43),
%   ported per design.md section 2 as a Constant schema table (TAG) + one-line
%   delegating members calling the BaseOxmlElement child engine (getChild /
%   getOrAddChild / insertChildInSequence / addChild / removeChild). Because
%   every successor list is empty, insertion always appends -- so NO_SUCCESSORS
%   (a 1x0 string) is shared by every insert/add/get_or_add member.
%
%   ACCESSORS: the Dublin-Core string/datetime/int accessor properties
%   (author_text, category_text, ... created_datetime, modified_datetime,
%   revision_number, ...) that CoreProperties delegates to. These read/write the
%   descriptor child elements' text through the private helpers text_of_element_
%   / set_element_text_ / datetime_of_element_ / set_element_datetime_ /
%   get_or_add_ -- faithful ports of the Python `_text_of_element` /
%   `_set_element_text` / `_datetime_of_element` / `_set_element_datetime` /
%   `_get_or_add` (which use getattr reflection; MATLAB dynamic member access --
%   obj.(name) for the descriptor property and feval("get_or_add_"+name, obj)
%   for the generated adder -- is the faithful analogue, design.md section 2).
%
%   DATES (H14): W3CDTF parse/format. parse_W3CDTF_to_datetime_ mirrors the
%   Python strptime loop (four templates, LAST successful match wins, whole
%   parseable-part must match) with a STRPTIME-FAITHFUL grammar (Mat2Ppt Gate-2
%   F2): each template is an anchored regex replicating CPython _strptime's field
%   widths (%Y exactly four digits; %m/%d/%H/%M/%S one-or-two digits), and
%   datetime construction validates the calendar/range by component round-trip
%   (a rolled-over field -> reject, as strptime's datetime(...) raises) plus a
%   year>=1 floor (Mat2Ppt Gate-2 F2b: CPython datetime MINYEAR=1 rejects year 0,
%   but MATLAB datetime accepts it and would re-emit "0001-..."). Plus the manual
%   numeric-offset handling (offset_dt_ / offset_pattern_). set_element_datetime_
%   serializes with strftime("%Y-%m-%dT%H:%M:%SZ") and, for created/modified,
%   stamps the xsi:type="dcterms:W3CDTF" attribute (see the XSI HACK note).
%
%   XSI HACK (coreprops.py 260-275): created/modified require an explicit
%   xsi:type attribute AND the xsi namespace declared ONCE on the ROOT element
%   (not repeated per child). Python achieves the root placement with an opaque
%   lxml reconciliation trick (set a throwaway xsi:foo attr on the root, set
%   xsi:type on the child, delete xsi:foo -- lines 273-275) whose DOCUMENTED
%   intent is exactly "add the xsi namespace to the root element rather than each
%   child element". Our serializer would otherwise INVENT an nsN prefix on the
%   child (the xsi URI is absent from the new() template, which declares only
%   cp/dc/dcterms). The literal lxml trick does not reproduce in our tree (set
%   does not reconcile decls), so this port translates the trick to its stated
%   intent: XmlElement.set_nsdecl_("xsi", ...) declares xsi on the root, after
%   which the serializer renders the child attribute as xsi:type. Trailing-of-
%   decls placement -> cp,dc,dcterms,xsi matches lxml. This path is DEAD for the
%   default.docx round-trip (the template core.xml already carries the xsi decl
%   and xsi:type, parsed verbatim) and LIVE only when creating core-properties on
%   a core-less package or setting created/modified. DEVIATION D-serializer-nsdecl
%   (adopt-only; deviation_ledger.md).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1 contract):
%   forwards ALL positional args to BaseOxmlElement via varargin with no nsmap
%   re-validation -- REQUIRED because the parser instantiates this class as
%   feval(cls, name, ownDecls[, resolvedUri]) with ownDecls an Nx2 decl-pair when
%   it hits the registered <cp:coreProperties> root of a real core.xml.
%
%   DEVIATIONS (adopt-only, signed, ledgered -- validation\summary\deviation_ledger.md):
%     * D-002 (ASCII-only int/W3CDTF grammar). Both the revision reader
%       parse_int_ and the W3CDTF date grammar parse_W3CDTF_to_datetime_ accept
%       ASCII digits [0-9] only, whereas CPython int() / _strptime additionally
%       accept Unicode decimal digits, single underscores between digits, and
%       arbitrary precision. Unicode-digit / underscore inputs yield 0
%       (revision) / [] (date) here; an ASCII digit string >2^53 yields the
%       nearest-double ROUNDED value (the ledgered "imprecise double" case,
%       e.g. "99999999999999999999" -> 1e20 where CPython is exact) -- the
%       SAFE under-accept direction, and DEAD on any real core.xml (a
%       cp:revision / dcterms:* value is a small ASCII literal; probe:
%       default.docx revision "1"). parse_int_ NEVER emits NaN/Inf (MATLAB \d is
%       Unicode-aware, so a Unicode digit matched \d yet str2double->NaN leaked;
%       the [0-9] grammar + isfinite guard prevents this).
%     * D-serializer-nsdecl (xsi nsdecl hoist; see XSI HACK note above).
%     * D-005 (Python type-token in an error message). set_element_datetime_'s
%       non-datetime error reports the MATLAB class token (e.g. "string") where
%       CPython reports type(value) (e.g. "<class 'str'>"). Exception CLASS +
%       message template faithful; dead path (values arrive as datetime). The
%       W3CDTF-format wall-clock deviation (D-coreprops-time) lives on the
%       CorePropertiesPart.default() caller, not here.
%
%   None SENTINEL: this class uses inline `isequal(x, [])` for `x is None`
%   tests, matching the established Mat2Doc oxml convention (XmlElement.m,
%   BaseOxmlElement.m, CT_Relationship.m all inline isequal(...,[])); Mat2Ppt's
%   coreprops used mat2ppt.util.isNone (a helper Mat2Doc's +shared does not have).
%   "" / "" is NOT None (H3 tri-state); text is a "string scalar or []" per
%   XmlElement's contract.
%
%   Example:
%       e = mat2doc.oxml.coreprops.CT_CoreProperties.new();
%       e.title_text = "Report";         % creates <dc:title>Report</dc:title>
%       disp(e.title_text)               % "Report"
%       disp(e.revision_number)          % 0 (no <cp:revision> yet)
%
%   Ported from python-docx v1.2.0: src/docx/oxml/coreprops.py::CT_CoreProperties
%   (registered for <cp:coreProperties>, oxml/__init__.py:96)

    properties (Constant, Hidden)  % schema table (from the Python ZeroOrOne declarations)
        NO_SUCCESSORS = string.empty(1, 0)   % every descriptor has successors=()
        CATEGORY_TAG       = "cp:category"        % ZeroOrOne @ coreprops.py:27
        CONTENTSTATUS_TAG  = "cp:contentStatus"   % ZeroOrOne @ coreprops.py:28
        CREATED_TAG        = "dcterms:created"    % ZeroOrOne @ coreprops.py:29
        CREATOR_TAG        = "dc:creator"         % ZeroOrOne @ coreprops.py:30
        DESCRIPTION_TAG    = "dc:description"      % ZeroOrOne @ coreprops.py:31
        IDENTIFIER_TAG     = "dc:identifier"      % ZeroOrOne @ coreprops.py:32
        KEYWORDS_TAG       = "cp:keywords"        % ZeroOrOne @ coreprops.py:33
        LANGUAGE_TAG       = "dc:language"        % ZeroOrOne @ coreprops.py:34
        LASTMODIFIEDBY_TAG = "cp:lastModifiedBy"  % ZeroOrOne @ coreprops.py:35
        LASTPRINTED_TAG    = "cp:lastPrinted"     % ZeroOrOne @ coreprops.py:36
        MODIFIED_TAG       = "dcterms:modified"   % ZeroOrOne @ coreprops.py:37
        REVISION_TAG       = "cp:revision"        % ZeroOrOne @ coreprops.py:38-40
        SUBJECT_TAG        = "dc:subject"         % ZeroOrOne @ coreprops.py:41
        TITLE_TAG          = "dc:title"           % ZeroOrOne @ coreprops.py:42
        VERSION_TAG        = "cp:version"         % ZeroOrOne @ coreprops.py:43
    end

    properties (Dependent)  % generated descriptor properties (ZeroOrOne getters -> child or [])
        category
        contentStatus
        created
        creator
        description
        identifier
        keywords
        language
        lastModifiedBy
        lastPrinted
        modified
        revision
        subject
        title
        version
    end

    properties (Dependent)  % Dublin-Core string/datetime/int accessors (read/write @property)
        author_text
        category_text
        comments_text
        contentStatus_text
        identifier_text
        keywords_text
        language_text
        lastModifiedBy_text
        subject_text
        title_text
        version_text
        created_datetime
        lastPrinted_datetime
        modified_datetime
        revision_number
    end

    methods
        function obj = CT_CoreProperties(varargin)
            % CT_COREPROPERTIES Construct a loose <cp:coreProperties> element.
            %   TRANSPARENT PASS-THROUGH (design.md section 2 INT-1): forward all
            %   positional args verbatim; the base XmlElement ctor is the single
            %   point accepting both nsmap currencies (struct / Nx2 decl-pair).
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ============ generated ZeroOrOne descriptor members ============
        % Each block: get.<name> (child or []), get_or_add_<name>, new_<name>_,
        % insert_<name>_, add_<name>_, remove_<name>_ -- one-line delegation to
        % the BaseOxmlElement engine. Successors are empty for all (append).

        % ---- category ----
        function child = get.category(obj);            child = obj.getChild(obj.CATEGORY_TAG); end
        function child = get_or_add_category(obj);     child = obj.getOrAddChild(obj.CATEGORY_TAG, obj.NO_SUCCESSORS); end
        function child = new_category_(obj);           child = obj.newChild(obj.CATEGORY_TAG); end
        function child = insert_category_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_category_(obj, varargin); child = obj.addChild(obj.CATEGORY_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function remove_category_(obj);                obj.removeChild(obj.CATEGORY_TAG); end

        % ---- contentStatus ----
        function child = get.contentStatus(obj);            child = obj.getChild(obj.CONTENTSTATUS_TAG); end
        function child = get_or_add_contentStatus(obj);     child = obj.getOrAddChild(obj.CONTENTSTATUS_TAG, obj.NO_SUCCESSORS); end
        function child = new_contentStatus_(obj);           child = obj.newChild(obj.CONTENTSTATUS_TAG); end
        function child = insert_contentStatus_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_contentStatus_(obj, varargin); child = obj.addChild(obj.CONTENTSTATUS_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function remove_contentStatus_(obj);                obj.removeChild(obj.CONTENTSTATUS_TAG); end

        % ---- created ----
        function child = get.created(obj);            child = obj.getChild(obj.CREATED_TAG); end
        function child = get_or_add_created(obj);     child = obj.getOrAddChild(obj.CREATED_TAG, obj.NO_SUCCESSORS); end
        function child = new_created_(obj);           child = obj.newChild(obj.CREATED_TAG); end
        function child = insert_created_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_created_(obj, varargin); child = obj.addChild(obj.CREATED_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function remove_created_(obj);                obj.removeChild(obj.CREATED_TAG); end

        % ---- creator ----
        function child = get.creator(obj);            child = obj.getChild(obj.CREATOR_TAG); end
        function child = get_or_add_creator(obj);     child = obj.getOrAddChild(obj.CREATOR_TAG, obj.NO_SUCCESSORS); end
        function child = new_creator_(obj);           child = obj.newChild(obj.CREATOR_TAG); end
        function child = insert_creator_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_creator_(obj, varargin); child = obj.addChild(obj.CREATOR_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function remove_creator_(obj);                obj.removeChild(obj.CREATOR_TAG); end

        % ---- description ----
        function child = get.description(obj);            child = obj.getChild(obj.DESCRIPTION_TAG); end
        function child = get_or_add_description(obj);     child = obj.getOrAddChild(obj.DESCRIPTION_TAG, obj.NO_SUCCESSORS); end
        function child = new_description_(obj);           child = obj.newChild(obj.DESCRIPTION_TAG); end
        function child = insert_description_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_description_(obj, varargin); child = obj.addChild(obj.DESCRIPTION_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function remove_description_(obj);                obj.removeChild(obj.DESCRIPTION_TAG); end

        % ---- identifier ----
        function child = get.identifier(obj);            child = obj.getChild(obj.IDENTIFIER_TAG); end
        function child = get_or_add_identifier(obj);     child = obj.getOrAddChild(obj.IDENTIFIER_TAG, obj.NO_SUCCESSORS); end
        function child = new_identifier_(obj);           child = obj.newChild(obj.IDENTIFIER_TAG); end
        function child = insert_identifier_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_identifier_(obj, varargin); child = obj.addChild(obj.IDENTIFIER_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function remove_identifier_(obj);                obj.removeChild(obj.IDENTIFIER_TAG); end

        % ---- keywords ----
        function child = get.keywords(obj);            child = obj.getChild(obj.KEYWORDS_TAG); end
        function child = get_or_add_keywords(obj);     child = obj.getOrAddChild(obj.KEYWORDS_TAG, obj.NO_SUCCESSORS); end
        function child = new_keywords_(obj);           child = obj.newChild(obj.KEYWORDS_TAG); end
        function child = insert_keywords_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_keywords_(obj, varargin); child = obj.addChild(obj.KEYWORDS_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function remove_keywords_(obj);                obj.removeChild(obj.KEYWORDS_TAG); end

        % ---- language ----
        function child = get.language(obj);            child = obj.getChild(obj.LANGUAGE_TAG); end
        function child = get_or_add_language(obj);     child = obj.getOrAddChild(obj.LANGUAGE_TAG, obj.NO_SUCCESSORS); end
        function child = new_language_(obj);           child = obj.newChild(obj.LANGUAGE_TAG); end
        function child = insert_language_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_language_(obj, varargin); child = obj.addChild(obj.LANGUAGE_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function remove_language_(obj);                obj.removeChild(obj.LANGUAGE_TAG); end

        % ---- lastModifiedBy ----
        function child = get.lastModifiedBy(obj);            child = obj.getChild(obj.LASTMODIFIEDBY_TAG); end
        function child = get_or_add_lastModifiedBy(obj);     child = obj.getOrAddChild(obj.LASTMODIFIEDBY_TAG, obj.NO_SUCCESSORS); end
        function child = new_lastModifiedBy_(obj);           child = obj.newChild(obj.LASTMODIFIEDBY_TAG); end
        function child = insert_lastModifiedBy_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_lastModifiedBy_(obj, varargin); child = obj.addChild(obj.LASTMODIFIEDBY_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function remove_lastModifiedBy_(obj);                obj.removeChild(obj.LASTMODIFIEDBY_TAG); end

        % ---- lastPrinted ----
        function child = get.lastPrinted(obj);            child = obj.getChild(obj.LASTPRINTED_TAG); end
        function child = get_or_add_lastPrinted(obj);     child = obj.getOrAddChild(obj.LASTPRINTED_TAG, obj.NO_SUCCESSORS); end
        function child = new_lastPrinted_(obj);           child = obj.newChild(obj.LASTPRINTED_TAG); end
        function child = insert_lastPrinted_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_lastPrinted_(obj, varargin); child = obj.addChild(obj.LASTPRINTED_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function remove_lastPrinted_(obj);                obj.removeChild(obj.LASTPRINTED_TAG); end

        % ---- modified ----
        function child = get.modified(obj);            child = obj.getChild(obj.MODIFIED_TAG); end
        function child = get_or_add_modified(obj);     child = obj.getOrAddChild(obj.MODIFIED_TAG, obj.NO_SUCCESSORS); end
        function child = new_modified_(obj);           child = obj.newChild(obj.MODIFIED_TAG); end
        function child = insert_modified_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_modified_(obj, varargin); child = obj.addChild(obj.MODIFIED_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function remove_modified_(obj);                obj.removeChild(obj.MODIFIED_TAG); end

        % ---- revision ----
        function child = get.revision(obj);            child = obj.getChild(obj.REVISION_TAG); end
        function child = get_or_add_revision(obj);     child = obj.getOrAddChild(obj.REVISION_TAG, obj.NO_SUCCESSORS); end
        function child = new_revision_(obj);           child = obj.newChild(obj.REVISION_TAG); end
        function child = insert_revision_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_revision_(obj, varargin); child = obj.addChild(obj.REVISION_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function remove_revision_(obj);                obj.removeChild(obj.REVISION_TAG); end

        % ---- subject ----
        function child = get.subject(obj);            child = obj.getChild(obj.SUBJECT_TAG); end
        function child = get_or_add_subject(obj);     child = obj.getOrAddChild(obj.SUBJECT_TAG, obj.NO_SUCCESSORS); end
        function child = new_subject_(obj);           child = obj.newChild(obj.SUBJECT_TAG); end
        function child = insert_subject_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_subject_(obj, varargin); child = obj.addChild(obj.SUBJECT_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function remove_subject_(obj);                obj.removeChild(obj.SUBJECT_TAG); end

        % ---- title ----
        function child = get.title(obj);            child = obj.getChild(obj.TITLE_TAG); end
        function child = get_or_add_title(obj);     child = obj.getOrAddChild(obj.TITLE_TAG, obj.NO_SUCCESSORS); end
        function child = new_title_(obj);           child = obj.newChild(obj.TITLE_TAG); end
        function child = insert_title_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_title_(obj, varargin); child = obj.addChild(obj.TITLE_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function remove_title_(obj);                obj.removeChild(obj.TITLE_TAG); end

        % ---- version ----
        function child = get.version(obj);            child = obj.getChild(obj.VERSION_TAG); end
        function child = get_or_add_version(obj);     child = obj.getOrAddChild(obj.VERSION_TAG, obj.NO_SUCCESSORS); end
        function child = new_version_(obj);           child = obj.newChild(obj.VERSION_TAG); end
        function child = insert_version_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_version_(obj, varargin); child = obj.addChild(obj.VERSION_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function remove_version_(obj);                obj.removeChild(obj.VERSION_TAG); end

        % ================= Dublin-Core string accessors =================
        % Each maps a public *_text property to its descriptor element's text
        % via the text_of_element_ / set_element_text_ helpers (coreprops.py
        % 54-191). The property NAME passed is the ZeroOrOne descriptor name.

        function v = get.author_text(obj);                 v = obj.text_of_element_("creator"); end          % coreprops.py 54-61
        function set.author_text(obj, value);              obj.set_element_text_("creator", value); end
        function v = get.category_text(obj);               v = obj.text_of_element_("category"); end         % coreprops.py 63-69
        function set.category_text(obj, value);            obj.set_element_text_("category", value); end
        function v = get.comments_text(obj);               v = obj.text_of_element_("description"); end      % coreprops.py 71-77
        function set.comments_text(obj, value);            obj.set_element_text_("description", value); end
        function v = get.contentStatus_text(obj);          v = obj.text_of_element_("contentStatus"); end    % coreprops.py 79-85
        function set.contentStatus_text(obj, value);       obj.set_element_text_("contentStatus", value); end
        function v = get.identifier_text(obj);             v = obj.text_of_element_("identifier"); end       % coreprops.py 95-101
        function set.identifier_text(obj, value);          obj.set_element_text_("identifier", value); end
        function v = get.keywords_text(obj);               v = obj.text_of_element_("keywords"); end         % coreprops.py 103-109
        function set.keywords_text(obj, value);            obj.set_element_text_("keywords", value); end
        function v = get.language_text(obj);               v = obj.text_of_element_("language"); end         % coreprops.py 111-117
        function set.language_text(obj, value);            obj.set_element_text_("language", value); end
        function v = get.lastModifiedBy_text(obj);         v = obj.text_of_element_("lastModifiedBy"); end   % coreprops.py 119-125
        function set.lastModifiedBy_text(obj, value);      obj.set_element_text_("lastModifiedBy", value); end
        function v = get.subject_text(obj);                v = obj.text_of_element_("subject"); end          % coreprops.py 169-175
        function set.subject_text(obj, value);             obj.set_element_text_("subject", value); end
        function v = get.title_text(obj);                  v = obj.text_of_element_("title"); end            % coreprops.py 177-183
        function set.title_text(obj, value);               obj.set_element_text_("title", value); end
        function v = get.version_text(obj);                v = obj.text_of_element_("version"); end          % coreprops.py 185-191
        function set.version_text(obj, value);             obj.set_element_text_("version", value); end

        % ================ Dublin-Core datetime accessors ================
        function v = get.created_datetime(obj);            v = obj.datetime_of_element_("created"); end      % coreprops.py 87-93
        function set.created_datetime(obj, value);         obj.set_element_datetime_("created", value); end
        function v = get.lastPrinted_datetime(obj);        v = obj.datetime_of_element_("lastPrinted"); end  % coreprops.py 127-133
        function set.lastPrinted_datetime(obj, value);     obj.set_element_datetime_("lastPrinted", value); end
        function v = get.modified_datetime(obj);           v = obj.datetime_of_element_("modified"); end     % coreprops.py 135-141
        function set.modified_datetime(obj, value);        obj.set_element_datetime_("modified", value); end

        % ===================== revision (int) ===========================
        function value = get.revision_number(obj)
            % REVISION_NUMBER Integer value of revision property (coreprops.py 143-158).
            %   Non-integer / negative revision strings resolve to 0.
            revision = obj.revision;                          % child or [] (H3)
            if isequal(revision, [])                          % if revision is None
                value = 0;   return
            end
            % Python: revision_str = str(revision.text). When the element is
            % present but empty (<cp:revision/>), revision.text is None and
            % str(None) == "None" -- which then fails int() -> 0 (below). We
            % replicate that str() including the None->"None" case, rather than
            % short-circuiting, so the path matches coreprops.py:149 exactly.
            revision_str = revision.text;                     % string or [] (H3)
            if isequal(revision_str, [])
                revision_str = "None";                        % str(None) == "None"
            end
            n = mat2doc.oxml.coreprops.CT_CoreProperties.parse_int_(revision_str);
            if isempty(n)                                     % int() ValueError -> 0
                value = 0;   return
            end
            if n < 0                                          % negative integers -> 0
                value = 0;   return
            end
            value = n;
        end

        function set.revision_number(obj, value)
            % set revision property to string value of integer `value`
            %   (coreprops.py 160-167). Requires a positive int.
            %   DEAD-PATH NOTE (D-002 family): CPython accepts a bool here (bool
            %   subclasses int, so revision=True writes text "True"); the port's
            %   isnumeric guard rejects a MATLAB logical. No ported caller passes
            %   a logical to revision -- absurd dead path, ledgered not fixed.
            if ~(isnumeric(value) && isscalar(value) && isfinite(value) && ...
                    floor(value) == value) || value < 1
                error("mat2doc:ValueError", ...
                    "revision property requires positive int, got '%s'", ...
                    mat2doc.oxml.coreprops.CT_CoreProperties.repr_value_(value));
            end
            revision = obj.get_or_add_revision();
            revision.text = mat2doc.shared.pyStr(value, "int");   % str(value) (H14)
        end
    end

    methods (Static)
        function element = new()
            % NEW Return a new <cp:coreProperties> element (coreprops.py 47-52):
            %   parse_xml of the cp/dc/dcterms template. (Named `new` to match the
            %   Python classmethod and the Mat2Doc CT_Relationship.new convention.)
            tmpl = "<cp:coreProperties " + ...
                mat2doc.oxml.nsdecls("cp", "dc", "dcterms") + "/>" + newline;
            element = mat2doc.oxml.parse_xml(tmpl);
        end
    end

    methods (Access = private)
        function value = text_of_element_(obj, property_name)
            % TEXT_OF_ELEMENT_ Text of the descriptor child, or "" (coreprops.py 288-298).
            element = obj.(property_name);                    % descriptor getter
            if isequal(element, [])                           % if element is None
                value = "";   return
            end
            if isequal(element.text, [])                      % if element.text is None
                value = "";   return
            end
            value = element.text;
        end

        function set_element_text_(obj, prop_name, value)
            % SET_ELEMENT_TEXT_ Set the descriptor child's text (coreprops.py 277-286).
            %   Python: if not isinstance(value, str): value = str(value). The
            %   non-str branch IS reachable through the public CoreProperties
            %   setters (e.g. cp.comments = 1/3), so the str() must be pyStr
            %   (H14), not raw string(): string(1/3) is "0.33333" where Python
            %   str(1/3) is "0.3333333333333333", and string(true) is "true"
            %   where str(True) is "True". pyStr passes strings/chars through
            %   unchanged and errors loudly on types it cannot render as Python
            %   str() would (never a silent divergence).
            %   (Gate-2 P1-7 audit fix, probe T03.)
            value = mat2doc.shared.pyStr(value);              % value = str(value)
            % Count CODE POINTS, not UTF-16 code units (H2): Python len(str)
            % counts Unicode code points, but strlength counts UTF-16 code units,
            % so an astral char (surrogate pair, e.g. an emoji) would count as 2
            % and a 255-code-point emoji-bearing title Python accepts would
            % wrongly raise here.
            if mat2doc.oxml.coreprops.CT_CoreProperties.codepoint_length_(value) > 255
                error("mat2doc:ValueError", ...
                    "exceeded 255 char limit for property, got:\n\n'%s'", value);
            end
            element = obj.get_or_add_(prop_name);
            element.text = value;
        end

        function value = datetime_of_element_(obj, property_name)
            % DATETIME_OF_ELEMENT_ Datetime of the descriptor child, or [] (coreprops.py 193-202).
            element = obj.(property_name);                    % descriptor getter
            if isequal(element, [])                           % if element is None
                value = [];   return
            end
            datetime_str = element.text;
            if isequal(datetime_str, [])                      % if datetime_str is None
                value = [];   return
            end
            try
                value = mat2doc.oxml.coreprops.CT_CoreProperties ...
                    .parse_W3CDTF_to_datetime_(datetime_str);
            catch ME
                if ME.identifier == "mat2doc:ValueError"      % invalid strings ignored
                    value = [];
                else
                    rethrow(ME)
                end
            end
        end

        function set_element_datetime_(obj, prop_name, value)
            % SET_ELEMENT_DATETIME_ Set date/time of the descriptor child (coreprops.py 260-275).
            if ~isa(value, "datetime")
                % D-005 (type-token divergence): the message reports the MATLAB
                % class token (e.g. "string") where CPython reports type(value)
                % (e.g. "<class 'str'>"). Exception CLASS + template faithful;
                % dead path (values arrive as datetime). See deviation_ledger.md.
                error("mat2doc:ValueError", ...
                    "property requires <type 'datetime.datetime'> object, got %s", ...
                    class(value));
            end
            element = obj.get_or_add_(prop_name);
            % strftime("%Y-%m-%dT%H:%M:%SZ") -- whole-second, literal-Z-marked
            % stamp. MATLAB datetime formatting emits the wall-clock fields; the
            % 'Z' is a literal suffix (matching Python, whose strftime ignores
            % tzinfo for %S and appends the literal Z). VERIFY-tz: the getter
            % returns a NAIVE datetime (no TimeZone), mirroring the Mat2Ppt
            % precedent, where Python returns a tz-aware UTC datetime -- an
            % API-VALUE difference that does NOT affect XML bytes (see audit).
            dt_str = string(value, "yyyy-MM-dd'T'HH:mm:ss") + "Z";
            element.text = dt_str;
            if prop_name == "created" || prop_name == "modified"
                % coreprops.py 269-275 XSI HACK (see class header): declare xsi
                % ONCE on the root, then stamp xsi:type on the child. The Python
                % set/del-throwaway lxml trick is translated to its documented
                % intent (add xsi to the root element). D-serializer-nsdecl.
                m = mat2doc.oxml.nsmap();
                obj.set_nsdecl_("xsi", m.xsi);
                element.set(mat2doc.oxml.qn("xsi:type"), "dcterms:W3CDTF");
            end
        end

        function element = get_or_add_(obj, prop_name)
            % GET_OR_ADD_ Element from the get_or_add_<prop_name> member (coreprops.py 204-209).
            %   Python getattr reflection -> MATLAB dynamic method dispatch.
            element = feval("get_or_add_" + prop_name, obj);
        end
    end

    methods (Static, Access = private)
        function ts = parse_W3CDTF_to_datetime_(w3cdtf_str)
            % PARSE_W3CDTF_TO_DATETIME_ (coreprops.py 229-258). Mirrors the
            %   strptime loop over four templates (LAST successful match wins; the
            %   whole parseable part must match), then the manual numeric-offset
            %   handling (a literal '-07:30' strptime cannot parse).
            %
            %   STRPTIME-FAITHFUL GRAMMAR (Mat2Ppt Gate-2 F2): each template is an
            %   ANCHORED regex replicating CPython _strptime's field widths --
            %   %Y is EXACTLY four digits, %m/%d/%H/%M/%S are one OR two digits --
            %   then MATLAB datetime construction validates the calendar/range by
            %   round-trip: a rolled-over component (Feb-31 -> Mar-3, sec-60 ->
            %   +1min, month-13 -> next year) means a field was out of range, so
            %   the template is rejected -- exactly as strptime's underlying
            %   datetime(...) raises ValueError.
            %
            %   ASCII digits only ([0-9]) -- DEVIATION D-002: CPython _strptime's
            %   %Y/%m/... use a Unicode-aware \d, so it ACCEPTS Unicode-digit
            %   dates; this ASCII grammar returns [] for those. SAFE direction
            %   (under-accept), dead path (a cp:created/modified value in any real
            %   core.xml is ASCII). Ledgered under D-002.
            s = string(w3cdtf_str);
            if strlength(s) >= 20
                parseable_part = extractBefore(s, 20);          % w3cdtf_str[:19]
                offset_str = extractAfter(s, 19);               % w3cdtf_str[19:]
            else
                parseable_part = s;
                offset_str = "";
            end
            YY = "([0-9]{4})";   FF = "([0-9]{1,2})";
            % CPython _strptime compiles the format regex with re.IGNORECASE, so
            % the literal "T" separator matches "t" too (oracle-verified:
            % strptime("2003-12-31t10:14:55", "%Y-%m-%dT%H:%M:%S") parses).
            % [Tt] replicates that; the other literals (-,:) have no case.
            % (Gate-2 P1-7 audit fix, probe W35.)
            templates = [ ...
                "^" + YY+"-"+FF+"-"+FF+"[Tt]"+FF+":"+FF+":"+FF + "$"; ... % %Y-%m-%dT%H:%M:%S
                "^" + YY+"-"+FF+"-"+FF + "$"; ...                      % %Y-%m-%d
                "^" + YY+"-"+FF + "$"; ...                             % %Y-%m
                "^" + YY + "$"];                                       % %Y
            timestamp = [];
            for k = 1:numel(templates)
                tok = regexp(parseable_part, templates(k), "tokens", "once");
                if isempty(tok)
                    continue
                end
                vals = str2double(tok);                 % leading fields Y[,m,d,H,M,S]
                f = [1 1 1 0 0 0];                       % absent month/day=1, H/M/S=0
                f(1:numel(vals)) = vals;
                if f(1) < 1
                    % YEAR FLOOR (Mat2Ppt Gate-2 F2b): CPython datetime MINYEAR=1,
                    % so datetime(0,...) raises "year 0 is out of range" ->
                    % strptime ValueError -> None. MATLAB datetime ACCEPTS year 0
                    % and would re-serialize it as "0001-..." ('yyyy' is
                    % year-of-era) -- an over-accept + silent year-shift. Reject
                    % year 0 like Python; the %Y regex ([0-9]{4}) already caps at 9999.
                    continue
                end
                cand = datetime(f(1), f(2), f(3), f(4), f(5), f(6));
                [cy, cmo, cd] = ymd(cand);
                [ch, cmi, cs] = hms(cand);
                if ~isequal([cy cmo cd ch cmi cs], f)   % rollover -> out-of-range field
                    continue
                end
                timestamp = cand;   % no break: last successful template wins
            end
            if isempty(timestamp)
                error("mat2doc:ValueError", ...
                    "could not parse W3CDTF datetime string '%s'", s);
            end
            if strlength(offset_str) == 6
                ts = mat2doc.oxml.coreprops.CT_CoreProperties ...
                    .offset_dt_(timestamp, offset_str);
            else
                ts = timestamp;
            end
        end

        function out = offset_dt_(datetime_in, offset_str)
            % OFFSET_DT_ Offset a datetime by a '-07:00'-style string (coreprops.py 211-225).
            %   NOTE the INVERTED sign in Python: sign_factor = -1 if '+' else 1.
            pat = "([+-])(\d\d):(\d\d)";
            tok = regexp(offset_str, pat, "tokens", "once");
            if isempty(tok)
                error("mat2doc:ValueError", ...
                    "'%s' is not a valid offset string", offset_str);
            end
            sign = tok(1);
            sign_factor = -1;                       % Python: -1 if sign=='+' else 1
            if sign ~= "+"
                sign_factor = 1;
            end
            hrs = str2double(tok(2)) * sign_factor;
            mins = str2double(tok(3)) * sign_factor;
            out = datetime_in + hours(hrs) + minutes(mins);
        end

        function n = parse_int_(str_value)
            % PARSE_INT_ Python int(str) analogue over the ASCII decimal grammar;
            %   returns [] when `str_value` is not an ASCII integer literal (the
            %   caller maps [] -> 0, matching int()'s ValueError -> 0 in
            %   revision_number, coreprops.py 150-154). Accepts optional
            %   surrounding whitespace, an optional leading +/- sign, then ASCII
            %   digits only (int() rejects '1.5').
            %
            %   ASCII-ONLY ([0-9], NOT \d) -- DEVIATION D-002: MATLAB regex \d is
            %   Unicode-aware, so a Unicode-digit revision string MATCHED \d but
            %   str2double returned NaN, which escaped the revision_number guards
            %   and leaked NaN out of a property documented as a non-negative int.
            %   [0-9] rejects it (-> [] -> 0), never emitting NaN. The finite guard
            %   below likewise rejects an absurdly long ASCII digit string
            %   (str2double overflow -> Inf), never leaking Inf. Dead on real
            %   core.xml (a <cp:revision> value is a small ASCII integer; probe:
            %   default.docx revision "1").
            tok = regexp(string(str_value), "^\s*([+-]?[0-9]+)\s*$", "tokens", "once");
            if isempty(tok)
                n = [];
                return
            end
            n = str2double(tok(1));
            if ~isfinite(n)   % overflow of an absurdly long digit string -> Inf
                n = [];       % (D-002 >2^53 family); never leak Inf/NaN
            end
        end

        function n = codepoint_length_(value)
            % CODEPOINT_LENGTH_ Number of Unicode code points (scalar values) in a
            %   string scalar, matching Python len(str) (coreprops.py 282; H2).
            %   MATLAB strings are UTF-16 and strlength counts UTF-16 CODE UNITS,
            %   so an astral char (surrogate pair, e.g. an emoji U+1F600) counts as
            %   2 there but 1 in Python. Count the code units that are NOT a low
            %   surrogate (0xDC00-0xDFFF): each surrogate pair then contributes
            %   exactly 1 (its high surrogate) and every BMP unit contributes 1.
            u = double(char(value));                          % UTF-16 code units
            n = sum(~(u >= 0xDC00 & u <= 0xDFFF));
        end

        function s = repr_value_(value)
            % REPR_VALUE_ Best-effort str(value) for the revision error message
            %   (coreprops.py 164 "got '%s'" % value). Routes numerics through
            %   pyStr (H14) so an invalid float shows e.g. "2.5", not "2".
            if isnumeric(value) && isscalar(value)
                s = mat2doc.shared.pyStr(value);
            elseif (isstring(value) && isscalar(value)) || ischar(value)
                s = string(value);
            else
                s = string(class(value));
            end
        end
    end
end

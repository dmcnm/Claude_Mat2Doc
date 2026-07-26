classdef Document < mat2doc.shared.ElementProxy
% DOCUMENT WordprocessingML (WML) document -- the top API proxy object.
%
%   Not intended to be constructed directly. Use the package-level factory
%   MAT2DOC.DOCUMENT (`+mat2doc\Document.m`) to open or create a document.
%
%   THIN M1 SLICE: this is the walking-skeleton proxy. Only `save` and
%   `core_properties` (and the trivial `part` accessor) are LIVE; every content
%   member (add_paragraph / add_heading / add_table / add_picture / paragraphs /
%   sections / styles / settings / inline_shapes / tables / comments / ...) is a
%   mat2doc:notYetPorted stub. NONE is on the open->save path.
%
%   VERIFY-M1-DOC-BASE (RESOLVED in P2-1): in python-docx Document extends
%   ElementProxy (document.py 28, `class Document(ElementProxy)`), which
%   supplies element identity (`eq`/`ne`) and the `element` accessor. At M1
%   this class stored `_element`/`_part` directly and derived `handle` only.
%   P2-1 ported the shared proxy tier and retrofitted this class onto the real
%   base (mat2doc.shared.ElementProxy), mirroring the Mat2Ppt VERIFY-M1-C
%   precedent (WP8 reparented Presentation onto PartElementProxy, byte-neutral).
%   The wrapped element is now held by the base (protected `element_`), and this
%   class gains the inherited `element()` accessor and H5 element-identity
%   `eq`/`ne`. BYTE-NEUTRAL: the base adds hierarchy/identity/accessor behavior,
%   not serialized output -- re-proven by the 17-part M1 sweep at P2-1 Gate 1
%   (mat2doc.Document().save == references\s0001, 17/17). `part` is still
%   OVERRIDDEN below to return this document's own _part (document.py 193-196),
%   NOT the base ElementProxy.part (whose _parent is None here); the base
%   None-guard therefore never fires for a Document.
%
%   REFERENCE SEMANTICS (design.md section 2): a handle class -- the proxy wraps
%   a shared element tree and part, exactly like the Python proxy model.
%
%   SAVE-TO-STREAM (noted, not stubbed): python-docx accepts a path OR a
%   file-like object; save delegates unchanged to the part -> package ->
%   PackageWriter chain (P1-6a). M1 requires save-to-PATH, which works today.
%   Stream currency is inherited from the OpcPackage.save contract; any
%   stream-specific hardening is a P2-3 concern, not a P1-8 stub.
%
%   UNDERSCORE ROTATION (design.md section 2): Python `_element`/`_part` ->
%   element_/part_. Post-retrofit, element_ is the base's protected property;
%   only part_ is declared here. (Python's `__body` cache belongs to the
%   feature `_body` accessor, which is P2-3; it is not carried at M1.)
%
%   Example:
%       d = mat2doc.Document();     % opens the bundled default template
%       d.save("out.docx");
%
%   Ported from python-docx v1.2.0: src/docx/document.py::Document
%   (the M1 slice: __init__ 35-39, core_properties 165-168, part 193-196,
%   save 198-204. Base class ElementProxy is retrofitted in P2-1.)

    properties (Access = private)
        part_           % the owning mat2doc.parts.DocumentPart
        % NOTE: the wrapped w:document CT_Document root is now held by the base
        % mat2doc.shared.ElementProxy (protected `element_`); this class no
        % longer declares its own element_ (P2-1 VERIFY-M1-DOC-BASE retrofit).
    end

    methods
        function obj = Document(element, part)
            % DOCUMENT Construct over the w:document root element and its part
            %   (document.py 35-39). Python:
            %       super(Document, self).__init__(element)  # ElementProxy, parent=None
            %       self._element = element                  # redundant re-set
            %       self._part = part
            %       self.__body = None                       # P2-3 body cache
            %   The redundant `self._element = element` is a no-op (the base
            %   already stored it) and is not repeated here. `__body` is the
            %   P2-3 body-cache concern, not carried at M1.
            %
            %   Inputs:  element - the w:document root element.
            %            part    - the owning mat2doc.parts.DocumentPart.
            %   Outputs: obj     - a scalar Document handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/document.py::Document.__init__
            obj@mat2doc.shared.ElementProxy(element);   % base: element_=element, parent_=[] (None)
            obj.part_ = part;
        end

        function p = part(obj)
            % PART (document.py 193-196, @property): the DocumentPart of this
            %   document. Python: return self._part. Trivial accessor, LIVE.
            p = obj.part_;
        end

        function save(obj, path_or_stream)
            % SAVE (document.py 198-204): save this document to `path_or_stream`,
            %   a path (string) or file-like object. Python: self._part.save(...).
            %   LIVE (the M1 headline path). See SAVE-TO-STREAM note.
            obj.part_.save(path_or_stream);
        end

        function cp = core_properties(obj)
            % CORE_PROPERTIES (document.py 165-168): a CoreProperties object with
            %   Dublin Core properties of the document. Python:
            %   return self._part.core_properties. LIVE (P1-7 real).
            cp = obj.part_.core_properties();
        end

        % ------------------------------------------------------------------
        % Feature stubs (mat2doc:notYetPorted) -- P2-3 (content) / P2 tiers
        % un-stub these. NONE is on the open->save path.
        % ------------------------------------------------------------------

        function comment = add_comment(obj, runs, text, author, initials) %#ok<INUSD,MANU,STOUT>
            % ADD_COMMENT STUB (document.py 41-88). Owner: P8-2 comments tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.document.Document.add_comment (owning WP: P8-2 comments " + ...
                "tier) is not yet ported");
        end

        function p = add_heading(obj, text, level) %#ok<INUSD,MANU,STOUT>
            % ADD_HEADING STUB (document.py 90-101). P2-3 document content tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.document.Document.add_heading (owning WP: P2-3 document " + ...
                "content tier) is not yet ported");
        end

        function p = add_page_break(obj) %#ok<MANU,STOUT>
            % ADD_PAGE_BREAK STUB (document.py 103-107). P2-3 document content tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.document.Document.add_page_break (owning WP: P2-3 " + ...
                "document content tier) is not yet ported");
        end

        function p = add_paragraph(obj, text, style) %#ok<INUSD,MANU,STOUT>
            % ADD_PARAGRAPH STUB (document.py 109-119). P2-3 document content tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.document.Document.add_paragraph (owning WP: P2-3 " + ...
                "document content tier) is not yet ported");
        end

        function shape = add_picture(obj, image_path_or_stream, width, height) %#ok<INUSD,MANU,STOUT>
            % ADD_PICTURE STUB (document.py 121-138). P7 image tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.document.Document.add_picture (owning WP: P7 image " + ...
                "tier) is not yet ported");
        end

        function section = add_section(obj, start_type) %#ok<INUSD,MANU,STOUT>
            % ADD_SECTION STUB (document.py 140-148). Owner: P5 section tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.document.Document.add_section (owning WP: P5 section " + ...
                "tier) is not yet ported");
        end

        function table = add_table(obj, rows, cols, style) %#ok<INUSD,MANU,STOUT>
            % ADD_TABLE STUB (document.py 150-158). Owner: P6 table tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.document.Document.add_table (owning WP: P6 table tier) " + ...
                "is not yet ported");
        end

        function c = comments(obj) %#ok<MANU,STOUT>
            % COMMENTS STUB (document.py 160-163). Owner: P8-2 comments tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.document.Document.comments (owning WP: P8-2 comments " + ...
                "tier) is not yet ported");
        end

        function s = inline_shapes(obj) %#ok<MANU,STOUT>
            % INLINE_SHAPES STUB (document.py 170-178). Owner: P7 image/shape tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.document.Document.inline_shapes (owning WP: P7 image/shape " + ...
                "tier) is not yet ported");
        end

        function it = iter_inner_content(obj) %#ok<MANU,STOUT>
            % ITER_INNER_CONTENT STUB (document.py 180-182). P2-3 document content tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.document.Document.iter_inner_content (owning WP: P2-3 " + ...
                "document content tier) is not yet ported");
        end

        function p = paragraphs(obj) %#ok<MANU,STOUT>
            % PARAGRAPHS STUB (document.py 184-191). P2-3 document content tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.document.Document.paragraphs (owning WP: P2-3 document " + ...
                "content tier) is not yet ported");
        end

        function s = sections(obj) %#ok<MANU,STOUT>
            % SECTIONS STUB (document.py 206-209). Owner: P5 section tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.document.Document.sections (owning WP: P5 section tier) " + ...
                "is not yet ported");
        end

        function s = settings(obj) %#ok<MANU,STOUT>
            % SETTINGS STUB (document.py 211-214). Owner: P5-1 settings tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.document.Document.settings (owning WP: P5-1 settings " + ...
                "tier) is not yet ported");
        end

        function s = styles(obj) %#ok<MANU,STOUT>
            % STYLES STUB (document.py 216-219). Owner: P4-7 styles tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.document.Document.styles (owning WP: P4-7 styles tier) " + ...
                "is not yet ported");
        end

        function t = tables(obj) %#ok<MANU,STOUT>
            % TABLES STUB (document.py 221-230). Owner: P6 table tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.document.Document.tables (owning WP: P6 table tier) " + ...
                "is not yet ported");
        end
    end
end

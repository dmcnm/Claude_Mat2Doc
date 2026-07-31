classdef Document < mat2doc.shared.ElementProxy
% DOCUMENT WordprocessingML (WML) document -- the top API proxy object.
%
%   Not intended to be constructed directly. Use the package-level factory
%   MAT2DOC.DOCUMENT (`+mat2doc\Document.m`) to open or create a document.
%
%   P2-3 SLICE (+ P4-7a/P4-7b/P5-1/P5-3a): `save`, `core_properties`, `part`,
%   `styles` (P4-7a), `settings` (P5-1), the paragraph adders `add_heading` /
%   `add_paragraph` (P4-7b, the M2 critical path), and -- UN-STUBBED at P5-3a
%   (C1) -- `sections` and `add_section` are LIVE, as are the private
%   object-graph accessors `body_` / `block_width_` (block_width_ now reaches the
%   live sections). The remaining content members (add_table / add_picture /
%   add_page_break / paragraphs / inline_shapes / tables / comments / ...) are
%   mat2doc:notYetPorted stubs. NONE of the stubs is on the open->save path.
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
%   only part_ is declared here. (Python's `__body` cache rotates to the
%   body__ property, carried since P2-3.)
%
%   Example:
%       d = mat2doc.Document();     % opens the bundled default template
%       d.save("out.docx");
%
%   Ported from python-docx v1.2.0: src/docx/document.py::Document
%   (live members: __init__ 35-39, core_properties 165-168, part 193-196,
%   save 198-204, _block_width 232-239, _body 241-246. Base class ElementProxy
%   retrofitted in P2-1; body_/block_width_ added in P2-3.)

    properties (Access = private)
        part_           % the owning mat2doc.parts.DocumentPart
        % NOTE: the wrapped w:document CT_Document root is now held by the base
        % mat2doc.shared.ElementProxy (protected `element_`); this class no
        % longer declares its own element_ (P2-1 VERIFY-M1-DOC-BASE retrofit).
        body__ = []     % __body: the cached _Body proxy, or [] (None) until first
                        % _body access. Python `self.__body` (name-mangled private
                        % cache); double leading underscore rotates -> body__
                        % (design.md section 2). NOT a lazyproperty: the manual
                        % is-None cache mirrors document.py 242-246 exactly. The
                        % sentinel is the None-literal [] (a _Body handle is never
                        % [], so isequal(.,[]) is a sound cache-empty test, H3).
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
            %   already stored it) and is not repeated here. `__body = None` is
            %   the body_ accessor's cache, ported at P2-3 -> body__ = [].
            %
            %   Inputs:  element - the w:document root element.
            %            part    - the owning mat2doc.parts.DocumentPart.
            %   Outputs: obj     - a scalar Document handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/document.py::Document.__init__
            obj@mat2doc.shared.ElementProxy(element);   % base: element_=element, parent_=[] (None)
            obj.part_ = part;
            obj.body__ = [];                            % self.__body = None (H3)
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
        % LIVE object-graph SHELL (P2-3 + P5-3a): the private _body / _block_width
        % accessors. NEITHER is on the clean open->save path (proven: a bare
        % Document().save() fires ZERO stubs). body_ and (as of P5-3a, with
        % sections now LIVE) block_width_ are both fully functional; block_width_
        % is still reached only by add_table (a P6 stub), so it has no live caller
        % yet, but it now computes a real value rather than raising.
        % ------------------------------------------------------------------

        function b = body_(obj)
            % _BODY (document.py 241-246, @property): the _Body instance containing
            %   the content for this document. Python:
            %       if self.__body is None:
            %           self.__body = _Body(self._element.body, self)
            %       return self.__body
            %   MANUAL is-None cache (NOT a lazyproperty): body__ starts [] (None)
            %   and is populated on first access, then returned as-is. Because the
            %   cache holds the SAME _Body handle across calls, repeated body_()
            %   reads return one identical proxy (H5/H9 -- matches Python's cached
            %   @property; contrast DocumentPart.document, which is UNcached and
            %   yields a fresh Document each call). `self._element.body` is the
            %   CT_Document.body descriptor getter (the live <w:body> CT_Body).
            %   Underscore rotation: _body -> body_, __body -> body__.
            %
            %   Ported from python-docx v1.2.0: src/docx/document.py::Document._body
            if isequal(obj.body__, [])                  % if self.__body is None (H3)
                obj.body__ = mat2doc.document.Body_(obj.element_.body, obj);
            end
            b = obj.body__;
        end

        function w = block_width_(obj)
            % _BLOCK_WIDTH (document.py 232-239, @property): a Length giving the
            %   space between margins in the last section. Python:
            %       section = self.sections[-1]
            %       page_width  = section.page_width  or Inches(8.5)
            %       left_margin = section.left_margin or Inches(1)
            %       right_margin= section.right_margin or Inches(1)
            %       return Emu(page_width - left_margin - right_margin)
            %   Ported FAITHFULLY. As of P5-3a `sections` is LIVE (C1), so this
            %   computes a real value; it is still reached only via add_table (a P6
            %   stub), so there is no live caller yet. The Python `x or Inches(...)`
            %   falsy-default (H4: Length(0) is falsy) is ported; Emu/Inches are the
            %   shared Length subclasses (arithmetic in EMU, H6). Underscore
            %   rotation: _block_width -> block_width_.
            %
            %   Ported from python-docx v1.2.0: src/docx/document.py::Document._block_width
            %   P5-3a: sections is now LIVE (C1). `self.sections[-1]` -> the
            %   0-based getitem_(-1) on the Sections collection (negative wrap to
            %   the last section), replacing the stub-era `sections()(end)`
            %   placeholder that assumed sections returned a plain array.
            s = obj.sections();                              % Sections collection (LIVE, C1)
            section = s.getitem_(-1);                        % Python: self.sections[-1]
            page_width = section.page_width;                 % Python: section.page_width
            if isequal(page_width, []) || page_width == 0    % `or Inches(8.5)` (H4)
                page_width = mat2doc.shared.Inches(8.5);
            end
            left_margin = section.left_margin;               % Python: section.left_margin
            if isequal(left_margin, []) || left_margin == 0  % `or Inches(1)` (H4)
                left_margin = mat2doc.shared.Inches(1);
            end
            right_margin = section.right_margin;             % Python: section.right_margin
            if isequal(right_margin, []) || right_margin == 0 % `or Inches(1)` (H4)
                right_margin = mat2doc.shared.Inches(1);
            end
            w = mat2doc.shared.Emu(page_width - left_margin - right_margin);
        end

        % ------------------------------------------------------------------
        % Feature stubs (mat2doc:notYetPorted) -- the P4/P5/P6/P7/P8 tiers
        % un-stub these. NONE is on the open->save path.
        % ------------------------------------------------------------------

        function comment = add_comment(obj, runs, text, author, initials) %#ok<INUSD,MANU,STOUT>
            % ADD_COMMENT STUB (document.py 41-88). Owner: P8-2 comments tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.document.Document.add_comment (owning WP: P8-2 comments " + ...
                "tier) is not yet ported");
        end

        function p = add_heading(obj, text, level)
            % ADD_HEADING Return a heading paragraph newly added to the end of the
            %   document (document.py 90-101). The heading contains `text` and its
            %   paragraph style is determined by `level`: level 0 -> "Title";
            %   level 1 (or omitted) -> "Heading 1"; otherwise "Heading {level}".
            %   Raises mat2doc:ValueError if `level` is outside 0-9. UN-STUBBED at
            %   P4-7b.
            %
            %   Python (document.py 98-101):
            %     if not 0 <= level <= 9:
            %         raise ValueError("level must be in range 0-9, got %d" % level)
            %     style = "Title" if level == 0 else "Heading %d" % level
            %     return self.add_paragraph(text, style)
            %
            %   H13 default fidelity: add_heading(text="", level=1). H1: `level` is
            %   a DATA value (heading number used verbatim in "Heading %d" and the
            %   0-9 bound check), NOT a collection index -- no +1/-1 shift. The
            %   ValueError message is the VERBATIM Python string with id
            %   mat2doc:ValueError.
            %
            %   Ported from python-docx v1.2.0: src/docx/document.py::Document.add_heading
            arguments
                obj
                text  = ""   % Python default ""
                level = 1    % Python default 1
            end
            if ~(0 <= level && level <= 9)   % Python: if not 0 <= level <= 9
                error("mat2doc:ValueError", "%s", ...
                    sprintf("level must be in range 0-9, got %d", level));
            end
            if level == 0                    % Python: "Title" if level == 0
                style = "Title";
            else                             % Python: else "Heading %d" % level
                style = sprintf("Heading %d", level);
            end
            p = obj.add_paragraph(text, style);   % Python: return self.add_paragraph(text, style)
        end

        function p = add_page_break(obj) %#ok<MANU,STOUT>
            % ADD_PAGE_BREAK STUB (document.py 103-107). Owner: post-P4 content
            %   follow-up (out of P4-7b's named scope: add_heading + add_paragraph).
            %   Faithful body: paragraph = self.add_paragraph();
            %   paragraph.add_run().add_break(WD_BREAK.PAGE); return paragraph. All
            %   deps are now LIVE (add_paragraph P4-7b, add_run + add_break P4-5b);
            %   left stubbed only to stay within P4-7b's explicit scope. Clean
            %   un-stub candidate for the next content WP. See audit VERIFY note.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.document.Document.add_page_break (owning WP: post-P4 " + ...
                "content follow-up; deps live) is not yet ported");
        end

        function p = add_paragraph(obj, text, style)
            % ADD_PARAGRAPH Return a paragraph newly added to the end of the
            %   document (document.py 109-119), populated with `text` and given
            %   paragraph style `style`. Delegates to the body BlockItemContainer.
            %   UN-STUBBED at P4-7b.
            %
            %   Python: return self._body.add_paragraph(text, style).
            %   H13 default fidelity: add_paragraph(text="", style=None) -> text="",
            %   style=[]. Both defaults are forwarded unchanged to _body.
            %
            %   Ported from python-docx v1.2.0: src/docx/document.py::Document.add_paragraph
            arguments
                obj
                text  = ""   % Python default ""
                style = []   % Python default None
            end
            p = obj.body_().add_paragraph(text, style);   % Python: self._body.add_paragraph(text, style)
        end

        function shape = add_picture(obj, image_path_or_stream, width, height) %#ok<INUSD,MANU,STOUT>
            % ADD_PICTURE STUB (document.py 121-138). P7 image tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.document.Document.add_picture (owning WP: P7 image " + ...
                "tier) is not yet ported");
        end

        function section = add_section(obj, start_type)
            % ADD_SECTION Return a Section newly added at the end of the document
            %   (document.py 140-148). UN-STUBBED at P5-3a (C1). `start_type` must
            %   be a WD_SECTION (== WD_SECTION_START) member and defaults to
            %   WD_SECTION.NEW_PAGE.
            %
            %   Python:
            %     new_sectPr = self._element.body.add_section_break()
            %     new_sectPr.start_type = start_type
            %     return Section(new_sectPr, self._part)
            %
            %   H13 default fidelity: add_section(start_type=WD_SECTION.NEW_PAGE).
            %   add_section_break (CT_Body, LIVE P5-3a) clones the prior sentinel
            %   sectPr into a new trailing paragraph and returns the sentinel; its
            %   start_type is then set, and the Section proxy wraps it.
            %
            %   Ported from python-docx v1.2.0: src/docx/document.py::Document.add_section
            arguments
                obj
                start_type = mat2doc.enum.section.WD_SECTION.NEW_PAGE   % Python default
            end
            new_sectPr = obj.element_.body.add_section_break();   % Python: self._element.body.add_section_break()
            new_sectPr.start_type = start_type;                   % Python: new_sectPr.start_type = start_type
            section = mat2doc.section.Section(new_sectPr, obj.part_);  % Python: return Section(new_sectPr, self._part)
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
            % ITER_INNER_CONTENT STUB (document.py 180-182). Owner: P6 table tier.
            %   Delegates to self._body.iter_inner_content, itself stubbed at the
            %   Table boundary (Paragraph is now live P4-7b; Table is P6).
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.table.Table (owning WP: P6 table tier) required by " + ...
                "mat2doc.document.Document.iter_inner_content");
        end

        function p = paragraphs(obj) %#ok<MANU,STOUT>
            % PARAGRAPHS STUB (document.py 184-191). Owner: post-P4 content
            %   follow-up (out of P4-7b's named scope). Faithful body:
            %   return self._body.paragraphs -- and _body.paragraphs is now LIVE
            %   (P4-7b). Left stubbed only to stay within P4-7b's explicit scope
            %   (Document un-stubs = add_heading + add_paragraph). Clean un-stub
            %   candidate; the paragraphs list is reachable now via body_().paragraphs.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.document.Document.paragraphs (owning WP: post-P4 content " + ...
                "follow-up; deps live via _body.paragraphs) is not yet ported");
        end

        function s = sections(obj)
            % SECTIONS A Sections object providing access to each section in this
            %   document (document.py 206-209, @property). UN-STUBBED at P5-3a (C1).
            %   Python: return Sections(self._element, self._part).
            %
            %   Ported from python-docx v1.2.0: src/docx/document.py::Document.sections
            s = mat2doc.section.Sections(obj.element_, obj.part_);
        end

        function s = settings(obj)
            % SETTINGS A Settings object providing access to the document-level
            %   settings (document.py 211-214, @property). Python:
            %   return self._part.settings. UN-STUBBED at P5-1.
            %
            %   Ported from python-docx v1.2.0: src/docx/document.py::Document.settings
            s = obj.part_.settings();
        end

        function s = styles(obj)
            % STYLES A Styles object providing access to the styles in this
            %   document (document.py 216-219, @property). Python:
            %   return self._part.styles. UN-STUBBED at P4-7a.
            %
            %   Ported from python-docx v1.2.0: src/docx/document.py::Document.styles
            s = obj.part_.styles();
        end

        function t = tables(obj) %#ok<MANU,STOUT>
            % TABLES STUB (document.py 221-230). Owner: P6 table tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.document.Document.tables (owning WP: P6 table tier) " + ...
                "is not yet ported");
        end
    end
end

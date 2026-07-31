classdef CT_SectPr < mat2doc.oxml.BaseOxmlElement
% CT_SECTPR `w:sectPr` element, the container element for section properties.
%
%   The section-properties root. Registered for w:sectPr (oxml/__init__.py:129);
%   default.docx's word/document.xml carries a <w:sectPr> (with pgSz + pgMar +
%   cols + docGrid children), so this class transits the CENTRAL M1 part on every
%   load. The parse path is byte-neutral (registering CT classes changes only the
%   CLASS of parsed nodes, never their content or order -- P4-6/P5-1 precedent).
%
%   ============================ H11 (child ordering) ============================
%   _tag_seq (section.py 112-133, VERBATIM, 20 tags) is stored as TAG_SEQ. The
%   ZeroOrOne/ZeroOrMore successor slices map Python `_tag_seq[N:]` ->
%   `TAG_SEQ(N+1:end)` (0-based slice start N -> 1-based start N+1, the P5-1
%   lesson). Per-descriptor slice evidence (own-tag 1-based index / successor
%   slice / first successor tag):
%     headerReference ZeroOrMore successors=_tag_seq       -> TAG_SEQ (full)      first "w:footnotePr"
%     footerReference ZeroOrMore successors=_tag_seq       -> TAG_SEQ (full)      first "w:footnotePr"
%     type            ZeroOrOne  successors=_tag_seq[3:]   -> TAG_SEQ(4:end)      first "w:pgSz"
%     pgSz            ZeroOrOne  successors=_tag_seq[4:]   -> TAG_SEQ(5:end)      first "w:pgMar"
%     pgMar           ZeroOrOne  successors=_tag_seq[5:]   -> TAG_SEQ(6:end)      first "w:paperSrc"
%     titlePg         ZeroOrOne  successors=_tag_seq[14:]  -> TAG_SEQ(15:end)     first "w:textDirection"
%   headerReference/footerReference are NOT members of _tag_seq (it begins at
%   w:footnotePr), so their successors span the WHOLE sequence -- a new hdr/ftr
%   ref inserts after any existing refs but before type/pgSz/pgMar/... A wrong
%   slice here inserts a child in the wrong position -> Word repair / byte
%   divergence.
%
%   type/pgSz/pgMar/titlePg own tags sit at 1-based TAG_SEQ indices 3/4/5/14, so
%   each successor slice starts one past its own tag (N+1): 4/5/6/15. Confirmed
%   against TAG_SEQ below.
%
%   ===================== GENERATED DESCRIPTOR FAMILIES ==========================
%   headerReference / footerReference are ZeroOrMore (section.py 134-135). docx
%   ZeroOrMore generates x_lst, _new_x/_insert_x/_add_x (-> new_x_/insert_x_/
%   add_x_) AND a PUBLIC add_x (D-delta-4). Here the PUBLIC add_headerReference /
%   add_footerReference are SHADOWED: xmlchemy _add_to_class skips a generated
%   name when the class already defines it (xmlchemy.py:357), and CT_SectPr
%   defines add_headerReference(type_, rId) / add_footerReference(type_, rId)
%   (section.py 150-168) explicitly. So the public adder is the geometry method;
%   only the PRIVATE add_headerReference_ (the _add_x adder) is generated. The
%   Callable annotations at section.py 107-110 (_add_footerReference,
%   _add_headerReference, _remove_titlePg, _remove_type) are type hints only.
%
%   type/pgSz/pgMar/titlePg are ZeroOrOne (section.py 136-147): get.x,
%   get_or_add_x, new_x_, insert_x_, add_x_, remove_x_ (underscore rotation of
%   _new_x/_insert_x/_add_x/_remove_x; get_or_add_x public).
%
%   CHILD-CLASS REGISTRATION (this WP): w:type -> CT_SectType, w:pgSz ->
%   CT_PageSz, w:pgMar -> CT_PageMar, w:titlePg -> CT_OnOff (deferral from P5-1
%   closed here), w:headerReference/w:footerReference -> CT_HdrFtrRef.
%
%   ===================== H6 / H3 (page geometry accessors) ======================
%   The margin / page-size @property accessors read/write the pgMar/pgSz children.
%   FAITHFUL ASYMMETRY (transcribed verbatim, NOT normalized): five margin
%   setters (bottom_margin, footer, gutter, header, left_margin) accept
%   `int | Length | None` and wrap a bare int with Length(value)
%   (section.py 182/210/242/259/282); the other four (top_margin, right_margin,
%   page_height, page_width) accept `Length | None` and assign the value directly
%   (section.py 314/330/372/420). All are H6/H3: [] (None) round-trips, ints are
%   exact EMU held in doubles.
%
%   ===================== H10 / H4 (orientation, start_type) =====================
%   orientation getter: PORTRAIT when pgSz absent, else pgSz.orient. Setter:
%   Python `pgSz.orient = value if value else WD_ORIENTATION.PORTRAIT`. H4:
%   BaseXmlEnum subclasses `int`, and PORTRAIT's int value is 0, so `if value` is
%   FALSY for BOTH None AND PORTRAIT. Both fall through to PORTRAIT; because the
%   else-branch (value) also yields PORTRAIT when value==PORTRAIT, the observable
%   result is identical, but the truthiness is expanded faithfully (isFalsy =
%   None OR double(value.value)==0).
%   start_type getter: NEW_PAGE when type child or its @val is absent, else
%   type.val. Setter: Python `if value is None or value is WD_SECTION_START
%   .NEW_PAGE` -- IDENTITY (isequal on singleton members), removes the type child;
%   else get_or_add_type().val = value.
%
%   ===================== H3 (titlePg_val) =======================================
%   titlePg_val getter: False when titlePg absent, else titlePg.val. Setter:
%   Python `if value in [None, False]` -- MEMBERSHIP uses `==` (contrast
%   CT_Settings.evenAndOddHeaders_val which uses IDENTITY `is None or is False`),
%   so None, False, AND any x==False (e.g. 0) remove the child; else
%   get_or_add_titlePg().val = True (CT_OnOff.val=True equals its default ->
%   removes @val -> emits <w:titlePg/>, D-delta-1). Ported as
%   isequal(value,[]) || isequal(value,false) (isequal(0,false) is true, so the
%   numeric-0 breadth of Python's `in [None, False]` is reproduced).
%
%   H5 (identity): get_headerReference/get_footerReference return the LIVE child
%   handle (xpath, first match); remove_* detach it by handle identity.
%
%   iter_inner_content (P5-2b, LIVE): delegates to
%   mat2doc.oxml.section.SectBlockElementIterator_.iter_sect_block_elements --
%   tag-based, so no CT_Tbl dependency (un-stubbed fully). CT_HdrFtr (w:hdr/w:ftr
%   roots) also ported at P5-2b.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on document.xml's <w:sectPr> on every M1 load.
%
%   Example:
%       sp = mat2doc.oxml.OxmlElement("w:sectPr");   % a CT_SectPr
%       sp.page_width = mat2doc.shared.Twips(12240); % -> <w:pgSz w:w="12240"/>
%       sp.start_type                                % NEW_PAGE (no w:type child)
%
%   Ported from python-docx v1.2.0: src/docx/oxml/section.py::CT_SectPr
%   (lines 100-423; registered for w:sectPr)

    properties (Constant, Hidden)  % _tag_seq VERBATIM (section.py 112-133; 20 tags)
        TAG_SEQ = [ ...
            "w:footnotePr", "w:endnotePr", "w:type", ...          %  1- 3  (w:type own @3)
            "w:pgSz", "w:pgMar", "w:paperSrc", ...                %  4- 6  (pgSz@4, pgMar@5)
            "w:pgBorders", "w:lnNumType", "w:pgNumType", ...      %  7- 9
            "w:cols", "w:formProt", "w:vAlign", ...               % 10-12
            "w:noEndnote", "w:titlePg", "w:textDirection", ...    % 13-15  (titlePg own @14)
            "w:bidi", "w:rtlGutter", "w:docGrid", ...             % 16-18
            "w:printerSettings", "w:sectPrChange" ]               % 19-20
    end

    properties (Dependent)  % generated descriptors + @property members
        headerReference_lst  % ZeroOrMore list of <w:headerReference> children (doc order)
        footerReference_lst  % ZeroOrMore list of <w:footerReference> children (doc order)
        type                 % ZeroOrOne <w:type> child or [] (read-only; use get_or_add/remove)
        pgSz                 % ZeroOrOne <w:pgSz> child or []
        pgMar                % ZeroOrOne <w:pgMar> child or []
        titlePg              % ZeroOrOne <w:titlePg> child or []
        bottom_margin        % @w:bottom of w:pgMar as Length, or [] (element/attr absent)
        footer               % @w:footer of w:pgMar as Length, or []
        gutter               % @w:gutter of w:pgMar as Length, or []
        header               % @w:header of w:pgMar as Length, or []
        left_margin          % @w:left of w:pgMar as Length, or []
        right_margin         % @w:right of w:pgMar as Length, or []
        top_margin           % @w:top of w:pgMar as Length, or []
        orientation          % WD_ORIENTATION (PORTRAIT when pgSz absent)
        page_height          % @w:h of w:pgSz as Length, or []
        page_width           % @w:w of w:pgSz as Length, or []
        start_type           % WD_SECTION_START (NEW_PAGE when type/@val absent)
        titlePg_val          % bool from w:titlePg presence (False if absent)
        preceding_sectPr     % sectPr immediately preceding this one, or []
    end

    methods
        function obj = CT_SectPr(varargin)
            % CT_SECTPR Construct a loose <w:sectPr> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ============ ZeroOrMore headerReference (successors=_tag_seq -> TAG_SEQ full) ============
        function lst = get.headerReference_lst(obj);          lst = obj.getChildList("w:headerReference"); end
        function child = new_headerReference_(obj);           child = obj.newChild("w:headerReference"); end
        function child = insert_headerReference_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ); end
        function child = add_headerReference_(obj, varargin); child = obj.addChild("w:headerReference", obj.TAG_SEQ, varargin{:}); end

        % ============ ZeroOrMore footerReference (successors=_tag_seq -> TAG_SEQ full) ============
        function lst = get.footerReference_lst(obj);          lst = obj.getChildList("w:footerReference"); end
        function child = new_footerReference_(obj);           child = obj.newChild("w:footerReference"); end
        function child = insert_footerReference_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ); end
        function child = add_footerReference_(obj, varargin); child = obj.addChild("w:footerReference", obj.TAG_SEQ, varargin{:}); end

        % ============ ZeroOrOne type (successors=_tag_seq[3:] -> TAG_SEQ(4:end)) ============
        function child = get.type(obj);            child = obj.getChild("w:type"); end
        function child = get_or_add_type(obj);     child = obj.getOrAddChild("w:type", obj.TAG_SEQ(4:end)); end
        function child = new_type_(obj);           child = obj.newChild("w:type"); end
        function child = insert_type_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(4:end)); end
        function child = add_type_(obj, varargin); child = obj.addChild("w:type", obj.TAG_SEQ(4:end), varargin{:}); end
        function remove_type_(obj);                obj.removeChild("w:type"); end

        % ============ ZeroOrOne pgSz (successors=_tag_seq[4:] -> TAG_SEQ(5:end)) ============
        function child = get.pgSz(obj);            child = obj.getChild("w:pgSz"); end
        function child = get_or_add_pgSz(obj);     child = obj.getOrAddChild("w:pgSz", obj.TAG_SEQ(5:end)); end
        function child = new_pgSz_(obj);           child = obj.newChild("w:pgSz"); end
        function child = insert_pgSz_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(5:end)); end
        function child = add_pgSz_(obj, varargin); child = obj.addChild("w:pgSz", obj.TAG_SEQ(5:end), varargin{:}); end
        function remove_pgSz_(obj);                obj.removeChild("w:pgSz"); end

        % ============ ZeroOrOne pgMar (successors=_tag_seq[5:] -> TAG_SEQ(6:end)) ============
        function child = get.pgMar(obj);            child = obj.getChild("w:pgMar"); end
        function child = get_or_add_pgMar(obj);     child = obj.getOrAddChild("w:pgMar", obj.TAG_SEQ(6:end)); end
        function child = new_pgMar_(obj);           child = obj.newChild("w:pgMar"); end
        function child = insert_pgMar_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(6:end)); end
        function child = add_pgMar_(obj, varargin); child = obj.addChild("w:pgMar", obj.TAG_SEQ(6:end), varargin{:}); end
        function remove_pgMar_(obj);                obj.removeChild("w:pgMar"); end

        % ============ ZeroOrOne titlePg (successors=_tag_seq[14:] -> TAG_SEQ(15:end)) ============
        function child = get.titlePg(obj);            child = obj.getChild("w:titlePg"); end
        function child = get_or_add_titlePg(obj);     child = obj.getOrAddChild("w:titlePg", obj.TAG_SEQ(15:end)); end
        function child = new_titlePg_(obj);           child = obj.newChild("w:titlePg"); end
        function child = insert_titlePg_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(15:end)); end
        function child = add_titlePg_(obj, varargin); child = obj.addChild("w:titlePg", obj.TAG_SEQ(15:end), varargin{:}); end
        function remove_titlePg_(obj);                obj.removeChild("w:titlePg"); end

        % ===================== methods (section.py 150-423) =====================

        function footerReference = add_footerReference(obj, type_, rId)
            % ADD_FOOTERREFERENCE Newly added `w:footerReference` of `type_` with `rId`.
            %   Ported from python-docx v1.2.0: section.py CT_SectPr.add_footerReference
            %   (lines 150-158).
            footerReference = obj.add_footerReference_();
            footerReference.type_ = type_;
            footerReference.rId = rId;
        end

        function headerReference = add_headerReference(obj, type_, rId)
            % ADD_HEADERREFERENCE Newly added `w:headerReference` of `type_` with `rId`.
            %   Ported from python-docx v1.2.0: section.py CT_SectPr.add_headerReference
            %   (lines 160-168).
            headerReference = obj.add_headerReference_();
            headerReference.type_ = type_;
            headerReference.rId = rId;
        end

        function value = get.bottom_margin(obj)
            % Python (section.py 176-179): pgMar = self.pgMar; if None return None; pgMar.bottom
            pgMar = obj.pgMar;
            if isequal(pgMar, [])
                value = [];
                return
            end
            value = pgMar.bottom;
        end
        function set.bottom_margin(obj, value)
            % Python (section.py 182-184): pgMar.bottom =
            %   value if value is None or isinstance(value, Length) else Length(value)
            pgMar = obj.get_or_add_pgMar();
            pgMar.bottom = mat2doc.oxml.section.CT_SectPr.asLengthOrNone_(value);
        end

        function cloned = clone(obj)
            % CLONE Exact duplicate of this <w:sectPr> tree, with all root attrs removed.
            %   Ported from python-docx v1.2.0: section.py CT_SectPr.clone
            %   (lines 186-194): cloned_sectPr = deepcopy(self);
            %   cloned_sectPr.attrib.clear(); return cloned_sectPr. Removes all
            %   rsid* (and any other) attributes from the root <w:sectPr>, suitable
            %   for a section break. attrib.clear() clears ATTRIBUTES only (lxml
            %   keeps namespace declarations separate), so nsdecls are untouched.
            cloned = obj.deepcopy();
            names = cloned.attrib_names();   % value copy; safe to iterate while removing
            for i = 1:numel(names)
                cloned.remove_attrib(names(i));
            end
        end

        function value = get.footer(obj)
            % Python (section.py 204-207): pgMar; if None None; pgMar.footer
            pgMar = obj.pgMar;
            if isequal(pgMar, [])
                value = [];
                return
            end
            value = pgMar.footer;
        end
        function set.footer(obj, value)
            % Python (section.py 210-212): pgMar.footer =
            %   value if value is None or isinstance(value, Length) else Length(value)
            pgMar = obj.get_or_add_pgMar();
            pgMar.footer = mat2doc.oxml.section.CT_SectPr.asLengthOrNone_(value);
        end

        function value = get_footerReference(obj, type_)
            % GET_FOOTERREFERENCE footerReference element of `type_`, or [] if absent.
            %   Ported from python-docx v1.2.0: section.py CT_SectPr.get_footerReference
            %   (lines 214-220): xpath("./w:footerReference[@w:type='%s']" %
            %   WD_HEADER_FOOTER.to_xml(type_)); if not: None; else [0]. H1: res(1);
            %   H3: [].
            xmlval = mat2doc.enum.section.WD_HEADER_FOOTER.to_xml(type_);
            res = obj.xpath("./w:footerReference[@w:type='" + xmlval + "']");
            if isempty(res)          % Python: if not footerReferences
                value = [];
                return
            end
            value = res(1);          % Python: footerReferences[0]
        end

        function value = get_headerReference(obj, type_)
            % GET_HEADERREFERENCE headerReference element of `type_`, or [] if absent.
            %   Ported from python-docx v1.2.0: section.py CT_SectPr.get_headerReference
            %   (lines 222-229): xpath("./w:headerReference[@w:type='%s']" %
            %   WD_HEADER_FOOTER.to_xml(type_)); if len==0: None; else [0]. H1:
            %   res(1); H3: [].
            xmlval = mat2doc.enum.section.WD_HEADER_FOOTER.to_xml(type_);
            res = obj.xpath("./w:headerReference[@w:type='" + xmlval + "']");
            if isempty(res)          % Python: if len(matching_headerReferences) == 0
                value = [];
                return
            end
            value = res(1);          % Python: matching_headerReferences[0]
        end

        function value = get.gutter(obj)
            % Python (section.py 236-239): pgMar; if None None; pgMar.gutter
            pgMar = obj.pgMar;
            if isequal(pgMar, [])
                value = [];
                return
            end
            value = pgMar.gutter;
        end
        function set.gutter(obj, value)
            % Python (section.py 242-244): pgMar.gutter =
            %   value if value is None or isinstance(value, Length) else Length(value)
            pgMar = obj.get_or_add_pgMar();
            pgMar.gutter = mat2doc.oxml.section.CT_SectPr.asLengthOrNone_(value);
        end

        function value = get.header(obj)
            % Python (section.py 253-256): pgMar; if None None; pgMar.header
            pgMar = obj.pgMar;
            if isequal(pgMar, [])
                value = [];
                return
            end
            value = pgMar.header;
        end
        function set.header(obj, value)
            % Python (section.py 259-261): pgMar.header =
            %   value if value is None or isinstance(value, Length) else Length(value)
            pgMar = obj.get_or_add_pgMar();
            pgMar.header = mat2doc.oxml.section.CT_SectPr.asLengthOrNone_(value);
        end

        function items = iter_inner_content(obj)
            % ITER_INNER_CONTENT All <w:p> and <w:tbl> elements in this section,
            %   in document order, as a 1xN heterogeneous XmlElement array (CT_P
            %   for <w:p> + generic XmlElement for <w:tbl>; the iterator is
            %   TAG-BASED, so an unregistered <w:tbl> is INCLUDED, not dropped --
            %   no CT_Tbl dependency, un-stubbed FULLY at P5-2b). Delegates to the
            %   block-item iterator helper. Elements shaded by nesting in a w:ins
            %   or other wrapper are NOT included. H9: the Python generator is
            %   materialized (no mutation during iteration).
            %
            %   Ported from python-docx v1.2.0: src/docx/oxml/section.py::
            %   CT_SectPr.iter_inner_content (263-269):
            %   return _SectBlockElementIterator.iter_sect_block_elements(self)
            items = mat2doc.oxml.section.SectBlockElementIterator_ ...
                .iter_sect_block_elements(obj);
        end

        function value = get.left_margin(obj)
            % Python (section.py 276-279): pgMar; if None None; pgMar.left
            pgMar = obj.pgMar;
            if isequal(pgMar, [])
                value = [];
                return
            end
            value = pgMar.left;
        end
        function set.left_margin(obj, value)
            % Python (section.py 282-284): pgMar.left =
            %   value if value is None or isinstance(value, Length) else Length(value)
            pgMar = obj.get_or_add_pgMar();
            pgMar.left = mat2doc.oxml.section.CT_SectPr.asLengthOrNone_(value);
        end

        function value = get.orientation(obj)
            % Python (section.py 293-296): pgSz = self.pgSz; if None return
            %   WD_ORIENTATION.PORTRAIT; return pgSz.orient
            pgSz = obj.pgSz;
            if isequal(pgSz, [])
                value = mat2doc.enum.section.WD_ORIENTATION.PORTRAIT;
                return
            end
            value = pgSz.orient;
        end
        function set.orientation(obj, value)
            % Python (section.py 299-301): pgSz = self.get_or_add_pgSz();
            %   pgSz.orient = value if value else WD_ORIENTATION.PORTRAIT
            % H4: BaseXmlEnum subclasses int; PORTRAIT's int is 0, so `if value` is
            % falsy for BOTH None AND PORTRAIT. Both branches yield PORTRAIT for
            % those inputs (else-branch assigns value, which IS PORTRAIT), so the
            % result is identical; truthiness is expanded faithfully.
            pgSz = obj.get_or_add_pgSz();
            isFalsy = isequal(value, []) || double(value.value) == 0;
            if isFalsy
                pgSz.orient = mat2doc.enum.section.WD_ORIENTATION.PORTRAIT;
            else
                pgSz.orient = value;
            end
        end

        function value = get.page_height(obj)
            % Python (section.py 309-312): pgSz; if None None; pgSz.h
            pgSz = obj.pgSz;
            if isequal(pgSz, [])
                value = [];
                return
            end
            value = pgSz.h;
        end
        function set.page_height(obj, value)
            % Python (section.py 315-317): pgSz = self.get_or_add_pgSz();
            %   pgSz.h = value   (NO Length() wrap -- typed Length | None)
            pgSz = obj.get_or_add_pgSz();
            pgSz.h = value;
        end

        function value = get.page_width(obj)
            % Python (section.py 325-328): pgSz; if None None; pgSz.w
            pgSz = obj.pgSz;
            if isequal(pgSz, [])
                value = [];
                return
            end
            value = pgSz.w;
        end
        function set.page_width(obj, value)
            % Python (section.py 331-333): pgSz = self.get_or_add_pgSz();
            %   pgSz.w = value   (NO Length() wrap -- typed Length | None)
            pgSz = obj.get_or_add_pgSz();
            pgSz.w = value;
        end

        function value = get.preceding_sectPr(obj)
            % Python (section.py 336-340): xpath("./preceding::w:sectPr[1]");
            %   return [0] if len>0 else None. The [1] predicate is a per-axis
            %   position on the reverse `preceding` axis (nearest first), already
            %   1-based -- never shifted (H1). H3: res(1) or [].
            res = obj.xpath("./preceding::w:sectPr[1]");
            if numel(res) > 0
                value = res(1);
            else
                value = [];
            end
        end

        function rId = remove_footerReference(obj, type_)
            % REMOVE_FOOTERREFERENCE Remove the w:footerReference of `type_`; return its rId.
            %   Ported from python-docx v1.2.0: section.py CT_SectPr.remove_footerReference
            %   (lines 342-350). ValueError (verbatim message) when absent -- a
            %   "should never happen" guard.
            footerReference = obj.get_footerReference(type_);
            if isequal(footerReference, [])
                error("mat2doc:ValueError", "CT_SectPr has no footer reference");
            end
            rId = footerReference.rId;
            obj.remove(footerReference);
        end

        function rId = remove_headerReference(obj, type_)
            % REMOVE_HEADERREFERENCE Remove the w:headerReference of `type_`; return its rId.
            %   Ported from python-docx v1.2.0: section.py CT_SectPr.remove_headerReference
            %   (lines 352-360). ValueError (verbatim message) when absent.
            headerReference = obj.get_headerReference(type_);
            if isequal(headerReference, [])
                error("mat2doc:ValueError", "CT_SectPr has no header reference");
            end
            rId = headerReference.rId;
            obj.remove(headerReference);
        end

        function value = get.right_margin(obj)
            % Python (section.py 367-370): pgMar; if None None; pgMar.right
            pgMar = obj.pgMar;
            if isequal(pgMar, [])
                value = [];
                return
            end
            value = pgMar.right;
        end
        function set.right_margin(obj, value)
            % Python (section.py 373-375): pgMar = self.get_or_add_pgMar();
            %   pgMar.right = value   (NO Length() wrap -- typed Length | None)
            pgMar = obj.get_or_add_pgMar();
            pgMar.right = value;
        end

        function value = get.start_type(obj)
            % Python (section.py 378-385): type = self.type; if type is None or
            %   type.val is None: return WD_SECTION_START.NEW_PAGE; return type.val
            t = obj.type;
            if isequal(t, []) || isequal(t.val, [])
                value = mat2doc.enum.section.WD_SECTION_START.NEW_PAGE;
                return
            end
            value = t.val;
        end
        function set.start_type(obj, value)
            % Python (section.py 388-393): if value is None or value is
            %   WD_SECTION_START.NEW_PAGE: self._remove_type(); return
            %   type = self.get_or_add_type(); type.val = value
            % H4: `is None or is NEW_PAGE` is IDENTITY -- isequal on singleton
            % members (NEW_PAGE's int is 2, so no int-falsy concern).
            if isequal(value, []) || isequal(value, mat2doc.enum.section.WD_SECTION_START.NEW_PAGE)
                obj.remove_type_();
                return
            end
            t = obj.get_or_add_type();
            t.val = value;
        end

        function value = get.titlePg_val(obj)
            % Python (section.py 397-401): titlePg = self.titlePg; if None return
            %   False; return titlePg.val
            titlePg = obj.titlePg;
            if isequal(titlePg, [])
                value = false;
                return
            end
            value = titlePg.val;
        end
        function set.titlePg_val(obj, value)
            % Python (section.py 404-408): if value in [None, False]:
            %   self._remove_titlePg(); else: self.get_or_add_titlePg().val = True
            % H3: `in [None, False]` uses `==` (contrast CT_Settings' IDENTITY
            % `is None or is False`), so None, False, and any x==False (e.g. 0)
            % remove. isequal(0,false) is true, reproducing the numeric-0 breadth.
            if isequal(value, []) || isequal(value, false)
                obj.remove_titlePg_();
            else
                tp = obj.get_or_add_titlePg();
                tp.val = true;   % Python: .val = True
            end
        end

        function value = get.top_margin(obj)
            % Python (section.py 415-418): pgMar; if None None; pgMar.top
            pgMar = obj.pgMar;
            if isequal(pgMar, [])
                value = [];
                return
            end
            value = pgMar.top;
        end
        function set.top_margin(obj, value)
            % Python (section.py 421-423): pgMar = self.get_or_add_pgMar();
            %   pgMar.top = value   (NO Length() wrap -- typed Length | None)
            pgMar = obj.get_or_add_pgMar();
            pgMar.top = value;
        end
    end

    methods (Static, Access = private)
        function out = asLengthOrNone_(value)
            % ASLENGTHORNONE_ Realizes the five wrapping margin setters' RHS:
            %   value if value is None or isinstance(value, Length) else Length(value)
            %   (section.py 182/210/242/259/282). [] (None) and any Length
            %   (subclass) pass through; anything else is wrapped Length(value).
            %   isa covers Length subclasses (Twips/Emu/...) exactly like
            %   isinstance(value, Length).
            if isequal(value, []) || isa(value, "mat2doc.shared.Length")
                out = value;
            else
                out = mat2doc.shared.Length(value);
            end
        end
    end
end

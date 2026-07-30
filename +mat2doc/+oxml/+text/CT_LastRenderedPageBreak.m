classdef CT_LastRenderedPageBreak < mat2doc.oxml.BaseOxmlElement
% CT_LASTRENDEREDPAGEBREAK `<w:lastRenderedPageBreak>` element, a renderer page break.
%
%   A rendered page-break is one inserted by the renderer when it runs out of room
%   on a page. It is an empty element (no attrs or children) and is a child of CT_R,
%   peer to CT_Text. Registered for <w:lastRenderedPageBreak>
%   (docx/oxml/__init__.py, text block).
%
%   NOTE (pagebreak.py 23-26): this complex-type name does not exist in the schema,
%   where w:lastRenderedPageBreak maps to CT_Empty. This name was added to give the
%   element distinguished behavior (the page-break split machinery below).
%
%   PUBLIC @property ACCESSORS (pagebreak.py) -- ported as no-arg METHODS so
%   display never evaluates them and so the ValueError guards fire only on an
%   explicit call (established Mat2Doc convention; these do xpath/deepcopy and can
%   raise):
%     * following_fragment_p  (28-50) -- loose CT_P with content BEFORE this break
%                                        removed; raises ValueError unless this is
%                                        the FIRST rendered page-break in its p.
%     * follows_all_content   (52-75) -- True when this break is the last "content"
%                                        in the paragraph (uncommon).
%     * precedes_all_content  (77-99) -- True when this break precedes all paragraph
%                                        content (common; break on even p boundary).
%     * preceding_fragment_p  (101-118) -- loose CT_P with content AFTER this break
%                                          removed; same first-break guard.
%
%   ALGORITHM (pagebreak.py) -- xpath-based sibling/descendant traversal, ported
%   faithfully over the +oxml XPath engine (which supports the parent::, ancestor::,
%   following-sibling::, preceding-sibling:: axes, last(), and self:: unions in
%   predicates -- see evaluate_xpath). H1: XPath positions ([1], [last()]) are
%   already 1-based -- never shifted. Element removal / insert / deepcopy are
%   handle-identity operations (H5). The split differs when the break is inside a
%   hyperlink ("atomic", stays with the page it starts on) vs a bare run.
%
%   @lazyproperty MEMBERS (pagebreak.py) -- design.md section 2: private cache
%   value + logical computed-flag (NEVER isempty as the sentinel; [] is a legal
%   value). Underscore rotation: _is_in_hyperlink -> is_in_hyperlink_, etc.
%     * _is_in_hyperlink        (202-205)
%     * _following_frag_in_hlink (143-170)
%     * _following_frag_in_run   (172-200)
%     * _preceding_frag_in_hlink (207-235)
%     * _preceding_frag_in_run   (237-265)
%     * _run_inner_content_xpath (267-278) -- a lazyproperty returning a FIXED
%       literal; ported as a Constant property (RUN_INNER_CONTENT_XPATH). Caching a
%       constant is a constant.
%   Caching is FAITHFUL here (not merely an optimization): each public @property in
%   Python re-runs and reads the cached frag, so after a hypothetical tree mutation
%   both Python and this port return the STALE first fragment. See VERIFY-deepcopy
%   in the audit for the deepcopy-drops-cache interaction (proven unobservable).
%
%   PRIVATE HELPERS (pagebreak.py): _enclosing_hyperlink (120-126, a method taking
%   lrpb), _enclosing_p (128-131, @property), _first_lrpb_in_p (133-141, a method
%   taking p). Underscore rotation applied.
%
%   H3 (None): inline isequal(x, []) where used; the frag/xpath results are typed
%   arrays (empty typed array == falsy), tested via isempty (H4).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on <w:lastRenderedPageBreak> nodes inside document.xml /
%   styles.xml body content.
%
%   Example:
%       % typical: page break at an even paragraph boundary (first run of p)
%       brks = p.xpath("./w:r/w:lastRenderedPageBreak");  % two-step index --
%       lrpb = brks(1);                                    % MATLAB has no f(...)(1)
%       tf = lrpb.precedes_all_content();   % often true
%
%   Ported from python-docx v1.2.0: src/docx/oxml/text/pagebreak.py::CT_LastRenderedPageBreak
%   (lines 16-278; registered for w:lastRenderedPageBreak)

    properties (Constant, Hidden)
        % _run_inner_content_xpath (pagebreak.py 267-278): XPath fragment matching
        % any run inner-content element. A lazyproperty over a fixed literal ->
        % Constant. Verbatim, including the " | " separators.
        RUN_INNER_CONTENT_XPATH = ...
            "self::w:br | self::w:cr | self::w:drawing | self::w:noBreakHyphen" + ...
            " | self::w:ptab | self::w:t | self::w:tab"
    end

    properties (Access = private)
        % @lazyproperty caches (design.md section 2: value property + logical
        % computed flag; NEVER isempty as the sentinel). Defaults: Computed_ =
        % false, Value_ = [] (unread). NOTE: deepcopy does NOT copy this instance
        % state (base XmlElement.deepcopy contract) -- unobservable here because a
        % CLONED lrpb's lazyproperties are never accessed (see VERIFY-deepcopy).
        is_in_hyperlink_Value_
        is_in_hyperlink_Computed_ = false
        following_frag_in_hlink_Value_
        following_frag_in_hlink_Computed_ = false
        following_frag_in_run_Value_
        following_frag_in_run_Computed_ = false
        preceding_frag_in_hlink_Value_
        preceding_frag_in_hlink_Computed_ = false
        preceding_frag_in_run_Value_
        preceding_frag_in_run_Computed_ = false
    end

    methods
        function obj = CT_LastRenderedPageBreak(varargin)
            % CT_LASTRENDEREDPAGEBREAK Construct a loose <w:lastRenderedPageBreak>
            %   -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ================= following_fragment_p (pagebreak.py 28-50) =============
        function p = following_fragment_p(obj)
            % FOLLOWING_FRAGMENT_P A loose CT_P with this break and all content
            %   preceding it removed. Raises ValueError unless this is the first
            %   rendered page-break in its paragraph.
            % Python: if not self == self._first_lrpb_in_p(self._enclosing_p):
            if ~(obj == obj.first_lrpb_in_p_(obj.enclosing_p_()))   % H5 handle identity
                error("mat2doc:ValueError", "%s", ...
                    "only defined on first rendered page-break in paragraph");
            end
            % -- splitting approach differs when the break is inside a hyperlink --
            if obj.is_in_hyperlink_()
                p = obj.following_frag_in_hlink_();
            else
                p = obj.following_frag_in_run_();
            end
        end

        % ================= follows_all_content (pagebreak.py 52-75) ==============
        function tf = follows_all_content(obj)
            % FOLLOWS_ALL_CONTENT True when this page-break is the last "content" in
            %   the paragraph (very uncommon; only in hand-edited XML).
            % -- a page-break inside a hyperlink never meets these criteria (it is
            % -- "atomic" and associated with the page it starts on).
            if obj.is_in_hyperlink_()
                tf = false;
                return
            end
            % XPath matches zero-or-one w:lastRenderedPageBreak: in the LAST run of
            % the paragraph, not followed by any content-bearing element.
            expr = "(./w:r)[last()]/w:lastRenderedPageBreak" + ...
                "[not(following-sibling::*[" + obj.RUN_INNER_CONTENT_XPATH + "])]";
            tf = ~isempty(obj.enclosing_p_().xpath(expr));   % Python bool(node-set)
        end

        % ================= precedes_all_content (pagebreak.py 77-99) =============
        function tf = precedes_all_content(obj)
            % PRECEDES_ALL_CONTENT True when a w:lastRenderedPageBreak precedes all
            %   paragraph content (common; page breaks on an even p boundary).
            % -- a page-break inside a hyperlink never meets these criteria because
            % -- there is always part of the hyperlink text before the page-break.
            if obj.is_in_hyperlink_()
                tf = false;
                return
            end
            % XPath matches zero-or-one w:lastRenderedPageBreak: in the FIRST run of
            % the paragraph, not preceded by any content-bearing element.
            expr = "./w:r[1]/w:lastRenderedPageBreak" + ...
                "[not(preceding-sibling::*[" + obj.RUN_INNER_CONTENT_XPATH + "])]";
            tf = ~isempty(obj.enclosing_p_().xpath(expr));   % Python bool(node-set)
        end

        % ================= preceding_fragment_p (pagebreak.py 101-118) ===========
        function p = preceding_fragment_p(obj)
            % PRECEDING_FRAGMENT_P A loose CT_P with this break and all its
            %   FOLLOWING siblings removed. Raises ValueError unless this is the
            %   first rendered page-break in its paragraph.
            % Python: if not self == self._first_lrpb_in_p(self._enclosing_p):
            if ~(obj == obj.first_lrpb_in_p_(obj.enclosing_p_()))   % H5 handle identity
                error("mat2doc:ValueError", "%s", ...
                    "only defined on first rendered page-break in paragraph");
            end
            % -- splitting approach differs when the break is inside a hyperlink --
            if obj.is_in_hyperlink_()
                p = obj.preceding_frag_in_hlink_();
            else
                p = obj.preceding_frag_in_run_();
            end
        end
    end

    methods (Access = private)
        % ================= _enclosing_hyperlink (pagebreak.py 120-126) ===========
        function hyperlink = enclosing_hyperlink_(~, lrpb)
            % ENCLOSING_HYPERLINK_ The w:hyperlink grandparent of `lrpb`. Raises
            %   when the page-break has a w:p grandparent, so only call when
            %   is_in_hyperlink_ is True (partial function; Python `[0]` on an empty
            %   node-set raises IndexError -- here indexing an empty typed array
            %   raises MATLAB:badsubscript, the faithful never-reached analogue).
            %   `obj` unused (Python self unused; the method is called on lrpb and
            %   passed lrpb).
            res = lrpb.xpath("./parent::w:r/parent::w:hyperlink");
            hyperlink = res(1);   % Python res[0] -> 1-based (H1)
        end

        % ================= _enclosing_p (pagebreak.py 128-131) ===================
        function p = enclosing_p_(obj)
            % ENCLOSING_P_ The w:p parent or grandparent of this break.
            res = obj.xpath("./ancestor::w:p[1]");
            p = res(1);   % Python res[0] -> 1-based (H1)
        end

        % ================= _first_lrpb_in_p (pagebreak.py 133-141) ===============
        function lrpb = first_lrpb_in_p_(~, p)
            % FIRST_LRPB_IN_P_ The first w:lastRenderedPageBreak element in `p`.
            %   Raises ValueError if there are none. `obj` unused (Python self
            %   unused; called on lrpb-or-self, operates purely on `p`).
            lrpbs = p.xpath( ...
                "./w:r/w:lastRenderedPageBreak | ./w:hyperlink/w:r/w:lastRenderedPageBreak");
            if isempty(lrpbs)   % Python: if not lrpbs (empty node-set is falsy, H4)
                error("mat2doc:ValueError", "%s", ...
                    "no rendered page-breaks in paragraph element");
            end
            lrpb = lrpbs(1);   % Python lrpbs[0] -> 1-based (H1)
        end

        % ================= _following_frag_in_hlink (pagebreak.py 143-170) =======
        function p = following_frag_in_hlink_(obj)
            % FOLLOWING_FRAG_IN_HLINK_ Following CT_P fragment when the break occurs
            %   within a hyperlink. Partial function: raises when NOT in a hyperlink.
            %   @lazyproperty -> cached (design.md section 2).
            if obj.following_frag_in_hlink_Computed_
                p = obj.following_frag_in_hlink_Value_;
                return
            end
            if ~obj.is_in_hyperlink_()
                error("mat2doc:ValueError", "%s", ...
                    "only defined on a rendered page-break in a hyperlink");
            end
            % -- work on a clone w:p so our mutations don't persist --
            p = obj.enclosing_p_().deepcopy();
            % -- get this w:lastRenderedPageBreak in the cloned w:p (not self) --
            lrpb = obj.first_lrpb_in_p_(p);
            % -- locate the w:hyperlink in which this break is found --
            hyperlink = lrpb.enclosing_hyperlink_(lrpb);
            % -- delete all w:p inner-content preceding the hyperlink (not w:pPr) --
            elms = hyperlink.xpath("./preceding-sibling::*[not(self::w:pPr)]");
            for k = 1:numel(elms)
                p.remove(elms(k));
            end
            % -- remove the whole hyperlink; it belongs to the preceding fragment --
            hyperlink.getparent().remove(hyperlink);
            obj.following_frag_in_hlink_Value_ = p;
            obj.following_frag_in_hlink_Computed_ = true;
        end

        % ================= _following_frag_in_run (pagebreak.py 172-200) =========
        function p = following_frag_in_run_(obj)
            % FOLLOWING_FRAG_IN_RUN_ Following CT_P fragment when the break does NOT
            %   occur in a hyperlink. Partial function: raises when in a hyperlink.
            %   @lazyproperty -> cached.
            if obj.following_frag_in_run_Computed_
                p = obj.following_frag_in_run_Value_;
                return
            end
            if obj.is_in_hyperlink_()
                error("mat2doc:ValueError", "%s", ...
                    "only defined on a rendered page-break not in a hyperlink");
            end
            % -- work on a clone w:p so our mutations don't persist --
            p = obj.enclosing_p_().deepcopy();
            lrpb = obj.first_lrpb_in_p_(p);
            % -- locate the w:r in which this break is found --
            r_res = lrpb.xpath("./parent::w:r");
            enclosing_r = r_res(1);   % Python [0] -> 1-based (H1)
            % -- delete all w:p inner-content preceding that run (but not w:pPr) --
            elms = enclosing_r.xpath("./preceding-sibling::*[not(self::w:pPr)]");
            for k = 1:numel(elms)
                p.remove(elms(k));
            end
            % -- then remove all run inner-content preceding this break in its run
            % -- (but not the w:rPr) and remove the page-break itself --
            elms2 = lrpb.xpath("./preceding-sibling::*[not(self::w:rPr)]");
            for k = 1:numel(elms2)
                enclosing_r.remove(elms2(k));
            end
            enclosing_r.remove(lrpb);
            obj.following_frag_in_run_Value_ = p;
            obj.following_frag_in_run_Computed_ = true;
        end

        % ================= _is_in_hyperlink (pagebreak.py 202-205) ===============
        function tf = is_in_hyperlink_(obj)
            % IS_IN_HYPERLINK_ True when this break is embedded in a hyperlink run.
            %   @lazyproperty -> cached.
            if obj.is_in_hyperlink_Computed_
                tf = obj.is_in_hyperlink_Value_;
                return
            end
            % Python: return bool(self.xpath("./parent::w:r/parent::w:hyperlink"))
            tf = ~isempty(obj.xpath("./parent::w:r/parent::w:hyperlink"));
            obj.is_in_hyperlink_Value_ = tf;
            obj.is_in_hyperlink_Computed_ = true;
        end

        % ================= _preceding_frag_in_hlink (pagebreak.py 207-235) =======
        function p = preceding_frag_in_hlink_(obj)
            % PRECEDING_FRAG_IN_HLINK_ Preceding CT_P fragment when the break occurs
            %   within a hyperlink. Partial function: raises when NOT in a hyperlink.
            %   @lazyproperty -> cached.
            if obj.preceding_frag_in_hlink_Computed_
                p = obj.preceding_frag_in_hlink_Value_;
                return
            end
            if ~obj.is_in_hyperlink_()
                error("mat2doc:ValueError", "%s", ...
                    "only defined on a rendered page-break in a hyperlink");
            end
            % -- work on a clone w:p so our mutations don't persist --
            p = obj.enclosing_p_().deepcopy();
            lrpb = obj.first_lrpb_in_p_(p);
            hyperlink = lrpb.enclosing_hyperlink_(lrpb);
            % -- delete all w:p inner-content following the hyperlink --
            elms = hyperlink.xpath("./following-sibling::*");
            for k = 1:numel(elms)
                p.remove(elms(k));
            end
            % -- remove this page-break from inside the hyperlink; the entire
            % -- hyperlink goes into the preceding fragment (not "split") --
            lrpb.getparent().remove(lrpb);
            obj.preceding_frag_in_hlink_Value_ = p;
            obj.preceding_frag_in_hlink_Computed_ = true;
        end

        % ================= _preceding_frag_in_run (pagebreak.py 237-265) =========
        function p = preceding_frag_in_run_(obj)
            % PRECEDING_FRAG_IN_RUN_ Preceding CT_P fragment when the break does NOT
            %   occur in a hyperlink. Partial function: raises when in a hyperlink.
            %   @lazyproperty -> cached.
            if obj.preceding_frag_in_run_Computed_
                p = obj.preceding_frag_in_run_Value_;
                return
            end
            if obj.is_in_hyperlink_()
                error("mat2doc:ValueError", "%s", ...
                    "only defined on a rendered page-break not in a hyperlink");
            end
            % -- work on a clone w:p so our mutations don't persist --
            p = obj.enclosing_p_().deepcopy();
            lrpb = obj.first_lrpb_in_p_(p);
            r_res = lrpb.xpath("./parent::w:r");
            enclosing_r = r_res(1);   % Python [0] -> 1-based (H1)
            % -- delete all w:p inner-content following that run --
            elms = enclosing_r.xpath("./following-sibling::*");
            for k = 1:numel(elms)
                p.remove(elms(k));
            end
            % -- then delete all w:r inner-content following this break in its run
            % -- and remove the page-break itself --
            elms2 = lrpb.xpath("./following-sibling::*");
            for k = 1:numel(elms2)
                enclosing_r.remove(elms2(k));
            end
            enclosing_r.remove(lrpb);
            obj.preceding_frag_in_run_Value_ = p;
            obj.preceding_frag_in_run_Computed_ = true;
        end
    end
end

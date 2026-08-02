classdef CT_R < mat2doc.oxml.BaseOxmlElement
% CT_R `<w:r>` element: the properties and text for a run.
%
%   THE M2 BYTE-CRITICAL run element (document.xml body content is w:p/w:r/w:t).
%   Registered for <w:r> (docx/oxml/__init__.py:77).
%
%   DESCRIPTORS (run.py 33-38):
%     rPr     = ZeroOrOne("w:rPr")   -- NO successors (successors=()); an added
%               rPr is forced to the FRONT via the _insert_rPr OVERRIDE below.
%     br      = ZeroOrMore("w:br")   -- successors=() -> APPEND
%     cr      = ZeroOrMore("w:cr")   -- successors=() -> APPEND
%     drawing = ZeroOrMore("w:drawing") -- successors=() -> APPEND
%     t       = ZeroOrMore("w:t")    -- successors=() -> APPEND
%     tab     = ZeroOrMore("w:tab")  -- successors=() -> APPEND
%
%   H11 CHILD ORDERING: rPr FIRST, then run inner-content in INSERTION order.
%   run.py OVERRIDES `_insert_rPr` (run.py 145-147: `self.insert(0, rPr)`) so
%   get_or_add_rPr / _add_rPr always place rPr at index 0 (MATLAB insert(1,...),
%   H1). Every content descriptor has successors=() (NO_SUCCESSORS), so
%   insertChildInSequence APPENDS -- content stays in the order added. (Verified
%   vs source: rPr has no successors arg; the front placement is the explicit
%   inserter override, NOT a successor list.)
%
%   xmlchemy member generation (xmlchemy.py, docx form):
%     ZeroOrOne (rPr): get.rPr, get_or_add_rPr, new_rPr_, insert_rPr_ (OVERRIDE),
%                      add_rPr_, remove_rPr_ (xmlchemy 540-576).
%     ZeroOrMore (br/cr/drawing/t/tab): x_lst, new_x_, insert_x_, add_x_,
%                      add_x (PUBLIC) -- NO bare `x` getter (delattr), NO
%                      get_or_add, NO remover (xmlchemy 526-537).
%   Underscore rotation (design.md section 2): _new_x -> new_x_, _insert_x ->
%   insert_x_, _add_x -> add_x_ (the PUBLIC add_x keeps its bare name).
%
%   EXPLICIT (non-generated) public members that WIN over the generated adder
%   because xmlchemy._add_to_class skips a name already defined on the class
%   (xmlchemy 354-359):
%     * add_t(text)      -- run.py 40-45; the class-body method survives, so the
%                           ZeroOrMore no-arg public add_t is NOT generated.
%     * add_drawing(...) -- run.py 47-54; likewise.
%     * insert_rPr_      -- run.py 145-147; the OVERRIDE inserter wins.
%   The pyright Callable annotations on run.py 27-31 (add_br/add_tab/
%   get_or_add_rPr/_add_drawing/_add_t) are type hints only -- those members ARE
%   generated (the annotation is not a definition).
%
%   add_t BYTE-CRITICAL (run.py 40-45): adds a <w:t> whose text is `text`, then
%   sets @xml:space="preserve" WHEN `len(text.strip()) < len(text)` (i.e. text
%   has leading or trailing whitespace). Without it Word strips those spaces.
%   `" hello "` -> `<w:t xml:space="preserve"> hello </w:t>`; `"hello"` -> no
%   xml:space. The attribute key is qn("xml:space") (the reserved xml namespace,
%   never xmlns-declared).
%
%   TEXT SHADOW (D10): CT_R.text (run.py 129-143) SHADOWS the lxml `.text`
%   attribute -- it is the run's concatenated inner-content text, not the
%   element char data. MATLAB forbids redefining a superclass property, so this
%   is ported by OVERRIDING the protected getText_/setText_ (the serializer reads
%   text_raw_, bypassing the override, so the w:r's own char data -- normally
%   None -- serializes unchanged). The getter joins str_() of each
%   w:br|w:cr|w:noBreakHyphen|w:ptab|w:t|w:tab child; the setter clears content
%   then appends via RunContentAppender_.
%
%   NOT-YET-PORTED dependencies (VERIFY, see audit):
%     * w:tab -> CT_TabStop (parfmt WP): add_tab() CREATES a generic <w:tab/>
%       correctly (write path OK), but str_() over a w:tab in the `text` getter
%       requires CT_TabStop.str_() ("\t") -- absent until parfmt lands, so a
%       `.text` read over a run containing a w:tab errors until then.
%     * w:drawing -> CT_Drawing (P7-3), w:lastRenderedPageBreak ->
%       CT_LastRenderedPageBreak (P4-3): inner_content_items dispatches on these
%       classes (isinstance) -- now LIVE at P8-3 (both registered).
%   These are out of the M2 write path (byte-critical add_t / _RunContentAppender
%   are LIVE and complete).
%
%   H3 (None): inline isequal(x, []) (established Mat2Doc oxml convention).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on the many <w:r> nodes inside document.xml.
%
%   Example:
%       r = mat2doc.oxml.OxmlElement("w:r");
%       r.text = "a\tb";            % <w:t>a</w:t><w:tab/><w:t>b</w:t>
%       r.get_or_add_rPr();         % <w:rPr/> forced to the front (H11)
%       t = r.add_t(" hi ");        % <w:t xml:space="preserve"> hi </w:t>
%
%   Ported from python-docx v1.2.0: src/docx/oxml/text/run.py::CT_R
%   (lines 24-164; registered for w:r)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        NO_SUCCESSORS = string.empty(1, 0)  % rPr + all content: successors=() -> append
        % PY_STRIP_WS: the EXACT set of code points Python str.strip() removes
        % (CPython 3.13 str.isspace set, enumerated mechanically over all of
        % Unicode by the Gate-2 audit, 2026-07-26: 29 BMP code points). MATLAB
        % strip() DIVERGES on U+0085/U+00A0/U+2007/U+202F (it keeps the
        % non-breaking spaces), which flipped the add_t xml:space="preserve"
        % decision vs python-docx for NBSP/NEL-boundary text -- a real byte
        % divergence caught by audit probe xmlspace_10/11/13 (Gate-2 fix).
        % All 29 are BMP, so a UTF-16 code-unit test is surrogate-safe (H2).
        PY_STRIP_WS = char([9:13, 28:31, 32, 133, 160, 5760, 8192:8202, ...
            8232, 8233, 8239, 8287, 12288])
        RPR_TAG     = "w:rPr"      % ZeroOrOne  @ run.py:33
        BR_TAG      = "w:br"       % ZeroOrMore @ run.py:34
        CR_TAG      = "w:cr"       % ZeroOrMore @ run.py:35
        DRAWING_TAG = "w:drawing"  % ZeroOrMore @ run.py:36
        T_TAG       = "w:t"        % ZeroOrMore @ run.py:37
        TAB_TAG     = "w:tab"      % ZeroOrMore @ run.py:38
    end

    properties (Dependent)  % ZeroOrOne getter + ZeroOrMore list getters + @property
        rPr           % <w:rPr> child or [] (None) if absent
        br_lst        % list of <w:br> children (document order)
        cr_lst        % list of <w:cr> children (document order)
        drawing_lst   % list of <w:drawing> children (document order)
        t_lst         % list of <w:t> children (document order)
        tab_lst       % list of <w:tab> children (document order)
        style         % ./w:rStyle/@val (via rPr.style) or [] (None)
    end

    methods
        function obj = CT_R(varargin)
            % CT_R Construct a loose <w:r> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ============ rPr (ZeroOrOne, successors=(); FRONT via override) ============
        function child = get.rPr(obj);            child = obj.getChild(obj.RPR_TAG); end
        function child = get_or_add_rPr(obj)
            % Python get_or_add_rPr (xmlchemy 557-562), routed through the
            % overridden inserter (front placement).
            child = obj.rPr;
            if isequal(child, [])   % Python: if child is None (H3)
                child = obj.add_rPr_();
            end
        end
        function child = new_rPr_(obj);           child = obj.newChild(obj.RPR_TAG); end
        function child = insert_rPr_(obj, rPr)
            % OVERRIDE (run.py 145-147): self.insert(0, rPr).
            obj.insert(1, rPr);    % Python insert(0,...) -> 1-based (H1)
            child = rPr;
        end
        function child = add_rPr_(obj, varargin)
            % _add_rPr (xmlchemy 284-291) routed through the OVERRIDE inserter.
            child = obj.new_rPr_();
            for k = 1:2:numel(varargin)
                child.(varargin{k}) = varargin{k + 1};
            end
            child = obj.insert_rPr_(child);
        end
        function remove_rPr_(obj);                obj.removeChild(obj.RPR_TAG); end

        % ============ br (ZeroOrMore, successors=() -> append) ============
        function lst = get.br_lst(obj);           lst = obj.getChildList(obj.BR_TAG); end
        function child = new_br_(obj);            child = obj.newChild(obj.BR_TAG); end
        function child = insert_br_(obj, child);  child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_br_(obj, varargin);  child = obj.addChild(obj.BR_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function child = add_br(obj);             child = obj.add_br_(); end   % public adder (xmlchemy 340-352)

        % ============ cr (ZeroOrMore, successors=() -> append) ============
        function lst = get.cr_lst(obj);           lst = obj.getChildList(obj.CR_TAG); end
        function child = new_cr_(obj);            child = obj.newChild(obj.CR_TAG); end
        function child = insert_cr_(obj, child);  child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_cr_(obj, varargin);  child = obj.addChild(obj.CR_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function child = add_cr(obj);             child = obj.add_cr_(); end   % public adder

        % ============ drawing (ZeroOrMore, successors=() -> append) ============
        %   public add_drawing is EXPLICIT (below); only the private add_drawing_
        %   is the generated adder.
        function lst = get.drawing_lst(obj);           lst = obj.getChildList(obj.DRAWING_TAG); end
        function child = new_drawing_(obj);            child = obj.newChild(obj.DRAWING_TAG); end
        function child = insert_drawing_(obj, child);  child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_drawing_(obj, varargin);  child = obj.addChild(obj.DRAWING_TAG, obj.NO_SUCCESSORS, varargin{:}); end

        % ============ t (ZeroOrMore, successors=() -> append) ============
        %   public add_t is EXPLICIT (below); only the private add_t_ is generated.
        function lst = get.t_lst(obj);            lst = obj.getChildList(obj.T_TAG); end
        function child = new_t_(obj);             child = obj.newChild(obj.T_TAG); end
        function child = insert_t_(obj, child);   child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_t_(obj, varargin);   child = obj.addChild(obj.T_TAG, obj.NO_SUCCESSORS, varargin{:}); end

        % ============ tab (ZeroOrMore, successors=() -> append) ============
        function lst = get.tab_lst(obj);          lst = obj.getChildList(obj.TAB_TAG); end
        function child = new_tab_(obj);           child = obj.newChild(obj.TAB_TAG); end
        function child = insert_tab_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_tab_(obj, varargin); child = obj.addChild(obj.TAB_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function child = add_tab(obj);            child = obj.add_tab_(); end  % public adder

        % ======================= add_t (EXPLICIT, BYTE-CRITICAL) =================
        function t = add_t(obj, text)
            % ADD_T Return a newly added <w:t> containing `text` (run.py 40-45).
            %   Sets @xml:space="preserve" when the text has leading/trailing
            %   whitespace, i.e. len(text.strip()) < len(text). BYTE-CRITICAL:
            %   without it Word strips the spaces.
            arguments
                obj
                text (1,1) string
            end
            % Python: t = self._add_t(text=text)  -- setattr(child, "text", text)
            t = obj.add_t_("text", text);
            % Python: if len(text.strip()) < len(text): t.set(qn("xml:space"),"preserve")
            % `len(text.strip()) < len(text)` is TRUE iff text is non-empty and
            % its first or last character is Python whitespace. Ported with the
            % EXACT CPython strip set (PY_STRIP_WS) -- NOT MATLAB strip(), which
            % keeps the non-breaking spaces U+0085/U+00A0/U+2007/U+202F that
            % Python removes (Gate-2 audit fix, probes xmlspace_10/11/13; see
            % PY_STRIP_WS). First/last UTF-16 code unit is surrogate-safe: all
            % 29 set members are BMP and surrogates can never equal them (H2).
            tc = char(text);
            if ~isempty(tc) && (any(tc(1) == obj.PY_STRIP_WS) || any(tc(end) == obj.PY_STRIP_WS))
                t.set(mat2doc.oxml.qn("xml:space"), "preserve");
            end
        end

        % ======================= add_drawing (EXPLICIT) =========================
        function drawing = add_drawing(obj, inline_or_anchor)
            % ADD_DRAWING Return a newly appended <w:drawing> whose child is
            %   `inline_or_anchor` (a CT_Inline | CT_Anchor) (run.py 47-54).
            drawing = obj.add_drawing_();      % Python: self._add_drawing()
            drawing.append(inline_or_anchor);  % Python: drawing.append(inline_or_anchor)
        end

        % ========================= clear_content (LIVE) =========================
        function clear_content(obj)
            % CLEAR_CONTENT Remove all child elements except a <w:rPr> if present
            %   (run.py 56-60).
            %
            %   H9/H5: xpath() materializes the matches before the loop, so
            %   removing children as we go does not disturb iteration.
            content_elms = obj.xpath("./*[not(self::w:rPr)]");
            for k = 1:numel(content_elms)
                obj.remove(content_elms(k));
            end
        end

        % ===================== inner_content_items (LIVE) =======================
        function items = inner_content_items(obj)
            % INNER_CONTENT_ITEMS The run's content as an ordered heterogeneous
            %   list of str | CT_Drawing | CT_LastRenderedPageBreak, with the
            %   plain-text run-content elements coalesced into single str runs
            %   (run.py 62-89, @property). UN-STUBBED at P8-3 -- CT_Drawing (P7-3)
            %   and CT_LastRenderedPageBreak (P4-3) are both registered.
            %
            %   Python:
            %     accum = TextAccumulator()
            %     def iter_items():
            %         for e in self.xpath("w:br | w:cr | w:drawing"
            %                 " | w:lastRenderedPageBreak | w:noBreakHyphen"
            %                 " | w:ptab | w:t | w:tab"):
            %             if isinstance(e, (CT_Drawing, CT_LastRenderedPageBreak)):
            %                 yield from accum.pop()
            %                 yield e
            %             else:
            %                 accum.push(str(e))
            %         yield from accum.pop()   # the trailing "tail"
            %     return list(iter_items())
            %
            %   H10 isinstance dispatch: isa() on the two registered element
            %   classes (tuple isinstance -> the OR of two isa tests). H9: the
            %   Python generator is materialized into a 1xN CELL array (the list is
            %   HETEROGENEOUS -- string scalars interleaved with element handles --
            %   so a cell, not a typed array). str(e) -> e.str_() (the element
            %   text-equivalent, as in the CT_R.text getText_ getter). `yield from
            %   accum.pop()` iterates the 1x0-or-1x1 string TextAccumulator.pop()
            %   returns (nothing for 1x0), so an empty run yields the empty list {}.
            %
            %   Ported from python-docx v1.2.0: src/docx/oxml/text/run.py::CT_R.inner_content_items
            accum = mat2doc.shared.TextAccumulator();   % Python: accum = TextAccumulator()
            items = {};                                 % list(iter_items()) seed (heterogeneous)
            elms = obj.xpath("w:br | w:cr | w:drawing | w:lastRenderedPageBreak" + ...
                " | w:noBreakHyphen | w:ptab | w:t | w:tab");
            for k = 1:numel(elms)                       % Python: for e in self.xpath(...)
                e = elms(k);
                if isa(e, "mat2doc.oxml.drawing.CT_Drawing") || ...
                        isa(e, "mat2doc.oxml.text.CT_LastRenderedPageBreak")
                    % Python: yield from accum.pop(); yield e
                    for t = accum.pop()                 % 1x0 (nothing) or 1x1 (one str)
                        items{end + 1} = t;             %#ok<AGROW>
                    end
                    items{end + 1} = e;                 %#ok<AGROW>
                else
                    accum.push(e.str_());               % Python: accum.push(str(e))
                end
            end
            % Python: yield from accum.pop()  (don't forget the "tail" string)
            for t = accum.pop()
                items{end + 1} = t;                     %#ok<AGROW>
            end
        end

        % ===================== comment range helpers (LIVE) =====================
        function insert_comment_range_end_and_reference_below(obj, comment_id)
            % INSERT_COMMENT_RANGE_END_AND_REFERENCE_BELOW Insert a
            %   <w:commentRangeEnd> immediately after this run, followed by a <w:r>
            %   holding <w:commentReference> (run.py 91-98).
            %   Order matches Python: addnext the reference-run first, then addnext
            %   the commentRangeEnd, so the End lands BETWEEN this run and the
            %   reference-run.
            arguments
                obj
                comment_id (1,1) double
            end
            obj.addnext(obj.new_comment_reference_run_(comment_id));
            obj.addnext(mat2doc.oxml.OxmlElement("w:commentRangeEnd", ...
                [mat2doc.oxml.qn("w:id"), mat2doc.shared.pyStr(comment_id)]));
        end

        function insert_comment_range_start_above(obj, comment_id)
            % INSERT_COMMENT_RANGE_START_ABOVE Insert a <w:commentRangeStart> with
            %   `comment_id` immediately before this run (run.py 100-102).
            arguments
                obj
                comment_id (1,1) double
            end
            obj.addprevious(mat2doc.oxml.OxmlElement("w:commentRangeStart", ...
                [mat2doc.oxml.qn("w:id"), mat2doc.shared.pyStr(comment_id)]));
        end

        % ===================== lastRenderedPageBreaks (LIVE) ====================
        function breaks = lastRenderedPageBreaks(obj)
            % LASTRENDEREDPAGEBREAKS All <w:lastRenderedPageBreak> descendants of
            %   this run (run.py 104-107). Read-only accessor (Python @property);
            %   ported as a no-arg method so display never evaluates it.
            %   Returns generic XmlElement until CT_LastRenderedPageBreak registers
            %   (byte/structure identical; only the element CLASS differs).
            breaks = obj.xpath("./w:lastRenderedPageBreak");
        end

        % ============================ style (LIVE) ==============================
        function value = get.style(obj)
            % STYLE String in w:val of the w:rStyle grandchild, or [] if absent
            %   (run.py 109-118).
            rPr = obj.rPr;
            if isequal(rPr, [])   % Python: if rPr is None
                value = [];
                return
            end
            value = rPr.style;
        end
        function set.style(obj, style)
            % STYLE setter (run.py 120-127): set (or, when [], remove) the
            %   character style via the rPr. rPr.style handles the None removal.
            rPr = obj.get_or_add_rPr();
            rPr.style = style;
        end
    end

    methods (Access = private)
        function r = new_comment_reference_run_(obj, comment_id) %#ok<INUSD>
            % NEW_COMMENT_REFERENCE_RUN_ Return a new <w:r> referencing
            %   `comment_id` (run.py 149-164):
            %       <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
            %            <w:commentReference w:id="ID"/></w:r>
            %   OxmlElement("w:r") is registered -> a CT_R (the Python cast(CT_R,..)
            %   is a type hint, dropped). `obj` is unused (Python self unused).
            r = mat2doc.oxml.OxmlElement("w:r");
            rPr = r.get_or_add_rPr();
            rPr.style = "CommentReference";
            r.append(mat2doc.oxml.OxmlElement("w:commentReference", ...
                [mat2doc.oxml.qn("w:id"), mat2doc.shared.pyStr(comment_id)]));
        end
    end

    methods (Access = protected)
        % TEXT SHADOW (D10): CT_R.text is the run's concatenated inner-content
        % text (run.py 129-143), NOT the lxml char data. Overriding getText_/
        % setText_ (rather than the inherited `text` property) is the sanctioned
        % Mat2Doc mechanism; the serializer uses text_raw_ (bypass), so the w:r's
        % own char data (normally None) serializes unchanged.

        function value = getText_(obj)
            % Python CT_R.text getter (run.py 129-138): join str(e) of each
            %   inner-content child, in document order.
            %   VERIFY: str_() over a w:tab requires CT_TabStop (parfmt WP); until
            %   then a run containing a w:tab errors here (dependency-order gap,
            %   not a logic defect). str_() over br/cr/noBreakHyphen/ptab/t is LIVE.
            elms = obj.xpath("w:br | w:cr | w:noBreakHyphen | w:ptab | w:t | w:tab");
            value = "";                        % "".join(...) seed
            for k = 1:numel(elms)
                value = value + elms(k).str_();
            end
        end

        function setText_(obj, value)
            % Python CT_R.text setter (run.py 140-143): clear_content then append
            %   `value` as run content.
            obj.clear_content();
            mat2doc.oxml.text.RunContentAppender_.append_to_run_from_text(obj, string(value));
        end
    end
end

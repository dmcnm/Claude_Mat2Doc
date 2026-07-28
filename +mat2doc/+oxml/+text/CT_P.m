classdef CT_P < mat2doc.oxml.BaseOxmlElement
% CT_P `<w:p>` element: the properties and text for a paragraph.
%
%   THE M2 add_paragraph target (document.xml body content is w:p/w:pPr + w:r).
%   Registered for <w:p> (docx/oxml/__init__.py:229).
%
%   DESCRIPTORS (paragraph.py 29-31):
%     pPr       = ZeroOrOne("w:pPr")     -- NO successors (successors=()); an added
%                 pPr is forced to the FRONT via the _insert_pPr OVERRIDE below.
%     hyperlink = ZeroOrMore("w:hyperlink") -- successors=() -> APPEND
%     r         = ZeroOrMore("w:r")      -- successors=() -> APPEND
%
%   H11 CHILD ORDERING: pPr FIRST, then paragraph inner-content in INSERTION
%   order. paragraph.py OVERRIDES `_insert_pPr` (paragraph.py 104-106:
%   `self.insert(0, pPr)`) so get_or_add_pPr / _add_pPr always place pPr at index
%   0 (MATLAB insert(1,...), H1). Every content descriptor has successors=()
%   (NO_SUCCESSORS), so insertChildInSequence APPENDS -- content stays in the
%   order added. (Same shape as CT_R's rPr front-placement override.)
%
%   xmlchemy member generation (docx form):
%     ZeroOrOne (pPr): get.pPr, get_or_add_pPr, new_pPr_, insert_pPr_ (OVERRIDE),
%                      add_pPr_, remove_pPr_.
%     ZeroOrMore (hyperlink/r): x_lst, new_x_, insert_x_, add_x_, add_x (PUBLIC) --
%                      NO bare `x` getter, NO get_or_add, NO remover.
%   Underscore rotation: _new_x->new_x_, _insert_x->insert_x_, _add_x->add_x_
%   (the PUBLIC add_x keeps its bare name). The pyright Callable annotations
%   (paragraph.py 24-27: add_r/get_or_add_pPr/hyperlink_lst/r_lst) are type hints
%   only -- those members ARE generated.
%
%   TEXT SHADOW (D10): CT_P.text (paragraph.py 95-102) SHADOWS the lxml `.text`
%   attribute -- it is the paragraph's concatenated inner-content text
%   ("".join(e.text for e in w:r|w:hyperlink)), NOT the element char data. MATLAB
%   forbids redefining a superclass property, so this is ported by OVERRIDING the
%   protected getText_ (the serializer reads text_raw_, bypassing the override, so
%   the w:p's own char data -- normally None -- serializes unchanged). CT_P.text
%   is READ-ONLY in Python (getter only), so setText_ is overridden to RAISE
%   (mirroring the read-only property; the parser sets char data via
%   setTextRaw_/deepcopy copies via text_ raw fields, never through setText_).
%
%   NOT-YET-PORTED dependencies (VERIFY, see audit):
%     * w:hyperlink -> CT_Hyperlink (oxml/text/hyperlink.py): the `text` getter
%       reads e.text over w:r|w:hyperlink children. For a w:r child, e.text is the
%       LIVE CT_R.text (works now that w:tab is registered). For a w:hyperlink
%       child, w:hyperlink resolves to a generic XmlElement (its .text is the
%       element char data, NOT CT_Hyperlink's concatenated run text) until the
%       hyperlink WP lands -- so `.text` over a paragraph CONTAINING a hyperlink
%       does not reproduce python-docx (dependency-order gap, not a logic defect).
%       inner_content_elements / lastRenderedPageBreaks likewise return generic
%       elements for w:hyperlink until then (byte/structure identical; only the
%       element CLASS differs). NONE of these is on the M2 write path (the default
%       template body has one empty w:p; add_paragraph/add_heading write w:pPr +
%       w:r only).
%
%   H3 (None): inline isequal(x, []) (established Mat2Doc oxml convention).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on the <w:p> node(s) inside document.xml.
%
%   Example:
%       p = mat2doc.oxml.OxmlElement("w:p");
%       p.style = "Heading1";        % <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
%       r = p.add_r();               % append a <w:r>
%       r.add_t("Hello");            % <w:r><w:t>Hello</w:t></w:r>
%
%   Ported from python-docx v1.2.0: src/docx/oxml/text/paragraph.py::CT_P
%   (lines 21-106; registered for w:p)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        NO_SUCCESSORS = string.empty(1, 0)  % pPr + content: successors=() -> front(override)/append
        PPR_TAG       = "w:pPr"        % ZeroOrOne  @ paragraph.py:29
        HYPERLINK_TAG = "w:hyperlink"  % ZeroOrMore @ paragraph.py:30
        R_TAG         = "w:r"          % ZeroOrMore @ paragraph.py:31
    end

    properties (Dependent)  % ZeroOrOne getter + ZeroOrMore list getters + @property
        pPr             % <w:pPr> child or [] (None) if absent
        hyperlink_lst   % list of <w:hyperlink> children (document order)
        r_lst           % list of <w:r> children (document order)
        alignment       % pPr.jc_val (WD_PARAGRAPH_ALIGNMENT) or [] (None)
        style           % ./w:pPr/w:pStyle/@val or [] (None)
    end

    methods
        function obj = CT_P(varargin)
            % CT_P Construct a loose <w:p> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ============ pPr (ZeroOrOne, successors=(); FRONT via override) ============
        function child = get.pPr(obj);            child = obj.getChild(obj.PPR_TAG); end
        function child = get_or_add_pPr(obj)
            % Python get_or_add_pPr (xmlchemy 557-562), routed through the
            % overridden inserter (front placement).
            child = obj.pPr;
            if isequal(child, [])   % Python: if child is None (H3)
                child = obj.add_pPr_();
            end
        end
        function child = new_pPr_(obj);           child = obj.newChild(obj.PPR_TAG); end
        function child = insert_pPr_(obj, pPr)
            % OVERRIDE (paragraph.py 104-106): self.insert(0, pPr).
            obj.insert(1, pPr);    % Python insert(0,...) -> 1-based (H1)
            child = pPr;
        end
        function child = add_pPr_(obj, varargin)
            % _add_pPr (xmlchemy 284-291) routed through the OVERRIDE inserter.
            child = obj.new_pPr_();
            for k = 1:2:numel(varargin)
                child.(varargin{k}) = varargin{k + 1};
            end
            child = obj.insert_pPr_(child);
        end
        function remove_pPr_(obj);                obj.removeChild(obj.PPR_TAG); end

        % ============ hyperlink (ZeroOrMore, successors=() -> append) ============
        function lst = get.hyperlink_lst(obj);          lst = obj.getChildList(obj.HYPERLINK_TAG); end
        function child = new_hyperlink_(obj);           child = obj.newChild(obj.HYPERLINK_TAG); end
        function child = insert_hyperlink_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_hyperlink_(obj, varargin); child = obj.addChild(obj.HYPERLINK_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function child = add_hyperlink(obj);            child = obj.add_hyperlink_(); end   % public adder

        % ============ r (ZeroOrMore, successors=() -> append) ============
        function lst = get.r_lst(obj);          lst = obj.getChildList(obj.R_TAG); end
        function child = new_r_(obj);           child = obj.newChild(obj.R_TAG); end
        function child = insert_r_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_r_(obj, varargin); child = obj.addChild(obj.R_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function child = add_r(obj);            child = obj.add_r_(); end   % public adder

        % ============ add_p_before (paragraph.py 33-37) ============
        function new_p = add_p_before(obj)
            % ADD_P_BEFORE Return a new <w:p> inserted directly prior to this one.
            new_p = mat2doc.oxml.OxmlElement("w:p");   % registered -> CT_P (cast dropped)
            obj.addprevious(new_p);
        end

        % ============ alignment (paragraph.py 39-51) ============
        function value = get.alignment(obj)
            pPr = obj.pPr;
            if isequal(pPr, [])   % Python: if pPr is None
                value = [];
                return
            end
            value = pPr.jc_val;
        end
        function set.alignment(obj, value)
            pPr = obj.get_or_add_pPr();
            pPr.jc_val = value;
        end

        % ============ clear_content (paragraph.py 52-55) ============
        function clear_content(obj)
            % CLEAR_CONTENT Remove all child elements except the <w:pPr> if present.
            %   H9/H5: xpath() materializes the matches before the loop, so
            %   removing children as we go does not disturb iteration.
            children = obj.xpath("./*[not(self::w:pPr)]");
            for k = 1:numel(children)
                obj.remove(children(k));
            end
        end

        % ============ inner_content_elements (paragraph.py 57-60) ============
        function items = inner_content_elements(obj)
            % INNER_CONTENT_ELEMENTS Run and hyperlink children in document order.
            %   Read-only @property; ported as a no-arg method so display never
            %   evaluates it. Returns generic XmlElement for w:hyperlink until
            %   CT_Hyperlink registers (byte/structure identical; class differs).
            items = obj.xpath("./w:r | ./w:hyperlink");
        end

        % ============ lastRenderedPageBreaks (paragraph.py 62-71) ============
        function breaks = lastRenderedPageBreaks(obj)
            % LASTRENDEREDPAGEBREAKS All <w:lastRenderedPageBreak> descendants in a
            %   run, or a run inside a hyperlink (paragraph.py 62-71). Read-only
            %   @property; ported as a no-arg method so display never evaluates it.
            breaks = obj.xpath( ...
                "./w:r/w:lastRenderedPageBreak | ./w:hyperlink/w:r/w:lastRenderedPageBreak");
        end

        % ============ set_sectPr (paragraph.py 73-77) ============
        function set_sectPr(obj, sectPr)
            % SET_SECTPR Unconditionally replace or add sectPr as grandchild in
            %   correct sequence (paragraph.py 73-77).
            pPr = obj.get_or_add_pPr();
            pPr.remove_sectPr_();     % Python: pPr._remove_sectPr()
            pPr.insert_sectPr_(sectPr);  % Python: pPr._insert_sectPr(sectPr)
        end

        % ============ style (paragraph.py 79-93) ============
        function value = get.style(obj)
            % STYLE String in w:val of ./w:pPr/w:pStyle grandchild, or [] if absent.
            pPr = obj.pPr;
            if isequal(pPr, [])   % Python: if pPr is None
                value = [];
                return
            end
            value = pPr.style;
        end
        function set.style(obj, style)
            pPr = obj.get_or_add_pPr();
            pPr.style = style;   % pPr.style handles the [] (None) removal
        end
    end

    methods (Access = protected)
        % TEXT SHADOW (D10): CT_P.text is the paragraph's concatenated
        % inner-content text (paragraph.py 95-102), NOT the lxml char data.
        % Overriding getText_ (rather than the inherited `text` property) is the
        % sanctioned Mat2Doc mechanism; the serializer uses text_raw_ (bypass), so
        % the w:p's own char data (normally None) serializes unchanged. The Python
        % property is READ-ONLY, so setText_ raises.

        function value = getText_(obj)
            % Python CT_P.text getter (paragraph.py 95-102):
            %   "".join(e.text for e in self.xpath("w:r | w:hyperlink"))
            %   VERIFY: e.text over a w:hyperlink requires CT_Hyperlink (hyperlink
            %   WP); until then w:hyperlink is generic and its .text is the element
            %   char data, not the concatenated run text. e.text over a w:r is the
            %   LIVE CT_R.text.
            elms = obj.xpath("w:r | w:hyperlink");
            value = "";                        % "".join(...) seed
            for k = 1:numel(elms)
                value = value + elms(k).text;
            end
        end

        function setText_(obj, value) %#ok<INUSD>
            % Python CT_P.text is a READ-ONLY property (paragraph.py 95-102 defines
            %   only a getter). Assigning it raises AttributeError in Python; the
            %   proxy Paragraph.text setter operates on runs (CT_R.text), never on
            %   CT_P.text. No internal path calls setText_ on a w:p (the parser
            %   uses setTextRaw_; deepcopy copies text_ raw fields).
            error("mat2doc:AttributeError", "%s", ...
                "property 'text' of 'CT_P' object has no setter");
        end
    end
end

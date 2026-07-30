classdef CT_Hyperlink < mat2doc.oxml.BaseOxmlElement
% CT_HYPERLINK `<w:hyperlink>` element, containing the text and address for a hyperlink.
%
%   Registered for <w:hyperlink> (docx/oxml/__init__.py, text block).
%
%   DESCRIPTORS (hyperlink.py 24-32):
%     rId     = OptionalAttribute("r:id",    XsdString)             (default None -> [])
%     anchor  = OptionalAttribute("w:anchor", ST_String)           (default None -> [])
%     history = OptionalAttribute("w:history", ST_OnOff, default=True)  (NON-None bool default)
%     r       = ZeroOrMore("w:r")   -- successors=() -> APPEND
%
%   H3 tri-state:
%     * rId / anchor: no default (None -> []): absent -> []; set [] removes attr.
%     * history has a NON-None default `true`: getAttrTyped returns true when
%       @w:history is absent; the setter (D-delta-1) removes @w:history when set to
%       None ([]) OR to the default `true`, and writes "0" for false. `true`/`false`
%       are MATLAB logicals (ST_OnOff.from_xml returns logical; to_xml maps
%       True->"1", False->"0"). isequal(value, true) is the `value == self._default`
%       analogue.
%   XsdString / ST_String / ST_OnOff are P3-2 simple types; referenced by bare
%   short name (resolveTypeCls_ prefixes +oxml.simpletypes).
%
%   xmlchemy member generation (docx form):
%     OptionalAttribute (rId/anchor/history): get.<name> / set.<name> delegating to
%       getAttrTyped/setAttrTyped.
%     ZeroOrMore (r): r_lst, new_r_, insert_r_, add_r_, add_r (PUBLIC) -- NO bare
%       `r` getter, NO get_or_add, NO remover (xmlchemy). Underscore rotation
%       (design.md section 2): _new_r->new_r_, _insert_r->insert_r_, _add_r->add_r_
%       (the PUBLIC add_r keeps its bare name). The `r_lst: List[CT_R]` at
%       hyperlink.py:22 is a TYPE ANNOTATION only; the descriptor is `r =
%       ZeroOrMore("w:r")` at :32.
%
%   TEXT SHADOW (D10): CT_Hyperlink.text (hyperlink.py 39-45) SHADOWS the lxml
%   `.text` attribute -- it is the hyperlink's concatenated run text
%   ("".join(r.text for r in self.xpath("w:r"))), NOT the element char data. MATLAB
%   forbids redefining a superclass property, so this is ported by OVERRIDING the
%   protected getText_ (the serializer reads text_raw_, bypassing the override, so
%   the w:hyperlink's own char data -- normally None -- serializes unchanged).
%   CT_Hyperlink.text is READ-ONLY in Python (getter only), so setText_ raises
%   (same shape as CT_P.text). r.text over each w:r child is the LIVE CT_R.text.
%
%   lastRenderedPageBreaks (hyperlink.py 34-37): all <w:lastRenderedPageBreak>
%   descendants (./w:r/w:lastRenderedPageBreak). Read-only @property; ported as a
%   no-arg method so display never evaluates it (established Mat2Doc convention;
%   cf. CT_P.lastRenderedPageBreaks). Now that CT_LastRenderedPageBreak is
%   registered (same WP) these resolve to CT_LastRenderedPageBreak instances.
%
%   VERIFY-hyperlink CLOSURE (P4-2): registering this class flips CT_P.text over a
%   paragraph containing a <w:hyperlink> from generic-XmlElement char data to this
%   concatenated run text -- the P4-2 carry-forward. No CT_P edit required; the
%   registry row alone closes it.
%
%   DEVIATION -- s7 created-element nsdecl residual (D-serializer-nsdecl, ACCEPTED;
%   NO new D-number): on a LOOSE created <w:hyperlink> serialized STANDALONE,
%   setting rId (r:id) mints an xmlns:ns0 for the relationships URI, then CLEARING
%   rId leaves it orphaned -- lxml KEEPS the now-unused xmlns:ns0 while the Mat2Doc
%   serializer recomputes the used namespaces and DROPS it. Element / attributes /
%   text / expanded-names are all identical (exclusive-C14N-equal); the difference
%   is namespace-declaration-only. This is a manifestation of D-serializer-nsdecl
%   whose decl-emission fix covered PARSED verbatim nodes only, so this
%   created-element minted-then-orphaned sub-case is a residual. UNREACHABLE on
%   every real path: real hyperlinks keep their rId, and a hyperlink inside
%   document.xml is rooted under w:document (which already declares xmlns:r) so
%   r:id uses the existing prefix and never mints ns0 -- M1 17/17 stays
%   byte-neutral. ACCEPTED / deferred; reopen-check at P4-5b (add_hyperlink) and P7
%   (add_picture). See validation\summary\decision_2026-07-28_nsdecl_created_element_orphan.md.
%
%   H3 (None): inline isequal(x, []) (established Mat2Doc oxml convention).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on <w:hyperlink> nodes inside document.xml body content.
%
%   Example:
%       h = mat2doc.oxml.OxmlElement("w:hyperlink");
%       disp(h.history);            % true  (default; @w:history absent)
%       h.rId = "rId7";             % <w:hyperlink r:id="rId7">
%       r = h.add_r();  r.add_t("click"); disp(h.text);   % "click"
%
%   Ported from python-docx v1.2.0: src/docx/oxml/text/hyperlink.py::CT_Hyperlink
%   (lines 19-45; registered for w:hyperlink)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        NO_SUCCESSORS   = string.empty(1, 0)  % r: successors=() -> append
        RID_ATTR        = "r:id"        % OptionalAttribute @ hyperlink.py:24
        RID_TYPE        = "XsdString"   % XsdString simple type (P3-2)
        RID_DEFAULT     = []            % Python default: None
        ANCHOR_ATTR     = "w:anchor"    % OptionalAttribute @ hyperlink.py:25-27
        ANCHOR_TYPE     = "ST_String"   % ST_String simple type (P3-2)
        ANCHOR_DEFAULT  = []            % Python default: None
        HISTORY_ATTR    = "w:history"   % OptionalAttribute @ hyperlink.py:28-30
        HISTORY_TYPE    = "ST_OnOff"    % ST_OnOff simple type (P3-2)
        HISTORY_DEFAULT = true          % Python default=True (NON-None)
        R_TAG           = "w:r"         % ZeroOrMore @ hyperlink.py:32
    end

    properties (Dependent)  % generated OptionalAttribute properties + ZeroOrMore list
        rId       % r:id value, or [] (None) when @r:id absent
        anchor    % w:anchor value, or [] (None) when @w:anchor absent
        history   % w:history bool; true when @w:history absent (default)
        r_lst     % list of <w:r> children (document order)
    end

    methods
        function obj = CT_Hyperlink(varargin)
            % CT_HYPERLINK Construct a loose <w:hyperlink> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ============ rId (OptionalAttribute r:id, default None) ============
        function value = get.rId(obj)
            value = obj.getAttrTyped(obj.RID_ATTR, obj.RID_TYPE, obj.RID_DEFAULT);
        end
        function set.rId(obj, value)
            obj.setAttrTyped(obj.RID_ATTR, obj.RID_TYPE, value, obj.RID_DEFAULT);
        end

        % ============ anchor (OptionalAttribute w:anchor, default None) ============
        function value = get.anchor(obj)
            value = obj.getAttrTyped(obj.ANCHOR_ATTR, obj.ANCHOR_TYPE, obj.ANCHOR_DEFAULT);
        end
        function set.anchor(obj, value)
            obj.setAttrTyped(obj.ANCHOR_ATTR, obj.ANCHOR_TYPE, value, obj.ANCHOR_DEFAULT);
        end

        % ============ history (OptionalAttribute w:history, default=True) ============
        function value = get.history(obj)
            value = obj.getAttrTyped(obj.HISTORY_ATTR, obj.HISTORY_TYPE, obj.HISTORY_DEFAULT);
        end
        function set.history(obj, value)
            obj.setAttrTyped(obj.HISTORY_ATTR, obj.HISTORY_TYPE, value, obj.HISTORY_DEFAULT);
        end

        % ============ r (ZeroOrMore, successors=() -> append) ============
        function lst = get.r_lst(obj);          lst = obj.getChildList(obj.R_TAG); end
        function child = new_r_(obj);           child = obj.newChild(obj.R_TAG); end
        function child = insert_r_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_r_(obj, varargin); child = obj.addChild(obj.R_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function child = add_r(obj);            child = obj.add_r_(); end   % public adder

        % ============ lastRenderedPageBreaks (hyperlink.py 34-37) ============
        function breaks = lastRenderedPageBreaks(obj)
            % LASTRENDEREDPAGEBREAKS All <w:lastRenderedPageBreak> descendants of
            %   this hyperlink (hyperlink.py 34-37): self.xpath("./w:r/w:lastRenderedPageBreak").
            %   Read-only @property; ported as a no-arg method so display never
            %   evaluates it.
            breaks = obj.xpath("./w:r/w:lastRenderedPageBreak");
        end
    end

    methods (Access = protected)
        % TEXT SHADOW (D10): CT_Hyperlink.text is the hyperlink's concatenated run
        % text (hyperlink.py 39-45), NOT the lxml char data. Overriding getText_
        % (rather than the inherited `text` property) is the sanctioned Mat2Doc
        % mechanism; the serializer uses text_raw_ (bypass), so the w:hyperlink's
        % own char data (normally None) serializes unchanged. The Python property is
        % READ-ONLY, so setText_ raises.

        function value = getText_(obj)
            % Python CT_Hyperlink.text getter (hyperlink.py 39-45):
            %   "".join(r.text for r in self.xpath("w:r"))
            %   Each r is a CT_R and r.text is the LIVE CT_R.text (concatenated run
            %   content). xpath("w:r") selects only DIRECT w:r children.
            elms = obj.xpath("w:r");
            value = "";                        % "".join(...) seed
            for k = 1:numel(elms)
                value = value + elms(k).text;
            end
        end

        function setText_(obj, value) %#ok<INUSD>
            % Python CT_Hyperlink.text is a READ-ONLY property (hyperlink.py 39-45
            %   defines only a getter). Assigning it raises AttributeError in Python.
            %   No internal path sets text on a w:hyperlink (the parser uses
            %   setTextRaw_; deepcopy copies text_ raw fields).
            error("mat2doc:AttributeError", "%s", ...
                "property 'text' of 'CT_Hyperlink' object has no setter");
        end
    end
end

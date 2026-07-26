classdef CT_Body < mat2doc.oxml.BaseOxmlElement
% CT_BODY Custom element class for the <w:body> element.
%
%   The container element for the main document story in document.xml.
%   Registered for <w:body> (docx/oxml/__init__.py:100).
%
%   BYTE-NEUTRAL (P2-3): like CT_Document, registering w:body -> CT_Body changes
%   only the CLASS of the parsed w:body node (generic XmlElement -> CT_Body, both
%   BaseOxmlElement-tier serializers), not its bytes. Re-proven by the 17-part M1
%   sweep. See audit_P2-3_document_shell.md.
%
%   DESCRIPTORS (document.py 45-49):
%     p    = ZeroOrMore("w:p",    successors=("w:sectPr",))
%     tbl  = ZeroOrMore("w:tbl",  successors=("w:sectPr",))
%     sectPr = ZeroOrOne("w:sectPr", successors=())
%   Ported per design.md section 2 as a Constant schema table (P_TAG / TBL_TAG /
%   SECTPR_TAG + the successor Constants) + the generated one-line delegating
%   members. Per xmlchemy (docx form, P1-3b D-delta-4):
%     ZeroOrMore generates:  x_lst, _new_x, _insert_x, _add_x, add_x (PUBLIC) --
%                            NO bare `x` getter (delattr), NO get_or_add, NO
%                            remover  (xmlchemy.py 526-537).
%     ZeroOrOne  generates:  get.sectPr, get_or_add_sectPr, _new_sectPr,
%                            _insert_sectPr, _add_sectPr, _remove_sectPr
%                            (xmlchemy.py 540-573).
%   Underscore rotation (design.md section 2): _new_p -> new_p_, _insert_p ->
%   insert_p_, _add_p -> add_p_ (the PUBLIC add_p keeps its bare public name).
%
%   H11 CHILD ORDERING (the P2-3 hazard): p and tbl carry successors=("w:sectPr",)
%   -- a newly added w:p or w:tbl inserts BEFORE any existing w:sectPr (else
%   appends). The successor sequence is ported EXACTLY (PTBL_SUCCESSORS =
%   "w:sectPr"); the H11 slice logic lives in the P1-3a tree-ops
%   (insert_element_before / first_child_found_in), which insertChildInSequence
%   expands PTBL_SUCCESSORS into. This is dormant until P4/P6 wire add_paragraph/
%   add_table live, but the wiring is correct now. sectPr has successors=()
%   (append at end -> it is the sentinel, always last).
%
%   FORWARD DEPS (out of scope, resolve generic): CT_P (P4), CT_Tbl (P6),
%   CT_SectPr (P5) are NOT registered yet. The ZeroOrMore/ZeroOrOne descriptors
%   create children via OxmlElement (newChild), which falls back to a generic
%   XmlElement for these unregistered tags (P1-3b behavior). So add_p()/add_tbl()/
%   get_or_add_sectPr() produce/return generic w:p/w:tbl/w:sectPr elements that
%   still serialize correctly and insert in the right sequence -- only the element
%   CLASS is generic. The descriptors need ONLY the nsptag strings, all present.
%
%   LIVE vs STUB (P2-3):
%     LIVE  -- every descriptor member (all pure oxml, no proxy forward dep),
%              clear_content (xpath-based, document.py 73-79),
%              inner_content_elements (xpath-based, document.py 81-88).
%     STUB  -- add_section_break (document.py 51-71): it needs
%              sentinel_sectPr.clone() (CT_SectPr, P5) and self.add_p().set_sectPr
%              (CT_P, P5/P4). Owner: P5 section tier. Note add_p() itself is LIVE;
%              only the section-break orchestration stubs.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): forwards all
%   positional args verbatim -- REQUIRED because the parser instantiates this
%   class as feval(cls, name, ownDecls[, resolvedUri]) when it hits the registered
%   <w:body> child of a real document.xml.
%
%   None SENTINEL (H3): inline `isequal(x, [])` (established Mat2Doc oxml
%   convention). The descriptor engine handles None/absent; this class adds none.
%
%   Example:
%       d = mat2doc.oxml.OxmlElement("w:document");
%       body = d.get_or_add_body();      % a CT_Body
%       p = body.add_p();                % new <w:p>, before any <w:sectPr> (H11)
%       body.clear_content();            % removes all but a <w:sectPr> if present
%
%   Ported from python-docx v1.2.0: src/docx/oxml/document.py::CT_Body
%   (registered for <w:body>, oxml/__init__.py:100)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        NO_SUCCESSORS   = string.empty(1, 0)   % sectPr: successors=() -> append (sentinel)
        PTBL_SUCCESSORS = "w:sectPr"           % p/tbl: successors=("w:sectPr",) (H11)
        P_TAG      = "w:p"        % ZeroOrMore @ document.py:45
        TBL_TAG    = "w:tbl"      % ZeroOrMore @ document.py:46
        SECTPR_TAG = "w:sectPr"   % ZeroOrOne  @ document.py:47-49
    end

    properties (Dependent)  % generated list getters (ZeroOrMore) + ZeroOrOne getter + @property
        p_lst                    % list of <w:p> children (document order)
        tbl_lst                  % list of <w:tbl> children (document order)
        sectPr                   % <w:sectPr> child or [] (None) if absent
        inner_content_elements   % all <w:p> and <w:tbl>, document order
    end

    methods
        function obj = CT_Body(varargin)
            % CT_BODY Construct a loose <w:body> element.
            %   TRANSPARENT PASS-THROUGH (design.md section 2 INT-1).
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ============ generated ZeroOrMore descriptor members (p) ============
        % ZeroOrMore -> p_lst, new_p_, insert_p_, add_p_, add_p (public). NO bare
        % `p` getter (delattr). successors=("w:sectPr",) -> PTBL_SUCCESSORS (H11).
        function lst = get.p_lst(obj);           lst = obj.getChildList(obj.P_TAG); end
        function child = new_p_(obj);            child = obj.newChild(obj.P_TAG); end
        function child = insert_p_(obj, child);  child = obj.insertChildInSequence(child, obj.PTBL_SUCCESSORS); end
        function child = add_p_(obj, varargin);  child = obj.addChild(obj.P_TAG, obj.PTBL_SUCCESSORS, varargin{:}); end
        function child = add_p(obj);             child = obj.add_p_(); end   % public adder (xmlchemy 340-352)

        % ============ generated ZeroOrMore descriptor members (tbl) ============
        function lst = get.tbl_lst(obj);          lst = obj.getChildList(obj.TBL_TAG); end
        function child = new_tbl_(obj);           child = obj.newChild(obj.TBL_TAG); end
        function child = insert_tbl_(obj, child); child = obj.insertChildInSequence(child, obj.PTBL_SUCCESSORS); end
        function child = add_tbl_(obj, varargin); child = obj.addChild(obj.TBL_TAG, obj.PTBL_SUCCESSORS, varargin{:}); end
        function child = add_tbl(obj);            child = obj.add_tbl_(); end % public adder

        % ============ generated ZeroOrOne descriptor members (sectPr) ============
        % successors=() -> append (the sentinel sectPr is always last).
        function child = get.sectPr(obj);            child = obj.getChild(obj.SECTPR_TAG); end
        function child = get_or_add_sectPr(obj);     child = obj.getOrAddChild(obj.SECTPR_TAG, obj.NO_SUCCESSORS); end
        function child = new_sectPr_(obj);           child = obj.newChild(obj.SECTPR_TAG); end
        function child = insert_sectPr_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_sectPr_(obj, varargin); child = obj.addChild(obj.SECTPR_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function remove_sectPr_(obj);                obj.removeChild(obj.SECTPR_TAG); end

        % ======================= add_section_break (STUB) =======================
        function sectPr = add_section_break(obj) %#ok<MANU,STOUT>
            % ADD_SECTION_BREAK STUB (document.py 51-71). Owner: P5 section tier.
            %   Faithful body needs sentinel_sectPr.clone() (CT_SectPr.clone, P5)
            %   and self.add_p().set_sectPr(...) (CT_P.set_sectPr, P4/P5) plus
            %   sentinel_sectPr.xpath("w:headerReference|w:footerReference") removal.
            %   add_p()/get_or_add_sectPr() are LIVE here; only the clone/set_sectPr
            %   orchestration is unported. Stubbed to avoid a half-applied mutation.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.oxml.section.CT_SectPr.clone / mat2doc.oxml.text.CT_P.set_sectPr " + ...
                "(owning WP: P5 section tier) required by " + ...
                "mat2doc.oxml.document.CT_Body.add_section_break");
        end

        % ========================= clear_content (LIVE) =========================
        function clear_content(obj)
            % CLEAR_CONTENT Remove all content child elements, leaving a
            %   <w:sectPr> if present (document.py 73-79).
            %
            %   H9/H5: xpath() materializes the matches into a fixed typed array
            %   before the loop, so removing children as we go does not disturb the
            %   iteration (identical to Python's list-then-remove). Returns nothing
            %   (Python returns None); the proxy _Body.clear_content returns self.
            %
            %   Ported from python-docx v1.2.0: src/docx/oxml/document.py::
            %   CT_Body.clear_content
            content_elms = obj.xpath("./*[not(self::w:sectPr)]");
            for k = 1:numel(content_elms)
                obj.remove(content_elms(k));
            end
        end

        % ===================== inner_content_elements (@property) ===============
        function elms = get.inner_content_elements(obj)
            % INNER_CONTENT_ELEMENTS All <w:p> and <w:tbl> elements in this body,
            %   in document order (document.py 81-88). Elements shaded by nesting
            %   in a w:ins or other wrapper are NOT included (the xpath is a fixed
            %   two-branch child union, not a descendant scan).
            %
            %   H9: materialized typed array (Python returns a list); empty typed
            %   array when none present (never [] / None).
            %
            %   Ported from python-docx v1.2.0: src/docx/oxml/document.py::
            %   CT_Body.inner_content_elements
            elms = obj.xpath("./w:p | ./w:tbl");
        end
    end
end

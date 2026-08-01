classdef CT_HdrFtr < mat2doc.oxml.BaseOxmlElement
% CT_HDRFTR `w:hdr` and `w:ftr`, the root element for a header/footer PART.
%
%   CT_HdrFtr is the ROOT of a SEPARATE package part (word/header1.xml,
%   word/footer1.xml, ...), NOT a child of document.xml. Registered for BOTH
%   <w:hdr> and <w:ftr> (docx/oxml/__init__.py:124-125). It holds block content
%   (w:p / w:tbl children), exactly like CT_Body holds the main story.
%
%   M1-NEUTRAL (P5-2b): default.docx contains 17 parts and NO header/footer part;
%   the strings "header"/"footer" do not occur in its [Content_Types].xml, and no
%   M1 part contains a <w:hdr>/<w:ftr> tag. So registering w:hdr/w:ftr -> CT_HdrFtr
%   cannot alter any M1 parse path (the row only lights up when a header/footer
%   part is actually loaded, first materialized by P5-3b). Byte-neutral AND
%   flip-neutral: no existing exact-class XmlElement pin can see a w:hdr/w:ftr
%   class (planaudit_2026-07-31_p5-crosspart-hdrftr.md finding (e)).
%
%   DESCRIPTORS (section.py 38-39):
%     p   = ZeroOrMore("w:p",   successors=())
%     tbl = ZeroOrMore("w:tbl", successors=())
%   The section.py 32-36 annotations (add_p / p_lst / tbl_lst / _insert_tbl) are
%   type hints only for these dynamically generated members. Ported per design.md
%   section 2 as a Constant schema table (P_TAG / TBL_TAG + NO_SUCCESSORS) + the
%   generated one-line delegating members. Per xmlchemy (docx form, D-delta-4):
%     ZeroOrMore generates: x_lst, _new_x, _insert_x, _add_x, add_x (PUBLIC) --
%                           NO bare `x` getter (delattr), NO get_or_add, NO remover.
%   Underscore rotation (design.md section 2): _new_p -> new_p_, _insert_p ->
%   insert_p_, _add_p -> add_p_ (the PUBLIC add_p keeps its bare public name).
%   Both p and tbl carry successors=() -> NO_SUCCESSORS (append at end), exactly
%   like CT_Body's sectPr sentinel. (Mirror of the CT_Body descriptor pattern.)
%
%   ===================== CT_Tbl TAG-BASED INCLUSION (P5-2b) =====================
%   inner_content_elements is xpath("./w:p | ./w:tbl") -- a TAG-BASED child union,
%   NOT isinstance dispatch (contrast Section.iter_inner_content, P5-3a, which
%   DOES isinstance-dispatch and carries the table-branch debt). CT_Tbl is now
%   REGISTERED (P6-3b, w:tbl->CT_Tbl), so a <w:tbl> child resolves to a CT_Tbl (the
%   un-defer sweep auto-upgraded this tag-based site; through P2-3..P6-2 it was a
%   generic XmlElement -- the CT_Body.tbl_lst precedent). Because the xpath selects
%   by TAG it always INCLUDED the node (never dropped, never crashed); registering
%   CT_Tbl only flips the matched element's CLASS generic->CT_Tbl, no code change.
%   So this class ports COMPLETE with ZERO stubs; no notYetPorted CT_Tbl branch
%   belongs here.
%
%   Likewise the generated tbl descriptors (new_tbl_/insert_tbl_/add_tbl_/add_tbl)
%   need ONLY the "w:tbl" nsptag string; add_tbl() produces a generic <w:tbl> that
%   serializes and sequences correctly. add_tbl has no caller until P6.
%
%   H5 (element identity): inner_content_elements returns the LIVE child handles
%   (the xpath wrappers ARE the nodes), so two reads return == elements. H9: the
%   Python @property returns a list; the MATLAB surface is the materialized 1xN
%   heterogeneous XmlElement array the xpath engine yields (CT_P for <w:p> +
%   generic XmlElement for <w:tbl>), empty typed array when none present (never []
%   / None). Elements shaded by nesting in a w:ins or other wrapper are NOT
%   included (the xpath is a fixed two-branch CHILD union, not a descendant scan).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this as feval(cls, name, ownDecls[, resolvedUri]) when it parses
%   the <w:hdr>/<w:ftr> root of a real header/footer part (P5-3b onward).
%
%   Example:
%       h = mat2doc.oxml.OxmlElement("w:hdr");   % a CT_HdrFtr
%       p = h.add_p();                           % new <w:p>, appended
%       elms = h.inner_content_elements;         % 1xN [CT_P | XmlElement(tbl)]
%
%   Ported from python-docx v1.2.0: src/docx/oxml/section.py::CT_HdrFtr
%   (lines 29-48; registered for w:hdr and w:ftr, oxml/__init__.py:124-125)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        NO_SUCCESSORS = string.empty(1, 0)   % p/tbl: successors=() -> append at end
        P_TAG   = "w:p"     % ZeroOrMore @ section.py:38
        TBL_TAG = "w:tbl"   % ZeroOrMore @ section.py:39
    end

    properties (Dependent)  % generated ZeroOrMore list getters + @property
        p_lst                    % list of <w:p> children (document order)
        tbl_lst                  % list of <w:tbl> children (document order)
        inner_content_elements   % all <w:p> and <w:tbl>, document order
    end

    methods
        function obj = CT_HdrFtr(varargin)
            % CT_HDRFTR Construct a loose <w:hdr>/<w:ftr> element.
            %   TRANSPARENT PASS-THROUGH (design.md section 2 INT-1).
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ============ generated ZeroOrMore descriptor members (p) ============
        % ZeroOrMore -> p_lst, new_p_, insert_p_, add_p_, add_p (public). NO bare
        % `p` getter (delattr). successors=() -> NO_SUCCESSORS (append at end).
        function lst = get.p_lst(obj);           lst = obj.getChildList(obj.P_TAG); end
        function child = new_p_(obj);            child = obj.newChild(obj.P_TAG); end
        function child = insert_p_(obj, child);  child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_p_(obj, varargin);  child = obj.addChild(obj.P_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function child = add_p(obj);             child = obj.add_p_(); end   % public adder (xmlchemy 340-352)

        % ============ generated ZeroOrMore descriptor members (tbl) ============
        % CT_Tbl registered (P6-3b): add_tbl()/tbl_lst return CT_Tbl.
        function lst = get.tbl_lst(obj);          lst = obj.getChildList(obj.TBL_TAG); end
        function child = new_tbl_(obj);           child = obj.newChild(obj.TBL_TAG); end
        function child = insert_tbl_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_tbl_(obj, varargin); child = obj.addChild(obj.TBL_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function child = add_tbl(obj);            child = obj.add_tbl_(); end % public adder

        % ===================== inner_content_elements (@property) ===============
        function elms = get.inner_content_elements(obj)
            % INNER_CONTENT_ELEMENTS All <w:p> and <w:tbl> elements in this header
            %   or footer, in document order (section.py 41-48). TAG-BASED child
            %   union -- a <w:tbl> resolves to CT_Tbl (registered P6-3b) and
            %   is INCLUDED (H5). Elements shaded by nesting in a w:ins or other
            %   wrapper are NOT included. H9: materialized 1xN heterogeneous
            %   XmlElement array (Python returns a list); empty typed array when
            %   none present (never [] / None). Identical form to
            %   CT_Body.inner_content_elements (document.py 81-88).
            %
            %   Ported from python-docx v1.2.0: src/docx/oxml/section.py::
            %   CT_HdrFtr.inner_content_elements
            elms = obj.xpath("./w:p | ./w:tbl");
        end
    end
end

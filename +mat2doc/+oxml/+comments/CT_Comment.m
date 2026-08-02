classdef CT_Comment < mat2doc.oxml.BaseOxmlElement
% CT_COMMENT `<w:comment>` element, representing a single comment.
%
%   A comment is a "story" and, like a table-cell, can contain paragraphs and
%   tables. While most often a single sentence or phrase, a comment can carry
%   rich content: multiple rich-text paragraphs, hyperlinks, images, and tables.
%
%   Registered for <w:comment> (docx/oxml/__init__.py:92, P8-2).
%
%   ATTRIBUTES (comments.py 99-107):
%     id       = RequiredAttribute("w:id",       ST_DecimalNumber)
%     author   = RequiredAttribute("w:author",   ST_String)
%     initials = OptionalAttribute("w:initials", ST_String)     -- None -> absent
%     date     = OptionalAttribute("w:date",      ST_DateTime)    -- None -> absent
%
%   xmlchemy member generation:
%     * RequiredAttribute id       -> get.id/set.id       (getAttrRequired /
%       setAttrRequired, ST_DecimalNumber; InvalidXmlError if @w:id absent).
%     * RequiredAttribute author   -> get.author/set.author (ST_String).
%     * OptionalAttribute initials -> get.initials/set.initials (getAttrTyped /
%       setAttrTyped, ST_String, default None ([]); H3: assign [] removes @w:initials).
%     * OptionalAttribute date     -> get.date/set.date (ST_DateTime, default
%       None ([])). ***This is the FIRST reachable ST_DateTime consumer in the
%       project*** (see the class header of ST_DateTime + audit_P8-2). H3: assign
%       [] removes @w:date.
%
%   CHILDREN (comments.py 111-112):
%     p   = ZeroOrMore("w:p",   successors=())   -- append (comment is a block
%                                                   container; sentinel is end)
%     tbl = ZeroOrMore("w:tbl", successors=())   -- append
%   xmlchemy (docx form) generates for each ZeroOrMore: x_lst, _new_x, _insert_x,
%   _add_x (private, **attrs) AND the PUBLIC add_x (comments.py declares add_p /
%   p_lst / tbl_lst / _insert_tbl only as metaclass type-hints, so NO explicit
%   body suppresses the generated public adders). Underscore rotation (design.md
%   section 2): _new_p->new_p_, _insert_p->insert_p_, _add_p->add_p_ (public add_p
%   keeps its bare name); same for tbl. successors=() -> NO_SUCCESSORS -> append.
%
%   These generated members are exactly what the C3 BlockItemContainer seam calls
%   through Comment: BlockItemContainer._add_paragraph -> element_().add_p();
%   BlockItemContainer.add_table -> element_()._insert_tbl(tbl) (insert_tbl_);
%   .paragraphs -> element_().p_lst; .tables -> element_().tbl_lst.
%
%   inner_content_elements (@property, comments.py 121-124): all <w:p> and <w:tbl>
%   in document order via xpath("./w:p | ./w:tbl") (the union is in the verified
%   xpath subset). Materialized typed array (H9); empty typed array when none.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on <w:comment> nodes inside a real word/comments.xml.
%
%   None SENTINEL (H3): inline `isequal(x, [])`; the OptionalAttribute engine
%   handles absent/None for initials/date.
%
%   Example:
%       c = mat2doc.oxml.OxmlElement("w:comments").add_comment();  % a CT_Comment
%       c.author = "Amy";                       % RequiredAttribute setter
%       c.initials = "AJ";                      % OptionalAttribute setter
%       p = c.add_p();                          % a new <w:p> (append)
%
%   Ported from python-docx v1.2.0: src/docx/oxml/comments.py::CT_Comment
%   (lines 91-124; registered for w:comment)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        ID_ATTR        = "w:id"              % RequiredAttribute @ comments.py:100
        ID_TYPE        = "ST_DecimalNumber"  % simple type (+oxml\+simpletypes)
        AUTHOR_ATTR    = "w:author"          % RequiredAttribute @ comments.py:101
        AUTHOR_TYPE    = "ST_String"         % simple type
        INITIALS_ATTR  = "w:initials"        % OptionalAttribute @ comments.py:102-104
        INITIALS_TYPE  = "ST_String"         % simple type
        DATE_ATTR      = "w:date"            % OptionalAttribute @ comments.py:105-107
        DATE_TYPE      = "ST_DateTime"       % simple type (FIRST reachable consumer)
        P_TAG          = "w:p"               % ZeroOrMore @ comments.py:111
        TBL_TAG        = "w:tbl"             % ZeroOrMore @ comments.py:112
        NO_SUCCESSORS  = string.empty(1, 0)  % successors=() -> append
    end

    properties (Dependent)  % generated descriptor members
        id                       % RequiredAttribute('w:id', ST_DecimalNumber) -> double (int)
        author                   % RequiredAttribute('w:author', ST_String) -> string
        initials                 % OptionalAttribute('w:initials', ST_String) -> string | []
        date                     % OptionalAttribute('w:date', ST_DateTime) -> datetime | []
        p_lst                    % list of <w:p> children (document order)
        tbl_lst                  % list of <w:tbl> children (document order)
        inner_content_elements   % all <w:p> and <w:tbl>, document order
    end

    methods
        function obj = CT_Comment(varargin)
            % CT_COMMENT Construct a loose <w:comment> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ============ id (RequiredAttribute ST_DecimalNumber) ============
        function value = get.id(obj);   value = obj.getAttrRequired(obj.ID_ATTR, obj.ID_TYPE); end
        function set.id(obj, value);    obj.setAttrRequired(obj.ID_ATTR, obj.ID_TYPE, value); end

        % ============ author (RequiredAttribute ST_String) ============
        function value = get.author(obj); value = obj.getAttrRequired(obj.AUTHOR_ATTR, obj.AUTHOR_TYPE); end
        function set.author(obj, value);  obj.setAttrRequired(obj.AUTHOR_ATTR, obj.AUTHOR_TYPE, value); end

        % ============ initials (OptionalAttribute ST_String, default None) ============
        function value = get.initials(obj); value = obj.getAttrTyped(obj.INITIALS_ATTR, obj.INITIALS_TYPE); end
        function set.initials(obj, value);  obj.setAttrTyped(obj.INITIALS_ATTR, obj.INITIALS_TYPE, value); end

        % ============ date (OptionalAttribute ST_DateTime, default None) ============
        function value = get.date(obj); value = obj.getAttrTyped(obj.DATE_ATTR, obj.DATE_TYPE); end
        function set.date(obj, value);  obj.setAttrTyped(obj.DATE_ATTR, obj.DATE_TYPE, value); end

        % ============ p (ZeroOrMore, successors=() -> append) ============
        function lst = get.p_lst(obj);           lst = obj.getChildList(obj.P_TAG); end
        function child = new_p_(obj);            child = obj.newChild(obj.P_TAG); end
        function child = insert_p_(obj, child);  child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_p_(obj, varargin);  child = obj.addChild(obj.P_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function child = add_p(obj);             child = obj.add_p_(); end   % public adder (xmlchemy 340-352)

        % ============ tbl (ZeroOrMore, successors=() -> append) ============
        function lst = get.tbl_lst(obj);          lst = obj.getChildList(obj.TBL_TAG); end
        function child = new_tbl_(obj);           child = obj.newChild(obj.TBL_TAG); end
        function child = insert_tbl_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_tbl_(obj, varargin); child = obj.addChild(obj.TBL_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function child = add_tbl(obj);            child = obj.add_tbl_(); end % public adder

        % ============ inner_content_elements (@property) ============
        function elms = get.inner_content_elements(obj)
            % INNER_CONTENT_ELEMENTS All <w:p> and <w:tbl> in this comment, in
            %   document order (comments.py 121-124). Python: return
            %   self.xpath("./w:p | ./w:tbl"). H9: materialized typed array; empty
            %   typed array when none (never [] / None).
            %
            %   Ported from python-docx v1.2.0: src/docx/oxml/comments.py::CT_Comment.inner_content_elements
            elms = obj.xpath("./w:p | ./w:tbl");
        end
    end
end

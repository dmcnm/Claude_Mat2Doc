classdef CT_TabStops < mat2doc.oxml.BaseOxmlElement
% CT_TABSTOPS `<w:tabs>` element, container for a sorted sequence of tab stops.
%
%   DESCRIPTOR (parfmt.py 381):
%     tab = OneOrMore("w:tab", successors=())  -- successors=() -> APPEND.
%
%   xmlchemy OneOrMore member generation (docx form): tab_lst, new_tab_,
%   insert_tab_, add_tab_, add_tab (PUBLIC). No get_or_add, no bare `tab` getter,
%   no remover. Underscore rotation: _new_tab->new_tab_, _insert_tab->insert_tab_,
%   _add_tab->add_tab_ (the PUBLIC add_tab keeps its bare name).
%
%   insert_tab_in_order(pos, align, leader) (parfmt.py 383-392): create a w:tab,
%   set pos/val/leader, then insert it before the FIRST existing tab whose pos is
%   greater (keeping the sequence sorted by pos), else append. H1: the Python
%   loop over self.tab_lst is document order; the port scans tab_lst 1-based.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on <w:tabs> nodes inside a real part.
%
%   Example:
%       tabs = mat2doc.oxml.OxmlElement("w:tabs");
%       tabs.insert_tab_in_order(mat2doc.shared.Twips(720), ...
%           mat2doc.enum.text.WD_TAB_ALIGNMENT.LEFT, ...
%           mat2doc.enum.text.WD_TAB_LEADER.SPACES);
%
%   Ported from python-docx v1.2.0: src/docx/oxml/text/parfmt.py::CT_TabStops
%   (lines 378-392; registered for w:tabs)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        NO_SUCCESSORS = string.empty(1, 0)  % OneOrMore successors=() -> append
        TAB_TAG       = "w:tab"             % OneOrMore @ parfmt.py:381
    end

    properties (Dependent)  % OneOrMore list getter
        tab_lst   % list of <w:tab> children (document order)
    end

    methods
        function obj = CT_TabStops(varargin)
            % CT_TABSTOPS Construct a loose <w:tabs> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ============ tab (OneOrMore, successors=() -> append) ============
        function lst = get.tab_lst(obj);          lst = obj.getChildList(obj.TAB_TAG); end
        function child = new_tab_(obj);           child = obj.newChild(obj.TAB_TAG); end
        function child = insert_tab_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_tab_(obj, varargin); child = obj.addChild(obj.TAB_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function child = add_tab(obj);            child = obj.add_tab_(); end  % public adder

        % ============ insert_tab_in_order (parfmt.py 383-392) ============
        function new_tab = insert_tab_in_order(obj, pos, align, leader)
            % INSERT_TAB_IN_ORDER Insert a newly created w:tab child in `pos` order.
            new_tab = obj.new_tab_();
            % Python: new_tab.pos, new_tab.val, new_tab.leader = pos, align, leader
            new_tab.pos = pos;
            new_tab.val = align;
            new_tab.leader = leader;
            tabs = obj.tab_lst;
            for k = 1:numel(tabs)
                if new_tab.pos < tabs(k).pos
                    tabs(k).addprevious(new_tab);
                    return
                end
            end
            obj.append(new_tab);
        end
    end
end

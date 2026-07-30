classdef TabStops < mat2doc.shared.ElementProxy
% TABSTOPS A sequence of TabStop objects: the tab stops of a paragraph/style.
%
%   Supports iteration, indexed access, del, and len() -- in Python a
%   collections-style sequence accessed via ParagraphFormat.tab_stops; it is
%   not intended to be constructed directly. An ElementProxy subclass:
%   reference semantics (handle) and H5 element-identity eq/ne are inherited.
%
%   WRAPPED ELEMENT (tabstops.py 16-18): TabStops wraps the CT_PPr (`w:pPr`),
%   NOT the `w:tabs`. Its `self._pPr` (-> pPr_) reaches the tabs container on
%   demand via pPr.tabs / pPr.get_or_add_tabs() / pPr._remove_tabs() (the P4-2
%   CT_PPr descriptor family). super().__init__(element, None) sets both the
%   inherited _element and pPr_ to the SAME CT_PPr.
%
%   VERIFY-COLLECTION (design.md section 2 "Collections -> shared RedefinesParen
%   base"): the shared collection base (RedefinesParen; dunder mapping
%   x[i] -> x(i+1), len(x) -> numel(x), for s in x -> for s = x.to_array()) is a
%   FUTURE work package and does not exist yet in Mat2Doc. Following the
%   established cross-toolbox precedent (Mat2Ppt _GradientStops ->
%   GradientStops_), the Python sequence surface is ported here as EXPLICIT
%   methods -- getitem_ (__getitem__), len_ (__len__), to_array (__iter__),
%   delitem_ (__delitem__) -- keeping line-for-line fidelity. When the
%   collections base lands, TabStops should derive from it and expose native
%   1-based () indexing (tab_stops(1) == getitem_(0)); the methods below are the
%   interim faithful surface. FLAGGED for the auditor/validator.
%
%   H1 (indexing): getitem_/delitem_ take the PYTHON 0-based index (mirroring
%   __getitem__/__delitem__), and hit the 1-based MATLAB child list with an
%   explicit `+1` (the H1 site). Full Python list int-index semantics are
%   replicated: a negative index counts from the end; an out-of-range index
%   raises mat2doc:IndexError with the faithful message.
%
%   REFERENCE SEMANTICS (H5): each getitem_/to_array element mints a FRESH
%   TabStop view of the corresponding <w:tab> (Python does the same -- TabStop
%   objects are not cached).
%
%   Example:
%       pf = mat2doc.text.ParagraphFormat(mat2doc.oxml.OxmlElement("w:p"));
%       ts = pf.tab_stops;
%       ts.add_tab_stop(mat2doc.shared.Inches(1));   % LEFT / SPACES defaults
%       ts.len_()                                    % 1
%       first = ts.getitem_(0);                      % Python ts[0] (H1)
%       ts.clear_all();
%
%   Ported from python-docx v1.2.0: src/docx/text/tabstops.py::TabStops

    properties (Access = private)
        pPr_   % Python self._pPr: the wrapped <w:pPr> (CT_PPr)
    end

    methods
        function obj = TabStops(element)
            % TABSTOPS Wrap a CT_PPr (tabstops.py 16-18).
            %
            %   Inputs:  element - a mat2doc.oxml.text.CT_PPr (the `w:pPr` whose
            %                      `w:tabs` carries the tab stops).
            %   Outputs: obj     - a scalar TabStops handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/text/tabstops.py::TabStops.__init__
            obj@mat2doc.shared.ElementProxy(element, []);  % Python: super().__init__(element, None)
            obj.pPr_ = element;                            % Python: self._pPr = element
        end

        function delitem_(obj, idx)
            % DELITEM_ Remove the tab at 0-based offset `idx` (tabstops.py 20-29).
            %   Python __delitem__:
            %     tabs = self._pPr.tabs
            %     try: tabs.remove(tabs[idx])
            %     except (AttributeError, IndexError): raise IndexError("tab index out of range")
            %     if len(tabs) == 0: self._pPr.remove(tabs)
            %   AttributeError (tabs is None) and IndexError (out of range) both
            %   collapse to IndexError("tab index out of range").
            %
            %   Ported from python-docx v1.2.0: src/docx/text/tabstops.py::TabStops.__delitem__
            tabs = obj.pPr_.tabs;
            child = [];
            ok = false;
            if ~isequal(tabs, [])                     % Python tabs is None -> AttributeError branch
                % tabs[idx]: lxml element indexing over the <w:tabs> children
                % (all <w:tab>). Full Python list int-index semantics.
                children = tabs.to_array();           % all children == the w:tab list
                n = numel(children);
                i = idx;                              % Python 0-based index
                if i < 0                              % Python negative-index wrap
                    i = i + n;
                end
                if i >= 0 && i < n
                    child = children(i + 1);          % IDX: Python x[idx] -> x(idx+1)
                    ok = true;
                end
            end
            if ~ok                                    % Python: raise IndexError(...)
                error("mat2doc:IndexError", "%s", "tab index out of range");
            end
            tabs.remove(child);                       % Python: tabs.remove(tabs[idx])
            if numel(tabs.to_array()) == 0            % Python: if len(tabs) == 0
                obj.pPr_.remove(tabs);                % Python: self._pPr.remove(tabs)
            end
        end

        function tab = getitem_(obj, idx)
            % GETITEM_ The TabStop at 0-based position `idx` (tabstops.py 31-37).
            %   Python __getitem__:
            %     tabs = self._pPr.tabs
            %     if tabs is None: raise IndexError("TabStops object is empty")
            %     tab = tabs.tab_lst[idx]
            %     return TabStop(tab)
            %   An empty collection raises IndexError("TabStops object is empty");
            %   an out-of-range index raises the Python list message.
            %
            %   Ported from python-docx v1.2.0: src/docx/text/tabstops.py::TabStops.__getitem__
            tabs = obj.pPr_.tabs;
            if isequal(tabs, [])                      % Python: if tabs is None
                error("mat2doc:IndexError", "%s", "TabStops object is empty");
            end
            tab_lst = tabs.tab_lst;
            n = numel(tab_lst);
            i = idx;                                  % Python 0-based index
            if i < 0                                  % Python negative-index wrap
                i = i + n;
            end
            if i < 0 || i >= n                        % Python list out-of-range
                error("mat2doc:IndexError", "%s", "list index out of range");
            end
            % Python: TabStop(tabs.tab_lst[idx])
            tab = mat2doc.text.TabStop(tab_lst(i + 1));   % IDX: Python x[idx] -> x(idx+1)
        end

        function result = to_array(obj)
            % TO_ARRAY A TabStop for each <w:tab>, in XML document order (tabstops.py 39-45).
            %   The iteration idiom (design.md section 2): Python `for t in tab_stops`
            %   -> `for t = tab_stops.to_array()`. Python __iter__ yields nothing
            %   when tabs is None -> a 1x0 TabStop array here.
            %
            %   Ported from python-docx v1.2.0: src/docx/text/tabstops.py::TabStops.__iter__
            result = mat2doc.text.TabStop.empty(1, 0);
            tabs = obj.pPr_.tabs;
            if ~isequal(tabs, [])                     % Python: if tabs is not None
                tab_lst = tabs.tab_lst;
                for k = 1:numel(tab_lst)              % Python: for tab in tabs.tab_lst
                    result(k) = mat2doc.text.TabStop(tab_lst(k));
                end
            end
        end

        function n = len_(obj)
            % LEN_ Number of tab stops (tabstops.py 47-51).
            %   Python __len__: 0 when tabs is None, else len(tabs.tab_lst).
            %
            %   Ported from python-docx v1.2.0: src/docx/text/tabstops.py::TabStops.__len__
            tabs = obj.pPr_.tabs;
            if isequal(tabs, [])                      % Python: if tabs is None
                n = 0;
                return
            end
            n = numel(tabs.tab_lst);                  % Python: len(tabs.tab_lst)
        end

        function tab = add_tab_stop(obj, position, alignment, leader)
            % ADD_TAB_STOP Add a tab stop at `position`, inserted in pos order.
            %   (tabstops.py 53-65). alignment defaults to LEFT, leader to SPACES;
            %   a negative position is valid (hanging indentation).
            %
            %   Inputs:  position  - a mat2doc.shared.Length (offset from the
            %                        paragraph edge; may be negative).
            %            alignment - (optional) WD_TAB_ALIGNMENT; default LEFT.
            %            leader    - (optional) WD_TAB_LEADER; default SPACES.
            %   Outputs: tab       - the newly added TabStop.
            %
            %   Ported from python-docx v1.2.0: src/docx/text/tabstops.py::TabStops.add_tab_stop
            arguments
                obj
                position
                alignment = mat2doc.enum.text.WD_TAB_ALIGNMENT.LEFT    % Python default
                leader    = mat2doc.enum.text.WD_TAB_LEADER.SPACES     % Python default
            end
            tabs = obj.pPr_.get_or_add_tabs();
            tab_el = tabs.insert_tab_in_order(position, alignment, leader);
            tab = mat2doc.text.TabStop(tab_el);
        end

        function clear_all(obj)
            % CLEAR_ALL Remove all custom tab stops (tabstops.py 67-69).
            %   Python: self._pPr._remove_tabs().
            %
            %   Ported from python-docx v1.2.0: src/docx/text/tabstops.py::TabStops.clear_all
            obj.pPr_.remove_tabs_();                   % Python: self._pPr._remove_tabs()
        end
    end
end

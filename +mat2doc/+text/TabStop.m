classdef TabStop < mat2doc.shared.ElementProxy
% TABSTOP An individual tab stop applying to a paragraph or style.
%
%   Accessed using list semantics on its containing TabStops object. An
%   ElementProxy subclass: reference semantics (handle) and H5 element-identity
%   eq/ne are inherited unchanged.
%
%   WRAPPED ELEMENT (tabstops.py 78-80): wraps a single <w:tab> (CT_TabStop).
%   super().__init__(element, None) sets the inherited _element; self._tab
%   (-> tab_) is set to the SAME <w:tab> and is the accessor used by every
%   property. All three properties delegate to the P4-2 CT_TabStop descriptors
%   (val / leader / pos). TabStop adds NO oxml logic (API/proxy tier).
%
%   position SET (tabstops.py 118-123) is the one non-trivial accessor: it
%   RE-INSERTS a fresh <w:tab> at the new pos-order position and removes the old
%   element, so the parent <w:tabs> stays sorted by pos. Only self._tab is
%   reassigned to the new element -- self._element still references the OLD
%   (now-detached) <w:tab>, exactly as in python-docx (faithful quirk: eq()/
%   element() reflect the original element, the three properties reflect the new
%   one). Read/write for all three properties.
%
%   Example:
%       tab = ts.getitem_(0);                          % from a TabStops
%       tab.position                                   % a Length
%       tab.alignment = mat2doc.enum.text.WD_TAB_ALIGNMENT.CENTER;
%       tab.leader = mat2doc.enum.text.WD_TAB_LEADER.DOTS;
%       tab.position = mat2doc.shared.Inches(2);       % re-sorted in place
%
%   Ported from python-docx v1.2.0: src/docx/text/tabstops.py::TabStop

    properties (Access = private)
        tab_   % Python self._tab: the wrapped <w:tab> (CT_TabStop)
    end

    properties (Dependent)
        alignment  % WD_TAB_ALIGNMENT -- w:tab/@w:val (RequiredAttribute)
        leader     % WD_TAB_LEADER -- w:tab/@w:leader (OptionalAttribute, default SPACES)
        position   % Length -- w:tab/@w:pos (RequiredAttribute, signed twips)
    end

    methods
        function obj = TabStop(element)
            % TABSTOP Wrap a <w:tab> element (tabstops.py 78-80).
            %
            %   Inputs:  element - a mat2doc.oxml.text.CT_TabStop (the `w:tab`).
            %   Outputs: obj     - a scalar TabStop handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/text/tabstops.py::TabStop.__init__
            obj@mat2doc.shared.ElementProxy(element, []);  % Python: super().__init__(element, None)
            obj.tab_ = element;                            % Python: self._tab = element
        end

        % ---- alignment (tabstops.py 82-93) ----
        function value = get.alignment(obj)
            % Python: return self._tab.val
            value = obj.tab_.val;
        end
        function set.alignment(obj, value)
            % Python: self._tab.val = value
            obj.tab_.val = value;
        end

        % ---- leader (tabstops.py 95-107) ----
        function value = get.leader(obj)
            % Python: return self._tab.leader
            value = obj.tab_.leader;
        end
        function set.leader(obj, value)
            % Python: self._tab.leader = value (assigning None == assigning SPACES,
            %   resolved inside CT_TabStop's OptionalAttribute w:leader setter).
            obj.tab_.leader = value;
        end

        % ---- position (tabstops.py 109-123) ----
        function value = get.position(obj)
            % Python: return self._tab.pos
            value = obj.tab_.pos;
        end
        function set.position(obj, value)
            % Python (tabstops.py 118-123):
            %   tab = self._tab
            %   tabs = tab.getparent()
            %   self._tab = tabs.insert_tab_in_order(value, tab.val, tab.leader)
            %   tabs.remove(tab)
            %   Only self._tab is reassigned; self._element keeps the old element.
            tab = obj.tab_;
            tabs = tab.getparent();
            obj.tab_ = tabs.insert_tab_in_order(value, tab.val, tab.leader);
            tabs.remove(tab);
        end
    end
end

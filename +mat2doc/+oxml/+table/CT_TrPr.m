classdef CT_TrPr < mat2doc.oxml.BaseOxmlElement
% CT_TRPR `<w:trPr>` element, defining table-row properties.
%
%   Holds the row's grid-before / grid-after skip counts and its row height +
%   height rule. Registered for w:trPr (oxml/__init__.py:184).
%
%   ============================ H11 (child ordering) ============================
%   _tag_seq (table.py 897-913, VERBATIM, 15 tags) is stored as the Constant
%   TAG_SEQ. Each ZeroOrOne descriptor's successors = Python `_tag_seq[s0:]` is
%   expressed as `TAG_SEQ(s0+1:end)` (H1: 0-based Python slice start s0 ->
%   1-based MATLAB start s0+1). Per-descriptor slice (own 1-based idx / Python
%   slice / MATLAB slice / first successor tag):
%     gridBefore = 3 : successors=_tag_seq[3:] -> TAG_SEQ(4:end)  first "w:gridAfter"
%     gridAfter  = 4 : successors=_tag_seq[4:] -> TAG_SEQ(5:end)  first "w:wBefore"
%     trHeight   = 8 : successors=_tag_seq[8:] -> TAG_SEQ(9:end)  first "w:tblHeader"
%   NOTE gridBefore's own tag (1-based idx 3) precedes gridAfter's (idx 4), yet
%   gridBefore is DECLARED after gridAfter in the Python source -- declaration
%   order is irrelevant, only the successor slice drives insertion position, so
%   a scrambled add of gridBefore lands before gridAfter (schema order). A wrong
%   slice would place gridBefore/gridAfter out of order -> Word repair / byte
%   divergence.
%
%   ===================== GENERATED DESCRIPTOR FAMILIES ==========================
%   gridAfter / gridBefore / trHeight are ZeroOrOne (table.py 914-922), generic
%   engine (no _new_x/_insert_x override on CT_TrPr): get.x, get_or_add_x,
%   new_x_, insert_x_, add_x_, remove_x_ (underscore rotation of _new_x/_insert_x/
%   _add_x/_remove_x; get_or_add_x public). The Callable annotation at table.py:895
%   (get_or_add_trHeight) is a type hint only.
%
%   CHILD-CLASS REGISTRATION (this WP closes two P4-6 deferrals):
%     w:gridAfter  -> CT_DecimalNumber (oxml/__init__.py:169)
%     w:gridBefore -> CT_DecimalNumber (oxml/__init__.py:170)
%   grid_after/grid_before read .val on these children (a decimal number), so
%   they MUST resolve to CT_DecimalNumber (registered by this WP). w:trHeight ->
%   CT_Height was registered at P6-1.
%
%   ===================== H3 / H6 (@property members) ============================
%   grid_after / grid_before (table.py 925-935) are READ-ONLY @property: 0 when
%   the child is absent (H3: gridAfter/gridBefore is None -> 0), else child.val
%   (a double int, CT_DecimalNumber.val, H6). trHeight_hRule / trHeight_val
%   (table.py 937-961) get+set: [] (None) when w:trHeight absent; the setter
%   guards `value is None and self.trHeight is None -> return` (do NOT create an
%   empty <w:trHeight> just to assign None), then get_or_add_trHeight + assign
%   (CT_Height.hRule / CT_Height.val, H3/H6/H10).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on any <w:trPr> inside a real table row.
%
%   M1-NEUTRAL: default.docx contains ZERO <w:trPr> (and zero gridAfter/
%   gridBefore) elements, so nothing transits this class on the M1 parse path.
%
%   Example:
%       trPr = mat2doc.oxml.OxmlElement("w:trPr");   % a CT_TrPr
%       trPr.grid_before                             % 0 (no <w:gridBefore>)
%       trPr.trHeight_val = mat2doc.shared.Twips(360);  % <w:trHeight w:val="360"/>
%
%   Ported from python-docx v1.2.0: src/docx/oxml/table.py::CT_TrPr
%   (lines 892-961; registered for w:trPr)

    properties (Constant, Hidden)  % _tag_seq VERBATIM (table.py 897-913; 15 tags)
        TAG_SEQ = [ ...
            "w:cnfStyle", "w:divId", "w:gridBefore", ...        %  1- 3  (gridBefore own @3)
            "w:gridAfter", "w:wBefore", "w:wAfter", ...         %  4- 6  (gridAfter own @4)
            "w:cantSplit", "w:trHeight", "w:tblHeader", ...     %  7- 9  (trHeight own @8)
            "w:tblCellSpacing", "w:jc", "w:hidden", ...         % 10-12
            "w:ins", "w:del", "w:trPrChange" ]                  % 13-15
    end

    properties (Dependent)  % generated ZeroOrOne getters + @property members
        gridAfter      % ZeroOrOne <w:gridAfter> child or [] (read-only; use get_or_add/remove)
        gridBefore     % ZeroOrOne <w:gridBefore> child or []
        trHeight       % ZeroOrOne <w:trHeight> child or []
        grid_after     % @property: 0 when w:gridAfter absent, else its .val (int)
        grid_before    % @property: 0 when w:gridBefore absent, else its .val (int)
        trHeight_hRule % @property: w:trHeight/@w:hRule (WD_ROW_HEIGHT_RULE) or []
        trHeight_val   % @property: w:trHeight/@w:val (Length) or []
    end

    methods
        function obj = CT_TrPr(varargin)
            % CT_TRPR Construct a loose <w:trPr> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ============ gridAfter (ZeroOrOne, successors=_tag_seq[4:] -> TAG_SEQ(5:end)) ============
        function child = get.gridAfter(obj);            child = obj.getChild("w:gridAfter"); end
        function child = get_or_add_gridAfter(obj);     child = obj.getOrAddChild("w:gridAfter", obj.TAG_SEQ(5:end)); end
        function child = new_gridAfter_(obj);           child = obj.newChild("w:gridAfter"); end
        function child = insert_gridAfter_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(5:end)); end
        function child = add_gridAfter_(obj, varargin); child = obj.addChild("w:gridAfter", obj.TAG_SEQ(5:end), varargin{:}); end
        function remove_gridAfter_(obj);                obj.removeChild("w:gridAfter"); end

        % ============ gridBefore (ZeroOrOne, successors=_tag_seq[3:] -> TAG_SEQ(4:end)) ============
        function child = get.gridBefore(obj);            child = obj.getChild("w:gridBefore"); end
        function child = get_or_add_gridBefore(obj);     child = obj.getOrAddChild("w:gridBefore", obj.TAG_SEQ(4:end)); end
        function child = new_gridBefore_(obj);           child = obj.newChild("w:gridBefore"); end
        function child = insert_gridBefore_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(4:end)); end
        function child = add_gridBefore_(obj, varargin); child = obj.addChild("w:gridBefore", obj.TAG_SEQ(4:end), varargin{:}); end
        function remove_gridBefore_(obj);                obj.removeChild("w:gridBefore"); end

        % ============ trHeight (ZeroOrOne, successors=_tag_seq[8:] -> TAG_SEQ(9:end)) ============
        function child = get.trHeight(obj);            child = obj.getChild("w:trHeight"); end
        function child = get_or_add_trHeight(obj);     child = obj.getOrAddChild("w:trHeight", obj.TAG_SEQ(9:end)); end
        function child = new_trHeight_(obj);           child = obj.newChild("w:trHeight"); end
        function child = insert_trHeight_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(9:end)); end
        function child = add_trHeight_(obj, varargin); child = obj.addChild("w:trHeight", obj.TAG_SEQ(9:end), varargin{:}); end
        function remove_trHeight_(obj);                obj.removeChild("w:trHeight"); end

        % ===================== @property members (table.py 925-961) =====================

        % ---- grid_after (read-only, table.py 925-929) ----
        function value = get.grid_after(obj)
            % Python: gridAfter = self.gridAfter; return 0 if None else gridAfter.val
            gridAfter = obj.gridAfter;
            if isequal(gridAfter, [])   % Python: if gridAfter is None (H3)
                value = 0;
                return
            end
            value = gridAfter.val;
        end

        % ---- grid_before (read-only, table.py 931-935) ----
        function value = get.grid_before(obj)
            % Python: gridBefore = self.gridBefore; return 0 if None else gridBefore.val
            gridBefore = obj.gridBefore;
            if isequal(gridBefore, [])   % Python: if gridBefore is None (H3)
                value = 0;
                return
            end
            value = gridBefore.val;
        end

        % ---- trHeight_hRule (get+set, table.py 937-948) ----
        function value = get.trHeight_hRule(obj)
            % Python: trHeight = self.trHeight; return None if None else trHeight.hRule
            trHeight = obj.trHeight;
            if isequal(trHeight, [])   % Python: if trHeight is None (H3)
                value = [];
                return
            end
            value = trHeight.hRule;
        end
        function set.trHeight_hRule(obj, value)
            % Python (table.py 944-948): if value is None and self.trHeight is None:
            %   return; trHeight = self.get_or_add_trHeight(); trHeight.hRule = value
            if isequal(value, []) && isequal(obj.trHeight, [])
                return
            end
            trHeight = obj.get_or_add_trHeight();
            trHeight.hRule = value;
        end

        % ---- trHeight_val (get+set, table.py 950-961) ----
        function value = get.trHeight_val(obj)
            % Python: trHeight = self.trHeight; return None if None else trHeight.val
            trHeight = obj.trHeight;
            if isequal(trHeight, [])   % Python: if trHeight is None (H3)
                value = [];
                return
            end
            value = trHeight.val;
        end
        function set.trHeight_val(obj, value)
            % Python (table.py 957-961): if value is None and self.trHeight is None:
            %   return; trHeight = self.get_or_add_trHeight(); trHeight.val = value
            if isequal(value, []) && isequal(obj.trHeight, [])
                return
            end
            trHeight = obj.get_or_add_trHeight();
            trHeight.val = value;
        end
    end
end

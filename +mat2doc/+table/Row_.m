classdef Row_ < mat2doc.shared.Parented
% ROW_ Table row.
%
%   python-docx `_Row` (table.py 387, FLAG-3 trailing underscore for the
%   leading-underscore proxy class). A parented proxy over a single <w:tr>.
%   `class _Row(Parented)` -> mat2doc.shared.Parented (default handle identity,
%   no eq/ne).
%
%   ATTRIBUTES (table.py 390-393): Python `super().__init__(parent);
%   self._parent = parent; self._tr = self._element = tr` -- two names (_tr and
%   _element) for the SAME <w:tr>. Parented.__init__ stores _parent (the
%   explicit re-assignment is redundant, kept in python-docx). Ported:
%   obj@Parented(parent) sets parent_; tr_ AND element_ both hold the CT_Row.
%
%   MEMBERS (table.py 395-504):
%     cells            (@property) -> STUB P6-4b (a _Cell tuple; _Cell is P6-4b).
%     grid_cols_after  (@property) -> self._tr.grid_after (int, read-only).
%     grid_cols_before (@property) -> self._tr.grid_before (int, read-only).
%     height           (get/set)   -> self._tr.trHeight_val (Length | []; H6/H3).
%     height_rule      (get/set)   -> self._tr.trHeight_hRule (WD_ROW_HEIGHT_RULE
%                                     | []; H10/H3).
%     table            (@property) -> self._parent.table (the owning Table).
%     _index           (@property) -> self._tr.tr_idx (0-based; H1 DATA).
%
%   H3 (None): height/height_rule are [] when absent; assigning [] removes the
%   underlying attribute (delegated to CT_Row/CT_TrPr). H1: _index is 0-based row
%   DATA (Python list.index()), ported via CT_Row.tr_idx (already -1 adjusted).
%   H10: height_rule is a WD_ROW_HEIGHT_RULE member (CT_Row.trHeight_hRule).
%   Underscore rotation: _index -> index_ (PUBLIC method -- python-docx exposes
%   _Row._index and its tests read it).
%
%   Example:
%       row = tbl.rows.getitem_(0);   % a Row_
%       row.height = mat2doc.shared.Pt(20);
%       row.height_rule = mat2doc.enum.table.WD_ROW_HEIGHT_RULE.EXACTLY;
%       row.index_()                  % 0
%
%   Ported from python-docx v1.2.0: src/docx/table.py::_Row (lines 387-504)

    properties (Access = private)
        tr_         % _tr (table.py 393): the wrapped <w:tr> (a mat2doc.oxml.table.CT_Row)
        element_    % _element (table.py 393): same handle; two names for the <w:tr>
    end

    properties (Dependent)
        grid_cols_after     % int (read-only) -- unpopulated grid-cols after the last cell
        grid_cols_before    % int (read-only) -- unpopulated grid-cols before the first cell
        height              % Length | [] -- the row height (read/write)
        height_rule         % WD_ROW_HEIGHT_RULE | [] -- the height rule (read/write)
    end

    methods
        function obj = Row_(tr, parent)
            % ROW_ Wrap a <w:tr> and its parent (table.py 390-393).
            %
            %   Inputs:  tr     - the <w:tr> (a CT_Row).
            %            parent - the collection/Table owning this row (a
            %                     TableParent -- provides `table`/`part`).
            %   Outputs: obj    - a scalar Row_ handle.
            %
            %   Python: super().__init__(parent); self._parent = parent;
            %           self._tr = self._element = tr
            %   Ported from python-docx v1.2.0: src/docx/table.py::_Row.__init__
            obj@mat2doc.shared.Parented(parent);   % Python: super().__init__(parent) + self._parent = parent
            % Python: self._tr = self._element = tr (one element, two names)
            obj.tr_ = tr;
            obj.element_ = tr;
        end

        function c = cells(obj)
            % CELLS The Cell_ sequence for this row (table.py 395-438). UN-STUBBED
            %   at P6-4b. Walks self._tr.tc_lst, generating a Cell_ per layout-grid
            %   cell: a horizontal span (gridSpan) yields the SAME Cell_ handle once
            %   per spanned grid column (H5 identity), and a vertical-merge
            %   CONTINUATION cell (vMerge="continue") delegates to the cell in the
            %   prior row (its _tc_above) so the returned cell holds the span's
            %   actual content. Word allows a row to start late / end early, so only
            %   cells actually present are returned (length may differ per row).
            %
            %   Python (table.py 410-438): two nested generators --
            %     def iter_tc_cells(tc):
            %         if tc.vMerge == "continue":
            %             yield from iter_tc_cells(tc._tc_above); return
            %         cell = _Cell(tc, self.table)
            %         for _ in range(tc.grid_span): yield cell
            %     def _iter_row_cells():
            %         for tc in self._tr.tc_lst: yield from iter_tc_cells(tc)
            %     return tuple(_iter_row_cells())
            %   H9: the generators are materialized (no tree mutation while reading).
            %   iter_tc_cells is ported as the private recursive helper
            %   iter_tc_cells_ below (parent = self.table, table.py 429).
            %
            %   Ported from python-docx v1.2.0: src/docx/table.py::_Row.cells
            c = mat2doc.table.Cell_.empty(1, 0);
            tcs = obj.tr_.tc_lst;                    % Python: for tc in self._tr.tc_lst
            for i = 1:numel(tcs)
                c = [c, obj.iter_tc_cells_(tcs(i))]; %#ok<AGROW> % Python: yield from iter_tc_cells(tc)
            end
        end

        function value = get.grid_cols_after(obj)
            % GRID_COLS_AFTER get (table.py 440-455): count of unpopulated
            %   grid-columns after the last cell. Python: return self._tr.grid_after.
            value = obj.tr_.grid_after;   % Python: return self._tr.grid_after
        end

        function value = get.grid_cols_before(obj)
            % GRID_COLS_BEFORE get (table.py 457-472): count of unpopulated
            %   grid-columns before the first cell. Python: self._tr.grid_before.
            value = obj.tr_.grid_before;   % Python: return self._tr.grid_before
        end

        function value = get.height(obj)
            % HEIGHT get (table.py 474-478): the row height as a Length, or [] if
            %   no explicit height is set. Python: return self._tr.trHeight_val.
            value = obj.tr_.trHeight_val;   % Python: return self._tr.trHeight_val
        end
        function set.height(obj, value)
            % HEIGHT set (table.py 480-482): self._tr.trHeight_val = value (Length
            %   | []). CT_Row handles the H3/H6 write.
            obj.tr_.trHeight_val = value;   % Python: self._tr.trHeight_val = value
        end

        function value = get.height_rule(obj)
            % HEIGHT_RULE get (table.py 484-490): the height rule as a
            %   WD_ROW_HEIGHT_RULE member, or [] if not set. Python: return
            %   self._tr.trHeight_hRule.
            value = obj.tr_.trHeight_hRule;   % Python: return self._tr.trHeight_hRule
        end
        function set.height_rule(obj, value)
            % HEIGHT_RULE set (table.py 492-494): self._tr.trHeight_hRule = value (a
            %   WD_ROW_HEIGHT_RULE member | []). CT_Row handles the H10/H3 write.
            obj.tr_.trHeight_hRule = value;   % Python: self._tr.trHeight_hRule = value
        end

        function value = table(obj)
            % TABLE The Table this row belongs to (table.py 496-499, @property).
            %   Python: return self._parent.table. Property-as-method.
            value = obj.parent_.table;   % Python: return self._parent.table
        end

        function value = index_(obj)
            % _INDEX Index of this row in its table, from zero (table.py 501-504).
            %   Python: return self._tr.tr_idx (already 0-based, H1). Underscore
            %   rotation: _index -> index_ (public method).
            value = obj.tr_.tr_idx;   % Python: return self._tr.tr_idx
        end
    end

    methods (Access = private)
        function cells = iter_tc_cells_(obj, tc)
            % ITER_TC_CELLS_ A Cell_ for each layout-grid cell in `tc` (table.py
            %   410-431 -- the inner generator of _Row.cells, ported as a private
            %   recursive helper). A horizontal span yields the SAME Cell_ handle
            %   once per grid column (H5). A vertical-merge continuation cell
            %   (vMerge="continue") delegates recursively to the guaranteed-same-
            %   width cell directly above (tc._tc_above), so its content resolves to
            %   the span's root cell.
            %
            %   Python (table.py 410-431):
            %     if tc.vMerge == "continue":
            %         yield from iter_tc_cells(tc._tc_above); return
            %     cell = _Cell(tc, self.table)
            %     for _ in range(tc.grid_span): yield cell
            %   H5: the loop yields the ONE `cell` handle grid_span times (identity
            %   across the span). self.table is the owning Table (table.py 429).
            if isequal(tc.vMerge, "continue")   % Python: if tc.vMerge == "continue"
                cells = obj.iter_tc_cells_(tc.tc_above_());   % yield from iter_tc_cells(tc._tc_above)
                return
            end
            cell = mat2doc.table.Cell_(tc, obj.table);   % Python: cell = _Cell(tc, self.table)
            cells = mat2doc.table.Cell_.empty(1, 0);
            for k = 1:tc.grid_span   % Python: for _ in range(tc.grid_span): yield cell
                cells(k) = cell;     % SAME handle at each grid position (H5)
            end
        end
    end
end

classdef Column_ < mat2doc.shared.Parented
% COLUMN_ Table column.
%
%   python-docx `_Column` (table.py 314, FLAG-3 trailing underscore for the
%   leading-underscore proxy class). A parented proxy over a single <w:gridCol>.
%   `class _Column(Parented)` -> mat2doc.shared.Parented (default handle
%   identity, no eq/ne).
%
%   ATTRIBUTES (table.py 317-320): Python `super().__init__(parent);
%   self._parent = parent; self._gridCol = gridCol`. Parented.__init__ stores
%   _parent (the explicit re-assignment is redundant, kept in python-docx).
%   Ported: obj@Parented(parent) sets parent_; gridCol_ holds the CT_TblGridCol.
%
%   MEMBERS (table.py 322-344):
%     cells  (@property)  -> STUB P6-4b (returns a _Cell tuple via
%                            table.column_cells(self._index); _Cell is P6-4b).
%     table  (@property)  -> self._parent.table (the owning Table).
%     width  (get/set)    -> self._gridCol.w (Length | []; H6 EMU, H3 tri-state).
%     _index (@property)  -> self._gridCol.gridCol_idx (0-based; H1 DATA).
%
%   H3 (None): width is [] when @w:w is absent; assigning [] removes @w:w
%   (delegated to CT_TblGridCol.w). H1: _index is 0-based grid DATA (Python
%   list.index()), ported via CT_TblGridCol.gridCol_idx (already -1 adjusted).
%   Underscore rotation: _index -> index_ (PUBLIC method -- python-docx exposes
%   _Column._index and its tests read it).
%
%   Example:
%       col = tbl.columns.getitem_(0);   % a Column_
%       col.width = mat2doc.shared.Inches(1.5);
%       col.index_()                     % 0
%
%   Ported from python-docx v1.2.0: src/docx/table.py::_Column (lines 314-344)

    properties (Access = private)
        gridCol_    % _gridCol (table.py 320): the wrapped <w:gridCol> (a CT_TblGridCol)
    end

    properties (Dependent)
        width    % Length | [] -- the column width in EMU (read/write)
    end

    methods
        function obj = Column_(gridCol, parent)
            % COLUMN_ Wrap a <w:gridCol> and its parent (table.py 317-320).
            %
            %   Inputs:  gridCol - the <w:gridCol> (a CT_TblGridCol).
            %            parent  - the collection/Table owning this column (a
            %                      TableParent -- provides `table`/`part`).
            %   Outputs: obj     - a scalar Column_ handle.
            %
            %   Python: super().__init__(parent); self._parent = parent;
            %           self._gridCol = gridCol
            %   Ported from python-docx v1.2.0: src/docx/table.py::_Column.__init__
            obj@mat2doc.shared.Parented(parent);   % Python: super().__init__(parent) + self._parent = parent
            obj.gridCol_ = gridCol;                 % Python: self._gridCol = gridCol
        end

        function c = cells(obj) %#ok<MANU,STOUT>
            % CELLS STUB (table.py 322-325). Owner: P6-4b (_Cell).
            %   Faithful body: return tuple(self.table.column_cells(self._index)).
            %   column_cells builds _Cell objects, so this defers to P6-4b.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.table.Cell_ (owning WP: P6-4b) required by mat2doc.table.Column_.cells");
        end

        function value = table(obj)
            % TABLE The Table this column belongs to (table.py 327-330, @property).
            %   Python: return self._parent.table. Property-as-method.
            value = obj.parent_.table;   % Python: return self._parent.table
        end

        function value = get.width(obj)
            % WIDTH get (table.py 332-335): the column width in EMU or [] if no
            %   explicit width is set. Python: return self._gridCol.w.
            value = obj.gridCol_.w;   % Python: return self._gridCol.w
        end
        function set.width(obj, value)
            % WIDTH set (table.py 337-339): self._gridCol.w = value. Assigning []
            %   (None) removes @w:w (CT_TblGridCol handles H3/H6).
            obj.gridCol_.w = value;   % Python: self._gridCol.w = value
        end

        function value = index_(obj)
            % _INDEX Index of this column in its table, from zero (table.py
            %   341-344). Python: return self._gridCol.gridCol_idx (already 0-based,
            %   H1). Underscore rotation: _index -> index_ (public method).
            value = obj.gridCol_.gridCol_idx;   % Python: return self._gridCol.gridCol_idx
        end
    end
end

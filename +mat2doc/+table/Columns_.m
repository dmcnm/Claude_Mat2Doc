classdef Columns_ < mat2doc.shared.Parented
% COLUMNS_ Sequence of Column_ instances -- the columns in a table.
%
%   python-docx `_Columns` (table.py 347, FLAG-3 trailing underscore for the
%   leading-underscore proxy class). Supports len(), iteration and indexed
%   access. Accessed via Table.columns (@lazyproperty). In python-docx it is
%   `class _Columns(Parented)` -- a PLAIN parented proxy holding the <w:tbl> and
%   its parent Table; it has no single wrapped element that is its identity, so
%   it derives from mat2doc.shared.Parented (default handle identity, no eq/ne).
%
%   ATTRIBUTES (table.py 353-356): Python `super().__init__(parent);
%   self._parent = parent; self._tbl = tbl`. Parented.__init__ already stores
%   _parent, so the explicit `self._parent = parent` is redundant (kept in
%   python-docx). Ported: obj@Parented(parent) sets parent_; tbl_ holds the
%   CT_Tbl (underscore rotation, design.md section 2).
%
%   VERIFY-COLLECTION (design.md section 2 "Collections -> shared RedefinesParen
%   base"): the shared 1-based () collection base is a FUTURE work package. Per
%   the established Mat2Doc precedent (Sections / TabStops / Styles), the Python
%   Sequence surface is ported as EXPLICIT methods keeping line-for-line
%   fidelity:
%       getitem_ (__getitem__)   to_array (__iter__)   len_ (__len__)
%   Dunder mapping (design.md): `columns[key]` -> columns.getitem_(key);
%   `for c in columns` -> `for c = columns.to_array()`; `len(columns)` ->
%   columns.len_(). FLAGGED for the auditor/validator.
%
%   ============================ H1 (getitem_ index) =============================
%   __getitem__ is INT-ONLY (no slice overload, unlike _Rows). Python
%   (table.py 358-365):
%       try:
%           gridCol = self._gridCol_lst[idx]
%       except IndexError:
%           msg = "column index [%d] is out of range" % idx
%           raise IndexError(msg)
%       return _Column(gridCol, self)
%   getitem_ takes the PYTHON 0-based key; a negative key counts from the end
%   (list indexing wraps). An out-of-range key raises mat2doc:IndexError with the
%   CUSTOM message "column index [%d] is out of range" formatted with the
%   ORIGINAL idx (H1 -- NOT the wrapped value; e.g. columns.getitem_(-9) on a
%   3-column table -> "column index [-9] is out of range"). The `+1` converts the
%   0-based Python position to the 1-based MATLAB one (% IDX).
%
%   H5 (identity): every getitem_/to_array element mints a FRESH Column_ view of
%   its <w:gridCol> (python-docx does not cache Column_ objects); the wrapped
%   CT_TblGridCol is the shared identity. The parent of each Column_ is THIS
%   Columns_ (table.py 365/369: _Column(gridCol, self)).
%
%   Example:
%       cols = tbl.columns;               % a Columns_
%       cols.len_()                       % number of columns
%       c0   = cols.getitem_(0);          % Python columns[0]
%       for c = cols.to_array(); disp(c.width); end
%
%   Ported from python-docx v1.2.0: src/docx/table.py::_Columns (lines 347-384)

    properties (Access = private)
        tbl_    % _tbl (table.py 356): the wrapped <w:tbl> (a mat2doc.oxml.table.CT_Tbl)
    end

    methods
        function obj = Columns_(tbl, parent)
            % COLUMNS_ Wrap the <w:tbl> and its parent Table (table.py 353-356).
            %
            %   Inputs:  tbl    - the <w:tbl> (a CT_Tbl).
            %            parent - the Table this column collection belongs to (a
            %                     TableParent -- provides `table`/`part`).
            %   Outputs: obj    - a scalar Columns_ handle.
            %
            %   Python: super().__init__(parent); self._parent = parent;
            %           self._tbl = tbl
            %   Ported from python-docx v1.2.0: src/docx/table.py::_Columns.__init__
            obj@mat2doc.shared.Parented(parent);   % Python: super().__init__(parent) + self._parent = parent
            obj.tbl_ = tbl;                         % Python: self._tbl = tbl
        end

        function result = getitem_(obj, idx)
            % GETITEM_ Indexed access by int (table.py 358-365). Python 0-based key
            %   with negative-index wrap; out-of-range -> mat2doc:IndexError with
            %   the CUSTOM message (see class header H1). H1: `+1` for the 1-based
            %   MATLAB list.
            %
            %   Ported from python-docx v1.2.0: src/docx/table.py::_Columns.__getitem__
            lst = obj.gridCol_lst_();   % Python: self._gridCol_lst
            n = numel(lst);
            i = idx;                    % Python 0-based index
            if i < 0                    % Python negative-index wrap
                i = i + n;
            end
            if i < 0 || i >= n          % Python list out-of-range -> caught IndexError
                % Python: msg = "column index [%d] is out of range" % idx (ORIGINAL idx)
                msg = "column index [" + mat2doc.shared.pyStr(idx, "int") + "] is out of range";
                error("mat2doc:IndexError", "%s", msg);
            end
            % Python: return _Column(gridCol, self)
            result = mat2doc.table.Column_(lst(i + 1), obj);   % IDX
        end

        function result = to_array(obj)
            % TO_ARRAY A Column_ per <w:gridCol>, in document order (table.py
            %   367-369). Python __iter__:
            %     for gridCol in self._gridCol_lst:
            %         yield _Column(gridCol, self)
            %   Materialized (H9) into a 1xN Column_ array; no gridCols -> a 1x0
            %   Column_ array. Iteration idiom: `for c in columns` ->
            %   `for c = columns.to_array()`.
            %
            %   Ported from python-docx v1.2.0: src/docx/table.py::_Columns.__iter__
            lst = obj.gridCol_lst_();
            result = mat2doc.table.Column_.empty(1, 0);
            for k = 1:numel(lst)   % Python: for gridCol in self._gridCol_lst
                result(k) = mat2doc.table.Column_(lst(k), obj);   % _Column(gridCol, self)
            end
        end

        function n = len_(obj)
            % LEN_ Number of columns (table.py 371-372).
            %   Python __len__: return len(self._gridCol_lst).
            %
            %   Ported from python-docx v1.2.0: src/docx/table.py::_Columns.__len__
            n = numel(obj.gridCol_lst_());
        end

        function value = table(obj)
            % TABLE The Table this column collection belongs to (table.py 374-377,
            %   @property). Python: return self._parent.table. Property-as-method.
            value = obj.parent_.table;   % Python: return self._parent.table
        end
    end

    methods (Access = private)
        function lst = gridCol_lst_(obj)
            % _GRIDCOL_LST The <w:gridCol> elements for this table (table.py
            %   379-384, @property). Python: tblGrid = self._tbl.tblGrid; return
            %   tblGrid.gridCol_lst. Underscore rotation: _gridCol_lst ->
            %   gridCol_lst_ (private -- internal helper only).
            tblGrid = obj.tbl_.tblGrid;   % Python: tblGrid = self._tbl.tblGrid
            lst = tblGrid.gridCol_lst;    % Python: return tblGrid.gridCol_lst
        end
    end
end

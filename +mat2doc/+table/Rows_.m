classdef Rows_ < mat2doc.shared.Parented
% ROWS_ Sequence of Row_ objects -- the rows in a table.
%
%   python-docx `_Rows` (table.py 507, FLAG-3 trailing underscore for the
%   leading-underscore proxy class). Supports len(), iteration, indexed access
%   AND slicing. Accessed via Table.rows (@lazyproperty). In python-docx it is
%   `class _Rows(Parented)` -- a PLAIN parented proxy holding the <w:tbl> and its
%   parent Table; it has no single wrapped element that is its identity, so it
%   derives from mat2doc.shared.Parented (default handle identity, no eq/ne).
%
%   ATTRIBUTES (table.py 513-516): Python `super().__init__(parent);
%   self._parent = parent; self._tbl = tbl`. Parented.__init__ already stores
%   _parent (the explicit re-assignment is redundant, kept in python-docx).
%   Ported: obj@Parented(parent) sets parent_; tbl_ holds the CT_Tbl.
%
%   VERIFY-COLLECTION (design.md section 2): as with Columns_/Sections, the
%   Python Sequence surface is ported as EXPLICIT methods (the shared 1-based ()
%   RedefinesParen base is a future WP):
%       getitem_ (__getitem__)   to_array (__iter__)   len_ (__len__)
%   Dunder mapping: `rows[key]` -> rows.getitem_(key); `for r in rows` ->
%   `for r = rows.to_array()`; `len(rows)` -> rows.len_(). FLAGGED for auditor.
%
%   ===================== __getitem__ (INT + SLICE, table.py 518-526) ============
%   Python `def __getitem__(self, idx): return list(self)[idx]` with @overload
%   int -> _Row and slice -> list[_Row]. list(self) materializes all Row_ objects
%   then indexes:
%     * INT key: Python 0-based with negative-index wrap; out-of-range raises the
%       standard CPython "list index out of range" (mat2doc:IndexError). H1: `+1`.
%     * SLICE key: represented as a STRUCT with fields start/stop/step (each a
%       scalar double or [] for None) -- the interim currency until the
%       RedefinesParen base lands. sliceIndices_ is the FAITHFUL port of CPython
%       slice.indices(n) + range(...) (identical to Sections.sliceIndices_);
%       returns a 1xN Row_ array (empty slice -> 1x0).
%
%   H5 (identity): every getitem_/to_array element mints a FRESH Row_ view of its
%   <w:tr> (python-docx does not cache Row_ objects); the wrapped CT_Row is the
%   shared identity. The parent of each Row_ is THIS Rows_ (table.py 526/529:
%   list(self) -> _Row(tr, self)).
%
%   Example:
%       rows = tbl.rows;                  % a Rows_
%       rows.len_()                       % number of rows
%       r0   = rows.getitem_(0);          % Python rows[0]
%       mid  = rows.getitem_(struct("start",1,"stop",3,"step",[]));  % rows[1:3]
%       for r = rows.to_array(); disp(r.height); end
%
%   Ported from python-docx v1.2.0: src/docx/table.py::_Rows (lines 507-537)

    properties (Access = private)
        tbl_    % _tbl (table.py 516): the wrapped <w:tbl> (a mat2doc.oxml.table.CT_Tbl)
    end

    methods
        function obj = Rows_(tbl, parent)
            % ROWS_ Wrap the <w:tbl> and its parent Table (table.py 513-516).
            %
            %   Inputs:  tbl    - the <w:tbl> (a CT_Tbl).
            %            parent - the Table this row collection belongs to (a
            %                     TableParent -- provides `table`/`part`).
            %   Outputs: obj    - a scalar Rows_ handle.
            %
            %   Python: super().__init__(parent); self._parent = parent;
            %           self._tbl = tbl
            %   Ported from python-docx v1.2.0: src/docx/table.py::_Rows.__init__
            obj@mat2doc.shared.Parented(parent);   % Python: super().__init__(parent) + self._parent = parent
            obj.tbl_ = tbl;                         % Python: self._tbl = tbl
        end

        function result = getitem_(obj, key)
            % GETITEM_ Indexed access by int OR slice (table.py 524-526). Python:
            %   `return list(self)[idx]`. INT key is Python 0-based (negative wrap;
            %   out-of-range "list index out of range"). SLICE key is a
            %   struct(start,stop,step). H1: `+1` converts the 0-based Python
            %   position to the 1-based MATLAB one.
            %
            %   Ported from python-docx v1.2.0: src/docx/table.py::_Rows.__getitem__
            tr_lst = obj.tbl_.tr_lst;   % LIVE CT_Tbl.tr_lst (the list(self) source)
            if isstruct(key)            % Python: slice key
                idxs = mat2doc.table.Rows_.sliceIndices_(key, numel(tr_lst));  % 1-based
                result = mat2doc.table.Row_.empty(1, 0);
                for j = 1:numel(idxs)   % Python list slice
                    result(j) = mat2doc.table.Row_(tr_lst(idxs(j)), obj);
                end
                return
            end
            % --- INT overload: Python list(self)[idx] (0-based list indexing) ---
            n = numel(tr_lst);
            i = key;                    % Python 0-based index
            if i < 0                    % Python negative-index wrap
                i = i + n;
            end
            if i < 0 || i >= n          % Python list out-of-range
                error("mat2doc:IndexError", "%s", "list index out of range");
            end
            % Python: _Row(tr, self) for the selected tr
            result = mat2doc.table.Row_(tr_lst(i + 1), obj);   % IDX
        end

        function result = to_array(obj)
            % TO_ARRAY A Row_ per <w:tr>, in document order (table.py 528-529).
            %   Python __iter__:
            %     return (_Row(tr, self) for tr in self._tbl.tr_lst)
            %   Materialized (H9) into a 1xN Row_ array; no trs -> a 1x0 Row_ array.
            %   Iteration idiom: `for r in rows` -> `for r = rows.to_array()`.
            %
            %   Ported from python-docx v1.2.0: src/docx/table.py::_Rows.__iter__
            tr_lst = obj.tbl_.tr_lst;
            result = mat2doc.table.Row_.empty(1, 0);
            for k = 1:numel(tr_lst)   % Python: for tr in self._tbl.tr_lst
                result(k) = mat2doc.table.Row_(tr_lst(k), obj);   % _Row(tr, self)
            end
        end

        function n = len_(obj)
            % LEN_ Number of rows (table.py 531-532).
            %   Python __len__: return len(self._tbl.tr_lst).
            %
            %   Ported from python-docx v1.2.0: src/docx/table.py::_Rows.__len__
            n = numel(obj.tbl_.tr_lst);
        end

        function value = table(obj)
            % TABLE The Table this row collection belongs to (table.py 534-537,
            %   @property). Python: return self._parent.table. Property-as-method.
            value = obj.parent_.table;   % Python: return self._parent.table
        end
    end

    methods (Static, Access = private)
        function idxs = sliceIndices_(sl, n)
            % SLICEINDICES_ 1-based MATLAB positions selected by a Python slice.
            %   FAITHFUL port of CPython slice.indices(length) + range(start, stop,
            %   step) -- byte-for-byte identical to mat2doc.section.Sections's
            %   private helper (same interim struct-slice currency). `sl` is a
            %   struct with fields start/stop/step, each a scalar double or []
            %   (None). Realizes list(self)[key] for a slice key (table.py 524-526).
            %   Returns a 1xM double of 1-based indices (M may be 0).

            % step (default 1); step == 0 is a ValueError, as in CPython.
            if isfield(sl, "step") && ~isequal(sl.step, [])
                step = sl.step;
            else
                step = 1;
            end
            if step == 0
                error("mat2doc:ValueError", "%s", "slice step cannot be zero");
            end

            % lower/upper bounds depend on step sign (CPython slice.indices).
            if step > 0
                lower = 0;      upper = n;
            else
                lower = -1;     upper = n - 1;
            end

            % start
            if ~isfield(sl, "start") || isequal(sl.start, [])   % None
                if step < 0, start = upper; else, start = lower; end
            else
                start = sl.start;
                if start < 0
                    start = max(start + n, lower);
                else
                    start = min(start, upper);
                end
            end

            % stop
            if ~isfield(sl, "stop") || isequal(sl.stop, [])     % None
                if step < 0, stop = lower; else, stop = upper; end
            else
                stop = sl.stop;
                if stop < 0
                    stop = max(stop + n, lower);
                else
                    stop = min(stop, upper);
                end
            end

            % range(start, stop, step): Python EXCLUDES stop; MATLAB colon is
            % inclusive, so cap at stop-/+1 (integers). Then 0-based -> +1 (IDX).
            if step > 0
                zeroBased = start : step : (stop - 1);
            else
                zeroBased = start : step : (stop + 1);
            end
            idxs = zeroBased + 1;
        end
    end
end

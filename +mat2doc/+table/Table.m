classdef Table < mat2doc.shared.StoryChild
% TABLE Proxy class for a WordprocessingML `<w:tbl>` element.
%
%   The public API tier over the oxml table layer (CT_Tbl, P6-3b). Adds NO oxml
%   logic, NO registry rows and NO serialization code: every accessor delegates
%   to the wrapped CT_Tbl / CT_TblPr (P6-2/P6-3) or up to the DocumentPart
%   (style resolution). Equivalence is BEHAVIORAL (proxy return values +
%   serialized bytes) except the add_table authoring path, whose constructed
%   document.xml is a byte oracle (new_tbl is byte-proven P6-3b).
%
%   TIER (table.py 29 `class Table(StoryChild)`): StoryChild < handle (the
%   parent-only tier whose `part` delegates up to a StoryPart, P2-1). So Table
%   derives from mat2doc.shared.StoryChild -- NOT ElementProxy -- and, like
%   StoryChild/Paragraph, does NOT define eq/ne (default handle identity ==
%   Python default object identity; a Table is NOT compared by wrapped-element
%   identity).
%
%   ATTRIBUTES (table.py 32-35): Python `self._element = tbl; self._tbl = tbl` --
%   two names for the SAME <w:tbl> element (the Paragraph p_/element_ precedent).
%   Ported verbatim: element_ (Python _element, used by the table_direction
%   setter) and tbl_ (Python _tbl, the working handle used by every other
%   method). Both hold the same CT_Tbl handle.
%
%   LAZYPROPERTY (table.py 99-102 columns / 114-117 rows): @lazyproperty -> the
%   Dependent read-only property + private cache + computed-flag idiom
%   (mat2doc.shared.lazyproperty; NEVER isempty as the sentinel, H3). columns ->
%   Columns_(self._tbl, self); rows -> Rows_(self._tbl, self), cached on first
%   access and returned unchanged thereafter.
%
%   ========================= A2 CROSS-ENUM (alignment) ==========================
%   alignment get -> self._tblPr.alignment (CT_TblPr.alignment, table.py 65). Per
%   the A2 rule (design.md section 2) CT_TblPr.alignment returns a paragraph-
%   alignment member (CT_Jc.val) -- at runtime a mat2doc.enum.text.
%   WD_PARAGRAPH_ALIGNMENT instance (WD_ALIGN_PARAGRAPH is an alias class of that
%   canonical enum), exactly as Paragraph.alignment returns -- and python-docx's
%   cast(WD_TABLE_ALIGNMENT, ...) is a runtime NO-OP. This port returns that
%   member VERBATIM -- it does NOT convert to WD_TABLE_ALIGNMENT (a converting
%   getter would crash on the legal <w:jc w:val="both"> that WD_TABLE_ALIGNMENT
%   cannot represent). CONSEQUENCE: a MATLAB user's
%   `table.alignment == WD_TABLE_ALIGNMENT.CENTER` is FALSE where Python
%   (int-subclass duck-equality) is True; the NAME ("CENTER"), int value (1) and
%   serialized bytes ("center") are all identical. Binding user idiom: compare by
%   NAME (`== "CENTER"`) or `.value`, never cross-class `==`. NOT a D-number (no
%   byte/output divergence; consistent with the ratified section-2 A2 note). The
%   setter accepts a WD_TABLE_ALIGNMENT value and writes it through CT_TblPr (the
%   cross-enum to_xml resolves by int value -> correct "left"/"center"/"right").
%
%   table_direction (table.py 150-161): get -> self._tbl.bidiVisual_val (WD_TABLE_
%   DIRECTION is a REAL distinct enum, no A2 issue). python-docx's
%   cast(WD_TABLE_DIRECTION, ...) is again a runtime no-op; bidiVisual_val returns
%   [] (None) or a LOGICAL (CT_OnOff.val), so this getter returns that verbatim
%   (as python-docx returns the bool cast). set -> self._element.bidiVisual_val =
%   value (a WD_TABLE_DIRECTION member; CT_Tbl handles bool(value) H4/H10).
%
%   ========================= P6-4a vs P6-4b BOUNDARY ============================
%   LIVE (P6-4a, this WP): alignment (A2), autofit, columns, rows, style,
%   table (returns self), table_direction, _column_count (column_count_),
%   _tblPr (tblPr_). Everything reading the table's read structure ports fully.
%   STUB -> P6-4b (need _Cell, which is P6-4b): add_column, add_row (the
%   mutators), cell, column_cells, row_cells, _cells. Each raises a clean
%   mat2doc:notYetPorted naming its owner (P6-4b). _Cell itself is NOT ported in
%   this WP (no Cell_.m file); it lands with P6-4b.
%
%   UNDERSCORE ROTATION (design.md section 2): _element -> element_, _tbl -> tbl_,
%   _column_count -> column_count_, _tblPr -> tblPr_. Python single-underscore
%   @property members `_column_count` / `_tblPr` are notionally-internal but
%   accessible (python-docx tests read them), so they are ported as PUBLIC
%   zero-arg methods (property-as-method house convention).
%
%   Example:
%       d   = mat2doc.Document();
%       tbl = d.add_table(2, 3, mat2doc.shared.Inches(6));  % a Table
%       tbl.alignment = mat2doc.enum.table.WD_TABLE_ALIGNMENT.CENTER;
%       n   = tbl.rows.len_();          % number of rows (2)
%       c0  = tbl.columns.getitem_(0);  % Python table.columns[0] -> a Column_
%       tbl.autofit = false;            % <w:tblLayout w:type="fixed"/>
%
%   Ported from python-docx v1.2.0: src/docx/table.py::Table (lines 29-189)

    properties (Access = private)
        element_    % _element (table.py 34): the <w:tbl> (a mat2doc.oxml.table.CT_Tbl)
        tbl_        % _tbl (table.py 35): same handle; the working <w:tbl>
        % @lazyproperty caches for columns/rows (table.py 99/114). Manual
        % computed-flag caching (design.md @lazyproperty rule; NEVER isempty as
        % the sentinel -- a Columns_/Rows_ is always a valid non-empty value).
        columns_cache_
        columns_computed_ (1,1) logical = false
        rows_cache_
        rows_computed_ (1,1) logical = false
    end

    properties (Dependent)
        alignment        % WD_ALIGN_PARAGRAPH member (A2 cross-enum) | [] -- read/write
        autofit          % logical -- True unless a fixed table layout is set; read/write
        columns          % Columns_ (read-only, @lazyproperty cached) -- the columns
        rows             % Rows_ (read-only, @lazyproperty cached) -- the rows
        style            % TableStyle_ | [] -- the applied table style; read/write
        table            % Table (read-only) -- returns self (child-parent terminus)
        table_direction  % WD_TABLE_DIRECTION | [] -- cell-ordering direction; read/write
    end

    methods
        function obj = Table(tbl, parent)
            % TABLE Wrap a `<w:tbl>` element (table.py 32-35).
            %
            %   Inputs:  tbl    - a mat2doc.oxml.table.CT_Tbl (the `w:tbl` element).
            %            parent - the parent proxy (a ProvidesStoryPart) providing
            %                     `part`.
            %   Outputs: obj    - a scalar Table handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/table.py::Table.__init__
            obj@mat2doc.shared.StoryChild(parent);   % Python: super().__init__(parent)
            % Python: self._element = tbl; self._tbl = tbl (one element, two names)
            obj.element_ = tbl;
            obj.tbl_ = tbl;
        end

        % ============================ add_column (STUB -> P6-4b) ============================
        function col = add_column(obj, width) %#ok<INUSD,STOUT>
            % ADD_COLUMN STUB (table.py 37-45). Owner: P6-4b (the mutators).
            %   Faithful body: tblGrid = self._tbl.tblGrid; gridCol =
            %   tblGrid.add_gridCol(); gridCol.w = width; for tr in self._tbl.tr_lst:
            %   tc = tr.add_tc(); tc.width = width; return _Column(gridCol, self).
            %   CT_Tc.width and the returned _Cell-free _Column exist, but add_column
            %   is grouped with the P6-4b mutators per the WP split.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.table.Table.add_column (owning WP: P6-4b table mutators)");
        end

        % ============================ add_row (STUB -> P6-4b) ============================
        function row = add_row(obj) %#ok<MANU,STOUT>
            % ADD_ROW STUB (table.py 47-55). Owner: P6-4b (the mutators).
            %   Faithful body: tbl = self._tbl; tr = tbl.add_tr(); for gridCol in
            %   tbl.tblGrid.gridCol_lst: tc = tr.add_tc(); if gridCol.w is not None:
            %   tc.width = gridCol.w; return _Row(tr, self).
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.table.Table.add_row (owning WP: P6-4b table mutators)");
        end

        % ============================ alignment (A2 cross-enum) ============================
        function value = get.alignment(obj)
            % ALIGNMENT get (table.py 57-65): return self._tblPr.alignment. A2: the
            %   CT_TblPr getter returns a WD_ALIGN_PARAGRAPH member (or []); this
            %   port passes it through VERBATIM (see class header A2 note).
            value = obj.tblPr_().alignment;   % Python: return self._tblPr.alignment
        end
        function set.alignment(obj, value)
            % ALIGNMENT set (table.py 67-69): self._tblPr.alignment = value.
            obj.tblPr_().alignment = value;   % Python: self._tblPr.alignment = value
        end

        % ============================ autofit ============================
        function value = get.autofit(obj)
            % AUTOFIT get (table.py 71-79): return self._tblPr.autofit (logical).
            value = obj.tblPr_().autofit;     % Python: return self._tblPr.autofit
        end
        function set.autofit(obj, value)
            % AUTOFIT set (table.py 81-83): self._tblPr.autofit = value (bool).
            obj.tblPr_().autofit = value;     % Python: self._tblPr.autofit = value
        end

        % ============================ cell (STUB -> P6-4b) ============================
        function c = cell(obj, row_idx, col_idx) %#ok<INUSD,STOUT>
            % CELL STUB (table.py 85-91). Owner: P6-4b (_Cell).
            %   Faithful body: cell_idx = col_idx + (row_idx * self._column_count);
            %   return self._cells[cell_idx]. Needs _cells (a _Cell sequence), P6-4b.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.table.Cell_ (owning WP: P6-4b) required by mat2doc.table.Table.cell");
        end

        % ============================ column_cells (STUB -> P6-4b) ============================
        function cells = column_cells(obj, column_idx) %#ok<INUSD,STOUT>
            % COLUMN_CELLS STUB (table.py 93-97). Owner: P6-4b (_Cell).
            %   Faithful body walks self._cells at self._column_count stride. Needs
            %   _cells (a _Cell sequence), P6-4b.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.table.Cell_ (owning WP: P6-4b) required by mat2doc.table.Table.column_cells");
        end

        % ============================ columns (@lazyproperty) ============================
        function value = get.columns(obj)
            % COLUMNS get (table.py 99-102, @lazyproperty): a Columns_ over this
            %   table's <w:gridCol> sequence. Python: return _Columns(self._tbl,
            %   self). Cached on first access (computed-flag idiom); read-only.
            if ~obj.columns_computed_
                obj.columns_cache_ = mat2doc.table.Columns_(obj.tbl_, obj);
                obj.columns_computed_ = true;
            end
            value = obj.columns_cache_;
        end

        % ============================ row_cells (STUB -> P6-4b) ============================
        function cells = row_cells(obj, row_idx) %#ok<INUSD,STOUT>
            % ROW_CELLS STUB (table.py 104-112). Owner: P6-4b (_Cell).
            %   DEPRECATED in python-docx (use table.rows[row_idx].cells). Faithful
            %   body slices self._cells[start:end]. Needs _cells (a _Cell sequence),
            %   P6-4b.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.table.Cell_ (owning WP: P6-4b) required by mat2doc.table.Table.row_cells");
        end

        % ============================ rows (@lazyproperty) ============================
        function value = get.rows(obj)
            % ROWS get (table.py 114-117, @lazyproperty): a Rows_ over this table's
            %   <w:tr> sequence. Python: return _Rows(self._tbl, self). Cached on
            %   first access (computed-flag idiom); read-only.
            if ~obj.rows_computed_
                obj.rows_cache_ = mat2doc.table.Rows_(obj.tbl_, obj);
                obj.rows_computed_ = true;
            end
            value = obj.rows_cache_;
        end

        % ============================ style (LIVE delegation, P4-7a styles) ============================
        function value = get.style(obj)
            % STYLE get (table.py 119-133): style_id = self._tbl.tblStyle_val; then
            %   self.part.get_style(style_id, WD_STYLE_TYPE.TABLE). The default table
            %   style is returned when no style is directly applied. part is the
            %   DocumentPart whose get_style is LIVE at P4-7a. cast is a no-op.
            style_id = obj.tbl_.tblStyle_val;              % Python: style_id = self._tbl.tblStyle_val
            value = obj.part().get_style( ...              % Python: self.part.get_style(...)
                style_id, mat2doc.enum.style.WD_STYLE_TYPE.TABLE);
        end
        function set.style(obj, style_or_name)
            % STYLE set (table.py 135-138): style_id = self.part.get_style_id(value,
            %   WD_STYLE_TYPE.TABLE); self._tbl.tblStyle_val = style_id. Assigning []
            %   (None) removes any directly-applied style. get_style_id is LIVE at P4-7a.
            style_id = obj.part().get_style_id( ...        % Python: self.part.get_style_id(...)
                style_or_name, mat2doc.enum.style.WD_STYLE_TYPE.TABLE);
            obj.tbl_.tblStyle_val = style_id;              % Python: self._tbl.tblStyle_val = style_id
        end

        % ============================ table ============================
        function value = get.table(obj)
            % TABLE get (table.py 140-148): the terminus of a child's parent._table
            %   chain -- returns THIS Table. Python: return self.
            value = obj;   % Python: return self
        end

        % ============================ table_direction ============================
        function value = get.table_direction(obj)
            % TABLE_DIRECTION get (table.py 150-157): cell-ordering direction, or []
            %   when inherited. Python: return cast("WD_TABLE_DIRECTION | None",
            %   self._tbl.bidiVisual_val). The cast is a runtime no-op; bidiVisual_val
            %   is [] or a logical (CT_OnOff.val), returned verbatim.
            value = obj.tbl_.bidiVisual_val;   % Python: self._tbl.bidiVisual_val
        end
        function set.table_direction(obj, value)
            % TABLE_DIRECTION set (table.py 159-161): self._element.bidiVisual_val =
            %   value (a WD_TABLE_DIRECTION member or []). CT_Tbl applies bool(value)
            %   (LTR=0 -> "0", RTL=1 -> bare) H4/H10.
            obj.element_.bidiVisual_val = value;   % Python: self._element.bidiVisual_val = value
        end
    end

    methods  % property-as-method ports of the single-underscore @property members
        function cells = cells_(obj) %#ok<MANU,STOUT>
            % _CELLS STUB (table.py 163-180). Owner: P6-4b (_Cell).
            %   A _Cell per layout-grid cell (repeated references for spans). The
            %   whole body constructs _Cell objects, so it defers to P6-4b.
            %   Underscore rotation: _cells -> cells_.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.table.Cell_ (owning WP: P6-4b) required by mat2doc.table.Table._cells");
        end

        function n = column_count_(obj)
            % _COLUMN_COUNT The number of grid columns in this table (table.py
            %   182-185). Python: return self._tbl.col_count. Public method
            %   (python-docx exposes _column_count and its tests read it).
            %   Underscore rotation: _column_count -> column_count_.
            n = obj.tbl_.col_count;   % Python: return self._tbl.col_count
        end

        function p = tblPr_(obj)
            % _TBLPR The <w:tblPr> of this table (table.py 187-189). Python: return
            %   self._tbl.tblPr (a CT_TblPr; OneAndOnlyOne, created if absent by the
            %   getter). Underscore rotation: _tblPr -> tblPr_.
            p = obj.tbl_.tblPr;   % Python: return self._tbl.tblPr
        end
    end
end

classdef Cell_ < mat2doc.BlockItemContainer
% CELL_ Table cell.
%
%   python-docx `_Cell` (table.py 192-311, FLAG-3 trailing underscore for the
%   leading-underscore proxy class). `class _Cell(BlockItemContainer)` ->
%   mat2doc.BlockItemContainer (the block-item container tier, P2-3/P4-7b). A
%   cell wraps a `<w:tc>` and, being a BlockItemContainer, contributes the
%   paragraph/table add+read surface plus the cell-specific members
%   (grid_span/merge/text/vertical_alignment/width).
%
%   ATTRIBUTES (table.py 195-198): Python
%     super(_Cell, self).__init__(tc, parent)   # BlockItemContainer: _element=tc, _parent=parent
%     self._parent = parent                     # redundant (already set by super)
%     self._tc = self._element = tc             # two names for the SAME <w:tc>
%   Ported: obj@mat2doc.BlockItemContainer(tc, parent) sets element_store_ = tc
%   (the C3 seam's concrete store) and parent_ = the `parent` argument (via StoryChild).
%   tc_ additionally holds the same CT_Tc handle (Python self._tc). So _tc and
%   _element are the SAME tc: the members reading Python's `self._tc`
%   (grid_span/merge/text-set/width) use obj.tc_, and the members reading
%   Python's `self._element` (vertical_alignment) route through the inherited
%   element_() C3 seam (obj.element_() == element_store_ == tc) -- both are the
%   one <w:tc>.
%
%   =========================== C3 ELEMENT SEAM (inherited) ======================
%   Unlike _BaseHeaderFooter, _Cell stores a CONCRETE element (the tc) and does
%   NOT override the element_() seam -- BlockItemContainer's base element_()
%   returning element_store_ is exactly right (the tc is passed to the ctor). The
%   inherited BlockItemContainer surface therefore operates directly on the tc:
%   add_paragraph_/paragraphs/add_table read tc via the seam; tables reads
%   tc.tbl_lst.
%
%   MEMBERS (table.py 200-311):
%     add_paragraph(text, style)  -> super().add_paragraph(text, style) (docstring
%                                    override only; behavior == BlockItemContainer).
%     add_table(rows, cols)       -> a NESTED table: width defaults to the cell's
%                                    width (Inches(1) when the cell has none), then
%                                    super().add_table(rows, cols, width); an empty
%                                    <w:p> is appended after (Word requires a
%                                    paragraph as a cell's last block child).
%     grid_span     (@property)   -> self._tc.grid_span (columns spanned; 1 normal).
%     merge(other)                -> self._tc.merge(other._tc) (CT_Tc.merge, the
%                                    byte-proven P6-3a engine); returns a NEW Cell_
%                                    over the merged tc, parented to self._parent.
%     paragraphs    (@property)   -> super().paragraphs (Paragraph list).
%     tables        (@property)   -> super().tables (Table list; the base
%                                    BlockItemContainer.tables is un-stubbed at
%                                    P6-4b now that Table is live).
%     text          (get/set)     -> get: "\n".join(p.text for p in self.paragraphs)
%                                    (NO strip -- H16 latent, not triggered here);
%                                    set: clear the tc's content, add one <w:p>/<w:r>
%                                    holding `text`.
%     vertical_alignment (get/set)-> ./w:tcPr/w:vAlign/@w:val
%                                    (WD_CELL_VERTICAL_ALIGNMENT | []; H10/H3).
%     width         (get/set)     -> self._tc.width (EMU Length | []; H6/H3).
%
%   ======================= H5 (span identity -- read via the grid) ==============
%   Cell_ is a handle class, so the SAME Cell_ handle appears at each grid
%   position a spanned cell covers (Table._cells / _Row.cells build the grid by
%   repeating one handle). Two accesses of a spanned position compare `==` (handle
%   identity), matching lxml's same-tc-same-_Cell semantics.
%
%   H3 (None): vertical_alignment/width are [] when absent; assigning [] removes
%   the underlying attr (delegated to CT_Tc/CT_TcPr). H16 (strip): the text getter
%   does NOT strip (python-docx `"\n".join(...)`, no `.strip()`), so the MATLAB
%   strip-set hazard is latent here -- newline join only.
%
%   Example:
%       t = mat2doc.Document().add_table(2, 2, mat2doc.shared.Inches(4));
%       c = t.cell(0, 0);
%       c.text = "A";                          % one <w:p>/<w:r> with "A"
%       m = t.cell(0, 0).merge(t.cell(0, 1));  % horizontal merge -> a merged Cell_
%       c.vertical_alignment = mat2doc.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER;
%
%   Ported from python-docx v1.2.0: src/docx/table.py::_Cell (lines 192-311)

    properties (Access = private)
        tc_   % _tc (table.py 198): the wrapped <w:tc> (a CT_Tc); == _element
    end

    properties (Dependent)
        grid_span           % int (read-only) -- columns this cell spans (1 default)
        text                % string -- get concatenates paragraphs (\n); set replaces
        vertical_alignment  % WD_CELL_VERTICAL_ALIGNMENT | [] -- read/write
        width               % Length | [] -- the cell width in EMU; read/write
    end

    methods
        function obj = Cell_(tc, parent)
            % CELL_ Wrap a `<w:tc>` and its parent (table.py 195-198).
            %
            %   Inputs:  tc     - the `<w:tc>` (a mat2doc.oxml.table.CT_Tc).
            %            parent - the parent proxy (a TableParent -- a Table /
            %                     Rows_ / Columns_; provides `part`/`table`).
            %   Outputs: obj    - a scalar Cell_ handle.
            %
            %   Python: super(_Cell, self).__init__(tc, parent); self._parent =
            %           parent; self._tc = self._element = tc
            %   Ported from python-docx v1.2.0: src/docx/table.py::_Cell.__init__
            obj@mat2doc.BlockItemContainer(tc, parent);   % super().__init__: _element=tc, _parent=parent
            obj.tc_ = tc;                                  % Python: self._tc = self._element = tc
        end

        function paragraph = add_paragraph(obj, text, style)
            % ADD_PARAGRAPH A paragraph newly added to the end of this cell's
            %   content (table.py 200-211). Delegates VERBATIM to the
            %   BlockItemContainer implementation (python-docx re-declares this only
            %   to enrich the docstring: `return super(_Cell, self).add_paragraph(
            %   text, style)`). H13: defaults text="" / style=[] (None).
            %
            %   Ported from python-docx v1.2.0: src/docx/table.py::_Cell.add_paragraph
            arguments
                obj
                text  = ""   % Python default ""
                style = []   % Python default None
            end
            % Python: return super(_Cell, self).add_paragraph(text, style)
            paragraph = add_paragraph@mat2doc.BlockItemContainer(obj, text, style);
        end

        function t = add_table(obj, rows, cols)
            % ADD_TABLE A NESTED table newly added to this cell after any existing
            %   content (table.py 213-226). The new table has `rows` rows and `cols`
            %   columns; its width defaults to the cell's width, or Inches(1) when
            %   the cell has no explicit width. An empty paragraph is appended after
            %   the table (Word requires a <w:p> as a cell's last block child).
            %
            %   Python (table.py 223-226):
            %     width = self.width if self.width is not None else Inches(1)
            %     table = super(_Cell, self).add_table(rows, cols, width)
            %     self.add_paragraph()
            %     return table
            %   H3: `self.width is not None` -> ~isequal(width, []).
            %
            %   Ported from python-docx v1.2.0: src/docx/table.py::_Cell.add_table
            width = obj.width;                         % Python: self.width
            if isequal(width, [])                      % Python: ... if self.width is not None else Inches(1)
                width = mat2doc.shared.Inches(1);
            end
            % Python: table = super(_Cell, self).add_table(rows, cols, width)
            t = add_table@mat2doc.BlockItemContainer(obj, rows, cols, width);
            obj.add_paragraph();                       % Python: self.add_paragraph()
        end

        function value = get.grid_span(obj)
            % GRID_SPAN get (table.py 228-235): number of layout-grid cells this
            %   cell spans horizontally (1 normal, 2+ horizontally merged).
            %   Python: return self._tc.grid_span.
            value = obj.tc_.grid_span;   % Python: return self._tc.grid_span
        end

        function merged = merge(obj, other_cell)
            % MERGE A merged cell spanning the rectangular region with this cell and
            %   `other_cell` as diagonal corners (table.py 237-245). Raises
            %   mat2doc:InvalidSpanError (via CT_Tc.merge) if the region is not
            %   rectangular. Returns a NEW Cell_ over the merged <w:tc>.
            %
            %   Python (table.py 243-245):
            %     tc, tc_2 = self._tc, other_cell._tc
            %     merged_tc = tc.merge(tc_2)
            %     return _Cell(merged_tc, self._parent)
            %   H5: tc.merge is the byte-proven CT_Tc merge engine (P6-3a); the
            %   returned Cell_ wraps the top-left merged tc, parented to self._parent.
            %
            %   Ported from python-docx v1.2.0: src/docx/table.py::_Cell.merge
            arguments
                obj (1,1) mat2doc.table.Cell_
                other_cell (1,1) mat2doc.table.Cell_
            end
            tc = obj.tc_;                    % Python: tc = self._tc
            tc_2 = other_cell.tc_;           % Python: tc_2 = other_cell._tc
            merged_tc = tc.merge(tc_2);      % Python: merged_tc = tc.merge(tc_2)
            % Python: return _Cell(merged_tc, self._parent)
            merged = mat2doc.table.Cell_(merged_tc, obj.parent_);
        end

        function ps = paragraphs(obj)
            % PARAGRAPHS The paragraphs in this cell, in document order (table.py
            %   247-254, @property, read-only). Delegates VERBATIM to
            %   BlockItemContainer.paragraphs (python-docx: `return super(_Cell,
            %   self).paragraphs`). A homogeneous 1xN Paragraph array; each read
            %   mints fresh Paragraph views (H5). Property-as-method (the
            %   BlockItemContainer precedent).
            %
            %   Ported from python-docx v1.2.0: src/docx/table.py::_Cell.paragraphs
            ps = paragraphs@mat2doc.BlockItemContainer(obj);   % Python: return super().paragraphs
        end

        function ts = tables(obj)
            % TABLES The tables in this cell, in document order (table.py 256-262,
            %   @property, read-only). Delegates VERBATIM to
            %   BlockItemContainer.tables (python-docx: `return super(_Cell,
            %   self).tables`), which is un-stubbed at P6-4b now that Table is live.
            %   Property-as-method.
            %
            %   Ported from python-docx v1.2.0: src/docx/table.py::_Cell.tables
            ts = tables@mat2doc.BlockItemContainer(obj);   % Python: return super().tables
        end

        function value = get.text(obj)
            % TEXT get (table.py 264-271): the entire contents of this cell as text.
            %   Python: return "\n".join(p.text for p in self.paragraphs).
            %   H16: python-docx does NOT strip here -- a plain newline join over the
            %   paragraph texts. strjoin over a 1x0 array returns "" (Python
            %   "".join([])), over 1x1 returns that text (Python "".join([x])).
            ps = obj.paragraphs();                       % Python: self.paragraphs
            texts = strings(1, numel(ps));
            for k = 1:numel(ps)                          % Python: p.text for p in ...
                texts(k) = ps(k).text;
            end
            value = strjoin(texts, string(newline));     % Python: "\n".join(...)
        end
        function set.text(obj, text)
            % TEXT set (table.py 273-284): replace ALL existing content with a single
            %   paragraph holding `text` in a single run. Python:
            %     tc = self._tc; tc.clear_content(); p = tc.add_p(); r = p.add_r();
            %     r.text = text
            tc = obj.tc_;              % Python: tc = self._tc
            tc.clear_content();        % Python: tc.clear_content()
            p = tc.add_p();            % Python: p = tc.add_p()
            r = p.add_r();             % Python: r = p.add_r()
            r.text = text;             % Python: r.text = text
        end

        function value = get.vertical_alignment(obj)
            % VERTICAL_ALIGNMENT get (table.py 286-297): a
            %   WD_CELL_VERTICAL_ALIGNMENT member, or [] (None) when inherited.
            %   Python: tcPr = self._element.tcPr; if tcPr is None: return None;
            %           return tcPr.vAlign_val
            %   Uses the RAW tcPr (NOT get_or_add) -- a read must not create it (H3).
            tcPr = obj.element_().tcPr;      % Python: self._element.tcPr (C3 seam == tc)
            if isequal(tcPr, [])             % Python: if tcPr is None (H3)
                value = [];
                return
            end
            value = tcPr.vAlign_val;         % Python: return tcPr.vAlign_val
        end
        function set.vertical_alignment(obj, value)
            % VERTICAL_ALIGNMENT set (table.py 299-302): write ./w:tcPr/w:vAlign;
            %   assigning [] (None) removes any explicit vertical alignment. Python:
            %     tcPr = self._element.get_or_add_tcPr(); tcPr.vAlign_val = value
            tcPr = obj.element_().get_or_add_tcPr();   % Python: self._element.get_or_add_tcPr()
            tcPr.vAlign_val = value;                   % Python: tcPr.vAlign_val = value
        end

        function value = get.width(obj)
            % WIDTH get (table.py 304-307): the cell width in EMU, or [] (None) if no
            %   explicit width is set. Python: return self._tc.width.
            value = obj.tc_.width;   % Python: return self._tc.width
        end
        function set.width(obj, value)
            % WIDTH set (table.py 309-311): self._tc.width = value (an EMU Length).
            %   CT_Tc handles the H6/H3 write (dxa twips via CT_TblWidth).
            obj.tc_.width = value;   % Python: self._tc.width = value
        end
    end
end

classdef BlockItemContainer < mat2doc.shared.StoryChild
% BLOCKITEMCONTAINER Base class for proxy objects that can contain block items.
%
%   Block-level items are things like paragraphs and tables. These containers
%   include _Body, _Cell, header, footer, footnote, endnote, comment, and text
%   box objects. Provides the shared functionality to add a block item like a
%   paragraph or table.
%
%   TIER (blkcntnr.py:33 `class BlockItemContainer(StoryChild)`):
%   BlockItemContainer < mat2doc.shared.StoryChild (the parent-only tier whose
%   `part` delegates up to a StoryPart, ported at P2-1). Adds the wrapped block
%   element (`_element`) and the block-item add/read surface.
%
%   MODULE MAPPING (design.md section 1): docx.blkcntnr is a TOP-LEVEL module, so
%   this class lands at the +mat2doc package root (mat2doc.BlockItemContainer),
%   mirroring `from docx.blkcntnr import BlockItemContainer`. It does NOT collide
%   with the +mat2doc\Document.m factory function (distinct names).
%
%   LIVE vs STUB (P2-3 + P4-7b + P6-4a + P6-4b):
%     LIVE  -- the container wiring: __init__(element, parent) storing _element
%              (rotated element_) and delegating parent to StoryChild; the
%              paragraph surface add_paragraph / _add_paragraph (add_paragraph_) /
%              paragraphs (P4-7b); add_table / iter_inner_content (P6-4a); AND
%              tables (P6-4b) -- un-stubbed now that Table (P6-4a) is live. This
%              tables() is the shared implementation _Cell.tables / Body_.tables /
%              (via delegation) Document.tables route through.
%     STUB  -- (none) the container block-item surface is fully ported.
%
%   UNDERSCORE ROTATION (design.md section 2): _element -> element_,
%   _add_paragraph -> add_paragraph_.
%
%   ============================ C3 ELEMENT-ACCESSOR SEAM (P5-3b) ================
%   python-docx's `_BaseHeaderFooter` (section.py 289-354) NEVER stores
%   `self._element`; it overrides `_element` as a LAZY @property returning
%   `self._get_or_add_definition().element` -- the header/footer part is created
%   on first content access. MATLAB CANNOT redefine a stored property in a
%   subclass, so a naive port (element_ as a plain stored property) would give the
%   header the wrong element. Resolution: `element_` is a zero-arg PROTECTED
%   METHOD (the house property-as-method convention) backed by the concrete store
%   `element_store_`. The base seam returns the stored handle; every internal
%   BlockItemContainer element use routes through `obj.element_()`. Because a
%   zero-arg method is read with the SAME `obj.element_` dot-syntax a property
%   would use, this refactor is BYTE-NEUTRAL for every existing subclass (Body_
%   passes a concrete element to the constructor; Document is NOT a
%   BlockItemContainer -- it extends ElementProxy). mat2doc.section.BaseHeaderFooter_
%   OVERRIDES element_() to return get_or_add_definition_().element() (lazy
%   part-creation), reproducing the python-docx semantics exactly. Proven
%   byte-neutral by the M1 17/17 + M2 s0033 sweeps (audit_P5-3b).
%
%   Example:
%       % A container wraps a block element (e.g. a CT_Body) and a story parent.
%       body_elm = d.element.body;                 % a CT_Body
%       c = mat2doc.BlockItemContainer(body_elm, d);   % d is a ProvidesStoryPart
%       p = c.part();                              % delegated up to the DocumentPart
%
%   Ported from python-docx v1.2.0: src/docx/blkcntnr.py::BlockItemContainer

    properties (Access = protected)
        element_store_  % _element backing store (READ via the element_() seam, C3)
    end

    methods
        function obj = BlockItemContainer(element, parent)
            % BLOCKITEMCONTAINER Store the wrapped block element and story parent
            %   (blkcntnr.py 41-43): super().__init__(parent); self._element = element.
            %
            %   Inputs:  element - the block element proxied (a mat2doc.oxml
            %                      BaseOxmlElement, e.g. a CT_Body).
            %            parent  - a ProvidesStoryPart (its `part` is a StoryPart).
            %   Outputs: obj     - a scalar BlockItemContainer handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/blkcntnr.py::
            %   BlockItemContainer.__init__
            obj@mat2doc.shared.StoryChild(parent);
            obj.element_store_ = element;   % C3: concrete store, read via element_()
        end

        function paragraph = add_paragraph(obj, text, style)
            % ADD_PARAGRAPH Return a paragraph newly added to the end of this
            %   container's content (blkcntnr.py 45-59). The paragraph has `text`
            %   in a single run if present, and paragraph style `style`. If `style`
            %   is [] (None) no style is applied (same effect as the 'Normal' style).
            %   UN-STUBBED at P4-7b (Paragraph / add_run / style all live).
            %
            %   Python (blkcntnr.py 54-59):
            %     paragraph = self._add_paragraph()
            %     if text:                 # non-empty string (H4)
            %         paragraph.add_run(text)
            %     if style is not None:    # None-identity (H3), NOT truthiness
            %         paragraph.style = style
            %     return paragraph
            %
            %   H13 default fidelity: add_paragraph(text="", style=None) -> text="",
            %   style=[]. The `if text:` guard mirrors Paragraph.add_run: the
            %   short-circuit `~isequal(text,[]) && strlength(text)>0` treats both
            %   "" and [] as falsy (no run) without calling strlength([]).
            %
            %   Ported from python-docx v1.2.0: src/docx/blkcntnr.py::BlockItemContainer.add_paragraph
            arguments
                obj
                text  = ""   % Python default ""
                style = []   % Python default None
            end
            paragraph = obj.add_paragraph_();               % Python: self._add_paragraph()
            if ~isequal(text, []) && strlength(text) > 0    % Python: if text:
                paragraph.add_run(text);                    % Python: paragraph.add_run(text)
            end
            if ~isequal(style, [])                          % Python: if style is not None:
                paragraph.style = style;                    % Python: paragraph.style = style
            end
        end

        function t = add_table(obj, rows, cols, width)
            % ADD_TABLE Return a table of `width` having `rows` rows and `cols`
            %   columns, appended at the end of this container's content
            %   (blkcntnr.py 61-72). `width` is evenly distributed between the
            %   columns. UN-STUBBED at P6-4a -- the FIRST end-to-end table
            %   authoring path (CT_Tbl.new_tbl byte-proven P6-3b; CT_Body._insert_tbl
            %   live P2-3; the Table proxy is now live P6-4a).
            %
            %   Python (blkcntnr.py 68-72):
            %     from docx.table import Table
            %     tbl = CT_Tbl.new_tbl(rows, cols, width)
            %     self._element._insert_tbl(tbl)
            %     return Table(tbl, self)
            %
            %   Ported from python-docx v1.2.0: src/docx/blkcntnr.py::BlockItemContainer.add_table
            tbl = mat2doc.oxml.table.CT_Tbl.new_tbl(rows, cols, width);  % Python: CT_Tbl.new_tbl(...)
            obj.element_().insert_tbl_(tbl);                             % Python: self._element._insert_tbl(tbl) (C3 seam)
            t = mat2doc.table.Table(tbl, obj);                          % Python: return Table(tbl, self)
        end

        function items = iter_inner_content(obj)
            % ITER_INNER_CONTENT Each Paragraph or Table in this container in
            %   document order (blkcntnr.py 74-79). UN-STUBBED at P6-4a (Table is
            %   now live). Returns a 1xN CELL array (heterogeneous Paragraph | Table
            %   share no matlab.mixin.Heterogeneous base -- the cell preserves
            %   document order, mirroring Paragraph/Section.iter_inner_content).
            %
            %   Python (blkcntnr.py 78-79):
            %     for element in self._element.inner_content_elements:
            %         yield (Paragraph(element, self) if isinstance(element, CT_P)
            %                else Table(element, self))
            %   H9: the Python generator is materialized (no tree mutation during
            %   iteration). H10: isinstance -> isa on the ported CT_P class name.
            %   CT_Body.inner_content_elements (xpath ./w:p|./w:tbl, LIVE) is already
            %   in document order (H1).
            %
            %   Ported from python-docx v1.2.0: src/docx/blkcntnr.py::BlockItemContainer.iter_inner_content
            elements = obj.element_().inner_content_elements;   % C3 seam; LIVE xpath list
            items = cell(1, numel(elements));
            for k = 1:numel(elements)
                element = elements(k);
                if isa(element, "mat2doc.oxml.text.CT_P")   % Python: isinstance(element, CT_P)
                    items{k} = mat2doc.text.Paragraph(element, obj);   % Python: Paragraph(element, self)
                else
                    items{k} = mat2doc.table.Table(element, obj);      % Python: Table(element, self)
                end
            end
        end

        function ps = paragraphs(obj)
            % PARAGRAPHS A list of the paragraphs in this container, in document
            %   order (blkcntnr.py 81-87, @property, read-only). UN-STUBBED at
            %   P4-7b. Python: [Paragraph(p, self) for p in self._element.p_lst].
            %   A homogeneous 1xN Paragraph array (LIST-PROPERTY SURFACE, per the
            %   Paragraph.runs precedent); each mints a FRESH Paragraph view (H5).
            %   CT_Body.p_lst is LIVE (P2-3); empty -> a 1x0 Paragraph array.
            %
            %   Ported from python-docx v1.2.0: src/docx/blkcntnr.py::BlockItemContainer.paragraphs
            plst = obj.element_().p_lst;   % C3 seam (lazy for header/footer)
            ps = mat2doc.text.Paragraph.empty(1, 0);
            for k = 1:numel(plst)   % Python: for p in self._element.p_lst
                ps(k) = mat2doc.text.Paragraph(plst(k), obj);   % Paragraph(p, self)
            end
        end

        function ts = tables(obj)
            % TABLES A list of the tables in this container, in document order
            %   (blkcntnr.py 89-97, @property, read-only). UN-STUBBED at P6-4b now
            %   that Table (P6-4a) is live -- this is the shared implementation
            %   _Cell.tables / Body_.tables / (via delegation) Document.tables all
            %   route through. Python: [Table(tbl, self) for tbl in
            %   self._element.tbl_lst]. A homogeneous 1xN Table array; each mints a
            %   FRESH Table view (H5). CT_Body.tbl_lst is LIVE (P2-3); empty -> a
            %   1x0 Table array.
            %
            %   Ported from python-docx v1.2.0: src/docx/blkcntnr.py::BlockItemContainer.tables
            tbllst = obj.element_().tbl_lst;   % C3 seam; LIVE tbl_lst
            ts = mat2doc.table.Table.empty(1, 0);
            for k = 1:numel(tbllst)   % Python: for tbl in self._element.tbl_lst
                ts(k) = mat2doc.table.Table(tbllst(k), obj);   % Table(tbl, self)
            end
        end
    end

    methods (Access = protected)
        function p = add_paragraph_(obj)
            % _ADD_PARAGRAPH Return a paragraph newly added to the end of this
            %   container's content (blkcntnr.py 99-101). UN-STUBBED at P4-7b.
            %   Python: return Paragraph(self._element.add_p(), self). CT_Body.add_p()
            %   is LIVE (P2-3); Paragraph is LIVE (P4-5b). Underscore rotation:
            %   _add_paragraph -> add_paragraph_.
            %
            %   Ported from python-docx v1.2.0: src/docx/blkcntnr.py::BlockItemContainer._add_paragraph
            p = mat2doc.text.Paragraph(obj.element_().add_p(), obj);   % C3 seam
        end

        function e = element_(obj)
            % ELEMENT_ C3 SEAM (P5-3b): the block element this container operates
            %   on. Base behavior returns the concrete store set at construction
            %   (element_store_). mat2doc.section.BaseHeaderFooter_ OVERRIDES this
            %   to return get_or_add_definition_().element() -- python-docx's lazy
            %   `_element` @property (section.py 351-354). Every internal
            %   BlockItemContainer element use (add_paragraph_/paragraphs) reads
            %   through this seam; `obj.element_` dot-syntax still works because a
            %   zero-arg method reads like a property (byte-neutral for Body_).
            e = obj.element_store_;
        end
    end
end

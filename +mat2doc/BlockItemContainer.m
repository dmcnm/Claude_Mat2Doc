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
%   LIVE vs STUB (P2-3):
%     LIVE  -- the container wiring: __init__(element, parent) storing _element
%              (rotated element_) and delegating parent to StoryChild.
%     STUB  -- add_paragraph / _add_paragraph / add_table / paragraphs / tables /
%              iter_inner_content. Each stubs EXACTLY at the item-CONSTRUCTION
%              boundary: the container structure is real, but the RETURNED item is
%              a Paragraph (needs CT_P + Paragraph, P4) or a Table (needs CT_Tbl +
%              Table, P6). blkcntnr.py:99-101 shows _add_paragraph reduces to
%              `Paragraph(self._element.add_p(), self)` -- the CT_Body.add_p() is
%              LIVE (P2-3), but Paragraph is not, so the whole method must raise
%              (calling add_p() first would leave a stray <w:p> in the tree).
%
%   UNDERSCORE ROTATION (design.md section 2): _element -> element_,
%   _add_paragraph -> add_paragraph_.
%
%   Example:
%       % A container wraps a block element (e.g. a CT_Body) and a story parent.
%       body_elm = d.element.body;                 % a CT_Body
%       c = mat2doc.BlockItemContainer(body_elm, d);   % d is a ProvidesStoryPart
%       p = c.part();                              % delegated up to the DocumentPart
%
%   Ported from python-docx v1.2.0: src/docx/blkcntnr.py::BlockItemContainer

    properties (Access = protected)
        element_        % _element: the wrapped block element (a CT_Body / CT_Tc / ...)
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
            obj.element_ = element;
        end

        function p = add_paragraph(obj, text, style) %#ok<INUSD,MANU,STOUT>
            % ADD_PARAGRAPH STUB (blkcntnr.py 45-59). Owner: P4 paragraph tier.
            %   Faithful body: paragraph = self._add_paragraph(); if text:
            %   paragraph.add_run(text); if style is not None: paragraph.style =
            %   style; return paragraph. Every step needs the Paragraph proxy /
            %   CT_P (P4); stubbed at the class boundary.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.text.paragraph.Paragraph (owning WP: P4 paragraph tier) " + ...
                "required by mat2doc.BlockItemContainer.add_paragraph");
        end

        function t = add_table(obj, rows, cols, width) %#ok<INUSD,MANU,STOUT>
            % ADD_TABLE STUB (blkcntnr.py 61-72). Owner: P6 table tier.
            %   Faithful body: tbl = CT_Tbl.new_tbl(rows, cols, width);
            %   self._element._insert_tbl(tbl); return Table(tbl, self). The
            %   CT_Body._insert_tbl (insert_tbl_) is LIVE (P2-3), but CT_Tbl.new_tbl
            %   and the Table proxy are P6; stubbed at the class boundary.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.oxml.table.CT_Tbl.new_tbl / mat2doc.table.Table (owning " + ...
                "WP: P6 table tier) required by mat2doc.BlockItemContainer.add_table");
        end

        function it = iter_inner_content(obj) %#ok<MANU,STOUT>
            % ITER_INNER_CONTENT STUB (blkcntnr.py 74-79). Owner: P4/P6.
            %   Faithful body iterates self._element.inner_content_elements (LIVE)
            %   yielding Paragraph(element, self) or Table(element, self). The
            %   Paragraph (P4) / Table (P6) construction is the stub boundary.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.text.paragraph.Paragraph (P4) / mat2doc.table.Table (P6) " + ...
                "required by mat2doc.BlockItemContainer.iter_inner_content");
        end

        function ps = paragraphs(obj) %#ok<MANU,STOUT>
            % PARAGRAPHS STUB (blkcntnr.py 81-87). Owner: P4 paragraph tier.
            %   Faithful body: [Paragraph(p, self) for p in self._element.p_lst].
            %   CT_Body.p_lst is LIVE; the Paragraph construction is the boundary.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.text.paragraph.Paragraph (owning WP: P4 paragraph tier) " + ...
                "required by mat2doc.BlockItemContainer.paragraphs");
        end

        function ts = tables(obj) %#ok<MANU,STOUT>
            % TABLES STUB (blkcntnr.py 89-97). Owner: P6 table tier.
            %   Faithful body: [Table(tbl, self) for tbl in self._element.tbl_lst].
            %   CT_Body.tbl_lst is LIVE; the Table construction is the boundary.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.table.Table (owning WP: P6 table tier) required by " + ...
                "mat2doc.BlockItemContainer.tables");
        end
    end

    methods (Access = protected)
        function p = add_paragraph_(obj) %#ok<MANU,STOUT>
            % _ADD_PARAGRAPH STUB (blkcntnr.py 99-101). Owner: P4 paragraph tier.
            %   Faithful body: return Paragraph(self._element.add_p(), self). The
            %   CT_Body.add_p() is LIVE (P2-3), but Paragraph is P4; the whole
            %   method raises (calling add_p() first would leave a stray <w:p>).
            %   Underscore rotation: _add_paragraph -> add_paragraph_.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.text.paragraph.Paragraph (owning WP: P4 paragraph tier) " + ...
                "required by mat2doc.BlockItemContainer._add_paragraph");
        end
    end
end

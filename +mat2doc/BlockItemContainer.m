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
%   LIVE vs STUB (P2-3 + P4-7b):
%     LIVE  -- the container wiring: __init__(element, parent) storing _element
%              (rotated element_) and delegating parent to StoryChild; AND the
%              paragraph surface add_paragraph / _add_paragraph (add_paragraph_) /
%              paragraphs, un-stubbed at P4-7b now that Paragraph (P4-5b),
%              add_run and style are all live over the LIVE CT_Body.add_p() /
%              p_lst (P2-3).
%     STUB  -- add_table / tables (need CT_Tbl + Table, P6) and iter_inner_content
%              (heterogeneous Paragraph|Table -- still stubs at the Table P6
%              boundary though Paragraph is now live).
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
            % ITER_INNER_CONTENT STUB (blkcntnr.py 74-79). Owner: P6 table tier.
            %   Faithful body iterates self._element.inner_content_elements (LIVE)
            %   yielding Paragraph(element, self) or Table(element, self). Paragraph
            %   is now LIVE (P4-7b) but Table is P6, so this heterogeneous iterator
            %   still stubs at the Table boundary.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.table.Table (owning WP: P6 table tier) " + ...
                "required by mat2doc.BlockItemContainer.iter_inner_content");
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
            plst = obj.element_.p_lst;
            ps = mat2doc.text.Paragraph.empty(1, 0);
            for k = 1:numel(plst)   % Python: for p in self._element.p_lst
                ps(k) = mat2doc.text.Paragraph(plst(k), obj);   % Paragraph(p, self)
            end
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
        function p = add_paragraph_(obj)
            % _ADD_PARAGRAPH Return a paragraph newly added to the end of this
            %   container's content (blkcntnr.py 99-101). UN-STUBBED at P4-7b.
            %   Python: return Paragraph(self._element.add_p(), self). CT_Body.add_p()
            %   is LIVE (P2-3); Paragraph is LIVE (P4-5b). Underscore rotation:
            %   _add_paragraph -> add_paragraph_.
            %
            %   Ported from python-docx v1.2.0: src/docx/blkcntnr.py::BlockItemContainer._add_paragraph
            p = mat2doc.text.Paragraph(obj.element_.add_p(), obj);
        end
    end
end

classdef InlineShapes < mat2doc.shared.Parented
% INLINESHAPES Sequence of InlineShape instances (len, iteration, indexed access).
%
%   python-docx `InlineShapes(Parented)` (shape.py 21) -- a PLAIN parented proxy
%   holding the document body element (`_body`) and its parent (a StoryPart). It
%   has no single wrapped element that is its identity, so it derives from
%   mat2doc.shared.Parented (default handle identity, no eq/ne).
%
%   ATTRIBUTES (shape.py 24-26): Python `super().__init__(parent);
%   self._body = body_elm`. Ported: obj@Parented(parent) sets parent_; body_
%   holds the <w:body> (a CT_Body).
%
%   VERIFY-COLLECTION (design.md section 2): the Python Sequence surface is
%   ported as EXPLICIT methods (the shared 1-based () RedefinesParen base is a
%   future WP), matching the Rows_/Columns_/Sections precedent:
%       getitem_ (__getitem__)   to_array (__iter__)   len_ (__len__)
%   Dunder mapping: `inline_shapes[idx]` -> inline_shapes.getitem_(idx);
%   `for s in inline_shapes` -> `for s = inline_shapes.to_array()`;
%   `len(inline_shapes)` -> inline_shapes.len_(). FLAGGED for auditor.
%
%   __getitem__ (shape.py 28-36): INT-only (no slice overload). Python
%   `self._inline_lst[idx]`; on IndexError re-raises with the custom message
%   "inline shape index [%d] out of range" % idx (the ORIGINAL idx, incl. a
%   negative one). Python list indexing is 0-based with negative-index wrap.
%   H1: `+1` converts the resolved 0-based position to the 1-based MATLAB one.
%
%   _inline_lst (shape.py 44-48): body.xpath("//w:p/w:r/w:drawing/wp:inline") --
%   every inline shape in the story, document order. H9: the xpath engine yields
%   a materialized 1xN element array (empty typed array when none). Ported as the
%   private property-as-method inline_lst_ (leading-underscore rotation).
%
%   H5 (identity): every getitem_/to_array element mints a FRESH InlineShape view
%   of its <wp:inline> (python-docx does not cache); the wrapped CT_Inline is the
%   shared identity.
%
%   UNDERSCORE ROTATION (design.md section 2): _body -> body_, _inline_lst ->
%   inline_lst_.
%
%   Ported from python-docx v1.2.0: src/docx/shape.py::InlineShapes (lines 21-48)

    properties (Access = private)
        body_   % _body (shape.py 26): the document body element (a CT_Body)
    end

    methods
        function obj = InlineShapes(body_elm, parent)
            % INLINESHAPES Wrap the document body and its parent (shape.py 24-26).
            %   Python: super().__init__(parent); self._body = body_elm.
            %
            %   Inputs:  body_elm - the <w:body> element (a CT_Body).
            %            parent   - the StoryPart this collection belongs to.
            %   Outputs: obj      - a scalar InlineShapes handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/shape.py::InlineShapes.__init__
            obj@mat2doc.shared.Parented(parent);   % Python: super().__init__(parent)
            obj.body_ = body_elm;                  % Python: self._body = body_elm
        end

        function result = getitem_(obj, idx)
            % GETITEM_ Indexed access, e.g. inline_shapes[idx] (shape.py 28-36).
            %   Python:
            %     try: inline = self._inline_lst[idx]
            %     except IndexError:
            %         msg = "inline shape index [%d] out of range" % idx
            %         raise IndexError(msg)
            %     return InlineShape(inline)
            %   INT key, 0-based with negative-index wrap; out-of-range raises the
            %   custom message carrying the ORIGINAL idx. H1: `+1`.
            %
            %   Ported from python-docx v1.2.0: src/docx/shape.py::InlineShapes.__getitem__
            lst = obj.inline_lst_();   % the list(self) source
            n = numel(lst);
            i = idx;                   % Python 0-based index
            if i < 0                   % Python negative-index wrap
                i = i + n;
            end
            if i < 0 || i >= n         % Python list IndexError
                % Python: "inline shape index [%d] out of range" % idx (ORIGINAL idx)
                error("mat2doc:IndexError", "%s", ...
                    "inline shape index [" + mat2doc.shared.pyStr(idx, "int") + "] out of range");
            end
            result = mat2doc.shape.InlineShape(lst(i + 1));   % IDX; InlineShape(inline)
        end

        function result = to_array(obj)
            % TO_ARRAY An InlineShape per <wp:inline>, in document order (shape.py 38-39).
            %   Python __iter__:
            %     return (InlineShape(inline) for inline in self._inline_lst)
            %   Materialized (H9) into a 1xN InlineShape array; none -> a 1x0 array.
            %   Iteration idiom: `for s in inline_shapes` ->
            %   `for s = inline_shapes.to_array()`.
            %
            %   Ported from python-docx v1.2.0: src/docx/shape.py::InlineShapes.__iter__
            lst = obj.inline_lst_();
            result = mat2doc.shape.InlineShape.empty(1, 0);
            for k = 1:numel(lst)   % Python: for inline in self._inline_lst
                result(k) = mat2doc.shape.InlineShape(lst(k));
            end
        end

        function n = len_(obj)
            % LEN_ Number of inline shapes (shape.py 41-42).
            %   Python __len__: return len(self._inline_lst).
            %
            %   Ported from python-docx v1.2.0: src/docx/shape.py::InlineShapes.__len__
            n = numel(obj.inline_lst_());
        end
    end

    methods (Access = private)
        function lst = inline_lst_(obj)
            % INLINE_LST_ Every inline shape element in the story, document order
            %   (shape.py 44-48, @property). Python:
            %     body = self._body
            %     xpath = "//w:p/w:r/w:drawing/wp:inline"
            %     return body.xpath(xpath)
            %   H9: materialized 1xN element array (empty typed array when none).
            %
            %   Ported from python-docx v1.2.0: src/docx/shape.py::InlineShapes._inline_lst
            body = obj.body_;
            lst = body.xpath("//w:p/w:r/w:drawing/wp:inline");
        end
    end
end

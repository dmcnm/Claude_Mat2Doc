classdef ParagraphStyle < mat2doc.styles.CharacterStyle
% PARAGRAPHSTYLE A paragraph style.
%
%   A paragraph style provides both character formatting (inherited from
%   CharacterStyle) and paragraph formatting such as indentation and
%   line-spacing (via the |ParagraphFormat| object in its `paragraph_format`
%   property). Adds `next_paragraph_style` (read/write) and `paragraph_format`
%   (read-only).
%
%   next_paragraph_style (style.py 206-226): the style applied automatically to a
%   new paragraph inserted after a paragraph of this style.
%     GET (style.py 214-219): returns SELF when no next style is defined
%       (`w:next` absent) OR when the referenced style is not a PARAGRAPH type;
%       otherwise StyleFactory over the referenced sibling. Note the H3/H4 chain:
%       `next_style_elm.type != WD_STYLE_TYPE.PARAGRAPH` is True when type is None
%       (a dangling/typeless ref) -> also returns self.
%     SET (style.py 222-226): assigning None or self (same style_id) REMOVES the
%       `w:next`; any other style writes its style_id into `w:next/@w:val`.
%
%   paragraph_format (read-only, style.py 228-232): a FRESH ParagraphFormat
%   wrapping the wrapped `w:style` each access (Python
%   `return ParagraphFormat(self._element)`; not cached). ParagraphFormat reaches
%   its `w:pPr` through the CT_Style pPr descriptors (P4-6).
%
%   NOTE (__repr__, style.py 203-204): Python overrides __repr__ to
%   "_ParagraphStyle('<name>') id: <id>". MATLAB object display is a separate
%   mechanism (disp/display) and is NOT ported (no output-visible effect); see
%   the audit. `_ParagraphStyle = ParagraphStyle` (style.py 236) is a non-public
%   back-compat alias with no MATLAB equivalent.
%
%   Example:
%       se = mat2doc.oxml.OxmlElement("w:style");
%       se.type = mat2doc.enum.style.WD_STYLE_TYPE.PARAGRAPH;
%       ps = mat2doc.styles.ParagraphStyle(se);
%       ps.paragraph_format.left_indent = mat2doc.shared.Pt(36);
%
%   Ported from python-docx v1.2.0: src/docx/styles/style.py::ParagraphStyle

    properties (Dependent)
        next_paragraph_style   % ParagraphStyle -- style for the following paragraph
        paragraph_format       % ParagraphFormat (read-only) -- paragraph formatting
    end

    methods
        function obj = ParagraphStyle(style_elm)
            % PARAGRAPHSTYLE Wrap a `w:style` (inherits BaseStyle.__init__).
            %   Ported from python-docx v1.2.0: styles/style.py::ParagraphStyle
            obj@mat2doc.styles.CharacterStyle(style_elm);
        end

        % ---- next_paragraph_style (style.py 206-226) ----
        function value = get.next_paragraph_style(obj)
            % Python (style.py 214-219):
            %   next_style_elm = self._element.next_style
            %   if next_style_elm is None: return self
            %   if next_style_elm.type != WD_STYLE_TYPE.PARAGRAPH: return self
            %   return StyleFactory(next_style_elm)
            next_style_elm = obj.element_.next_style;   % CT_Style or [] (not found -> [])
            if isequal(next_style_elm, [])   % Python: if next_style_elm is None
                value = obj;   % Python: return self
                return
            end
            % Python: type != PARAGRAPH -- True also when type is None (H3/H4).
            if ~isequal(next_style_elm.type, mat2doc.enum.style.WD_STYLE_TYPE.PARAGRAPH)
                value = obj;   % Python: return self
                return
            end
            value = mat2doc.styles.StyleFactory(next_style_elm);
        end
        function set.next_paragraph_style(obj, style)
            % Python (style.py 222-226):
            %   if style is None or style.style_id == self.style_id:
            %       self._element._remove_next()
            %   else:
            %       self._element.get_or_add_next().val = style.style_id
            %   `is None` short-circuits before style.style_id is read. The
            %   style_id equality is value equality over {string, None}; isequal
            %   faithfully reproduces Python `==` for that domain (None==None True).
            if isequal(style, []) || isequal(style.style_id, obj.style_id)
                obj.element_.remove_next_();   % Python: self._element._remove_next()
            else
                nx = obj.element_.get_or_add_next();
                nx.val = style.style_id;       % Python: ...get_or_add_next().val = style.style_id
            end
        end

        % ---- paragraph_format (read-only; style.py 228-232) ----
        function value = get.paragraph_format(obj)
            % Python: return ParagraphFormat(self._element) -- fresh each access.
            value = mat2doc.text.ParagraphFormat(obj.element_);
        end
    end
end

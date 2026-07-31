classdef CharacterStyle < mat2doc.styles.BaseStyle
% CHARACTERSTYLE A character style.
%
%   A character style is applied to a |Run| object and primarily provides
%   character-level formatting via the |Font| object in its `font` property.
%   Adds `base_style` (read/write) and `font` (read-only) to the inherited
%   BaseStyle surface.
%
%   H3 (base_style): the get returns a Style (via StyleFactory over the sibling
%   `w:style` referenced by `w:basedOn`) or [] when there is no basedOn or the
%   referenced style is not found (a dangling ref). The set writes the target's
%   style_id into `w:basedOn/@w:val`, or removes basedOn when the assigned style
%   is [] (None).
%
%   font (read-only, style.py 185-189): returns a FRESH Font wrapping the wrapped
%   `w:style` on each access (Python `return Font(self._element)`; not cached).
%   Font reaches its `w:rPr` through the CT_Style rPr descriptors (P4-6).
%
%   NOTE (alias): python-docx also binds `_CharacterStyle = CharacterStyle`
%   (style.py 193) "just in case someone uses the old name in an extension
%   function". It is not part of the public API and has no MATLAB equivalent
%   (see the audit).
%
%   Example:
%       se = mat2doc.oxml.OxmlElement("w:style");
%       se.type = mat2doc.enum.style.WD_STYLE_TYPE.CHARACTER;
%       cs = mat2doc.styles.CharacterStyle(se);
%       cs.font.bold = true;               % <w:rPr><w:b/></w:rPr> on the style
%
%   Ported from python-docx v1.2.0: src/docx/styles/style.py::CharacterStyle

    properties (Dependent)
        base_style   % BaseStyle|[] -- the style this inherits from (w:basedOn)
        font         % Font (read-only) -- character formatting of this style
    end

    methods
        function obj = CharacterStyle(style_elm)
            % CHARACTERSTYLE Wrap a `w:style` element (inherits BaseStyle.__init__).
            %   Ported from python-docx v1.2.0: styles/style.py::CharacterStyle
            obj@mat2doc.styles.BaseStyle(style_elm);
        end

        % ---- base_style (style.py 171-183) ----
        function value = get.base_style(obj)
            % Python (style.py 174-178): base_style = self._element.base_style;
            %   if base_style is None: return None; return StyleFactory(base_style).
            base_style = obj.element_.base_style;   % CT_Style or [] (dangling ref -> [])
            if isequal(base_style, [])   % Python: if base_style is None
                value = [];
                return
            end
            value = mat2doc.styles.StyleFactory(base_style);
        end
        function set.base_style(obj, style)
            % Python (style.py 181-183): style_id = style.style_id if style is not
            %   None else None; self._element.basedOn_val = style_id.
            if isequal(style, [])   % Python: style is None
                style_id = [];
            else
                style_id = style.style_id;
            end
            obj.element_.basedOn_val = style_id;
        end

        % ---- font (read-only; style.py 185-189) ----
        function value = get.font(obj)
            % Python: return Font(self._element) -- a fresh Font each access.
            value = mat2doc.text.Font(obj.element_);
        end
    end
end

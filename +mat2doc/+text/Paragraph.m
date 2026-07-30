classdef Paragraph < mat2doc.shared.StoryChild
% PARAGRAPH Proxy object wrapping a `<w:p>` element.
%
%   A paragraph is a block-level item made of runs and hyperlinks (its inner
%   content) plus paragraph-level formatting (its `w:pPr`). This is a pure
%   API/proxy tier over the already-registered CT_P (P4-2): it adds NO oxml
%   logic, NO registry rows and NO serialization code. Equivalence is therefore
%   BEHAVIORAL (proxy return values + serialized bytes), not byte-registry.
%
%   TIER (paragraph.py 23 `class Paragraph(StoryChild)`): StoryChild < handle
%   (the parent-only tier whose `part` delegates up to a StoryPart, P2-1). So
%   Paragraph derives from mat2doc.shared.StoryChild -- NOT ElementProxy -- and,
%   like StoryChild, does NOT define eq/ne (default handle identity == Python
%   default object identity; a Paragraph is NOT compared by wrapped-element
%   identity, unlike an ElementProxy subclass).
%
%   ATTRIBUTES (paragraph.py 26-28): Python `self._p = self._element = p` -- two
%   names for the SAME <w:p> element. Ported verbatim: p_ (private, Python _p,
%   the working handle used by every method) and element_ (private, Python
%   _element, what paragraph_format wraps). Both hold the same CT_P handle.
%
%   LIST-PROPERTY SURFACE (runs / hyperlinks / rendered_page_breaks): the Python
%   properties return PLAIN lists (list comprehensions), not collection classes.
%   They are homogeneous (all Run / all Hyperlink / all RenderedPageBreak), so
%   the faithful MATLAB surface is a 1xN object ARRAY seeded via Class.empty(1,0)
%   -- mirroring the plain Python list and the TabStops.to_array() precedent
%   (P4-5a). Native 1-based () indexing then realizes the design.md section-2
%   dunder mapping x[i] -> x(i+1); numel(x) == len(x). (Contrast TabStops, a
%   genuine collection CLASS with getitem_/len_.) FLAGGED for auditor/validator.
%
%   ITER_INNER_CONTENT SURFACE (heterogeneous): iter_inner_content yields a MIXED
%   Run|Hyperlink sequence in document order. Run < StoryChild and
%   Hyperlink < Parented share NO matlab.mixin.Heterogeneous base, so they cannot
%   inhabit one object array. The faithful surface is a 1xN CELL array preserving
%   document order (H1: the ./w:r | ./w:hyperlink xpath is already in document
%   order). Python is a generator; per H9 it is precomputed (the source performs
%   no mutation during iteration, so laziness is unobservable).
%
%   STYLE DELEGATION (paragraph.py 130-147): style get/set delegate through
%   self.part.get_style / get_style_id -- the P4-7 stubs that raise
%   mat2doc:notYetPorted (DocumentPart.get_style/get_style_id). The delegation is
%   ported EXACTLY; the getter and setter therefore both raise notYetPorted until
%   P4-7 wires styles. This is faithful stub PROPAGATION, not a stand-in. Ported
%   as a Dependent read/write property to preserve its true Python shape (both
%   `p.style` and `p.style = x` raise the same notYetPorted). Auto-display is safe
%   (MATLAB catches the erroring getter and omits the style row; cf. Run.style).
%
%   H3 (None): inline isequal(x, []) (established Mat2Doc None idiom -- no shared
%   isNone helper). H4 (truthiness): `if text:` is a non-empty-string test;
%   `if style:` / `if style is not None` are ported DISTINCTLY per paragraph.py.
%
%   Example:
%       p_elm = mat2doc.oxml.OxmlElement("w:p");
%       para  = mat2doc.text.Paragraph(p_elm, someStoryParent);
%       run   = para.add_run("Hello");            % append a run with text
%       para.alignment = mat2doc.enum.text.WD_ALIGN_PARAGRAPH.CENTER;
%       n     = numel(para.runs);                 % len(paragraph.runs)
%       txt   = para.text;                        % concatenated inner-content text
%
%   Ported from python-docx v1.2.0: src/docx/text/paragraph.py::Paragraph

    properties (Access = private)
        p_          % _p (paragraph.py 28): the working <w:p> (a mat2doc.oxml.text.CT_P)
        element_    % _element (paragraph.py 28): same handle; what paragraph_format wraps
    end

    properties (Dependent)
        alignment           % WD_PARAGRAPH_ALIGNMENT|[] -- ./w:pPr/w:jc/@val (CT_P.alignment)
        contains_page_break % bool -- true when >=1 rendered page-break occurs
        hyperlinks          % 1xN Hyperlink array -- a Hyperlink per <w:hyperlink> child
        paragraph_format    % ParagraphFormat (read-only) -- fresh each access
        rendered_page_breaks% 1xN RenderedPageBreak array -- all rendered page-breaks
        runs                % 1xN Run array -- a Run per <w:r> child
        style               % ParagraphStyle -- STUB delegation (P4-7)
        text                % string -- concatenated inner-content text (CT_P.text)
    end

    methods
        function obj = Paragraph(p, parent)
            % PARAGRAPH Wrap a `<w:p>` element (paragraph.py 26-28).
            %
            %   Inputs:  p      - a mat2doc.oxml.text.CT_P (the `w:p` element).
            %            parent - the parent proxy (a ProvidesStoryPart) providing
            %                     `part`.
            %   Outputs: obj    - a scalar Paragraph handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/text/paragraph.py::Paragraph.__init__
            obj@mat2doc.shared.StoryChild(parent);   % Python: super().__init__(parent)
            % Python: self._p = self._element = p (one element, two names)
            obj.p_ = p;
            obj.element_ = p;
        end

        % ============================ add_run ============================
        function run = add_run(obj, text, style)
            % ADD_RUN Append a run containing `text` with character-style `style`
            %   (paragraph.py 30-44). `text` may contain \t (-> <w:tab/>) and
            %   \n/\r (-> line breaks); when `text` is None (default []) the new
            %   run is empty. `style` (default None [] ) is a style name or a
            %   CharacterStyle. NOTE: assigning `style` reaches Run.style, the
            %   P4-7 STUB (raises mat2doc:notYetPorted) -- faithful propagation.
            %
            %   Ported from python-docx v1.2.0: src/docx/text/paragraph.py::Paragraph.add_run
            arguments
                obj
                text  = []   % Python default None
                style = []   % Python default None
            end
            r = obj.p_.add_r();                    % Python: r = self._p.add_r()
            run = mat2doc.text.Run(r, obj);        % Python: run = Run(r, self)
            % Python: if text: (None/'' falsy) -- a non-empty string
            if ~isequal(text, []) && strlength(text) > 0
                run.text = text;                   % Python: run.text = text
            end
            % Python: if style: -- a non-empty style name OR a (truthy) style object
            if mat2doc.text.Paragraph.truthy_(style)
                run.style = style;                 % Python: run.style = style
            end
        end

        % ============================ alignment ============================
        function value = get.alignment(obj)
            % ALIGNMENT get (paragraph.py 46-55): delegates to CT_P.alignment.
            value = obj.p_.alignment;              % Python: return self._p.alignment
        end
        function set.alignment(obj, value)
            % ALIGNMENT set (paragraph.py 57-59): delegates to CT_P.alignment.
            obj.p_.alignment = value;              % Python: self._p.alignment = value
        end

        % ============================ clear ============================
        function obj = clear(obj)
            % CLEAR Remove all this paragraph's content and return this SAME
            %   paragraph (paragraph.py 61-67). Paragraph-level formatting (the
            %   w:pPr, e.g. style) is preserved. H5: obj is the same handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/text/paragraph.py::Paragraph.clear
            obj.p_.clear_content();                % Python: self._p.clear_content()
            % Python: return self  (obj is the same handle)
        end

        % ============================ contains_page_break ============================
        function value = get.contains_page_break(obj)
            % CONTAINS_PAGE_BREAK true when one or more RENDERED page-breaks occur
            %   in this paragraph (paragraph.py 69-72). Python:
            %   `return bool(self._p.lastRenderedPageBreaks)` (H4: bool of the list
            %   -> ~isempty of the xpath result).
            value = ~isempty(obj.p_.lastRenderedPageBreaks());
        end

        % ============================ hyperlinks ============================
        function value = get.hyperlinks(obj)
            % HYPERLINKS A Hyperlink for each <w:hyperlink> child, in document
            %   order (paragraph.py 74-77). Python list comprehension ->
            %   homogeneous 1xN Hyperlink array (see LIST-PROPERTY SURFACE). Each
            %   element mints a FRESH Hyperlink view (Python does not cache; H5).
            hlst = obj.p_.hyperlink_lst;
            value = mat2doc.text.Hyperlink.empty(1, 0);
            for k = 1:numel(hlst)                  % Python: for hyperlink in self._p.hyperlink_lst
                value(k) = mat2doc.text.Hyperlink(hlst(k), obj);   % Hyperlink(hyperlink, self)
            end
        end

        % ============================ insert_paragraph_before ============================
        function paragraph = insert_paragraph_before(obj, text, style)
            % INSERT_PARAGRAPH_BEFORE Return a new paragraph inserted directly
            %   before this one (paragraph.py 79-92). If `text` supplied, the new
            %   paragraph holds it in a single run; if `style` is not None it is
            %   assigned (reaching Paragraph.style, the P4-7 STUB -> notYetPorted).
            %
            %   NOTE the DISTINCT guards (paragraph.py 88, 90): `if text:`
            %   (truthiness) vs `if style is not None:` (None-identity).
            %
            %   Ported from python-docx v1.2.0: src/docx/text/paragraph.py::Paragraph.insert_paragraph_before
            arguments
                obj
                text  = []   % Python default None
                style = []   % Python default None
            end
            paragraph = obj.insert_paragraph_before_();    % Python: self._insert_paragraph_before()
            % Python: if text: (None/'' falsy) -- a non-empty string
            if ~isequal(text, []) && strlength(text) > 0
                paragraph.add_run(text);                   % Python: paragraph.add_run(text)
            end
            % Python: if style is not None: (identity, NOT truthiness)
            if ~isequal(style, [])
                paragraph.style = style;                   % Python: paragraph.style = style
            end
        end

        % ============================ iter_inner_content ============================
        function items = iter_inner_content(obj)
            % ITER_INNER_CONTENT The runs and hyperlinks in this paragraph, in the
            %   order they appear (paragraph.py 94-107). Returns a 1xN CELL array
            %   (heterogeneous Run|Hyperlink -- see ITER_INNER_CONTENT SURFACE);
            %   each cell is a Run (for a <w:r>) or a Hyperlink (otherwise). H9:
            %   the Python generator is precomputed (no mutation during iteration).
            %   H1: inner_content_elements (./w:r | ./w:hyperlink) is already in
            %   document order.
            %
            %   Ported from python-docx v1.2.0: src/docx/text/paragraph.py::Paragraph.iter_inner_content
            elms = obj.p_.inner_content_elements();    % Python: self._p.inner_content_elements
            items = cell(1, numel(elms));
            for k = 1:numel(elms)                      % Python: for r_or_hlink in ...
                e = elms(k);
                % Python: Run(r_or_hlink, self) if isinstance(r_or_hlink, CT_R)
                %         else Hyperlink(r_or_hlink, self)
                if isa(e, "mat2doc.oxml.text.CT_R")
                    items{k} = mat2doc.text.Run(e, obj);
                else
                    items{k} = mat2doc.text.Hyperlink(e, obj);
                end
            end
        end

        % ============================ paragraph_format ============================
        function value = get.paragraph_format(obj)
            % PARAGRAPH_FORMAT The ParagraphFormat for this paragraph's formatting
            %   (paragraph.py 109-113). Python `return ParagraphFormat(self._element)`
            %   -- a FRESH ParagraphFormat each access (not cached), wrapping the
            %   <w:p> (element_).
            value = mat2doc.text.ParagraphFormat(obj.element_);
        end

        % ============================ rendered_page_breaks ============================
        function value = get.rendered_page_breaks(obj)
            % RENDERED_PAGE_BREAKS All rendered page-breaks in this paragraph
            %   (paragraph.py 115-122). Most often empty; a RenderedPageBreak per
            %   <w:lastRenderedPageBreak> -- homogeneous 1xN array (LIST-PROPERTY
            %   SURFACE). Each mints a FRESH view (H5).
            lrpbs = obj.p_.lastRenderedPageBreaks();
            value = mat2doc.text.RenderedPageBreak.empty(1, 0);
            for k = 1:numel(lrpbs)                 % Python: for lrpb in self._p.lastRenderedPageBreaks
                value(k) = mat2doc.text.RenderedPageBreak(lrpbs(k), obj);   % RenderedPageBreak(lrpb, self)
            end
        end

        % ============================ runs ============================
        function value = get.runs(obj)
            % RUNS A Run per <w:r> child, in document order (paragraph.py 124-128).
            %   Homogeneous 1xN Run array (LIST-PROPERTY SURFACE). Each mints a
            %   FRESH Run view (H5).
            rlst = obj.p_.r_lst;
            value = mat2doc.text.Run.empty(1, 0);
            for k = 1:numel(rlst)                  % Python: for r in self._p.r_lst
                value(k) = mat2doc.text.Run(rlst(k), obj);   % Run(r, self)
            end
        end

        % ============================ style (STUB delegation, P4-7) ============================
        function value = get.style(obj)
            % STYLE get (paragraph.py 130-142): style_id = self._p.style; then
            %   self.part.get_style(style_id, WD_STYLE_TYPE.PARAGRAPH). part is the
            %   DocumentPart whose get_style is the P4-7 STUB -> mat2doc:notYetPorted.
            %   Delegation ported EXACTLY (no stand-in).
            style_id = obj.p_.style;               % Python: style_id = self._p.style
            value = obj.part().get_style( ...      % Python: self.part.get_style(...)
                style_id, mat2doc.enum.style.WD_STYLE_TYPE.PARAGRAPH);
        end
        function set.style(obj, style_or_name)
            % STYLE set (paragraph.py 144-147): style_id =
            %   self.part.get_style_id(value, WD_STYLE_TYPE.PARAGRAPH);
            %   self._p.style = style_id. get_style_id is the P4-7 STUB ->
            %   mat2doc:notYetPorted (raised before self._p.style is touched).
            style_id = obj.part().get_style_id( ...% Python: self.part.get_style_id(...)
                style_or_name, mat2doc.enum.style.WD_STYLE_TYPE.PARAGRAPH);
            obj.p_.style = style_id;               % Python: self._p.style = style_id
        end

        % ============================ text ============================
        function value = get.text(obj)
            % TEXT get (paragraph.py 149-163): the paragraph's concatenated
            %   inner-content text (incl. hyperlink visible text); tabs -> \t,
            %   line breaks -> \n. Python `return self._p.text` -> CT_P.text (D10).
            value = obj.p_.text;                   % Python: return self._p.text
        end
        function set.text(obj, text)
            % TEXT set (paragraph.py 165-168): replace all content with a single
            %   run holding `text` (\t -> <w:tab/>, \n/\r -> line break);
            %   paragraph-level formatting preserved, run-level formatting removed.
            %   Python: self.clear(); self.add_run(text).
            obj.clear();                           % Python: self.clear()
            obj.add_run(text);                     % Python: self.add_run(text)
        end
    end

    methods (Access = protected)
        % ============================ _insert_paragraph_before ============================
        function paragraph = insert_paragraph_before_(obj)
            % _INSERT_PARAGRAPH_BEFORE Return a new paragraph inserted directly
            %   before this one (paragraph.py 170-173). Its parent is THIS
            %   paragraph's parent (self._parent), not this paragraph.
            %   Underscore rotation: _insert_paragraph_before -> insert_paragraph_before_.
            %
            %   Ported from python-docx v1.2.0: src/docx/text/paragraph.py::Paragraph._insert_paragraph_before
            p = obj.p_.add_p_before();             % Python: p = self._p.add_p_before()
            paragraph = mat2doc.text.Paragraph(p, obj.parent_);   % Paragraph(p, self._parent)
        end
    end

    methods (Static, Access = private)
        function tf = truthy_(style)
            % TRUTHY_ Python `if style:` for add_run (paragraph.py 42), where
            %   style is None | str | CharacterStyle. Python truthiness: None ->
            %   False; a str -> len(str) > 0; a style OBJECT -> True. ([] is the
            %   None sentinel, H3.)
            if isequal(style, [])                  % Python: None -> falsy
                tf = false;
            elseif isstring(style) || ischar(style)
                tf = strlength(string(style)) > 0; % Python: non-empty str is truthy
            else
                tf = true;                         % a (truthy) style object
            end
        end
    end
end

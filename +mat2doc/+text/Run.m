classdef Run < mat2doc.shared.StoryChild
% RUN Proxy object wrapping a `<w:r>` run element.
%
%   Several properties on Run take a tri-state value (H3): true, false, or []
%   (Python None). true/false correspond to on/off; [] means the property is
%   not specified directly on the run and its effective value is taken from the
%   style hierarchy.
%
%   REFERENCE SEMANTICS (design.md section 2): a handle class (StoryChild <
%   handle). Two Run proxies over the same `w:r` are views of one object. Run
%   adds NO oxml logic, NO registry rows and NO serialization code -- it is a
%   pure API/proxy over the already-registered CT_R (P4-1b). Equivalence is
%   therefore BEHAVIORAL, not byte-registry.
%
%   ATTRIBUTES (run.py 34-36): Python `__init__` sets
%   `self._r = self._element = self.element = r` -- three names for the SAME
%   run element. Ported verbatim: r_ (private, Python _r), element_ (private,
%   Python _element) and element (public, Python `self.element`). All hold the
%   same CT_R handle. `_r` is the working handle; `_element` is what `font`
%   wraps; `self.element` is a public attribute set but never read inside
%   run.py (ported for fidelity).
%
%   FONT DELEGATION (run.py 98-151, 227-249): bold/italic/underline delegate to
%   the Font proxy (P4-4a). `font` returns a FRESH Font(self._element) each
%   access (Python does not cache it), so the setters `self.font.bold = value`
%   mutate the run's `w:rPr` through a throwaway Font whose mutation persists on
%   the shared element.
%
%   STUBBED members (unported dependencies; design.md section 7 -- no silent
%   approximation):
%     * add_picture (run.py 59-81)         -> P7  (InlineShape / image inline shapes)
%     * iter_inner_content (run.py 153-174)-> P4-5b + P7 (RenderedPageBreak proxy /
%                                             Drawing; also CT_R.inner_content_items
%                                             is itself stubbed at P4-1b)
%     * style get/set (run.py 188-203)     -> LIVE at P4-7a (CharacterStyle via
%                                             DocumentPart.get_style / get_style_id)
%   NOTE (style): a Dependent read/write property delegating to
%   part().get_style / get_style_id with WD_STYLE_TYPE.CHARACTER. UN-STUBBED at
%   P4-7a: `run.style` resolves to a CharacterStyle (or the document default
%   character style) and `run.style = name/style/[]` applies/removes it.
%
%   Example:
%       r   = mat2doc.oxml.OxmlElement("w:r");
%       run = mat2doc.text.Run(r, someParagraph);
%       run.text = "a\tb";                    % <w:t>a</w:t><w:tab/><w:t>b</w:t>
%       run.bold = true;                      % <w:rPr><w:b/></w:rPr>
%       run.add_break(mat2doc.enum.text.WD_BREAK.PAGE);  % <w:br w:type="page"/>
%       tf  = run.contains_page_break;        % false
%
%   Ported from python-docx v1.2.0: src/docx/text/run.py::Run

    properties (Access = private)
        r_          % _r (run.py 36): the working <w:r> (a mat2doc.oxml.text.CT_R)
        element_    % _element (run.py 36): same handle; what `font` wraps
    end

    properties (SetAccess = private)
        element     % public `self.element` (run.py 36): same handle (set, never read in run.py)
    end

    properties (Dependent)
        bold                % bool|[] -- delegates to font.bold (w:rPr/w:b)
        italic              % bool|[] -- delegates to font.italic (w:rPr/w:i)
        underline           % bool|WD_UNDERLINE|[] -- delegates to font.underline (w:rPr/w:u)
        font                % Font (read-only) -- fresh Font(self._element) each access
        text                % string -- concatenated run inner-content text (CT_R.text)
        contains_page_break % bool -- any w:lastRenderedPageBreak descendant
        style               % CharacterStyle -- LIVE at P4-7a (resolves via the styles tier)
    end

    methods
        function obj = Run(r, parent)
            % RUN Wrap a `<w:r>` element (run.py 34-36).
            %
            %   Inputs:  r      - a mat2doc.oxml.text.CT_R (the `w:r` element).
            %            parent - the parent proxy (a ProvidesStoryPart) providing
            %                     `part`.
            %   Outputs: obj    - a scalar Run handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/text/run.py::Run.__init__
            obj@mat2doc.shared.StoryChild(parent);   % Python: super().__init__(parent)
            % Python: self._r = self._element = self.element = r (one element, three names)
            obj.r_ = r;
            obj.element_ = r;
            obj.element = r;
        end

        % ============================ add_break ============================
        function add_break(obj, break_type)
            % ADD_BREAK Add a `<w:br>` break of `break_type` to this run
            %   (run.py 38-57). `break_type` is a WD_BREAK member (P3-3),
            %   default WD_BREAK.LINE.
            %
            %   The (type_, clear) map (run.py 45-52):
            %     LINE             -> ([],           [])
            %     PAGE             -> ("page",       [])
            %     COLUMN           -> ("column",     [])
            %     LINE_CLEAR_LEFT  -> ("textWrapping","left")
            %     LINE_CLEAR_RIGHT -> ("textWrapping","right")
            %     LINE_CLEAR_ALL   -> ("textWrapping","all")
            %   ("textWrapping" is the CT_Br @w:type DEFAULT, so setting it
            %   REMOVES @w:type -- the LINE_CLEAR_* results carry only @w:clear,
            %   matching python-docx byte-for-byte.) Any other member reproduces
            %   Python's `{...}[break_type]` KeyError (mat2doc:KeyError).
            %
            %   Ported from python-docx v1.2.0: src/docx/text/run.py::Run.add_break
            arguments
                obj
                break_type = mat2doc.enum.text.WD_BREAK.LINE   % default WD_BREAK.LINE
            end
            WB = mat2doc.enum.text.WD_BREAK;
            if break_type == WB.LINE
                type_ = [];              clear = [];        % Python: (None, None)
            elseif break_type == WB.PAGE
                type_ = "page";          clear = [];        % Python: ("page", None)
            elseif break_type == WB.COLUMN
                type_ = "column";        clear = [];        % Python: ("column", None)
            elseif break_type == WB.LINE_CLEAR_LEFT
                type_ = "textWrapping";  clear = "left";    % Python: ("textWrapping","left")
            elseif break_type == WB.LINE_CLEAR_RIGHT
                type_ = "textWrapping";  clear = "right";   % Python: ("textWrapping","right")
            elseif break_type == WB.LINE_CLEAR_ALL
                type_ = "textWrapping";  clear = "all";     % Python: ("textWrapping","all")
            else
                % Python `{...}[break_type]` on an unmapped member -> KeyError.
                error("mat2doc:KeyError", "%s", string(break_type));
            end
            br = obj.r_.add_br();               % Python: br = self._r.add_br()
            if ~isequal(type_, [])              % Python: if type_ is not None
                br.type = type_;                %   br.type = type_
            end
            if ~isequal(clear, [])              % Python: if clear is not None
                br.clear = clear;               %   br.clear = clear
            end
        end

        % ============================ add_picture (STUB) ============================
        function inline = add_picture(obj, image_path_or_stream, width, height) %#ok<STOUT,INUSD>
            % ADD_PICTURE STUB (run.py 59-81). Owner: P7 (InlineShape /
            %   part.new_pic_inline / CT_R.add_drawing image path). Faithful body
            %   needs mat2doc.shape.InlineShape and StoryPart.new_pic_inline,
            %   neither ported -- raises mat2doc:notYetPorted.
            arguments
                obj %#ok<INUSA>
                image_path_or_stream %#ok<INUSA>
                width = []            %#ok<INUSA>  % Python default None
                height = []           %#ok<INUSA>  % Python default None
            end
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.shape.InlineShape + StoryPart.new_pic_inline (P7) " + ...
                "required by mat2doc.text.Run.add_picture");
        end

        % ============================ add_tab ============================
        function add_tab(obj)
            % ADD_TAB Append a `<w:tab/>` element, which Word interprets as a tab
            %   character (run.py 83-86).
            %
            %   Ported from python-docx v1.2.0: src/docx/text/run.py::Run.add_tab
            obj.r_.add_tab();                   % Python: self._r.add_tab()
        end

        % ============================ add_text ============================
        function t = add_text(obj, text)
            % ADD_TEXT Return a newly appended _Text (a new `<w:t>` child)
            %   containing `text` (run.py 88-96).
            %
            %   Inputs:  text - a string.
            %   Outputs: t    - a mat2doc.text.Text_ wrapping the new `w:t`.
            %
            %   Ported from python-docx v1.2.0: src/docx/text/run.py::Run.add_text
            arguments
                obj
                text (1,1) string
            end
            t = obj.r_.add_t(text);             % Python: t = self._r.add_t(text)
            t = mat2doc.text.Text_(t);          % Python: return _Text(t)
        end

        % ============================ bold ============================
        function value = get.bold(obj)
            % BOLD get (run.py 98-106): delegates to font.bold.
            value = obj.font.bold;              % Python: return self.font.bold
        end
        function set.bold(obj, value)
            % BOLD set (run.py 108-110): delegates to font.bold.
            obj.font.bold = value;              % Python: self.font.bold = value
        end

        % ============================ clear ============================
        function obj = clear(obj)
            % CLEAR Remove all run content, preserving run formatting (the
            %   `w:rPr`); return this run (run.py 112-118).
            %
            %   Ported from python-docx v1.2.0: src/docx/text/run.py::Run.clear
            obj.r_.clear_content();             % Python: self._r.clear_content()
            % Python: return self  (obj is the same handle)
        end

        % ============================ contains_page_break ============================
        function value = get.contains_page_break(obj)
            % CONTAINS_PAGE_BREAK true when one or more RENDERED page-breaks occur
            %   in this run (run.py 120-131). Hard (author) page-breaks are NOT
            %   counted. Returns a plain logical -- no RenderedPageBreak proxy
            %   needed. Python: `return bool(self._r.lastRenderedPageBreaks)`
            %   (H4: bool of the list -> ~isempty of the xpath result).
            value = ~isempty(obj.r_.lastRenderedPageBreaks());
        end

        % ============================ font ============================
        function value = get.font(obj)
            % FONT The Font proxy for this run's character formatting
            %   (run.py 133-137). Python `return Font(self._element)` -- a FRESH
            %   Font each access (not cached). Font(r) leaves parent = [] (None).
            value = mat2doc.text.Font(obj.element_);
        end

        % ============================ italic ============================
        function value = get.italic(obj)
            % ITALIC get (run.py 139-147): delegates to font.italic.
            value = obj.font.italic;            % Python: return self.font.italic
        end
        function set.italic(obj, value)
            % ITALIC set (run.py 149-151): delegates to font.italic.
            obj.font.italic = value;            % Python: self.font.italic = value
        end

        % ============================ iter_inner_content (STUB) ============================
        function items = iter_inner_content(obj) %#ok<STOUT,MANU>
            % ITER_INNER_CONTENT STUB (run.py 153-174). Yields str | Drawing |
            %   RenderedPageBreak in document order. Owner: P4-5b
            %   (RenderedPageBreak API proxy) + P7 (Drawing). It relies on
            %   CT_R.inner_content_items, which is itself stubbed at P4-1b
            %   (isinstance dispatch on CT_Drawing / CT_LastRenderedPageBreak).
            %   Stubbing the whole method (design.md section 7: no silent
            %   approximation) -- raises mat2doc:notYetPorted.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.text.RenderedPageBreak (P4-5b) + mat2doc.drawing.Drawing (P7) " + ...
                "required by mat2doc.text.Run.iter_inner_content");
        end

        % ============================ mark_comment_range ============================
        function mark_comment_range(obj, last_run, comment_id)
            % MARK_COMMENT_RANGE Mark the range of runs from this run to
            %   `last_run` (inclusive) as belonging to comment `comment_id`
            %   (run.py 176-186).
            %
            %   Ported from python-docx v1.2.0: src/docx/text/run.py::Run.mark_comment_range
            arguments
                obj
                last_run (1,1) mat2doc.text.Run
                comment_id (1,1) double
            end
            % -- insert w:commentRangeStart before this (first) run --
            obj.r_.insert_comment_range_start_above(comment_id);
            % -- insert w:commentRangeEnd + w:commentReference run after last_run --
            last_run.r_.insert_comment_range_end_and_reference_below(comment_id);
        end

        % ============================ style ============================
        function value = get.style(obj)
            % STYLE get (run.py 188-198): a CharacterStyle for the character style
            %   applied to this run. The document's default character style is
            %   returned when the run has no directly-applied character style.
            %   Python: style_id = self._r.style;
            %           return self.part.get_style(style_id, WD_STYLE_TYPE.CHARACTER)
            %   UN-STUBBED at P4-7a (DocumentPart.get_style is now live).
            %
            %   Ported from python-docx v1.2.0: src/docx/text/run.py::Run.style
            style_id = obj.r_.style;               % Python: style_id = self._r.style
            value = obj.part().get_style( ...      % Python: self.part.get_style(...)
                style_id, mat2doc.enum.style.WD_STYLE_TYPE.CHARACTER);
        end
        function set.style(obj, style_or_name)
            % STYLE set (run.py 200-203): apply character style `style_or_name`
            %   (a name or CharacterStyle); [] (None) removes any directly-applied
            %   character style. Python:
            %     style_id = self.part.get_style_id(style_or_name, WD_STYLE_TYPE.CHARACTER)
            %     self._r.style = style_id
            %   UN-STUBBED at P4-7a (DocumentPart.get_style_id is now live).
            %
            %   Ported from python-docx v1.2.0: src/docx/text/run.py::Run.style
            style_id = obj.part().get_style_id( ...% Python: self.part.get_style_id(...)
                style_or_name, mat2doc.enum.style.WD_STYLE_TYPE.CHARACTER);
            obj.r_.style = style_id;               % Python: self._r.style = style_id
        end

        % ============================ text ============================
        function value = get.text(obj)
            % TEXT get (run.py 205-221): the run's concatenated inner-content
            %   text. Python `return self._r.text` -- CT_R.text joins the text
            %   equivalent of each w:t/w:tab/w:br/... child (P4-1b).
            value = obj.r_.text;                % Python: return self._r.text
        end
        function set.text(obj, text)
            % TEXT set (run.py 223-225): replace run content from `text`; each
            %   \t -> <w:tab/>, each \n or \r -> <w:br/>; run formatting (w:rPr)
            %   preserved. Python `self._r.text = text` -> CT_R.text setter.
            obj.r_.text = text;                 % Python: self._r.text = text
        end

        % ============================ underline ============================
        function value = get.underline(obj)
            % UNDERLINE get (run.py 227-245): delegates to font.underline. Value
            %   is [] (inherited), true (single), false (no underline), or a
            %   WD_UNDERLINE member.
            value = obj.font.underline;         % Python: return self.font.underline
        end
        function set.underline(obj, value)
            % UNDERLINE set (run.py 247-249): delegates to font.underline.
            obj.font.underline = value;         % Python: self.font.underline = value
        end
    end
end

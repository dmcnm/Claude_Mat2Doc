classdef Styles < mat2doc.shared.ElementProxy
% STYLES Provides access to the styles defined in a document.
%
%   Accessed using the Document.styles property. Wraps the `<w:styles>` root
%   element (a CT_Styles) of the styles part. In python-docx it is a
%   collections-style sequence supporting len(), iteration, `in`, and
%   dictionary-style access by UI style name. An ElementProxy subclass: reference
%   semantics (handle) and H5 element-identity eq/ne are inherited unchanged.
%
%   VERIFY-COLLECTION (design.md section 2 "Collections -> RedefinesParen base"):
%   the shared 1-based () collection base is a FUTURE work package and does not
%   exist yet. Following the established precedent (TabStops, Mat2Ppt
%   _GradientStops), the Python sequence dunder surface is ported here as EXPLICIT
%   methods keeping line-for-line fidelity:
%       contains_  (__contains__)   getitem_ (__getitem__)
%       to_array   (__iter__)       len_     (__len__)
%   Dunder mapping (design.md): `name in styles` -> styles.contains_(name);
%   `styles[key]` -> styles.getitem_(key); `for s in styles` ->
%   `for s = styles.to_array()`; `len(styles)` -> styles.len_(). When the
%   collection base lands, Styles should derive from it and expose native ()
%   access. FLAGGED for the auditor/validator.
%
%   H1 (indexing): none of the sequence methods take an integer index -- keys are
%   style NAMES/IDS (strings), so there is no 0/1-based shift. `next(iter(xpath),
%   None)` first-match lookups live in CT_Styles (P4-6) and already return the
%   first element (1-based) or [].
%
%   H3 (tri-state): get_by_id(None,...) -> default; get_style_id(None,...) -> [];
%   default() -> [] when no default of that type; __getitem__ raises KeyError.
%
%   H5 (identity): to_array/getitem_ mint FRESH Style views via StyleFactory each
%   call (python-docx does not cache Style objects). to_array returns a
%   HETEROGENEOUS 1xN array typed BaseStyle (its subclasses ParagraphStyle/
%   CharacterStyle/TableStyle_/NumberingStyle_ mixed) -- enabled by the
%   matlab.mixin.Heterogeneous root on BaseStyle.
%
%   Example:
%       d  = mat2doc.Document();
%       st = d.styles;
%       st.len_()                          % number of styles
%       st.contains_("Heading 1")          % true if defined
%       h1 = st.getitem_("Heading 1");     % Python st["Heading 1"] -> ParagraphStyle
%       for s = st.to_array(); disp(s.style_id); end
%
%   Ported from python-docx v1.2.0: src/docx/styles/styles.py::Styles

    methods
        function obj = Styles(styles)
            % STYLES Wrap a `w:styles` element (styles.py 22-24).
            %
            %   Inputs:  styles - a mat2doc.oxml.styles.CT_Styles.
            %   Outputs: obj    - a scalar Styles handle.
            %
            %   Python: super().__init__(styles); self._element = styles.
            %   Ported from python-docx v1.2.0: styles/styles.py::Styles.__init__
            obj@mat2doc.shared.ElementProxy(styles);   % Python: super().__init__(styles)
            obj.element_ = styles;                     % Python: self._element = styles (redundant, faithful)
        end

        function tf = contains_(obj, name)
            % CONTAINS_ Enables the `in` operator on style name (styles.py 26-29).
            %   Python: internal_name = BabelFish.ui2internal(name);
            %           return any(style.name_val == internal_name
            %                      for style in self._element.style_lst)
            %   H4/H3: style.name_val is a string or None; `== internal_name`
            %   (a string) is value equality -- isequal reproduces it
            %   (None == "x" -> False).
            %
            %   Ported from python-docx v1.2.0: styles/styles.py::Styles.__contains__
            internal_name = mat2doc.styles.BabelFish.ui2internal(name);
            lst = obj.element_.style_lst;
            tf = false;
            for i = 1:numel(lst)   % Python: for style in self._element.style_lst
                if isequal(lst(i).name_val, internal_name)   % Python: style.name_val == internal_name
                    tf = true;
                    return
                end
            end
        end

        function style = getitem_(obj, key)
            % GETITEM_ Dictionary-style access by UI name (styles.py 31-47).
            %   Python:
            %     style_elm = self._element.get_by_name(BabelFish.ui2internal(key))
            %     if style_elm is not None: return StyleFactory(style_elm)
            %     style_elm = self._element.get_by_id(key)
            %     if style_elm is not None: warn(<deprecation>); return StyleFactory(style_elm)
            %     raise KeyError("no style with name '%s'" % key)
            %   Lookup by style id is deprecated (a UserWarning) and will be
            %   removed in a near-future release.
            %
            %   Ported from python-docx v1.2.0: styles/styles.py::Styles.__getitem__
            style_elm = obj.element_.get_by_name( ...
                mat2doc.styles.BabelFish.ui2internal(key));
            if ~isequal(style_elm, [])   % Python: if style_elm is not None
                style = mat2doc.styles.StyleFactory(style_elm);
                return
            end
            style_elm = obj.element_.get_by_id(key);
            if ~isequal(style_elm, [])   % Python: if style_elm is not None
                % Python: warn(msg, UserWarning, stacklevel=2). MATLAB warnings
                % differ in mechanism from Python's; the message text is preserved.
                warning("mat2doc:UserWarning", "%s", ...
                    "style lookup by style_id is deprecated. Use style name as key instead.");
                style = mat2doc.styles.StyleFactory(style_elm);
                return
            end
            % Python: raise KeyError("no style with name '%s'" % key)
            error("mat2doc:KeyError", "%s", ...
                sprintf("no style with name '%s'", key));
        end

        function result = to_array(obj)
            % TO_ARRAY A Style per `w:style` child, in document order (styles.py 49-50).
            %   Python __iter__: (StyleFactory(style) for style in
            %   self._element.style_lst). Materialized to a heterogeneous 1xN
            %   BaseStyle array (design.md iteration idiom: `for s in styles` ->
            %   `for s = styles.to_array()`). No styles -> a 1x0 BaseStyle array.
            %
            %   Ported from python-docx v1.2.0: styles/styles.py::Styles.__iter__
            lst = obj.element_.style_lst;
            result = mat2doc.styles.BaseStyle.empty(1, 0);
            for k = 1:numel(lst)   % Python: for style in self._element.style_lst
                result(k) = mat2doc.styles.StyleFactory(lst(k));
            end
        end

        function n = len_(obj)
            % LEN_ Number of styles (styles.py 52-53).
            %   Python __len__: len(self._element.style_lst).
            %
            %   Ported from python-docx v1.2.0: styles/styles.py::Styles.__len__
            n = numel(obj.element_.style_lst);
        end

        function style = add_style(obj, name, style_type, builtin)
            % ADD_STYLE Add and return a new style of `style_type` named `name` (styles.py 55-65).
            %   A builtin style is created by passing true for `builtin` (default false).
            %   Python:
            %     style_name = BabelFish.ui2internal(name)
            %     if style_name in self: raise ValueError("document already contains style '%s'" % name)
            %     style = self._element.add_style_of_type(style_name, style_type, builtin)
            %     return StyleFactory(style)
            %   NOTE: `style_name in self` re-applies ui2internal inside contains_
            %   (idempotent on the internal names produced here) -- ported verbatim.
            %
            %   Ported from python-docx v1.2.0: styles/styles.py::Styles.add_style
            arguments
                obj
                name
                style_type
                builtin = false   % Python default builtin=False
            end
            style_name = mat2doc.styles.BabelFish.ui2internal(name);
            if obj.contains_(style_name)   % Python: if style_name in self
                error("mat2doc:ValueError", "%s", ...
                    sprintf("document already contains style '%s'", name));
            end
            style_elm = obj.element_.add_style_of_type(style_name, style_type, builtin);
            style = mat2doc.styles.StyleFactory(style_elm);
        end

        function style = default(obj, style_type)
            % DEFAULT The default style for `style_type`, or [] if none (styles.py 67-73).
            %   Python:
            %     style = self._element.default_for(style_type)
            %     if style is None: return None
            %     return StyleFactory(style)
            %
            %   Ported from python-docx v1.2.0: styles/styles.py::Styles.default
            style_elm = obj.element_.default_for(style_type);
            if isequal(style_elm, [])   % Python: if style is None
                style = [];
                return
            end
            style = mat2doc.styles.StyleFactory(style_elm);
        end

        function style = get_by_id(obj, style_id, style_type)
            % GET_BY_ID The style of `style_type` matching `style_id` (styles.py 75-83).
            %   Returns the default for `style_type` if `style_id` is [] (None) or
            %   not found or of the wrong type. Python:
            %     if style_id is None: return self.default(style_type)
            %     return self._get_by_id(style_id, style_type)
            %
            %   Ported from python-docx v1.2.0: styles/styles.py::Styles.get_by_id
            if isequal(style_id, [])   % Python: if style_id is None
                style = obj.default(style_type);
                return
            end
            style = obj.get_by_id_(style_id, style_type);
        end

        function style_id = get_style_id(obj, style_or_name, style_type)
            % GET_STYLE_ID The id of the style matching `style_or_name`, or [] (styles.py 85-98).
            %   Python:
            %     if style_or_name is None: return None
            %     elif isinstance(style_or_name, BaseStyle):
            %         return self._get_style_id_from_style(style_or_name, style_type)
            %     else:
            %         return self._get_style_id_from_name(style_or_name, style_type)
            %   H10 (isinstance): isa(style_or_name, 'mat2doc.styles.BaseStyle')
            %   captures every Style subclass (ParagraphStyle/CharacterStyle/...).
            %
            %   Ported from python-docx v1.2.0: styles/styles.py::Styles.get_style_id
            if isequal(style_or_name, [])   % Python: if style_or_name is None
                style_id = [];
                return
            end
            if isa(style_or_name, "mat2doc.styles.BaseStyle")   % Python: isinstance(..., BaseStyle)
                style_id = obj.get_style_id_from_style_(style_or_name, style_type);
            else
                style_id = obj.get_style_id_from_name_(style_or_name, style_type);
            end
        end

        function ls = latent_styles(obj) %#ok<MANU,STOUT>
            % LATENT_STYLES STUB (styles.py 100-105, @property). Owner: P4-7b.
            %   Faithful body: return LatentStyles(self._element.get_or_add_latentStyles()).
            %   The LatentStyles proxy (docx/styles/latent.py) lands at P4-7b; this
            %   accessor stays a clean notYetPorted stub naming that owner. The
            %   underlying CT_Styles.get_or_add_latentStyles descriptor IS live (P4-6).
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.styles.LatentStyles (owning WP: P4-7b latent-styles tier) " + ...
                "required by mat2doc.styles.Styles.latent_styles");
        end
    end

    methods (Access = private)
        function style = get_by_id_(obj, style_id, style_type)
            % GET_BY_ID_ The style of `style_type` matching `style_id`, else default (styles.py 107-116).
            %   Python _get_by_id:
            %     style = self._element.get_by_id(style_id) if style_id else None
            %     if style is None or style.type != style_type: return self.default(style_type)
            %     return StyleFactory(style)
            %   H4: `if style_id` -- a falsy style_id (None or "") yields None
            %   (Python empty string is falsy). Reproduced with an explicit
            %   truthiness test over the {string, None} domain.
            %
            %   Ported from python-docx v1.2.0: styles/styles.py::Styles._get_by_id
            if ~isequal(style_id, []) && strlength(style_id) > 0   % Python: if style_id (H4: ""/None falsy)
                style_elm = obj.element_.get_by_id(style_id);
            else
                style_elm = [];
            end
            % Python: if style is None or style.type != style_type
            if isequal(style_elm, []) || ~isequal(style_elm.type, style_type)
                style = obj.default(style_type);
                return
            end
            style = mat2doc.styles.StyleFactory(style_elm);
        end

        function style_id = get_style_id_from_name_(obj, style_name, style_type)
            % GET_STYLE_ID_FROM_NAME_ Id of the style named `style_name` (styles.py 118-125).
            %   Python: return self._get_style_id_from_style(self[style_name], style_type)
            %   -- self[style_name] (getitem_) raises KeyError if the name is not
            %   found; _get_style_id_from_style raises ValueError on a type mismatch
            %   and returns None when the style is the default.
            %
            %   Ported from python-docx v1.2.0: styles/styles.py::Styles._get_style_id_from_name
            style_id = obj.get_style_id_from_style_( ...
                obj.getitem_(style_name), style_type);   % Python: self[style_name]
        end

        function style_id = get_style_id_from_style_(obj, style, style_type)
            % GET_STYLE_ID_FROM_STYLE_ Id of `style`, or [] if it is the default (styles.py 127-136).
            %   Python:
            %     if style.type != style_type:
            %         raise ValueError("assigned style is type %s, need type %s" % (style.type, style_type))
            %     if style == self.default(style_type): return None
            %     return style.style_id
            %   H5: `style == self.default(...)` is a SCALAR ElementProxy identity
            %   compare (same wrapped w:style element). The default may be [] (None);
            %   comparing a Style proxy to [] via ElementProxy.eq returns false
            %   (not both proxies) -- matching Python (a Style is never `== None`).
            %
            %   Ported from python-docx v1.2.0: styles/styles.py::Styles._get_style_id_from_style
            if ~isequal(style.type, style_type)   % Python: if style.type != style_type
                error("mat2doc:ValueError", "%s", sprintf( ...
                    "assigned style is type %s, need type %s", ...
                    mat2doc.styles.Styles.styletype_str_(style.type), ...
                    mat2doc.styles.Styles.styletype_str_(style_type)));
            end
            if style == obj.default(style_type)   % Python: if style == self.default(style_type)
                style_id = [];
                return
            end
            style_id = style.style_id;
        end
    end

    methods (Static, Access = private)
        function s = styletype_str_(style_type)
            % STYLETYPE_STR_ Python str() of a WD_STYLE_TYPE member for the
            %   ValueError message (styles.py 133). Python `"%s" % member` uses
            %   str(member); a docx BaseEnum's str is "<NAME> (<value>)", e.g.
            %   str(WD_STYLE_TYPE.PARAGRAPH) == "PARAGRAPH (1)" (verified against
            %   the pinned clone). A [] (None) type prints as "None".
            if isequal(style_type, [])
                s = "None";
                return
            end
            s = sprintf("%s (%d)", string(style_type), double(style_type.value));
        end
    end
end

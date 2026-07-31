classdef LatentStyles < mat2doc.shared.ElementProxy
% LATENTSTYLES Default behaviors for latent styles and the _LatentStyle overrides.
%
%   Provides access to the default behaviors for latent styles in this document
%   and to the collection of |LatentStyle_| objects that define overrides of those
%   defaults for a particular named latent style. Accessed via Styles.latent_styles.
%   Wraps the `<w:latentStyles>` root child of the styles part (a CT_LatentStyles).
%   An ElementProxy subclass: reference semantics (handle) and H5 element-identity
%   eq/ne are inherited unchanged.
%
%   VERIFY-COLLECTION (design.md section 2 "Collections -> RedefinesParen base"):
%   the shared 1-based () collection base is a FUTURE work package. Following the
%   established precedent (Styles, TabStops, Mat2Ppt _GradientStops), the Python
%   sequence dunder surface is ported here as EXPLICIT methods keeping line-for-
%   line fidelity:
%       getitem_ (__getitem__)   to_array (__iter__)   len_ (__len__)
%   Dunder mapping (design.md): `latent_styles[key]` -> ls.getitem_(key);
%   `for s in latent_styles` -> `for s = ls.to_array()`; `len(latent_styles)` ->
%   ls.len_(). When the collection base lands, LatentStyles should derive from it
%   and expose native () access. FLAGGED for the auditor/validator.
%
%   SCOPE NOTE (faithful to v1.2.0 SOURCE, not the WP brief prose): the actual
%   latent.py LatentStyles class (lines 7-107) exposes EXACTLY __getitem__,
%   __iter__, __len__, add_latent_style, default_priority, default_to_hidden,
%   default_to_locked, default_to_quick_style, default_to_unhide_when_used, and
%   load_count. It has NO `styles` list property, NO `count`/`element`/
%   `default_to_builtin` members (those named in the WP brief do not exist in
%   v1.2.0). `element` is inherited from ElementProxy; `load_count` is the member
%   that reads/writes the underlying `w:count` attribute. Ported to the SOURCE.
%
%   H1 (indexing): the sequence methods take a style NAME (string) key, never an
%   integer index -- no 0/1-based shift. to_array materializes the __iter__
%   generator to a 1xN LatentStyle_ array (a LatentStyle_ per <w:lsdException>,
%   document order); len_ counts those children; empty -> a 1x0 array.
%
%   H3 (tri-state): default_priority / load_count are ST_DecimalNumber optional
%   attrs -> int or [] (None). The default_to_* booleans read
%   CT_LatentStyles.bool_prop, which returns the logical FALSE (not []) when the
%   attribute is absent -- the "effective default" python-docx reports; their
%   setters go through set_bool_prop (F-1-fixed: bool(None)->False writes "0").
%   getitem_ raises KeyError when no latent style has the given name.
%
%   Example:
%       d  = mat2doc.Document();
%       ls = d.styles.latent_styles;
%       ls.len_()                              % number of lsdException overrides
%       n  = ls.getitem_("Normal");            % Python ls["Normal"] -> LatentStyle_
%       x  = ls.add_latent_style("Table Grid");% add and return a new override
%       ls.default_priority = 99;              % w:defUIPriority
%       ls.default_to_hidden = true;           % w:defSemiHidden
%
%   Ported from python-docx v1.2.0: src/docx/styles/latent.py::LatentStyles

    properties (Dependent)
        default_priority            % int|[] -- w:defUIPriority (default sort order)
        default_to_hidden           % bool -- w:defSemiHidden (effective default)
        default_to_locked           % bool -- w:defLockedState
        default_to_quick_style      % bool -- w:defQFormat
        default_to_unhide_when_used % bool -- w:defUnhideWhenUsed
        load_count                  % int|[] -- w:count (# built-in styles to init)
    end

    methods
        function obj = LatentStyles(element)
            % LATENTSTYLES Wrap a `w:latentStyles` element.
            %   python-docx has no custom __init__ on LatentStyles; it inherits
            %   ElementProxy.__init__(element, parent=None). Constructed by
            %   Styles.latent_styles as LatentStyles(get_or_add_latentStyles()).
            %
            %   Inputs:  element - a mat2doc.oxml.styles.CT_LatentStyles.
            %   Outputs: obj     - a scalar LatentStyles handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/shared.py::ElementProxy.__init__
            obj@mat2doc.shared.ElementProxy(element);   % Python: super().__init__(element) (parent None)
        end

        function latent_style = getitem_(obj, key)
            % GETITEM_ Dictionary-style access to a latent style by name (latent.py 12-18).
            %   Python:
            %     style_name = BabelFish.ui2internal(key)
            %     lsdException = self._element.get_by_name(style_name)
            %     if lsdException is None: raise KeyError("no latent style with name '%s'" % key)
            %     return _LatentStyle(lsdException)
            %
            %   Ported from python-docx v1.2.0: styles/latent.py::LatentStyles.__getitem__
            style_name = mat2doc.styles.BabelFish.ui2internal(key);
            lsdException = obj.element_.get_by_name(style_name);
            if isequal(lsdException, [])   % Python: if lsdException is None
                error("mat2doc:KeyError", "%s", ...
                    sprintf("no latent style with name '%s'", key));
            end
            latent_style = mat2doc.styles.LatentStyle_(lsdException);
        end

        function result = to_array(obj)
            % TO_ARRAY A LatentStyle_ per <w:lsdException> child, document order (latent.py 20-21).
            %   Python __iter__: (_LatentStyle(ls) for ls in self._element.lsdException_lst).
            %   Materialized to a 1xN LatentStyle_ array (design.md iteration idiom
            %   `for s in latent_styles` -> `for s = latent_styles.to_array()`). No
            %   children -> a 1x0 LatentStyle_ array.
            %
            %   Ported from python-docx v1.2.0: styles/latent.py::LatentStyles.__iter__
            lst = obj.element_.lsdException_lst;
            result = mat2doc.styles.LatentStyle_.empty(1, 0);
            for k = 1:numel(lst)   % Python: for ls in self._element.lsdException_lst
                result(k) = mat2doc.styles.LatentStyle_(lst(k));
            end
        end

        function n = len_(obj)
            % LEN_ Number of latent-style overrides (latent.py 23-24).
            %   Python __len__: len(self._element.lsdException_lst).
            %
            %   Ported from python-docx v1.2.0: styles/latent.py::LatentStyles.__len__
            n = numel(obj.element_.lsdException_lst);
        end

        function latent_style = add_latent_style(obj, name)
            % ADD_LATENT_STYLE Add and return a new _LatentStyle override for `name` (latent.py 26-31).
            %   Python:
            %     lsdException = self._element.add_lsdException()
            %     lsdException.name = BabelFish.ui2internal(name)
            %     return _LatentStyle(lsdException)
            %
            %   Ported from python-docx v1.2.0: styles/latent.py::LatentStyles.add_latent_style
            lsdException = obj.element_.add_lsdException();
            lsdException.name = mat2doc.styles.BabelFish.ui2internal(name);
            latent_style = mat2doc.styles.LatentStyle_(lsdException);
        end

        % ---- default_priority (latent.py 33-44) ----
        function value = get.default_priority(obj)
            % Python: return self._element.defUIPriority (int, or None when absent).
            value = obj.element_.defUIPriority;
        end
        function set.default_priority(obj, value)
            % Python: self._element.defUIPriority = value.
            obj.element_.defUIPriority = value;
        end

        % ---- default_to_hidden (latent.py 46-57) ----
        function value = get.default_to_hidden(obj)
            % Python: return self._element.bool_prop("defSemiHidden") (False when absent).
            value = obj.element_.bool_prop("defSemiHidden");
        end
        function set.default_to_hidden(obj, value)
            % Python: self._element.set_bool_prop("defSemiHidden", value).
            obj.element_.set_bool_prop("defSemiHidden", value);
        end

        % ---- default_to_locked (latent.py 59-72) ----
        function value = get.default_to_locked(obj)
            % Python: return self._element.bool_prop("defLockedState") (False when absent).
            value = obj.element_.bool_prop("defLockedState");
        end
        function set.default_to_locked(obj, value)
            % Python: self._element.set_bool_prop("defLockedState", value).
            obj.element_.set_bool_prop("defLockedState", value);
        end

        % ---- default_to_quick_style (latent.py 74-82) ----
        function value = get.default_to_quick_style(obj)
            % Python: return self._element.bool_prop("defQFormat") (False when absent).
            value = obj.element_.bool_prop("defQFormat");
        end
        function set.default_to_quick_style(obj, value)
            % Python: self._element.set_bool_prop("defQFormat", value).
            obj.element_.set_bool_prop("defQFormat", value);
        end

        % ---- default_to_unhide_when_used (latent.py 84-92) ----
        function value = get.default_to_unhide_when_used(obj)
            % Python: return self._element.bool_prop("defUnhideWhenUsed") (False when absent).
            value = obj.element_.bool_prop("defUnhideWhenUsed");
        end
        function set.default_to_unhide_when_used(obj, value)
            % Python: self._element.set_bool_prop("defUnhideWhenUsed", value).
            obj.element_.set_bool_prop("defUnhideWhenUsed", value);
        end

        % ---- load_count (latent.py 94-107) ----
        function value = get.load_count(obj)
            % Python: return self._element.count (int, or None when absent).
            value = obj.element_.count;
        end
        function set.load_count(obj, value)
            % Python: self._element.count = value.
            obj.element_.count = value;
        end
    end
end

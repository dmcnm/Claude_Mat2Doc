classdef BaseStyle < mat2doc.shared.ElementProxy & matlab.mixin.Heterogeneous
% BASESTYLE Base class for the paragraph/character/table/numbering style objects.
%
%   These properties and methods are inherited by all style objects. An
%   ElementProxy subclass: reference semantics (handle) and H5 element-identity
%   eq/ne are inherited unchanged. Every read/write property reaches through the
%   wrapped `w:style` element (a CT_Style) to its `.._val` helpers (P4-6);
%   BaseStyle adds NO oxml logic, NO registry rows and NO serialization code.
%
%   HETEROGENEOUS ROOT (design.md section 2 "Collections -> Heterogeneous"):
%   BaseStyle additionally derives matlab.mixin.Heterogeneous so that a mixed
%   vector of its subclasses (ParagraphStyle / CharacterStyle / TableStyle_ /
%   NumberingStyle_) -- as produced by Styles.to_array (Python Styles.__iter__)
%   and StyleFactory -- can be held in a single 1xN array typed BaseStyle. This
%   is the "downstream collection WP that needs a heterogeneous proxy vector"
%   anticipated in the ElementProxy note. SEALING: only methods actually invoked
%   ARRAY-WISE on a heterogeneous BaseStyle array need Sealed (design.md section
%   2, integration-critical). Styles uses eq ONLY scalar-to-scalar
%   (`style == self.default(...)`) and never array-wise, so ElementProxy's
%   unsealed eq/ne are sufficient and are left unsealed (verified: scalar eq over
%   a hetero array element works; array-wise `arr == x` is neither used here nor
%   by python-docx). If a future consumer needs array-wise eq on styles, seal it
%   here with an array-aware body then.
%
%   CONSTRUCTOR (style.py 34-36): `BaseStyle(style_elm)` ->
%   `super().__init__(style_elm); self._style_elm = style_elm`. The super call
%   passes only the element (parent defaults None). `_style_elm` (-> style_elm_)
%   duplicates the wrapped element and is read by the style_id/type getters
%   (which use `self._style_elm`, not `self._element`); both refer to the same
%   CT_Style. Ported verbatim for fidelity.
%
%   H3 (tri-state) across the properties (style.py 38-161):
%     * builtin (read-only): Python `not self._element.customStyle` -- a
%       truthiness negation over customStyle {None, True, False} (H4).
%     * name: BabelFish.internal2ui of name_val, or [] when name_val is None.
%     * priority: uiPriority_val -> int or [] (None).
%     * hidden/locked/quick_style/unhide_when_used: the CT_OnOff-backed booleans;
%       their GETTERS return the logical FALSE (not []) when absent -- matching
%       python-docx, whose semiHidden_val/locked_val/qFormat_val/
%       unhideWhenUsed_val return bool False when the child is missing.
%     * style_id: the raw @w:styleId (string or [] when absent).
%     * type: WD_STYLE_TYPE member; None defaults to PARAGRAPH (style.py 144-147).
%
%   DELETE_STYLE (style.py 49-57) -- named `delete_style` (by the kind of thing
%   it removes) so NO method named `delete` exists anywhere on the styles
%   surface. MATLAB's `delete` on a handle IS the destructor; with no `delete`
%   override in this hierarchy the destructor is never touched and no guarded-
%   destructor machinery is needed (H17 dissolved). Pure naming choice, byte-
%   identical to python-docx style.delete(). See the delete_style method below.
%
%   Example:
%       se = mat2doc.oxml.OxmlElement("w:style");
%       se.type = mat2doc.enum.style.WD_STYLE_TYPE.PARAGRAPH;
%       s = mat2doc.styles.ParagraphStyle(se);
%       s.name = "Heading 1";       % writes internal "heading 1" via BabelFish
%       s.style_id = "Heading1";
%       s.priority = 10; s.hidden = true; s.quick_style = true;
%
%   Ported from python-docx v1.2.0: src/docx/styles/style.py::BaseStyle

    properties (Access = protected)
        style_elm_   % Python self._style_elm (style.py 36): the wrapped CT_Style
    end

    properties (Dependent)
        builtin           % bool (read-only) -- not customStyle (H4)
        hidden            % bool -- w:semiHidden
        locked            % bool -- w:locked
        name              % string|[] -- <w:name> via BabelFish (UI name)
        priority          % int|[] -- w:uiPriority
        quick_style       % bool -- w:qFormat
        style_id          % string|[] -- @w:styleId
        type              % WD_STYLE_TYPE -- @w:type (None -> PARAGRAPH)
        unhide_when_used  % bool -- w:unhideWhenUsed
    end

    methods
        function obj = BaseStyle(style_elm)
            % BASESTYLE Wrap a `w:style` element `style_elm` (style.py 34-36).
            %
            %   Inputs:  style_elm - a mat2doc.oxml.styles.CT_Style.
            %   Outputs: obj       - a scalar BaseStyle (subclass) handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/styles/style.py::BaseStyle.__init__
            obj@mat2doc.shared.ElementProxy(style_elm);   % Python: super().__init__(style_elm)
            obj.style_elm_ = style_elm;                   % Python: self._style_elm = style_elm
        end

        % ---- builtin (read-only; style.py 38-47) ----
        function value = get.builtin(obj)
            % Python: return not self._element.customStyle (H4 truthiness).
            %   customStyle is [] (None), true, or false; `not None`=True,
            %   `not True`=False, `not False`=True.
            cs = obj.element_.customStyle;
            value = ~(~isequal(cs, []) && cs);
        end

        % ---- delete_style (style.py 49-57; renamed from delete to avoid the name) ----
        function delete_style(obj)
            % DELETE_STYLE Remove this style definition from the document.
            %
            %   Python (style.py 49-57, `BaseStyle.delete`):
            %       self._element.delete()   % detach the w:style from w:styles
            %       self._element = None
            %   Note calling this does NOT change the style applied to any content;
            %   content with the deleted style renders using the default.
            %
            %   NAMING (H17 dissolved): named `delete_style` -- by the kind of thing
            %   it removes -- so NO method named `delete` exists on the styles
            %   surface. MATLAB's `delete` on a `handle` IS the destructor; with no
            %   `delete` override anywhere in this hierarchy the destructor is never
            %   touched and no guarded-destructor machinery is needed. Pure naming
            %   choice, byte-identical to python-docx `style.delete()` -- NOT an
            %   output deviation, NO D-number. The wrapped element method is likewise
            %   `delete_style` (CT_Style), an ordinary getparent().remove() that is
            %   never GC's destructor, so the detached element survives (as in Python)
            %   rather than being destroyed.
            %
            %   Gate-4 rule: after delete_style the proxy's element_ is [] (Python
            %   None), so accessing any proxy property errors -- matching python-docx
            %   (AttributeError). Assert only the parent-side effect (child count /
            %   serialized bytes); never inspect the proxy after delete_style.
            %
            %   Ported from python-docx v1.2.0: src/docx/styles/style.py::BaseStyle.delete
            obj.element_.delete_style();   % Python: self._element.delete() (detach w:style)
            obj.element_ = [];             % Python: self._element = None
        end

        % ---- hidden (style.py 59-71) ----
        function value = get.hidden(obj)
            % Python: return self._element.semiHidden_val (False when absent).
            value = obj.element_.semiHidden_val;
        end
        function set.hidden(obj, value)
            obj.element_.semiHidden_val = value;
        end

        % ---- locked (style.py 73-86) ----
        function value = get.locked(obj)
            % Python: return self._element.locked_val (False when absent).
            value = obj.element_.locked_val;
        end
        function set.locked(obj, value)
            obj.element_.locked_val = value;
        end

        % ---- name (style.py 88-98) ----
        function value = get.name(obj)
            % Python (style.py 89-94): name = self._element.name_val;
            %   if name is None: return None; return BabelFish.internal2ui(name).
            nm = obj.element_.name_val;
            if isequal(nm, [])   % Python: if name is None
                value = [];
                return
            end
            value = mat2doc.styles.BabelFish.internal2ui(nm);
        end
        function set.name(obj, value)
            % Python (style.py 96-98): self._element.name_val = value.
            obj.element_.name_val = value;
        end

        % ---- priority (style.py 100-112) ----
        function value = get.priority(obj)
            % Python: return self._element.uiPriority_val ([] when absent).
            value = obj.element_.uiPriority_val;
        end
        function set.priority(obj, value)
            obj.element_.uiPriority_val = value;
        end

        % ---- quick_style (style.py 114-125) ----
        function value = get.quick_style(obj)
            % Python: return self._element.qFormat_val (False when absent).
            value = obj.element_.qFormat_val;
        end
        function set.quick_style(obj, value)
            obj.element_.qFormat_val = value;
        end

        % ---- style_id (style.py 127-138) ----
        function value = get.style_id(obj)
            % Python (style.py 134): return self._style_elm.styleId  (uses _style_elm).
            value = obj.style_elm_.styleId;
        end
        function set.style_id(obj, value)
            % Python (style.py 137-138): self._element.styleId = value  (uses _element).
            obj.element_.styleId = value;
        end

        % ---- type (read-only; style.py 140-147) ----
        function value = get.type(obj)
            % Python (style.py 144-147): type = self._style_elm.type;
            %   if type is None: return WD_STYLE_TYPE.PARAGRAPH; return type.
            t = obj.style_elm_.type;
            if isequal(t, [])   % Python: if type is None
                value = mat2doc.enum.style.WD_STYLE_TYPE.PARAGRAPH;
                return
            end
            value = t;
        end

        % ---- unhide_when_used (style.py 149-161) ----
        function value = get.unhide_when_used(obj)
            % Python: return self._element.unhideWhenUsed_val (False when absent).
            value = obj.element_.unhideWhenUsed_val;
        end
        function set.unhide_when_used(obj, value)
            obj.element_.unhideWhenUsed_val = value;
        end
    end
end

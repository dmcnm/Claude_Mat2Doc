classdef LatentStyle_ < mat2doc.shared.ElementProxy
% LATENTSTYLE_ Proxy for a `w:lsdException` element: one latent-style override.
%
%   Specifies display behaviors for a built-in style when no definition for that
%   style is stored yet in the styles.xml part. The values in this element
%   override the defaults specified in the parent `w:latentStyles` element. An
%   ElementProxy subclass: reference semantics (handle) and H5 element-identity
%   eq/ne are inherited unchanged.
%
%   FLAG-3 (naming): Python's private `_LatentStyle` maps to the trailing-
%   underscore convention LatentStyle_ (leading-underscore rotation:
%   _LatentStyle -> LatentStyle_), matching the sibling TableStyle_ /
%   NumberingStyle_ renames.
%
%   H3 (tri-state): hidden/locked/quick_style/unhide_when_used read
%   CT_LsdException.on_off_prop, which returns the RAW tri-state {[] (None),
%   true, false} -- [] meaning "inherit the parent latentStyles default"
%   (CONTRAST BaseStyle's *_val, which collapse absent to False). name reads the
%   REQUIRED @w:name (always present) through BabelFish.internal2ui; priority is
%   the ST_DecimalNumber @w:uiPriority (int or []).
%
%   DELETE_ (latent.py 119-128; H17 FLAG-3 method rename delete()->delete_()):
%   the MATLAB `delete` name is the handle destructor and is NOT overridden here
%   (GC-safe); the faithful element-removal is exposed as delete_(), byte-
%   identical to python-docx `_LatentStyle.delete()`. See the delete_ method and
%   the H17 addendum below. This mirrors BaseStyle.delete_ exactly (design.md
%   section 9 H17 addendum; decision_2026-07-30_h17_delete_destructor.md names
%   P4-7b explicitly).
%
%   Example:
%       d  = mat2doc.Document();
%       ls = d.styles.latent_styles;
%       x  = ls.add_latent_style("Table Grid");
%       x.priority = 59;          % w:uiPriority
%       x.hidden = false;         % w:semiHidden="0"
%       x.quick_style = true;     % w:qFormat="1"
%       x.delete_();              % remove the <w:lsdException> from its parent
%
%   Ported from python-docx v1.2.0: src/docx/styles/latent.py::_LatentStyle

    properties (Dependent)
        hidden           % bool|[] -- w:semiHidden (tri-state; [] inherits default)
        locked           % bool|[] -- w:locked
        name             % string (read-only) -- <w:name> via BabelFish (UI name)
        priority         % int|[] -- w:uiPriority
        quick_style      % bool|[] -- w:qFormat
        unhide_when_used % bool|[] -- w:unhideWhenUsed
    end

    methods
        function obj = LatentStyle_(lsdException)
            % LATENTSTYLE_ Wrap a `w:lsdException` element.
            %   python-docx has no custom __init__ on _LatentStyle; it inherits
            %   ElementProxy.__init__(element, parent=None).
            %
            %   Inputs:  lsdException - a mat2doc.oxml.styles.CT_LsdException.
            %   Outputs: obj          - a scalar LatentStyle_ handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/shared.py::ElementProxy.__init__
            obj@mat2doc.shared.ElementProxy(lsdException);   % Python: super().__init__ (parent None)
        end

        % ---- delete_ (latent.py 119-128; H17 FLAG-3 rename delete -> delete_) ----
        function delete_(obj)
            % DELETE_ Remove this latent-style definition so the containing
            %   LatentStyles defaults provide the effective value again (latent.py 119-128).
            %
            %   Python (latent.py 127-128, `_LatentStyle.delete`):
            %       self._element.delete()   % detach the w:lsdException from w:latentStyles
            %       self._element = None
            %   Python notes that accessing any attribute afterward raises
            %   AttributeError (self._element is now None).
            %
            %   H17 ADDENDUM (proxy-layer, SIGNED-PROVISIONAL 2026-07-30,
            %   design.md section 9 + decision_2026-07-30_h17_delete_destructor.md,
            %   which names P4-7b explicitly). In MATLAB `delete` on a `handle` IS
            %   the destructor; overriding it on a transient proxy is UNSAFE (GC
            %   would detach the live element). RESOLUTION (FLAG-3 method rename,
            %   same convention as _TableStyle -> TableStyle_): DO NOT override
            %   `delete`; expose the faithful element-removal as `delete_`. This
            %   class therefore does NOT declare a `delete` override -- the default
            %   handle destructor is left in place (GC-safe: no tree effect).
            %
            %   `delete_` is only ever called EXPLICITLY (never by GC) and detaches
            %   a PARENTED element via the P4-6/oxml CT_LsdException.delete (a
            %   guarded getparent().remove(self)) -- so the result is BYTE-IDENTICAL
            %   to python-docx `_LatentStyle.delete()`. Method-naming resolution
            %   (FLAG-3), NOT an output deviation -- NO D-number.
            %
            %   H17 Gate-4 rule: after delete_ the MATLAB element handle is invalid
            %   whereas Python's element survives detached -- NEVER inspect the
            %   handle after delete_; assert only the parent-side effect (child
            %   count / serialized bytes).
            %
            %   Ported from python-docx v1.2.0: styles/latent.py::_LatentStyle.delete
            obj.element_.delete();   % Python: self._element.delete() (CT_LsdException.delete)
            obj.element_ = [];       % Python: self._element = None
        end

        % ---- hidden (latent.py 130-142) ----
        function value = get.hidden(obj)
            % Python: return self._element.on_off_prop("semiHidden") (tri-state).
            value = obj.element_.on_off_prop("semiHidden");
        end
        function set.hidden(obj, value)
            % Python: self._element.set_on_off_prop("semiHidden", value).
            obj.element_.set_on_off_prop("semiHidden", value);
        end

        % ---- locked (latent.py 144-156) ----
        function value = get.locked(obj)
            % Python: return self._element.on_off_prop("locked") (tri-state).
            value = obj.element_.on_off_prop("locked");
        end
        function set.locked(obj, value)
            % Python: self._element.set_on_off_prop("locked", value).
            obj.element_.set_on_off_prop("locked", value);
        end

        % ---- name (read-only; latent.py 158-161) ----
        function value = get.name(obj)
            % Python: return BabelFish.internal2ui(self._element.name). @w:name is
            %   REQUIRED, so self._element.name always returns a string (no [] guard).
            value = mat2doc.styles.BabelFish.internal2ui(obj.element_.name);
        end

        % ---- priority (latent.py 163-170) ----
        function value = get.priority(obj)
            % Python: return self._element.uiPriority (int, or None when absent).
            value = obj.element_.uiPriority;
        end
        function set.priority(obj, value)
            % Python: self._element.uiPriority = value.
            obj.element_.uiPriority = value;
        end

        % ---- quick_style (latent.py 172-184) ----
        function value = get.quick_style(obj)
            % Python: return self._element.on_off_prop("qFormat") (tri-state).
            value = obj.element_.on_off_prop("qFormat");
        end
        function set.quick_style(obj, value)
            % Python: self._element.set_on_off_prop("qFormat", value).
            obj.element_.set_on_off_prop("qFormat", value);
        end

        % ---- unhide_when_used (latent.py 186-198) ----
        function value = get.unhide_when_used(obj)
            % Python: return self._element.on_off_prop("unhideWhenUsed") (tri-state).
            value = obj.element_.on_off_prop("unhideWhenUsed");
        end
        function set.unhide_when_used(obj, value)
            % Python: self._element.set_on_off_prop("unhideWhenUsed", value).
            obj.element_.set_on_off_prop("unhideWhenUsed", value);
        end
    end
end

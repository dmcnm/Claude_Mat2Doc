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
%   DELETE_LATENT_STYLE (latent.py 119-128): named `delete_latent_style` (by the
%   kind of thing it removes) so NO method named `delete` exists on the styles
%   surface. MATLAB's `delete` on a handle IS the destructor; with no `delete`
%   override the destructor is never touched and no guarded-destructor machinery
%   is needed (H17 dissolved). Byte-identical to python-docx
%   `_LatentStyle.delete()`. Mirrors BaseStyle.delete_style. See the method below.
%
%   Example:
%       d  = mat2doc.Document();
%       ls = d.styles.latent_styles;
%       x  = ls.add_latent_style("Table Grid");
%       x.priority = 59;          % w:uiPriority
%       x.hidden = false;         % w:semiHidden="0"
%       x.quick_style = true;     % w:qFormat="1"
%       x.delete_latent_style();  % remove the <w:lsdException> from its parent
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

        % ---- delete_latent_style (latent.py 119-128; renamed from delete) ----
        function delete_latent_style(obj)
            % DELETE_LATENT_STYLE Remove this latent-style definition so the
            %   containing LatentStyles defaults provide the effective value again
            %   (latent.py 119-128).
            %
            %   Python (latent.py 127-128, `_LatentStyle.delete`):
            %       self._element.delete()   % detach the w:lsdException from w:latentStyles
            %       self._element = None
            %   Python notes that accessing any attribute afterward raises
            %   AttributeError (self._element is now None).
            %
            %   NAMING (H17 dissolved): named `delete_latent_style` -- by the kind of
            %   thing it removes -- so NO method named `delete` exists on the styles
            %   surface. MATLAB's `delete` on a `handle` IS the destructor; with no
            %   `delete` override the destructor is never touched and no guarded-
            %   destructor machinery is needed. Pure naming choice, byte-identical to
            %   python-docx `_LatentStyle.delete()` -- NOT an output deviation, NO
            %   D-number. The wrapped element method is likewise `delete_lsd_exception`
            %   (CT_LsdException), an ordinary getparent().remove() that is never GC's
            %   destructor, so the detached element survives (as in Python).
            %
            %   Gate-4 rule: after delete_latent_style the proxy's element_ is []
            %   (Python None), so accessing any proxy property errors -- matching
            %   python-docx (AttributeError). Assert only the parent-side effect
            %   (child count / serialized bytes); never inspect the proxy after it.
            %
            %   Ported from python-docx v1.2.0: styles/latent.py::_LatentStyle.delete
            obj.element_.delete_lsd_exception();   % Python: self._element.delete() (detach w:lsdException)
            obj.element_ = [];                     % Python: self._element = None
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

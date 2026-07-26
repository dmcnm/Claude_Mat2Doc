function write_only_property(varargin) %#ok<VANUS>
% WRITE_ONLY_PROPERTY Idiom guide for python-docx @write_only_property. NOT CALLABLE.
%
%   Python's @write_only_property decorator (shared.py 266-274) creates a
%   property that accepts assignment but NOT getattr:
%
%       def write_only_property(f):
%           return property(fset=f, doc=f.__doc__)
%
%   i.e. `property(fset=f)` with no fget. Reading such an attribute raises
%   AttributeError ("unreadable attribute"); assigning invokes the setter `f`.
%
%   MATLAB has no decorators, so this ports as an IDIOM, not a callable helper.
%   The MATLAB expression of a set-only property is a Dependent property with a
%   `set.x` method and NO `get.x` method:
%
%       properties (Dependent)
%           x
%       end
%       methods
%           function set.x(obj, value)
%               <Python fset body>
%           end
%           % NO get.x -> reading obj.x raises
%           %   MATLAB:class:noGetMethod, the write-only analogue of Python's
%           %   AttributeError on read.
%       end
%
%   NEW FOR MAT2DOC (audit_P2-1_proxy_tier.md): there is no python-pptx
%   counterpart (pptx has no write_only_property); this file establishes the
%   Mat2Doc idiom guide. In python-docx it decorates a small number of
%   setter-only members (e.g. a text-setting convenience); those members port
%   at their owning WP using the pattern above.
%
%   READ/WRITE ASYMMETRY (fidelity note): Python's `property(fset=f)` raises on
%   read with message "unreadable attribute"; MATLAB's missing-get.x path
%   raises MATLAB:class:noGetMethod. The identifiers/messages differ, but both
%   are read-rejections of a write-only property; python-docx never relies on
%   catching that specific read error (it is a programmer-error guard), so the
%   behavior is observationally equivalent for all upstream uses.
%
%   This function exists only so that `help mat2doc.shared.write_only_property`
%   documents the idiom; calling it is an error.
%
%   Ported from python-docx v1.2.0: src/docx/shared.py::write_only_property

error("mat2doc:write_only_property:notCallable", ...
    "write_only_property is a porting idiom, not a callable function. " + ...
    "See 'help mat2doc.shared.write_only_property' for the set-only Dependent-property pattern.");
end

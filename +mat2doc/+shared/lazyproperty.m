function lazyproperty(varargin) %#ok<VANUS>
% LAZYPROPERTY Idiom guide for python-docx @lazyproperty. NOT CALLABLE.
%
%   Python's @lazyproperty is a data-descriptor decorator: like @property but
%   the getter runs only on first access; the result is cached on the instance
%   and returned unchanged on every later access. It is read-only (assignment
%   raises AttributeError "can't set attribute" unconditionally, shared.py
%   244-263).
%
%   MATLAB has no decorators, so @lazyproperty ports as an IDIOM, not a
%   callable helper. Per design.md section 2 (binding): a Dependent read-only
%   property backed by a private cache property plus a LOGICAL computed-flag.
%   NEVER use isempty on the cache as the sentinel - empty ([]) is a legal
%   cached value (it is the None sentinel, H3).
%
%   NEW FOR MAT2DOC (audit_P2-1_proxy_tier.md): P1 established no lazyproperty
%   home; this file establishes the Mat2Doc idiom guide. docx's lazyproperty
%   lives in shared.py (python-pptx's lived in util.py), so its Mat2Doc home is
%   mat2doc.shared.lazyproperty. The cache+computed-flag pattern already
%   appears across the P1 code (e.g. +mat2doc\+opc\CoreProperties.m,
%   +mat2doc\+oxml\+coreprops\CT_CoreProperties.m); this documents it centrally.
%
%   Pattern - Python:
%
%       @lazyproperty
%       def fget(self):
%           return <compute>
%
%   ports to (inside a classdef X < handle):
%
%       properties (Dependent)
%           fget
%       end
%       properties (Access = private)
%           fget_cache_                          % cached value
%           fget_isComputed_ (1,1) logical = false
%       end
%       methods
%           function value = get.fget(obj)
%               if ~obj.fget_isComputed_
%                   obj.fget_cache_ = <compute>;   % Python fget body
%                   obj.fget_isComputed_ = true;
%               end
%               value = obj.fget_cache_;
%           end
%       end
%
%   Read-only behavior: define NO set.fget method - MATLAB then errors on
%   assignment to the Dependent property, mirroring Python's AttributeError.
%
%   Fidelity note: Python's implementation uses the cached value's None-ness as
%   its sentinel (shared.py 235-236), so a getter that returns None is
%   re-evaluated on every access. The computed-flag idiom evaluates exactly
%   once regardless. lazyproperty's documented contract requires getters to be
%   idempotent and side-effect free, so the two are observationally equivalent
%   for all upstream uses; the flag idiom is mandated by design.md because
%   isempty-as-sentinel is a recurring bug source (H3).
%
%   This function exists only so that `help mat2doc.shared.lazyproperty`
%   documents the idiom; calling it is an error.
%
%   Ported from python-docx v1.2.0: src/docx/shared.py::lazyproperty

error("mat2doc:lazyproperty:notCallable", ...
    "lazyproperty is a porting idiom, not a callable function. " + ...
    "See 'help mat2doc.shared.lazyproperty' for the cache + computed-flag pattern.");
end

classdef CaseInsensitiveDict < handle
% CASEINSENSITIVEDICT Mapping that matches keys without respect to case (H15).
%
%   Mapping type that behaves like dict except that it matches without respect
%   to the case of the key -- e.g. cid.get("A") == cid.get("a"). Not
%   general-purpose, just complete enough to satisfy opc package needs: it
%   assumes str keys, and that it is created EMPTY (keys passed in a constructor
%   are not accounted for). Mutable, so a handle class (Python subclasses dict).
%
%   LOWERCASING SCOPE (exact, H15): only the three overridden dunders fold the
%   key to lowercase -- set (__setitem__), get (__getitem__), isKey
%   (__contains__). The STORED key is the lowercased form, and keys() returns
%   those lowercased keys (that is what the P1-6a consumer's sorted(keys()) sees,
%   and the lowercased ext is what add_default emits as the Extension attribute).
%   Values are stored verbatim (never case-folded).
%
%   USED SURFACE (docx opc consumers, P1-6a): _defaults[ext] = ct (set),
%   _defaults[ext] (get), _defaults.keys() (keys) in
%   pkgwriter._ContentTypesItem. This class exposes exactly set / get / keys,
%   plus isKey for the __contains__ dunder. The Python type is created empty and
%   populated only via __setitem__; this port mirrors that (a no-arg
%   constructor). H11: keys() ordering is unspecified here because the sole
%   consumer sorts it; keys are drawn from the backing dictionary.
%
%   Example:
%       d = mat2doc.opc.CaseInsensitiveDict();
%       d.set("PNG", "image/png");
%       disp(d.isKey("png"))   % true   (key folded to lowercase, H15)
%       disp(d.get("Png"))     % "image/png"
%       disp(d.keys())         % "png"  (stored lowercased)
%
%   Ported from python-docx v1.2.0: src/docx/opc/shared.py::CaseInsensitiveDict
%   (lines 10-26)

    properties (Access = private)
        map_   % dictionary(string -> string), keyed by lowercased key
    end

    methods
        function obj = CaseInsensitiveDict()
            % Created empty (the Python docstring: "it is created empty; keys
            % passed in constructor are not accounted for").
            obj.map_ = dictionary(string.empty(0, 1), string.empty(0, 1));
        end

        function set(obj, key, value)
            % __setitem__ (shared.py 25-26): store under key.lower().
            obj.map_(lower(string(key))) = string(value);
        end

        function value = get(obj, key)
            % __getitem__ (shared.py 22-23): look up key.lower(). Missing key
            % raises (dictionary indexing error), matching dict's KeyError.
            value = obj.map_(lower(string(key)));
        end

        function tf = isKey(obj, key)
            % __contains__ (shared.py 19-20): membership of key.lower().
            tf = isKey(obj.map_, lower(string(key)));
        end

        function k = keys(obj)
            % KEYS Stored (lowercased) keys -- dict.keys(). 1xN string, or 1x0
            %   when empty. The P1-6a consumer wraps this in sorted().
            if numEntries(obj.map_) == 0
                k = strings(1, 0);
                return
            end
            k = reshape(keys(obj.map_), 1, []);
        end
    end
end

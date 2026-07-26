classdef LazyHost_p2_1 < handle
% LAZYHOST_P2_1  Host realising the design.md @lazyproperty idiom for the P2-1
%   proxy-tier tests.
%
%   Realises the idiom documented by mat2doc.shared.lazyproperty (which is a
%   non-callable guide, not a callable helper): a Dependent read-only property
%   backed by a private cache + a LOGICAL computed-flag (NEVER isempty as the
%   sentinel). A side-effect `compute_count` proves the getter body runs exactly
%   ONCE across repeated access, and the absence of a set.val method makes
%   assignment raise -- mirroring Python's lazyproperty first-access-caches +
%   read-only (AttributeError) contract (shared.py 155-263).
%
%   Not a matlab.unittest.TestCase, so runtests ignores it.
%
%   Provenance: verbatim port of the Gate-3 twin helper
%   validation\mat2doc\scenarios\s0013_LazyHost.m (MATLAB twin of the Python
%   _LazyHost in s0013_proxy_tier_probes.py), copied in so the suite is
%   self-contained.
    properties (Dependent)
        val
    end
    properties (Access = private)
        val_cache_
        val_isComputed_ (1,1) logical = false
    end
    properties
        compute_count (1,1) double = 0   % side-effect: getter-body run count
    end
    methods
        function v = get.val(obj)
            if ~obj.val_isComputed_
                obj.compute_count = obj.compute_count + 1;   % Python fget body
                obj.val_cache_ = 42;
                obj.val_isComputed_ = true;
            end
            v = obj.val_cache_;
        end
        % NO set.val -> assignment to the Dependent property raises (read-only),
        % mirroring Python lazyproperty.__set__ -> AttributeError.
    end
end

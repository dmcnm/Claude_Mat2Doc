classdef PartStub_p2_1 < handle
% PARTSTUB_P2_1  Minimal `part` provider for the P2-1 proxy-tier tests.
%
%   A structural ProvidesXmlPart: exposes only a zero-argument `part()`
%   accessor returning a sentinel string, so that ElementProxy / Parented /
%   StoryChild `.part` delegation is observable without dragging in the whole
%   package graph. Not a matlab.unittest.TestCase, so runtests ignores it
%   (same convention as tests\opc\StubPart_p1_6a.m / FakePart_p1_5.m).
%
%   Provenance: verbatim port of the Gate-3 twin helper
%   validation\mat2doc\scenarios\s0013_PartStub.m (itself the MATLAB twin of the
%   Python _PartStub in s0013_proxy_tier_probes.py), copied in so the Gate-4
%   suite is self-contained.
    methods
        function p = part(~)
            p = "STUBPART";
        end
    end
end

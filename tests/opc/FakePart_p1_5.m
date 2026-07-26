classdef FakePart_p1_5 < handle
% FAKEPART_P1_5  Minimal Part surface for the P1-5 rels tests (Gate-4).
%
%   Mirrors the Gate-3 scenario stub validation\mat2doc\scenarios\s0004_FakePart.m.
%   The real docx Part class is not ported until P1-6; the rel classes read ONLY
%   `.partname` off a target part (mat2doc.opc.Relationship_.target_ref /
%   target_part), so a handle object carrying a PackURI partname is a faithful
%   stand-in. HANDLE class is load-bearing: get_or_add's identity match ("==" on
%   two target parts) must behave like Python's `is`, so two FakePart_p1_5 with
%   the SAME partname but DIFFERENT identity yield DISTINCT relationships (H5).
%
%   Not a matlab.unittest.TestCase, so testsuite never collects it as a test.
    properties
        partname
    end
    methods
        function obj = FakePart_p1_5(pn)
            obj.partname = pn;
        end
    end
end

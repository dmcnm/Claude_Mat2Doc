classdef s0016_VPlain < mat2doc.enum.base.BaseEnum
% s0016_VPlain -- Gate-4 sample plain BaseEnum FIXTURE (no XML mapping). NOT
%   shipped in +mat2doc. Provenance: copied VERBATIM from
%   validation\mat2doc\scenarios\s0016_VPlain.m (Gate-3 validator sample). The
%   value-0 member (CONTINUOUS) exercises the non-truthy MS-API value path. Not
%   a TestCase -> testsuite does not collect it.
    methods
        function obj = s0016_VPlain(ms_api_value, docstr)
            obj@mat2doc.enum.base.BaseEnum(ms_api_value, docstr);
        end
    end
    enumeration
        NEW_PAGE   (2, "New page.")
        CONTINUOUS (0, "Continuous.")
    end
end

classdef StubPart_p1_6a < handle
% STUBPART_P1_6A  Minimal stand-in for a real Part (P1-6b scope) used by the
%   Gate-4 writer pins in Test_p1_6a_pkgrw.
%
%   Mirrors the Gate-3 twin stub validation\mat2doc\scenarios\S0005StubPart.m
%   (and the Python StubPart in s0005_writer_dfs_order.py). Carries exactly the
%   surface PackageWriter / ContentTypesItem_ consume: partname (PackURI),
%   content_type (string), blob (uint8), rels (Relationships).
%
%   NOT a matlab.unittest.TestCase, so `testsuite`/`runtests` never collects it
%   as a test (the FakePart_p1_5 helper-stub pattern).

    properties
        partname       % mat2doc.opc.PackURI
        content_type   % string
        blob           % uint8
        rels           % mat2doc.opc.Relationships
    end

    methods
        function obj = StubPart_p1_6a(partname, content_type, blob)
            obj.partname = partname;
            obj.content_type = content_type;
            obj.blob = blob;
            obj.rels = mat2doc.opc.Relationships(partname.baseURI);
        end
    end
end

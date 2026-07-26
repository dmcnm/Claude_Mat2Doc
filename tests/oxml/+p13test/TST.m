classdef TST
% TST  Gate-4 P1-3a test-only simple-type test double (NOT toolbox code).
%
%   Mimics the BaseSimpleType from_xml/to_xml interface so the attribute
%   descriptor ENGINE (getAttrTyped/setAttrTyped/getAttrRequired/setAttrRequired
%   in mat2doc.oxml.BaseOxmlElement) can be exercised in isolation WITHOUT
%   depending on P3 code: +mat2doc/+oxml/+simpletypes is P3 and does not exist
%   yet, so a controlled converter is supplied here (mirroring the Python
%   gen_p1_3a_oracle.TST and the Gate-3 harness scratch p13probe.TST) with
%   IDENTICAL behavior, isolating any divergence to the ENGINE.
%
%   Dispatched via resolveTypeCls_'s dotted-name currency: the type token
%   "p13test.TST" is used VERBATIM (it contains "."), so the engine calls
%   feval("p13test.TST.from_xml" / "p13test.TST.to_xml", ...). Lives under
%   tests\oxml\+p13test (resolved by the TestClassSetup PathFixture on the
%   tests\oxml folder); it is a test fixture, never part of the Mat2Doc toolbox
%   and never on the toolbox path.
%
%     from_xml(s)  -> "F:" + s        (proves the getter invokes from_xml)
%     to_xml(x)    -> [] (None) for the sentinel "__TOXML_NONE__" (exercises
%                     D-delta-2 / D-delta-3), else string(x)
%
%   Provenance: harness\mat2doc\+p13probe\TST.m (Gate-3 scratch), copied here
%   renamed +p13test so the permanent suite is self-contained and never reaches
%   into the harness tree.

    methods (Static)
        function v = from_xml(s)
            v = "F:" + string(s);
        end

        function s = to_xml(x)
            if (isstring(x) || ischar(x)) && string(x) == "__TOXML_NONE__"
                s = [];                 % Python TST.to_xml -> None
            else
                s = string(x);
            end
        end
    end
end

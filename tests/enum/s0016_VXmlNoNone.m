classdef s0016_VXmlNoNone < mat2doc.enum.base.BaseXmlEnum
% s0016_VXmlNoNone -- Gate-4 sample BaseXmlEnum FIXTURE with NO None-xml member
%   (exercises the raise-on-None from_xml path). NOT shipped in +mat2doc.
%   Provenance: copied VERBATIM from validation\mat2doc\scenarios\s0016_VXmlNoNone.m
%   (Gate-3 validator sample); the class name is byte-identical to the Python
%   twin so its ValueError messages match the frozen oracle. Not a TestCase ->
%   testsuite does not collect it.
    methods
        function obj = s0016_VXmlNoNone(ms_api_value, xml_value, docstr)
            obj@mat2doc.enum.base.BaseXmlEnum(ms_api_value, xml_value, docstr);
        end
    end
    methods (Static)
        function member = from_xml(xml_value)
            member = mat2doc.enum.base.BaseXmlEnum.from_xml_("s0016_VXmlNoNone", xml_value);
        end
        function s = to_xml(value)
            s = mat2doc.enum.base.BaseXmlEnum.to_xml_("s0016_VXmlNoNone", value);
        end
    end
    enumeration
        A (1, "a", "A.")
        B (2, "b", "B.")
    end
end

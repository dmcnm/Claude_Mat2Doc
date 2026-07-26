classdef s0016_VXml < mat2doc.enum.base.BaseXmlEnum
% s0016_VXml -- Gate-4 sample BaseXmlEnum FIXTURE (NOT shipped in +mat2doc).
%   Provenance: copied VERBATIM from the Gate-3 validator sample
%   validation\mat2doc\scenarios\s0016_VXml.m so this permanent test is
%   self-contained AND its ValueError messages (which interpolate the concrete
%   class name via cls.__name__ / shortName_) are byte-identical to the frozen
%   python-docx 1.2.0 oracle output in references\s0016\probe.json. The class
%   name is deliberately kept as "s0016_VXml" so every message compares to the
%   validator's frozen reference byte-for-byte.
%
%   Carries a None-xml member (INHERIT) and an empty-xml member (BLANK) plus two
%   ordinary members -- the H3 tri-state fixture. Being a Base(Xml)Enum subclass
%   (NOT a matlab.unittest.TestCase), testsuite does NOT collect it as a test.
    methods
        function obj = s0016_VXml(ms_api_value, xml_value, docstr)
            obj@mat2doc.enum.base.BaseXmlEnum(ms_api_value, xml_value, docstr);
        end
    end
    methods (Static)
        function member = from_xml(xml_value)
            member = mat2doc.enum.base.BaseXmlEnum.from_xml_("s0016_VXml", xml_value);
        end
        function s = to_xml(value)
            s = mat2doc.enum.base.BaseXmlEnum.to_xml_("s0016_VXml", value);
        end
    end
    enumeration
        CENTER  (1, "center", "Center.")
        BOTH    (3, "both", "Both.")
        INHERIT (5, string(missing), "None xml_value member (inherited).")
        BLANK   (7, "", "Empty-string xml_value member.")
    end
end

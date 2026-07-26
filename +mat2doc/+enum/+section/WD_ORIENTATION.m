classdef WD_ORIENTATION < mat2doc.enum.base.BaseXmlEnum
% WD_ORIENTATION Specifies the page layout orientation.
%
%   Alias: WD_ORIENT (mat2doc.enum.section.WD_ORIENT).
%
%   Example:
%       % section.orientation = mat2doc.enum.section.WD_ORIENT.LANDSCAPE;
%       double(mat2doc.enum.section.WD_ORIENT.LANDSCAPE.value)          % 1
%       mat2doc.enum.section.WD_ORIENTATION.to_xml( ...
%           mat2doc.enum.section.WD_ORIENT.LANDSCAPE)                   % "landscape"
%
%   MS API name: WdOrientation
%   http://msdn.microsoft.com/en-us/library/office/ff837902.aspx
%
%   Ported from python-docx v1.2.0: src/docx/enum/section.py::WD_ORIENTATION

    methods
        function obj = WD_ORIENTATION(ms_api_value, xml_value, docstr)
            obj@mat2doc.enum.base.BaseXmlEnum(ms_api_value, xml_value, docstr);
        end
    end

    methods (Static)
        function member = from_xml(xml_value)
            member = mat2doc.enum.base.BaseXmlEnum.from_xml_( ...
                "mat2doc.enum.section.WD_ORIENTATION", xml_value);
        end
        function s = to_xml(value)
            s = mat2doc.enum.base.BaseXmlEnum.to_xml_( ...
                "mat2doc.enum.section.WD_ORIENTATION", value);
        end
    end

    enumeration
        % Portrait orientation.
        PORTRAIT  (0, "portrait",  "Portrait orientation.")
        % Landscape orientation.
        LANDSCAPE (1, "landscape", "Landscape orientation.")
    end
end

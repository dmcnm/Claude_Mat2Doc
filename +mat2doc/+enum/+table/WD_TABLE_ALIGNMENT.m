classdef WD_TABLE_ALIGNMENT < mat2doc.enum.base.BaseXmlEnum
% WD_TABLE_ALIGNMENT Specifies table justification type.
%
%   Example:
%       % table.alignment = mat2doc.enum.table.WD_TABLE_ALIGNMENT.CENTER;
%       m = mat2doc.enum.table.WD_TABLE_ALIGNMENT.from_xml("center");   % CENTER
%       double(m.value)                                                 % 1
%       mat2doc.enum.table.WD_TABLE_ALIGNMENT.to_xml(m)                 % "center"
%
%   MS API name: WdRowAlignment
%   http://office.microsoft.com/en-us/word-help/HV080607259.aspx
%
%   Ported from python-docx v1.2.0: src/docx/enum/table.py::WD_TABLE_ALIGNMENT

    methods
        function obj = WD_TABLE_ALIGNMENT(ms_api_value, xml_value, docstr)
            obj@mat2doc.enum.base.BaseXmlEnum(ms_api_value, xml_value, docstr);
        end
    end

    methods (Static)
        function member = from_xml(xml_value)
            member = mat2doc.enum.base.BaseXmlEnum.from_xml_( ...
                "mat2doc.enum.table.WD_TABLE_ALIGNMENT", xml_value);
        end
        function s = to_xml(value)
            s = mat2doc.enum.base.BaseXmlEnum.to_xml_( ...
                "mat2doc.enum.table.WD_TABLE_ALIGNMENT", value);
        end
    end

    enumeration
        % Left-aligned
        LEFT   (0, "left",   "Left-aligned")
        % Center-aligned.
        CENTER (1, "center", "Center-aligned.")
        % Right-aligned.
        RIGHT  (2, "right",  "Right-aligned.")
    end
end

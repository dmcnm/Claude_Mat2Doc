classdef WD_STYLE_TYPE < mat2doc.enum.base.BaseXmlEnum
% WD_STYLE_TYPE Specifies one of the four style types.
%
%   Paragraph, character, list, or table.
%
%   Example:
%       % styles(1).type == mat2doc.enum.style.WD_STYLE_TYPE.PARAGRAPH
%       m = mat2doc.enum.style.WD_STYLE_TYPE.from_xml("character");   % CHARACTER
%       double(m.value)                                              % 2
%       mat2doc.enum.style.WD_STYLE_TYPE.to_xml(m)                   % "character"
%
%   MS API name: WdStyleType
%   http://msdn.microsoft.com/en-us/library/office/ff196870.aspx
%
%   Ported from python-docx v1.2.0: src/docx/enum/style.py::WD_STYLE_TYPE

    methods
        function obj = WD_STYLE_TYPE(ms_api_value, xml_value, docstr)
            obj@mat2doc.enum.base.BaseXmlEnum(ms_api_value, xml_value, docstr);
        end
    end

    methods (Static)
        function member = from_xml(xml_value)
            member = mat2doc.enum.base.BaseXmlEnum.from_xml_( ...
                "mat2doc.enum.style.WD_STYLE_TYPE", xml_value);
        end
        function s = to_xml(value)
            s = mat2doc.enum.base.BaseXmlEnum.to_xml_( ...
                "mat2doc.enum.style.WD_STYLE_TYPE", value);
        end
    end

    enumeration
        % Character style.
        CHARACTER (2, "character", "Character style.")
        % List style.
        LIST      (4, "numbering", "List style.")
        % Paragraph style.
        PARAGRAPH (1, "paragraph", "Paragraph style.")
        % Table style.
        TABLE     (3, "table",     "Table style.")
    end
end

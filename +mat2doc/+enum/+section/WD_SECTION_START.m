classdef WD_SECTION_START < mat2doc.enum.base.BaseXmlEnum
% WD_SECTION_START Specifies the start type of a section break.
%
%   Alias: WD_SECTION (mat2doc.enum.section.WD_SECTION).
%
%   Example:
%       % section.start_type = mat2doc.enum.section.WD_SECTION.NEW_PAGE;
%       double(mat2doc.enum.section.WD_SECTION.NEW_PAGE.value)        % 2
%       mat2doc.enum.section.WD_SECTION_START.from_xml("evenPage")    % EVEN_PAGE
%
%   MS API name: WdSectionStart
%   http://msdn.microsoft.com/en-us/library/office/ff840975.aspx
%
%   Ported from python-docx v1.2.0: src/docx/enum/section.py::WD_SECTION_START

    methods
        function obj = WD_SECTION_START(ms_api_value, xml_value, docstr)
            obj@mat2doc.enum.base.BaseXmlEnum(ms_api_value, xml_value, docstr);
        end
    end

    methods (Static)
        function member = from_xml(xml_value)
            member = mat2doc.enum.base.BaseXmlEnum.from_xml_( ...
                "mat2doc.enum.section.WD_SECTION_START", xml_value);
        end
        function s = to_xml(value)
            s = mat2doc.enum.base.BaseXmlEnum.to_xml_( ...
                "mat2doc.enum.section.WD_SECTION_START", value);
        end
    end

    enumeration
        % Continuous section break.
        CONTINUOUS (0, "continuous", "Continuous section break.")
        % New column section break.
        NEW_COLUMN (1, "nextColumn", "New column section break.")
        % New page section break.
        NEW_PAGE   (2, "nextPage",   "New page section break.")
        % Even pages section break.
        EVEN_PAGE  (3, "evenPage",   "Even pages section break.")
        % Section begins on next odd page.
        ODD_PAGE   (4, "oddPage",    "Section begins on next odd page.")
    end
end

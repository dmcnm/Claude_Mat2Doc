classdef WD_HEADER_FOOTER_INDEX < mat2doc.enum.base.BaseXmlEnum
% WD_HEADER_FOOTER_INDEX One of the three header/footer definitions for a section.
%
%   Alias: WD_HEADER_FOOTER (mat2doc.enum.section.WD_HEADER_FOOTER).
%
%   For internal use only; not part of the python-docx API.
%
%   Example:
%       m = mat2doc.enum.section.WD_HEADER_FOOTER_INDEX.from_xml("first"); % FIRST_PAGE
%       double(m.value)                                                   % 2
%
%   MS API name: WdHeaderFooterIndex
%   https://docs.microsoft.com/en-us/office/vba/api/word.wdheaderfooterindex
%
%   Ported from python-docx v1.2.0: src/docx/enum/section.py::WD_HEADER_FOOTER_INDEX

    methods
        function obj = WD_HEADER_FOOTER_INDEX(ms_api_value, xml_value, docstr)
            obj@mat2doc.enum.base.BaseXmlEnum(ms_api_value, xml_value, docstr);
        end
    end

    methods (Static)
        function member = from_xml(xml_value)
            member = mat2doc.enum.base.BaseXmlEnum.from_xml_( ...
                "mat2doc.enum.section.WD_HEADER_FOOTER_INDEX", xml_value);
        end
        function s = to_xml(value)
            s = mat2doc.enum.base.BaseXmlEnum.to_xml_( ...
                "mat2doc.enum.section.WD_HEADER_FOOTER_INDEX", value);
        end
    end

    enumeration
        % Header for odd pages or all if no even header.
        PRIMARY    (1, "default", "Header for odd pages or all if no even header.")
        % Header for first page of section.
        FIRST_PAGE (2, "first",   "Header for first page of section.")
        % Header for even pages of recto/verso section.
        EVEN_PAGE  (3, "even",    "Header for even pages of recto/verso section.")
    end
end

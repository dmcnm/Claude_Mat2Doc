classdef WD_CELL_VERTICAL_ALIGNMENT < mat2doc.enum.base.BaseXmlEnum
% WD_CELL_VERTICAL_ALIGNMENT Vertical alignment of text in table cells.
%
%   Alias: WD_ALIGN_VERTICAL (mat2doc.enum.table.WD_ALIGN_VERTICAL).
%
%   Example:
%       % table.cell(1,1).vertical_alignment = ...
%       %     mat2doc.enum.table.WD_ALIGN_VERTICAL.BOTTOM;
%       m = mat2doc.enum.table.WD_CELL_VERTICAL_ALIGNMENT.from_xml("center"); % CENTER
%       double(m.value)                                                       % 1
%       mat2doc.enum.table.WD_CELL_VERTICAL_ALIGNMENT.to_xml(m)               % "center"
%
%   MS API name: WdCellVerticalAlignment
%   https://msdn.microsoft.com/en-us/library/office/ff193345.aspx
%
%   Ported from python-docx v1.2.0: src/docx/enum/table.py::WD_CELL_VERTICAL_ALIGNMENT

    methods
        function obj = WD_CELL_VERTICAL_ALIGNMENT(ms_api_value, xml_value, docstr)
            obj@mat2doc.enum.base.BaseXmlEnum(ms_api_value, xml_value, docstr);
        end
    end

    methods (Static)
        function member = from_xml(xml_value)
            member = mat2doc.enum.base.BaseXmlEnum.from_xml_( ...
                "mat2doc.enum.table.WD_CELL_VERTICAL_ALIGNMENT", xml_value);
        end
        function s = to_xml(value)
            s = mat2doc.enum.base.BaseXmlEnum.to_xml_( ...
                "mat2doc.enum.table.WD_CELL_VERTICAL_ALIGNMENT", value);
        end
    end

    enumeration
        % Text is aligned to the top border of the cell.
        TOP    (0,   "top",    "Text is aligned to the top border of the cell.")
        % Text is aligned to the center of the cell.
        CENTER (1,   "center", "Text is aligned to the center of the cell.")
        % Text is aligned to the bottom border of the cell.
        BOTTOM (3,   "bottom", "Text is aligned to the bottom border of the cell.")
        % This is an option in the OpenXml spec, but not in Word itself.
        BOTH   (101, "both", ...
            "This is an option in the OpenXml spec, but not in Word itself. It's not" + ...
            " clear what Word behavior this setting produces. If you find out please" + ...
            " let us know and we'll update this documentation. Otherwise, probably best" + ...
            " to avoid this option.")
    end
end

classdef WD_ROW_HEIGHT_RULE < mat2doc.enum.base.BaseXmlEnum
% WD_ROW_HEIGHT_RULE Rule for determining the height of a table row.
%
%   Alias: WD_ROW_HEIGHT (mat2doc.enum.table.WD_ROW_HEIGHT).
%
%   Example:
%       % table.rows(1).height_rule = mat2doc.enum.table.WD_ROW_HEIGHT_RULE.EXACTLY;
%       m = mat2doc.enum.table.WD_ROW_HEIGHT_RULE.from_xml("atLeast");   % AT_LEAST
%       double(m.value)                                                  % 1
%       mat2doc.enum.table.WD_ROW_HEIGHT_RULE.to_xml(m)                  % "atLeast"
%
%   MS API name: WdRowHeightRule
%   https://msdn.microsoft.com/en-us/library/office/ff193620.aspx
%
%   Ported from python-docx v1.2.0: src/docx/enum/table.py::WD_ROW_HEIGHT_RULE

    methods
        function obj = WD_ROW_HEIGHT_RULE(ms_api_value, xml_value, docstr)
            obj@mat2doc.enum.base.BaseXmlEnum(ms_api_value, xml_value, docstr);
        end
    end

    methods (Static)
        function member = from_xml(xml_value)
            member = mat2doc.enum.base.BaseXmlEnum.from_xml_( ...
                "mat2doc.enum.table.WD_ROW_HEIGHT_RULE", xml_value);
        end
        function s = to_xml(value)
            s = mat2doc.enum.base.BaseXmlEnum.to_xml_( ...
                "mat2doc.enum.table.WD_ROW_HEIGHT_RULE", value);
        end
    end

    enumeration
        % The row height is adjusted to accommodate the tallest value in the row.
        AUTO     (0, "auto", ...
            "The row height is adjusted to accommodate the tallest value in the row.")
        % The row height is at least a minimum specified value.
        AT_LEAST (1, "atLeast", "The row height is at least a minimum specified value.")
        % The row height is an exact value.
        EXACTLY  (2, "exact",   "The row height is an exact value.")
    end
end

classdef WD_TABLE_DIRECTION < mat2doc.enum.base.BaseEnum
% WD_TABLE_DIRECTION Direction in which cells are ordered in a table or row.
%
%   Return/lookup enumeration (python-docx BaseEnum, NO XML mapping); it has no
%   from_xml/to_xml.
%
%   Example:
%       % table.direction = mat2doc.enum.table.WD_TABLE_DIRECTION.RTL;
%       t = mat2doc.enum.table.WD_TABLE_DIRECTION.RTL;
%       double(t.value)   % 1  (Python int(WD_TABLE_DIRECTION.RTL))
%       string(t)         % "RTL"  (member name)
%
%   MS API name: WdTableDirection
%   http://msdn.microsoft.com/en-us/library/ff835141.aspx
%
%   Ported from python-docx v1.2.0: src/docx/enum/table.py::WD_TABLE_DIRECTION

    methods
        function obj = WD_TABLE_DIRECTION(ms_api_value, docstr)
            obj@mat2doc.enum.base.BaseEnum(ms_api_value, docstr);
        end
    end

    enumeration
        % The table or row is arranged with the first column in the leftmost position.
        LTR (0, ...
            "The table or row is arranged with the first column in the leftmost position.")
        % The table or row is arranged with the first column in the rightmost position.
        RTL (1, ...
            "The table or row is arranged with the first column in the rightmost position.")
    end
end

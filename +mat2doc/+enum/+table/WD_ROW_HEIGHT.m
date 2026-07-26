classdef WD_ROW_HEIGHT
% WD_ROW_HEIGHT Alias of WD_ROW_HEIGHT_RULE.
%
%   Mirrors the python-docx module-level assignment
%   ``WD_ROW_HEIGHT = WD_ROW_HEIGHT_RULE`` (table.py line 82). MATLAB has no
%   class aliasing, so this class re-exports the canonical enumeration's members
%   as Constant properties and forwards the static methods. The members ARE
%   mat2doc.enum.table.WD_ROW_HEIGHT_RULE instances, so identity (==), isa, and
%   from_xml/to_xml behave exactly as if the two names referred to one
%   enumeration.
%
%   Example:
%       mat2doc.enum.table.WD_ROW_HEIGHT.EXACTLY == ...
%           mat2doc.enum.table.WD_ROW_HEIGHT_RULE.EXACTLY   % true
%
%   Ported from python-docx v1.2.0: src/docx/enum/table.py::WD_ROW_HEIGHT
%   (alias of WD_ROW_HEIGHT_RULE)

    properties (Constant)
        AUTO     = mat2doc.enum.table.WD_ROW_HEIGHT_RULE.AUTO
        AT_LEAST = mat2doc.enum.table.WD_ROW_HEIGHT_RULE.AT_LEAST
        EXACTLY  = mat2doc.enum.table.WD_ROW_HEIGHT_RULE.EXACTLY
    end

    methods (Static)
        function member = from_xml(xml_value)
            member = mat2doc.enum.table.WD_ROW_HEIGHT_RULE.from_xml(xml_value);
        end
        function s = to_xml(value)
            s = mat2doc.enum.table.WD_ROW_HEIGHT_RULE.to_xml(value);
        end
    end
end

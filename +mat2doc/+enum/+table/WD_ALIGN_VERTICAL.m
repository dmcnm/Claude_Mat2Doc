classdef WD_ALIGN_VERTICAL
% WD_ALIGN_VERTICAL Alias of WD_CELL_VERTICAL_ALIGNMENT.
%
%   Mirrors the python-docx module-level assignment
%   ``WD_ALIGN_VERTICAL = WD_CELL_VERTICAL_ALIGNMENT`` (table.py line 48). MATLAB
%   has no class aliasing, so this class re-exports the canonical enumeration's
%   members as Constant properties and forwards the static methods. The members
%   ARE mat2doc.enum.table.WD_CELL_VERTICAL_ALIGNMENT instances, so identity
%   (==), isa, and from_xml/to_xml behave exactly as if the two names referred to
%   one enumeration.
%
%   Example:
%       mat2doc.enum.table.WD_ALIGN_VERTICAL.BOTTOM == ...
%           mat2doc.enum.table.WD_CELL_VERTICAL_ALIGNMENT.BOTTOM   % true
%
%   Ported from python-docx v1.2.0: src/docx/enum/table.py::WD_ALIGN_VERTICAL
%   (alias of WD_CELL_VERTICAL_ALIGNMENT)

    properties (Constant)
        TOP    = mat2doc.enum.table.WD_CELL_VERTICAL_ALIGNMENT.TOP
        CENTER = mat2doc.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER
        BOTTOM = mat2doc.enum.table.WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
        BOTH   = mat2doc.enum.table.WD_CELL_VERTICAL_ALIGNMENT.BOTH
    end

    methods (Static)
        function member = from_xml(xml_value)
            member = mat2doc.enum.table.WD_CELL_VERTICAL_ALIGNMENT.from_xml(xml_value);
        end
        function s = to_xml(value)
            s = mat2doc.enum.table.WD_CELL_VERTICAL_ALIGNMENT.to_xml(value);
        end
    end
end

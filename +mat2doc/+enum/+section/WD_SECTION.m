classdef WD_SECTION
% WD_SECTION Alias of WD_SECTION_START.
%
%   Mirrors the python-docx module-level assignment
%   ``WD_SECTION = WD_SECTION_START`` (section.py line 86). MATLAB has no class
%   aliasing, so this class re-exports the canonical enumeration's members as
%   Constant properties and forwards the static methods. The members ARE
%   mat2doc.enum.section.WD_SECTION_START instances, so identity (==), isa, and
%   from_xml/to_xml behave exactly as if the two names referred to one
%   enumeration.
%
%   Example:
%       mat2doc.enum.section.WD_SECTION.NEW_PAGE == ...
%           mat2doc.enum.section.WD_SECTION_START.NEW_PAGE   % true
%
%   Ported from python-docx v1.2.0: src/docx/enum/section.py::WD_SECTION
%   (alias of WD_SECTION_START)

    properties (Constant)
        CONTINUOUS = mat2doc.enum.section.WD_SECTION_START.CONTINUOUS
        NEW_COLUMN = mat2doc.enum.section.WD_SECTION_START.NEW_COLUMN
        NEW_PAGE   = mat2doc.enum.section.WD_SECTION_START.NEW_PAGE
        EVEN_PAGE  = mat2doc.enum.section.WD_SECTION_START.EVEN_PAGE
        ODD_PAGE   = mat2doc.enum.section.WD_SECTION_START.ODD_PAGE
    end

    methods (Static)
        function member = from_xml(xml_value)
            member = mat2doc.enum.section.WD_SECTION_START.from_xml(xml_value);
        end
        function s = to_xml(value)
            s = mat2doc.enum.section.WD_SECTION_START.to_xml(value);
        end
    end
end

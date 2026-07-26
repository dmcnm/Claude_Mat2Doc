classdef WD_HEADER_FOOTER
% WD_HEADER_FOOTER Alias of WD_HEADER_FOOTER_INDEX.
%
%   Mirrors the python-docx module-level assignment
%   ``WD_HEADER_FOOTER = WD_HEADER_FOOTER_INDEX`` (section.py line 27). MATLAB
%   has no class aliasing, so this class re-exports the canonical enumeration's
%   members as Constant properties and forwards the static methods. The members
%   ARE mat2doc.enum.section.WD_HEADER_FOOTER_INDEX instances, so identity (==),
%   isa, and from_xml/to_xml behave exactly as if the two names referred to one
%   enumeration.
%
%   Example:
%       mat2doc.enum.section.WD_HEADER_FOOTER.PRIMARY == ...
%           mat2doc.enum.section.WD_HEADER_FOOTER_INDEX.PRIMARY   % true
%
%   Ported from python-docx v1.2.0: src/docx/enum/section.py::WD_HEADER_FOOTER
%   (alias of WD_HEADER_FOOTER_INDEX)

    properties (Constant)
        PRIMARY    = mat2doc.enum.section.WD_HEADER_FOOTER_INDEX.PRIMARY
        FIRST_PAGE = mat2doc.enum.section.WD_HEADER_FOOTER_INDEX.FIRST_PAGE
        EVEN_PAGE  = mat2doc.enum.section.WD_HEADER_FOOTER_INDEX.EVEN_PAGE
    end

    methods (Static)
        function member = from_xml(xml_value)
            member = mat2doc.enum.section.WD_HEADER_FOOTER_INDEX.from_xml(xml_value);
        end
        function s = to_xml(value)
            s = mat2doc.enum.section.WD_HEADER_FOOTER_INDEX.to_xml(value);
        end
    end
end

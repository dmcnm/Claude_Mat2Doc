classdef WD_ORIENT
% WD_ORIENT Alias of WD_ORIENTATION.
%
%   Mirrors the python-docx module-level assignment
%   ``WD_ORIENT = WD_ORIENTATION`` (section.py line 52). MATLAB has no class
%   aliasing, so this class re-exports the canonical enumeration's members as
%   Constant properties and forwards the static methods. The members ARE
%   mat2doc.enum.section.WD_ORIENTATION instances, so identity (==), isa, and
%   from_xml/to_xml behave exactly as if the two names referred to one
%   enumeration.
%
%   Example:
%       mat2doc.enum.section.WD_ORIENT.PORTRAIT == ...
%           mat2doc.enum.section.WD_ORIENTATION.PORTRAIT   % true
%
%   Ported from python-docx v1.2.0: src/docx/enum/section.py::WD_ORIENT
%   (alias of WD_ORIENTATION)

    properties (Constant)
        PORTRAIT  = mat2doc.enum.section.WD_ORIENTATION.PORTRAIT
        LANDSCAPE = mat2doc.enum.section.WD_ORIENTATION.LANDSCAPE
    end

    methods (Static)
        function member = from_xml(xml_value)
            member = mat2doc.enum.section.WD_ORIENTATION.from_xml(xml_value);
        end
        function s = to_xml(value)
            s = mat2doc.enum.section.WD_ORIENTATION.to_xml(value);
        end
    end
end

classdef WD_ALIGN_PARAGRAPH
% WD_ALIGN_PARAGRAPH Alias of WD_PARAGRAPH_ALIGNMENT.
%
%   Mirrors the python-docx module-level assignment
%   ``WD_ALIGN_PARAGRAPH = WD_PARAGRAPH_ALIGNMENT`` (text.py line 67).
%   WD_ALIGN_PARAGRAPH is the spelling the python-docx API examples use, so
%   consuming code references mat2doc.enum.text.WD_ALIGN_PARAGRAPH. MATLAB has no
%   class aliasing, so this class re-exports the canonical enumeration's members
%   as Constant properties and forwards the static methods. The members ARE
%   mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT instances, so identity (==), isa,
%   from_xml/to_xml, and mixed-name comparisons behave exactly as if the two
%   names referred to one enumeration.
%
%   Example:
%       mat2doc.enum.text.WD_ALIGN_PARAGRAPH.CENTER == ...
%           mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER   % true
%
%   Ported from python-docx v1.2.0: src/docx/enum/text.py::WD_ALIGN_PARAGRAPH
%   (alias of WD_PARAGRAPH_ALIGNMENT)

    properties (Constant)
        LEFT         = mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.LEFT
        CENTER       = mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER
        RIGHT        = mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.RIGHT
        JUSTIFY      = mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        DISTRIBUTE   = mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.DISTRIBUTE
        JUSTIFY_MED  = mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.JUSTIFY_MED
        JUSTIFY_HI   = mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.JUSTIFY_HI
        JUSTIFY_LOW  = mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.JUSTIFY_LOW
        THAI_JUSTIFY = mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.THAI_JUSTIFY
    end

    methods (Static)
        function member = from_xml(xml_value)
            member = mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.from_xml(xml_value);
        end
        function s = to_xml(value)
            s = mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.to_xml(value);
        end
    end
end

classdef CT_Highlight < mat2doc.oxml.BaseOxmlElement
% CT_HIGHLIGHT `w:highlight` element: font highlighting / background color.
%
%   `val` RequiredAttribute("w:val", WD_COLOR_INDEX) -> a WD_COLOR_INDEX member;
%   InvalidXmlError if @w:val is absent. WD_COLOR_INDEX is an enum simple-type,
%   referenced by its FULLY QUALIFIED name so BaseOxmlElement.resolveTypeCls_
%   dispatches to +enum verbatim.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on <w:highlight> nodes inside a real part.
%
%   Example:
%       h = mat2doc.oxml.OxmlElement("w:highlight");
%       h.val = mat2doc.enum.text.WD_COLOR_INDEX.YELLOW;  % <w:highlight w:val="yellow"/>
%
%   Ported from python-docx v1.2.0: src/docx/oxml/text/font.py::CT_Highlight
%   (lines 49-52; registered for w:highlight)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        VAL_ATTR = "w:val"                            % RequiredAttribute @ font.py:52
        VAL_TYPE = "mat2doc.enum.text.WD_COLOR_INDEX" % enum simple-type (verbatim, resolveTypeCls_)
    end

    properties (Dependent)  % generated descriptor property
        val   % RequiredAttribute('w:val', WD_COLOR_INDEX) -> member; InvalidXmlError if absent
    end

    methods
        function obj = CT_Highlight(varargin)
            % CT_HIGHLIGHT Construct a loose <w:highlight> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        function value = get.val(obj)
            value = obj.getAttrRequired(obj.VAL_ATTR, obj.VAL_TYPE);
        end
        function set.val(obj, value)
            obj.setAttrRequired(obj.VAL_ATTR, obj.VAL_TYPE, value);
        end
    end
end

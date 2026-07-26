classdef CT_VerticalAlignRun < mat2doc.oxml.BaseOxmlElement
% CT_VERTICALALIGNRUN `<w:vertAlign>` element: subscript or superscript.
%
%   `val` RequiredAttribute("w:val", ST_VerticalAlignRun) -> a string
%   ("baseline" / "superscript" / "subscript"); InvalidXmlError if @w:val is
%   absent. ST_VerticalAlignRun is a string-enumeration simple type (returns the
%   XML token itself, validated on write).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on <w:vertAlign> nodes inside a real part.
%
%   Example:
%       va = mat2doc.oxml.OxmlElement("w:vertAlign");
%       va.val = "superscript";     % <w:vertAlign w:val="superscript"/>
%
%   Ported from python-docx v1.2.0: src/docx/oxml/text/font.py::CT_VerticalAlignRun
%   (lines 328-331; registered for w:vertAlign)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        VAL_ATTR = "w:val"                % RequiredAttribute @ font.py:331
        VAL_TYPE = "ST_VerticalAlignRun"  % simple type (+oxml\+simpletypes)
    end

    properties (Dependent)  % generated descriptor property
        val   % RequiredAttribute('w:val', ST_VerticalAlignRun) -> string; InvalidXmlError if absent
    end

    methods
        function obj = CT_VerticalAlignRun(varargin)
            % CT_VERTICALALIGNRUN Construct a loose <w:vertAlign> -- TRANSPARENT PASS-THROUGH.
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

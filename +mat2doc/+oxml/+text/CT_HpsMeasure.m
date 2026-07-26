classdef CT_HpsMeasure < mat2doc.oxml.BaseOxmlElement
% CT_HPSMEASURE Used for `<w:sz>` (and `<w:szCs>`): font size in half-points.
%
%   `val` RequiredAttribute("w:val", ST_HpsMeasure) -> a Length (a bare
%   half-point count becomes Pt(int(value)/2)); InvalidXmlError if @w:val is
%   absent. ST_HpsMeasure.to_xml serializes a Length back to a half-point count
%   string.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on <w:sz>/<w:szCs> nodes inside a real part.
%
%   Example:
%       sz = mat2doc.oxml.OxmlElement("w:sz");
%       sz.val = mat2doc.shared.Pt(12);   % <w:sz w:val="24"/>  (24 half-points)
%
%   Ported from python-docx v1.2.0: src/docx/oxml/text/font.py::CT_HpsMeasure
%   (lines 55-58; registered for w:sz and w:szCs)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        VAL_ATTR = "w:val"          % RequiredAttribute @ font.py:58
        VAL_TYPE = "ST_HpsMeasure"  % simple type (+oxml\+simpletypes)
    end

    properties (Dependent)  % generated descriptor property
        val   % RequiredAttribute('w:val', ST_HpsMeasure) -> Length; InvalidXmlError if absent
    end

    methods
        function obj = CT_HpsMeasure(varargin)
            % CT_HPSMEASURE Construct a loose <w:sz> -- TRANSPARENT PASS-THROUGH.
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

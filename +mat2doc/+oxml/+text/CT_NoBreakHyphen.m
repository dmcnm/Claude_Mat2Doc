classdef CT_NoBreakHyphen < mat2doc.oxml.BaseOxmlElement
% CT_NOBREAKHYPHEN `<w:noBreakHyphen>` element, a hyphen ineligible for a
%   line-wrap position. Maps to a plain-text dash ("-"). Registered for
%   <w:noBreakHyphen> (docx/oxml/__init__.py:75).
%
%   NOTE (run.py 214-217): the complex-type name CT_NoBreakHyphen does NOT exist
%   in the schema, where w:noBreakHyphen maps to CT_Empty. The distinguished
%   name exists only to give w:noBreakHyphen its "-" __str__ behavior.
%
%   __str__ (run.py 219-221) -> str_(): a single dash character ("-").
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1).
%
%   Example:
%       nbh = mat2doc.oxml.OxmlElement("w:noBreakHyphen");
%       nbh.str_()     % "-"
%
%   Ported from python-docx v1.2.0: src/docx/oxml/text/run.py::CT_NoBreakHyphen
%   (lines 209-221; registered for w:noBreakHyphen)

    methods
        function obj = CT_NoBreakHyphen(varargin)
            % CT_NOBREAKHYPHEN Construct a loose <w:noBreakHyphen> -- PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        function value = str_(obj) %#ok<MANU>
            % STR_ Text equivalent (run.py 219-221): a single dash ("-").
            value = "-";
        end
    end
end

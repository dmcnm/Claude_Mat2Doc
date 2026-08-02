classdef CT_RelativeRect < mat2doc.oxml.BaseOxmlElement
% CT_RELATIVERECT Custom element class for the <a:fillRect> element.
%
%   `<a:fillRect>` element, specifying picture should fill containing rectangle
%   shape. The class body is EMPTY in python-docx (a bare marker class). No
%   descriptors, attributes, or methods are declared.
%
%   DOCX vs pptx: python-docx's CT_RelativeRect is EMPTY; python-pptx's
%   CT_RelativeRect (oxml/dml/fill.py) carries l/t/r/b ST_Percentage attributes.
%   Ported to match DOCX (empty). NOT registered in oxml/__init__.py -- a parsed
%   <a:fillRect> stays a generic XmlElement.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/shape.py::CT_RelativeRect
%   (lines 217-219; not registered.)

    methods
        function obj = CT_RelativeRect(varargin)
            % CT_RELATIVERECT Construct a loose <a:fillRect> -- PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end
    end
end

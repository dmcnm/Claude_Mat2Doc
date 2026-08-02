classdef CT_StretchInfoProperties < mat2doc.oxml.BaseOxmlElement
% CT_STRETCHINFOPROPERTIES Custom element class for the <a:stretch> element.
%
%   `<a:stretch>` element, specifies how picture should fill its containing
%   shape. The class body is EMPTY in python-docx (a bare marker class). No
%   descriptors, attributes, or methods are declared.
%
%   NOT registered in oxml/__init__.py -- a parsed <a:stretch> stays a generic
%   XmlElement.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/shape.py::CT_StretchInfoProperties
%   (lines 266-268; not registered.)

    methods
        function obj = CT_StretchInfoProperties(varargin)
            % CT_STRETCHINFOPROPERTIES Construct a loose <a:stretch> -- PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end
    end
end

classdef CT_PresetGeometry2D < mat2doc.oxml.BaseOxmlElement
% CT_PRESETGEOMETRY2D Custom element class for the <a:prstGeom> element.
%
%   `<a:prstGeom>` element, specifies a preset autoshape geometry, such as
%   ``rect``. The class body is EMPTY in python-docx (a bare marker class). No
%   descriptors, attributes, or methods are declared.
%
%   DOCX vs pptx: python-docx's CT_PresetGeometry2D is EMPTY; python-pptx's
%   CT_PresetGeometry2D (oxml/shapes/autoshape.py) is a rich class (prst
%   attribute, avLst / gd children). Ported to match DOCX (empty). NOT registered
%   in oxml/__init__.py -- a parsed <a:prstGeom> stays a generic XmlElement.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/shape.py::CT_PresetGeometry2D
%   (lines 212-214; not registered.)

    methods
        function obj = CT_PresetGeometry2D(varargin)
            % CT_PRESETGEOMETRY2D Construct a loose <a:prstGeom> -- PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end
    end
end

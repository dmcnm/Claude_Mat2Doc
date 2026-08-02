classdef CT_GraphicalObject < mat2doc.oxml.BaseOxmlElement
% CT_GRAPHICALOBJECT Custom element class for the <a:graphic> element.
%
%   `<a:graphic>` element, container for a DrawingML object. Its single required
%   child is a:graphicData (CT_GraphicalObjectData), a OneAndOnlyOne descriptor.
%
%   DOCX vs pptx: python-docx's CT_GraphicalObject has ONLY the `graphicData`
%   required child (no `chart` @property). python-pptx's same-named class
%   (oxml/shapes/graphfrm.py) adds a `chart` grandchild @property. Ported to
%   match DOCX (graphicData only).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/shape.py::CT_GraphicalObject
%   (lines 53-58; registered for <a:graphic>, oxml/__init__.py:49)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        GRAPHICDATA_TAG = "a:graphicData"   % OneAndOnlyOne @ shape.py:56
    end

    properties (Dependent)  % generated descriptor properties
        graphicData  % OneAndOnlyOne('a:graphicData') -- required child
    end

    methods
        function obj = CT_GraphicalObject(varargin)
            % CT_GRAPHICALOBJECT Construct a loose <a:graphic> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- graphicData (OneAndOnlyOne) ----
        function child = get.graphicData(obj)
            child = obj.getRequiredChild(obj.GRAPHICDATA_TAG);
        end
    end
end

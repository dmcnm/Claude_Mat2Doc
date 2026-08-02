classdef CT_PictureNonVisual < mat2doc.oxml.BaseOxmlElement
% CT_PICTURENONVISUAL Custom element class for the <pic:nvPicPr> element.
%
%   `<pic:nvPicPr>` element, non-visual picture properties. A single
%   OneAndOnlyOne descriptor `cNvPr` ("pic:cNvPr").
%
%   DOCX vs pptx: python-docx's CT_PictureNonVisual has ONLY the `cNvPr`
%   required child. python-pptx's same-named class (oxml/shapes/picture.py) has
%   `cNvPr` AND an `nvPr` required child. Ported to match DOCX (cNvPr only --
%   the docx pic:nvPicPr has no pic:nvPr).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/shape.py::CT_PictureNonVisual
%   (lines 182-185; registered for <pic:nvPicPr>, oxml/__init__.py:55)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        CNVPR_TAG = "pic:cNvPr"   % OneAndOnlyOne @ shape.py:185
    end

    properties (Dependent)  % generated descriptor properties
        cNvPr  % OneAndOnlyOne('pic:cNvPr') -- required child (CT_NonVisualDrawingProps)
    end

    methods
        function obj = CT_PictureNonVisual(varargin)
            % CT_PICTURENONVISUAL Construct a loose <pic:nvPicPr> -- PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- cNvPr (OneAndOnlyOne) ----
        function child = get.cNvPr(obj)
            child = obj.getRequiredChild(obj.CNVPR_TAG);
        end
    end
end

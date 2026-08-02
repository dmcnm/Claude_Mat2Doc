classdef CT_NonVisualPictureProperties < mat2doc.oxml.BaseOxmlElement
% CT_NONVISUALPICTUREPROPERTIES Custom element class for the <pic:cNvPicPr> element.
%
%   `<pic:cNvPicPr>` element, specifies picture locking and resize behaviors.
%   The class body is EMPTY in python-docx (a bare marker class). No
%   descriptors, attributes, or methods are declared.
%
%   DOCX vs pptx: python-docx registers <pic:cNvPicPr> -> CT_NonVisualPictureProperties
%   (this EMPTY class). python-pptx has NO such class -- its picture uses the
%   <p:...> prefix family and does not model p:cNvPicPr with a dedicated class.
%   NOVEL to the docx port.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/shape.py::CT_NonVisualPictureProperties
%   (lines 131-132; not directly registered -- the <pic:cNvPicPr> node inside a
%   pic:nvPicPr transits it only structurally; no register_element_cls row.)

    methods
        function obj = CT_NonVisualPictureProperties(varargin)
            % CT_NONVISUALPICTUREPROPERTIES Construct a loose <pic:cNvPicPr> -- PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end
    end
end

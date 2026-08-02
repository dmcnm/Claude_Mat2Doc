classdef CT_Anchor < mat2doc.oxml.BaseOxmlElement
% CT_ANCHOR Custom element class for the <wp:anchor> element.
%
%   `<wp:anchor>` element, container for a "floating" shape. The class body is
%   EMPTY in python-docx (a bare marker class) -- registering the tag makes a
%   parsed <wp:anchor> node a CT_Anchor rather than a generic XmlElement. No
%   descriptors, attributes, or methods are declared.
%
%   NOVEL (docx wp:-namespace inline/floating shape structure).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/shape.py::CT_Anchor
%   (lines 29-30; registered for <wp:anchor>, oxml/__init__.py:59)

    methods
        function obj = CT_Anchor(varargin)
            % CT_ANCHOR Construct a loose <wp:anchor> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end
    end
end

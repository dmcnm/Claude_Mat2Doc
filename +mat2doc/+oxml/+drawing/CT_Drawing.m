classdef CT_Drawing < mat2doc.oxml.BaseOxmlElement
% CT_DRAWING Custom element class for the <w:drawing> element.
%
%   `<w:drawing>` element, containing a DrawingML object like a picture or
%   chart. The class body is EMPTY in python-docx (a bare
%   `class CT_Drawing(BaseOxmlElement)` marker) -- registering the tag simply
%   makes a parsed <w:drawing> node a CT_Drawing rather than a generic
%   XmlElement. No descriptors, attributes, or methods are declared.
%
%   NOVEL (docx-only): python-pptx has no CT_Drawing. Lives in
%   docx/oxml/drawing.py (not shape.py) for the legacy reasons noted in that
%   module's docstring.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this as feval(cls, name, ownDecls[, resolvedUri]) when it hits a
%   registered <w:drawing> in a real document.xml, so the constructor must
%   forward ALL positional args to the base with NO nsmap re-validation.
%
%   Ported from python-docx v1.2.0: src/docx/oxml/drawing.py::CT_Drawing
%   (lines 10-11; registered for <w:drawing>, oxml/__init__.py:58)

    methods
        function obj = CT_Drawing(varargin)
            % CT_DRAWING Construct a loose <w:drawing> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end
    end
end

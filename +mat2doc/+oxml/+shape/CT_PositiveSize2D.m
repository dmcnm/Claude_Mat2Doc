classdef CT_PositiveSize2D < mat2doc.oxml.BaseOxmlElement
% CT_POSITIVESIZE2D Custom element class for the <wp:extent> element (and <a:ext>).
%
%   Used for `<wp:extent>` element, and perhaps others later. Specifies the size
%   of a DrawingML drawing. Two RequiredAttribute descriptors cx, cy typed
%   ST_PositiveCoordinate (non-negative EMU extents, H6). No child elements.
%
%   DOCX/pptx PARITY: this class body is byte-identical between python-docx
%   (shape.py:198-209) and python-pptx (CT_PositiveSize2D) -- both declare cx/cy
%   RequiredAttribute ST_PositiveCoordinate. Registered for BOTH <a:ext> and
%   <wp:extent> in docx (oxml/__init__.py:48, 61). Re-ported faithfully into the
%   docx home package (design.md section 7).
%
%   REQUIREDATTRIBUTE (design.md section 2): GET before set raises
%   mat2doc:InvalidXmlError; SET always writes, never removes.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/shape.py::CT_PositiveSize2D
%   (lines 198-209; registered for <a:ext> and <wp:extent>,
%   oxml/__init__.py:48, 61)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        CX_ATTR = "cx"                       % RequiredAttribute @ shape.py:204
        CX_TYPE = "ST_PositiveCoordinate"    % simple type (+oxml\+simpletypes)
        CY_ATTR = "cy"                       % RequiredAttribute @ shape.py:207
        CY_TYPE = "ST_PositiveCoordinate"    % simple type (+oxml\+simpletypes)
    end

    properties (Dependent)  % generated descriptor properties
        cx  % RequiredAttribute('cx', ST_PositiveCoordinate) -> Length; InvalidXmlError if absent
        cy  % RequiredAttribute('cy', ST_PositiveCoordinate) -> Length; InvalidXmlError if absent
    end

    methods
        function obj = CT_PositiveSize2D(varargin)
            % CT_POSITIVESIZE2D Construct a loose <wp:extent>/<a:ext> -- PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        function value = get.cx(obj)
            value = obj.getAttrRequired(obj.CX_ATTR, obj.CX_TYPE);
        end
        function set.cx(obj, value)
            obj.setAttrRequired(obj.CX_ATTR, obj.CX_TYPE, value);
        end

        function value = get.cy(obj)
            value = obj.getAttrRequired(obj.CY_ATTR, obj.CY_TYPE);
        end
        function set.cy(obj, value)
            obj.setAttrRequired(obj.CY_ATTR, obj.CY_TYPE, value);
        end
    end
end

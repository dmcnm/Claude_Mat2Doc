classdef CT_Point2D < mat2doc.oxml.BaseOxmlElement
% CT_POINT2D Custom element class for the <a:off> element.
%
%   Used for `<a:off>` element, and perhaps others. Specifies an x, y coordinate
%   (point). Two RequiredAttribute descriptors x, y typed ST_Coordinate (EMU
%   coordinates, H6). No child elements.
%
%   DOCX/pptx PARITY: this class body is byte-identical between python-docx
%   (shape.py:188-195) and python-pptx (oxml/shapes/shared.py::CT_Point2D) --
%   both declare x/y RequiredAttribute ST_Coordinate. Re-ported faithfully into
%   the docx home package (design.md section 7: no shared code across toolboxes).
%
%   REQUIREDATTRIBUTE (design.md section 2): GET before set raises
%   mat2doc:InvalidXmlError; SET always writes, never removes.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/shape.py::CT_Point2D
%   (lines 188-195; registered for <a:off>, oxml/__init__.py:51)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        X_ATTR = "x"                 % RequiredAttribute @ shape.py:194
        X_TYPE = "ST_Coordinate"     % simple type (+oxml\+simpletypes)
        Y_ATTR = "y"                 % RequiredAttribute @ shape.py:195
        Y_TYPE = "ST_Coordinate"     % simple type (+oxml\+simpletypes)
    end

    properties (Dependent)  % generated descriptor properties
        x  % RequiredAttribute('x', ST_Coordinate) -> Length; InvalidXmlError if absent
        y  % RequiredAttribute('y', ST_Coordinate) -> Length; InvalidXmlError if absent
    end

    methods
        function obj = CT_Point2D(varargin)
            % CT_POINT2D Construct a loose <a:off> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        function value = get.x(obj)
            value = obj.getAttrRequired(obj.X_ATTR, obj.X_TYPE);
        end
        function set.x(obj, value)
            obj.setAttrRequired(obj.X_ATTR, obj.X_TYPE, value);
        end

        function value = get.y(obj)
            value = obj.getAttrRequired(obj.Y_ATTR, obj.Y_TYPE);
        end
        function set.y(obj, value)
            obj.setAttrRequired(obj.Y_ATTR, obj.Y_TYPE, value);
        end
    end
end

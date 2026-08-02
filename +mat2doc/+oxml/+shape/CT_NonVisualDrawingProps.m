classdef CT_NonVisualDrawingProps < mat2doc.oxml.BaseOxmlElement
% CT_NONVISUALDRAWINGPROPS Custom element class for the <wp:docPr> element (and others).
%
%   Used for `<wp:docPr>` element, and perhaps others (e.g. <pic:cNvPr>).
%   Specifies the id and name of a DrawingML drawing. Two RequiredAttribute
%   descriptors:
%     id   = RequiredAttribute("id",   ST_DrawingElementId)
%     name = RequiredAttribute("name", XsdString)
%   ported as a Constant schema table plus one-line get./set. members delegating
%   to the BaseOxmlElement typed-attribute engine.
%
%   DOCX vs pptx: python-docx's CT_NonVisualDrawingProps is this MINIMAL id/name
%   class registered for BOTH <pic:cNvPr> and <wp:docPr> (oxml/__init__.py:54,60).
%   python-pptx's same-named class (oxml/shapes/shared.py) is a much larger
%   BaseShapeElement-derived class (hlinkClick/hlinkHover children, descr/hidden/
%   title attributes, etc.). Ported to match DOCX (id + name only).
%
%   REQUIREDATTRIBUTE (design.md section 2): GET before the attribute is set
%   raises mat2doc:InvalidXmlError; SET always writes and never removes (a
%   required attribute has no default). `id` is int-valued (ST_DrawingElementId);
%   `name` is a plain string (XsdString).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this when it hits a registered <wp:docPr> or <pic:cNvPr>.
%
%   Ported from python-docx v1.2.0: src/docx/oxml/shape.py::CT_NonVisualDrawingProps
%   (lines 121-128; registered for <pic:cNvPr> and <wp:docPr>,
%   oxml/__init__.py:54, 60)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        ID_ATTR = "id"                    % RequiredAttribute @ shape.py:127
        ID_TYPE = "ST_DrawingElementId"   % simple type (+oxml\+simpletypes)
        NAME_ATTR = "name"                % RequiredAttribute @ shape.py:128
        NAME_TYPE = "XsdString"           % simple type (+oxml\+simpletypes)
    end

    properties (Dependent)  % generated descriptor properties
        id    % RequiredAttribute('id', ST_DrawingElementId) -> int; InvalidXmlError if absent
        name  % RequiredAttribute('name', XsdString) -> string; InvalidXmlError if absent
    end

    methods
        function obj = CT_NonVisualDrawingProps(varargin)
            % CT_NONVISUALDRAWINGPROPS Construct a loose element -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        function value = get.id(obj)
            value = obj.getAttrRequired(obj.ID_ATTR, obj.ID_TYPE);
        end
        function set.id(obj, value)
            obj.setAttrRequired(obj.ID_ATTR, obj.ID_TYPE, value);
        end

        function value = get.name(obj)
            value = obj.getAttrRequired(obj.NAME_ATTR, obj.NAME_TYPE);
        end
        function set.name(obj, value)
            obj.setAttrRequired(obj.NAME_ATTR, obj.NAME_TYPE, value);
        end
    end
end

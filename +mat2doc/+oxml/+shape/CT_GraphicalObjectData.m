classdef CT_GraphicalObjectData < mat2doc.oxml.BaseOxmlElement
% CT_GRAPHICALOBJECTDATA Custom element class for the <a:graphicData> element.
%
%   `<a:graphicData>` element, container for the XML of a DrawingML object. In
%   python-docx it models an inline PICTURE's graphic-data: a ZeroOrOne
%   `pic` ("pic:pic") child and a RequiredAttribute `uri` ("uri", XsdToken).
%
%   DOCX vs pptx: python-docx's CT_GraphicalObjectData exposes `pic` + `uri`
%   (XsdToken). python-pptx's same-named class (oxml/shapes/graphfrm.py) is a
%   BaseShapeElement subclass exposing `chart` / `tbl` children, a `uri`
%   (XsdString), and an OLE-object read surface. Ported to match DOCX (pic +
%   uri, uri typed XsdToken -- NOT XsdString).
%
%   `pic` has NO successors declared (successors=()) -> NO_SUCCESSORS: an inserted
%   pic appends at end. The generated `insert_pic_` (Python `_insert_pic`) is the
%   member CT_Inline.new calls to attach the built pic:pic subtree.
%
%   REQUIREDATTRIBUTE `uri`: GET before set raises mat2doc:InvalidXmlError; SET
%   always writes.
%
%   UNDERSCORE ROTATION (design.md section 2): _new_pic -> new_pic_,
%   _insert_pic -> insert_pic_, _add_pic -> add_pic_.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/shape.py::CT_GraphicalObjectData
%   (lines 61-65; registered for <a:graphicData>, oxml/__init__.py:50)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        NO_SUCCESSORS = string.empty(1, 0)   % pic: successors=() -> append at end
        PIC_TAG = "pic:pic"                  % ZeroOrOne @ shape.py:64
        URI_ATTR = "uri"                     % RequiredAttribute @ shape.py:65
        URI_TYPE = "XsdToken"                % simple type (+oxml\+simpletypes)
    end

    properties (Dependent)  % generated descriptor properties
        pic  % ZeroOrOne('pic:pic') -- child or [] (None)
        uri  % RequiredAttribute('uri', XsdToken) -> string; InvalidXmlError if absent
    end

    methods
        function obj = CT_GraphicalObjectData(varargin)
            % CT_GRAPHICALOBJECTDATA Construct a loose <a:graphicData> -- PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- pic (ZeroOrOne) ----
        function child = get.pic(obj)
            child = obj.getChild(obj.PIC_TAG);
        end
        function child = get_or_add_pic(obj)
            child = obj.getOrAddChild(obj.PIC_TAG, obj.NO_SUCCESSORS);
        end
        function child = new_pic_(obj)
            child = obj.newChild(obj.PIC_TAG);
        end
        function child = insert_pic_(obj, child)
            child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS);
        end
        function child = add_pic_(obj, varargin)
            child = obj.addChild(obj.PIC_TAG, obj.NO_SUCCESSORS, varargin{:});
        end
        function remove_pic_(obj)
            obj.removeChild(obj.PIC_TAG);
        end

        % ---- uri (RequiredAttribute, XsdToken) ----
        function value = get.uri(obj)
            value = obj.getAttrRequired(obj.URI_ATTR, obj.URI_TYPE);
        end
        function set.uri(obj, value)
            obj.setAttrRequired(obj.URI_ATTR, obj.URI_TYPE, value);
        end
    end
end

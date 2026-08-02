classdef CT_BlipFillProperties < mat2doc.oxml.BaseOxmlElement
% CT_BLIPFILLPROPERTIES Custom element class for the <pic:blipFill> element.
%
%   `<pic:blipFill>` element, specifies picture properties. A single ZeroOrOne
%   descriptor `blip` ("a:blip", successors ("a:srcRect","a:tile","a:stretch")),
%   ported as a Constant schema table plus one-line delegating members calling
%   the BaseOxmlElement child engine.
%
%   DOCX vs pptx: python-docx's CT_BlipFillProperties has ONLY the `blip`
%   descriptor (no `srcRect`, no `crop` method). python-pptx's same-named class
%   (oxml/dml/fill.py) adds a `srcRect` ZeroOrOne and a hand-written `crop`
%   method. Ported to match DOCX (blip only). Registered for <pic:blipFill> in
%   docx (oxml/__init__.py:53) -- pptx registers a:blipFill/p:blipFill.
%
%   UNDERSCORE ROTATION (design.md section 2): _new_blip -> new_blip_,
%   _insert_blip -> insert_blip_, _add_blip -> add_blip_.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/shape.py::CT_BlipFillProperties
%   (lines 45-50; registered for <pic:blipFill>, oxml/__init__.py:53)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        BLIP_TAG = "a:blip"                                     % ZeroOrOne @ shape.py:48
        BLIP_SUCCESSORS = ["a:srcRect", "a:tile", "a:stretch"]  % successors @ shape.py:49
    end

    properties (Dependent)  % generated descriptor properties
        blip  % ZeroOrOne('a:blip') -- child or [] (None)
    end

    methods
        function obj = CT_BlipFillProperties(varargin)
            % CT_BLIPFILLPROPERTIES Construct a loose <pic:blipFill> -- PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- blip (ZeroOrOne) ----
        function child = get.blip(obj)
            child = obj.getChild(obj.BLIP_TAG);
        end
        function child = get_or_add_blip(obj)
            child = obj.getOrAddChild(obj.BLIP_TAG, obj.BLIP_SUCCESSORS);
        end
        function child = new_blip_(obj)
            child = obj.newChild(obj.BLIP_TAG);
        end
        function child = insert_blip_(obj, child)
            child = obj.insertChildInSequence(child, obj.BLIP_SUCCESSORS);
        end
        function child = add_blip_(obj, varargin)
            child = obj.addChild(obj.BLIP_TAG, obj.BLIP_SUCCESSORS, varargin{:});
        end
        function remove_blip_(obj)
            obj.removeChild(obj.BLIP_TAG);
        end
    end
end

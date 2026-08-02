classdef CT_Blip < mat2doc.oxml.BaseOxmlElement
% CT_BLIP Custom element class for the <a:blip> element.
%
%   `<a:blip>` element, specifies image source and adjustments such as alpha and
%   tint. Two OptionalAttribute descriptors:
%     embed = OptionalAttribute("r:embed", ST_RelationshipId)  (default None)
%     link  = OptionalAttribute("r:link",  ST_RelationshipId)  (default None)
%   ported as a Constant schema table plus one-line get./set. members delegating
%   to the BaseOxmlElement typed-attribute engine.
%
%   DOCX vs pptx: python-docx's CT_Blip exposes `embed` AND `link` (both
%   r:embed / r:link). python-pptx's CT_Blip (oxml/dml/fill.py) exposes only a
%   single `rEmbed` ("r:embed") attribute. Ported to match DOCX (both attrs,
%   docx member names `embed` / `link`).
%
%   H3 (tri-state): OptionalAttribute default is None (-> [] here); a get when the
%   attribute is absent returns [], a set to [] removes it.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/shape.py::CT_Blip
%   (lines 33-42; registered for <a:blip>, oxml/__init__.py:47)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        EMBED_ATTR = "r:embed"           % OptionalAttribute @ shape.py:37
        EMBED_TYPE = "ST_RelationshipId" % simple type (+oxml\+simpletypes)
        EMBED_DEFAULT = []               % Python default: None
        LINK_ATTR = "r:link"             % OptionalAttribute @ shape.py:40
        LINK_TYPE = "ST_RelationshipId"  % simple type (+oxml\+simpletypes)
        LINK_DEFAULT = []                % Python default: None
    end

    properties (Dependent)  % generated descriptor properties
        embed  % OptionalAttribute('r:embed', ST_RelationshipId) -> typed value or [] (None)
        link   % OptionalAttribute('r:link',  ST_RelationshipId) -> typed value or [] (None)
    end

    methods
        function obj = CT_Blip(varargin)
            % CT_BLIP Construct a loose <a:blip> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        function value = get.embed(obj)
            value = obj.getAttrTyped(obj.EMBED_ATTR, obj.EMBED_TYPE, obj.EMBED_DEFAULT);
        end
        function set.embed(obj, value)
            obj.setAttrTyped(obj.EMBED_ATTR, obj.EMBED_TYPE, value, obj.EMBED_DEFAULT);
        end

        function value = get.link(obj)
            value = obj.getAttrTyped(obj.LINK_ATTR, obj.LINK_TYPE, obj.LINK_DEFAULT);
        end
        function set.link(obj, value)
            obj.setAttrTyped(obj.LINK_ATTR, obj.LINK_TYPE, value, obj.LINK_DEFAULT);
        end
    end
end

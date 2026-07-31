classdef CT_PageMar < mat2doc.oxml.BaseOxmlElement
% CT_PAGEMAR `<w:pgMar>` element, defining page margins.
%
%   Seven OptionalAttributes, all default None ([]) (section.py 60-83):
%     top    OptionalAttribute("w:top",    ST_SignedTwipsMeasure) -> Length
%     right  OptionalAttribute("w:right",  ST_TwipsMeasure)       -> Length
%     bottom OptionalAttribute("w:bottom", ST_SignedTwipsMeasure) -> Length
%     left   OptionalAttribute("w:left",   ST_TwipsMeasure)       -> Length
%     header OptionalAttribute("w:header", ST_TwipsMeasure)       -> Length
%     footer OptionalAttribute("w:footer", ST_TwipsMeasure)       -> Length
%     gutter OptionalAttribute("w:gutter", ST_TwipsMeasure)       -> Length
%
%   H6 (EMU/Length int arithmetic): all seven are Length-typed (twips). top and
%   bottom use ST_SignedTwipsMeasure (signed 32-bit); the other five use
%   ST_TwipsMeasure (unsigned). VERIFIED per attr against section.py 63-83 -- the
%   top/bottom SIGNED vs left/right/header/footer/gutter UNSIGNED split is
%   transcribed exactly (a wrong simple-type would change the accepted range and
%   the from_xml round/parse path).
%
%   H3 (tri-state): no attribute has a Python default (default None -> []); the
%   getter returns [] when absent, the setter removes the attribute when assigned
%   [] (None) -- the standard OptionalAttribute delta (BaseOxmlElement
%   getAttrTyped/setAttrTyped).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on the <w:pgMar> inside default.docx's document.xml sectPr
%   on every M1 load, so all positional args forward verbatim.
%
%   Example:
%       pm = mat2doc.oxml.OxmlElement("w:pgMar");
%       pm.top = mat2doc.shared.Twips(1440);   % <w:pgMar w:top="1440"/>
%
%   Ported from python-docx v1.2.0: src/docx/oxml/section.py::CT_PageMar
%   (lines 60-83; registered for w:pgMar)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        TOP_ATTR      = "w:top"                 % OptionalAttribute @ section.py:63-65
        TOP_TYPE      = "ST_SignedTwipsMeasure" % signed
        TOP_DEFAULT   = []                       % Python default: None
        RIGHT_ATTR    = "w:right"               % OptionalAttribute @ section.py:66-68
        RIGHT_TYPE    = "ST_TwipsMeasure"       % unsigned
        RIGHT_DEFAULT = []                       % Python default: None
        BOTTOM_ATTR   = "w:bottom"              % OptionalAttribute @ section.py:69-71
        BOTTOM_TYPE   = "ST_SignedTwipsMeasure" % signed
        BOTTOM_DEFAULT = []                      % Python default: None
        LEFT_ATTR     = "w:left"                % OptionalAttribute @ section.py:72-74
        LEFT_TYPE     = "ST_TwipsMeasure"       % unsigned
        LEFT_DEFAULT  = []                       % Python default: None
        HEADER_ATTR   = "w:header"              % OptionalAttribute @ section.py:75-77
        HEADER_TYPE   = "ST_TwipsMeasure"       % unsigned
        HEADER_DEFAULT = []                      % Python default: None
        FOOTER_ATTR   = "w:footer"              % OptionalAttribute @ section.py:78-80
        FOOTER_TYPE   = "ST_TwipsMeasure"       % unsigned
        FOOTER_DEFAULT = []                      % Python default: None
        GUTTER_ATTR   = "w:gutter"              % OptionalAttribute @ section.py:81-83
        GUTTER_TYPE   = "ST_TwipsMeasure"       % unsigned
        GUTTER_DEFAULT = []                      % Python default: None
    end

    properties (Dependent)  % generated descriptor properties
        top     % OptionalAttribute('w:top', ST_SignedTwipsMeasure) -> Length or []
        right   % OptionalAttribute('w:right', ST_TwipsMeasure) -> Length or []
        bottom  % OptionalAttribute('w:bottom', ST_SignedTwipsMeasure) -> Length or []
        left    % OptionalAttribute('w:left', ST_TwipsMeasure) -> Length or []
        header  % OptionalAttribute('w:header', ST_TwipsMeasure) -> Length or []
        footer  % OptionalAttribute('w:footer', ST_TwipsMeasure) -> Length or []
        gutter  % OptionalAttribute('w:gutter', ST_TwipsMeasure) -> Length or []
    end

    methods
        function obj = CT_PageMar(varargin)
            % CT_PAGEMAR Construct a loose <w:pgMar> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- top (OptionalAttribute, ST_SignedTwipsMeasure, default None) ----
        function value = get.top(obj)
            value = obj.getAttrTyped(obj.TOP_ATTR, obj.TOP_TYPE, obj.TOP_DEFAULT);
        end
        function set.top(obj, value)
            obj.setAttrTyped(obj.TOP_ATTR, obj.TOP_TYPE, value, obj.TOP_DEFAULT);
        end

        % ---- right (OptionalAttribute, ST_TwipsMeasure, default None) ----
        function value = get.right(obj)
            value = obj.getAttrTyped(obj.RIGHT_ATTR, obj.RIGHT_TYPE, obj.RIGHT_DEFAULT);
        end
        function set.right(obj, value)
            obj.setAttrTyped(obj.RIGHT_ATTR, obj.RIGHT_TYPE, value, obj.RIGHT_DEFAULT);
        end

        % ---- bottom (OptionalAttribute, ST_SignedTwipsMeasure, default None) ----
        function value = get.bottom(obj)
            value = obj.getAttrTyped(obj.BOTTOM_ATTR, obj.BOTTOM_TYPE, obj.BOTTOM_DEFAULT);
        end
        function set.bottom(obj, value)
            obj.setAttrTyped(obj.BOTTOM_ATTR, obj.BOTTOM_TYPE, value, obj.BOTTOM_DEFAULT);
        end

        % ---- left (OptionalAttribute, ST_TwipsMeasure, default None) ----
        function value = get.left(obj)
            value = obj.getAttrTyped(obj.LEFT_ATTR, obj.LEFT_TYPE, obj.LEFT_DEFAULT);
        end
        function set.left(obj, value)
            obj.setAttrTyped(obj.LEFT_ATTR, obj.LEFT_TYPE, value, obj.LEFT_DEFAULT);
        end

        % ---- header (OptionalAttribute, ST_TwipsMeasure, default None) ----
        function value = get.header(obj)
            value = obj.getAttrTyped(obj.HEADER_ATTR, obj.HEADER_TYPE, obj.HEADER_DEFAULT);
        end
        function set.header(obj, value)
            obj.setAttrTyped(obj.HEADER_ATTR, obj.HEADER_TYPE, value, obj.HEADER_DEFAULT);
        end

        % ---- footer (OptionalAttribute, ST_TwipsMeasure, default None) ----
        function value = get.footer(obj)
            value = obj.getAttrTyped(obj.FOOTER_ATTR, obj.FOOTER_TYPE, obj.FOOTER_DEFAULT);
        end
        function set.footer(obj, value)
            obj.setAttrTyped(obj.FOOTER_ATTR, obj.FOOTER_TYPE, value, obj.FOOTER_DEFAULT);
        end

        % ---- gutter (OptionalAttribute, ST_TwipsMeasure, default None) ----
        function value = get.gutter(obj)
            value = obj.getAttrTyped(obj.GUTTER_ATTR, obj.GUTTER_TYPE, obj.GUTTER_DEFAULT);
        end
        function set.gutter(obj, value)
            obj.setAttrTyped(obj.GUTTER_ATTR, obj.GUTTER_TYPE, value, obj.GUTTER_DEFAULT);
        end
    end
end

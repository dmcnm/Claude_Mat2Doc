classdef CT_Color < mat2doc.oxml.BaseOxmlElement
% CT_COLOR `w:color` element, specifying the color of a font (and other objects).
%
%   `val`        RequiredAttribute("w:val", ST_HexColor) -> an RGBColor object
%                (or the string "auto"); InvalidXmlError if @w:val is absent.
%   `themeColor` OptionalAttribute("w:themeColor", MSO_THEME_COLOR), default
%                None ([]) -> an MSO_THEME_COLOR member or [] when absent.
%
%   H3: `themeColor` has NO Python default, so its default is None ([]); the
%   getter returns [] when absent and the setter removes the attribute when
%   assigned [] (None). The enum simple-type is referenced by its FULLY
%   QUALIFIED name so BaseOxmlElement.resolveTypeCls_ dispatches to +enum
%   verbatim (font.py imports MSO_THEME_COLOR from docx.enum.dml).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on <w:color> nodes inside styles.xml/document.xml.
%
%   Example:
%       c = mat2doc.oxml.OxmlElement("w:color");
%       c.val = mat2doc.shared.RGBColor(60, 47, 128);   % <w:color w:val="3C2F80"/>
%       c.val                                           % an RGBColor
%
%   Ported from python-docx v1.2.0: src/docx/oxml/text/font.py::CT_Color
%   (lines 32-36; registered for w:color)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        VAL_ATTR          = "w:val"          % RequiredAttribute @ font.py:35
        VAL_TYPE          = "ST_HexColor"    % simple type (+oxml\+simpletypes)
        THEMECOLOR_ATTR   = "w:themeColor"   % OptionalAttribute @ font.py:36
        THEMECOLOR_TYPE   = "mat2doc.enum.dml.MSO_THEME_COLOR"  % enum simple-type (verbatim, resolveTypeCls_)
        THEMECOLOR_DEFAULT = []              % Python default: None (no default arg)
    end

    properties (Dependent)  % generated descriptor properties
        val         % RequiredAttribute('w:val', ST_HexColor) -> RGBColor / "auto"
        themeColor  % OptionalAttribute('w:themeColor', MSO_THEME_COLOR) -> member or [] (None)
    end

    methods
        function obj = CT_Color(varargin)
            % CT_COLOR Construct a loose <w:color> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- val (RequiredAttribute) ----
        function value = get.val(obj)
            value = obj.getAttrRequired(obj.VAL_ATTR, obj.VAL_TYPE);
        end
        function set.val(obj, value)
            obj.setAttrRequired(obj.VAL_ATTR, obj.VAL_TYPE, value);
        end

        % ---- themeColor (OptionalAttribute, MSO_THEME_COLOR, default None) ----
        function value = get.themeColor(obj)
            value = obj.getAttrTyped(obj.THEMECOLOR_ATTR, obj.THEMECOLOR_TYPE, obj.THEMECOLOR_DEFAULT);
        end
        function set.themeColor(obj, value)
            obj.setAttrTyped(obj.THEMECOLOR_ATTR, obj.THEMECOLOR_TYPE, value, obj.THEMECOLOR_DEFAULT);
        end
    end
end

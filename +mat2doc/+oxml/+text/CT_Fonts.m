classdef CT_Fonts < mat2doc.oxml.BaseOxmlElement
% CT_FONTS `<w:rFonts>` element: typeface names for the various language types.
%
%   python-docx exposes only the two common-case OptionalAttributes on this
%   element (the schema has more; only these are ported, matching font.py):
%     `ascii` OptionalAttribute("w:ascii", ST_String), default None ([])
%     `hAnsi` OptionalAttribute("w:hAnsi", ST_String), default None ([])
%
%   H3: neither attribute has a Python default, so the default is None ([]); the
%   getter returns [] when absent and the setter removes the attribute when
%   assigned [] (None). ST_String is the identity string transform.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on <w:rFonts> nodes inside styles.xml/document.xml.
%
%   Example:
%       f = mat2doc.oxml.OxmlElement("w:rFonts");
%       f.ascii = "Arial";              % <w:rFonts w:ascii="Arial"/>
%
%   Ported from python-docx v1.2.0: src/docx/oxml/text/font.py::CT_Fonts
%   (lines 39-46; registered for w:rFonts)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        ASCII_ATTR   = "w:ascii"    % OptionalAttribute @ font.py:45
        ASCII_TYPE   = "ST_String"  % simple type (+oxml\+simpletypes)
        ASCII_DEFAULT = []          % Python default: None (no default arg)
        HANSI_ATTR   = "w:hAnsi"    % OptionalAttribute @ font.py:46
        HANSI_TYPE   = "ST_String"  % simple type (+oxml\+simpletypes)
        HANSI_DEFAULT = []          % Python default: None (no default arg)
    end

    properties (Dependent)  % generated descriptor properties
        ascii   % OptionalAttribute('w:ascii', ST_String) -> string or [] (None)
        hAnsi   % OptionalAttribute('w:hAnsi', ST_String) -> string or [] (None)
    end

    methods
        function obj = CT_Fonts(varargin)
            % CT_FONTS Construct a loose <w:rFonts> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- ascii (OptionalAttribute, ST_String, default None) ----
        function value = get.ascii(obj)
            value = obj.getAttrTyped(obj.ASCII_ATTR, obj.ASCII_TYPE, obj.ASCII_DEFAULT);
        end
        function set.ascii(obj, value)
            obj.setAttrTyped(obj.ASCII_ATTR, obj.ASCII_TYPE, value, obj.ASCII_DEFAULT);
        end

        % ---- hAnsi (OptionalAttribute, ST_String, default None) ----
        function value = get.hAnsi(obj)
            value = obj.getAttrTyped(obj.HANSI_ATTR, obj.HANSI_TYPE, obj.HANSI_DEFAULT);
        end
        function set.hAnsi(obj, value)
            obj.setAttrTyped(obj.HANSI_ATTR, obj.HANSI_TYPE, value, obj.HANSI_DEFAULT);
        end
    end
end

classdef CT_Underline < mat2doc.oxml.BaseOxmlElement
% CT_UNDERLINE `<w:u>` element, specifying the underlining style for a run.
%
%   `val` OptionalAttribute("w:val", WD_UNDERLINE), default None ([]) -> a
%   WD_UNDERLINE member, or [] when @w:val is absent.
%
%   H3 tri-state: `val` has NO Python default (default None -> []). getAttrTyped
%   returns [] when the attribute is absent; the setter removes @w:val when
%   assigned [] (None). WD_UNDERLINE is an enum simple-type referenced by its
%   FULLY QUALIFIED name (resolveTypeCls_ dispatches to +enum verbatim). Note the
%   "single"/"none"/None -> True/False/None run-visible mapping lives on the Font
%   PROXY (font.py Font.underline, a later WP), NOT here: this element attribute
%   just (de)serializes the enum member. WD_UNDERLINE.from_xml("none") ->
%   WD_UNDERLINE.NONE; from_xml("single") -> WD_UNDERLINE.SINGLE.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on <w:u> nodes inside a real part.
%
%   Example:
%       u = mat2doc.oxml.OxmlElement("w:u");
%       u.val = mat2doc.enum.text.WD_UNDERLINE.SINGLE;  % <w:u w:val="single"/>
%       u.val                                           % [] when @w:val absent
%
%   Ported from python-docx v1.2.0: src/docx/oxml/text/font.py::CT_Underline
%   (lines 322-325; registered for w:u)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        VAL_ATTR    = "w:val"                          % OptionalAttribute @ font.py:325
        VAL_TYPE    = "mat2doc.enum.text.WD_UNDERLINE" % enum simple-type (verbatim, resolveTypeCls_)
        VAL_DEFAULT = []                               % Python default: None (no default arg)
    end

    properties (Dependent)  % generated descriptor property
        val   % OptionalAttribute('w:val', WD_UNDERLINE) -> member or [] (None)
    end

    methods
        function obj = CT_Underline(varargin)
            % CT_UNDERLINE Construct a loose <w:u> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        function value = get.val(obj)
            value = obj.getAttrTyped(obj.VAL_ATTR, obj.VAL_TYPE, obj.VAL_DEFAULT);
        end
        function set.val(obj, value)
            obj.setAttrTyped(obj.VAL_ATTR, obj.VAL_TYPE, value, obj.VAL_DEFAULT);
        end
    end
end

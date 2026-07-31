classdef CT_SectType < mat2doc.oxml.BaseOxmlElement
% CT_SECTTYPE `<w:type>` element, defining the section start type.
%
%   (The python-docx docstring reads "``<w:sectType>``", but the element is
%   registered and used as `w:type` -- oxml/__init__.py:130
%   register_element_cls("w:type", CT_SectType). The tag ported here is w:type.)
%
%   One OptionalAttribute (section.py 429-431):
%     val OptionalAttribute("w:val", WD_SECTION_START) -> member, default None
%
%   H10 (enum dispatch): WD_SECTION_START is referenced by its fully qualified
%   name so resolveTypeCls_ dispatches to +enum verbatim. H3 (tri-state): val has
%   NO Python default (None -> []); the getter returns [] when @w:val is absent,
%   and the setter removes @w:val when assigned [] (None). CT_SectPr.start_type
%   reads .val on this child (returning NEW_PAGE when the child or its @val is
%   absent).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on any <w:type> node inside a sectPr, so all positional
%   args forward verbatim.
%
%   Example:
%       t = mat2doc.oxml.OxmlElement("w:type");
%       t.val = mat2doc.enum.section.WD_SECTION_START.EVEN_PAGE;  % <w:type w:val="evenPage"/>
%
%   Ported from python-docx v1.2.0: src/docx/oxml/section.py::CT_SectType
%   (lines 426-431; registered for w:type)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        VAL_ATTR    = "w:val"                              % OptionalAttribute @ section.py:429-431
        VAL_TYPE    = "mat2doc.enum.section.WD_SECTION_START" % enum simple-type (verbatim, resolveTypeCls_)
        VAL_DEFAULT = []                                    % Python default: None
    end

    properties (Dependent)  % generated descriptor property
        val   % OptionalAttribute('w:val', WD_SECTION_START) -> member or []
    end

    methods
        function obj = CT_SectType(varargin)
            % CT_SECTTYPE Construct a loose <w:type> -- TRANSPARENT PASS-THROUGH.
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

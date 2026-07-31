classdef CT_VerticalJc < mat2doc.oxml.BaseOxmlElement
% CT_VERTICALJC `<w:vAlign>` element: vertical alignment of a table cell.
%
%   One RequiredAttribute (table.py 967-969):
%     val RequiredAttribute("w:val", WD_CELL_VERTICAL_ALIGNMENT) -> a member
%
%   H10 (enum dispatch): WD_CELL_VERTICAL_ALIGNMENT is referenced by its fully
%   qualified name so resolveTypeCls_ dispatches to +enum\+table verbatim
%   (getAttrRequired/setAttrRequired route from_xml/to_xml through it).
%
%   H3 (RequiredAttribute): `val` is REQUIRED -- reading it when @w:val is absent
%   raises mat2doc:InvalidXmlError (getAttrRequired), never a default; the setter
%   always writes @w:val (setAttrRequired), never removes. (CT_TcPr.vAlign_val --
%   OUT of scope, P6-3 -- guards for the absent CHILD, returning None when there
%   is no <w:vAlign> at all; the @w:val on an EXISTING <w:vAlign> is required.)
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on any <w:vAlign> node inside a real table cell's <w:tcPr>,
%   so all positional args forward verbatim.
%
%   Example:
%       va = mat2doc.oxml.OxmlElement("w:vAlign");
%       va.val = mat2doc.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER;
%       % <w:vAlign w:val="center"/>
%
%   Ported from python-docx v1.2.0: src/docx/oxml/table.py::CT_VerticalJc
%   (lines 964-969; registered for w:vAlign)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        VAL_ATTR = "w:val"                                        % RequiredAttribute @ table.py:967-969
        VAL_TYPE = "mat2doc.enum.table.WD_CELL_VERTICAL_ALIGNMENT"  % enum simple-type (verbatim, resolveTypeCls_)
    end

    properties (Dependent)  % generated descriptor property
        val   % RequiredAttribute('w:val', WD_CELL_VERTICAL_ALIGNMENT) -> member; InvalidXmlError if absent
    end

    methods
        function obj = CT_VerticalJc(varargin)
            % CT_VERTICALJC Construct a loose <w:vAlign> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- val (RequiredAttribute, WD_CELL_VERTICAL_ALIGNMENT) ----
        function value = get.val(obj)
            value = obj.getAttrRequired(obj.VAL_ATTR, obj.VAL_TYPE);
        end
        function set.val(obj, value)
            obj.setAttrRequired(obj.VAL_ATTR, obj.VAL_TYPE, value);
        end
    end
end

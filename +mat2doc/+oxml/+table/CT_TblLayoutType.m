classdef CT_TblLayoutType < mat2doc.oxml.BaseOxmlElement
% CT_TBLLAYOUTTYPE `<w:tblLayout>` element: fixed vs autofit column widths.
%
%   Specifies whether column widths are fixed or auto-adjusted to content.
%
%   One OptionalAttribute (table.py 292-294):
%     type OptionalAttribute("w:type", ST_TblLayoutType) -> string, default None
%
%   H3 (tri-state): `type` has NO Python default (None -> []); the getter returns
%   [] when @w:type is absent, the setter removes @w:type when assigned [] (None).
%   ST_TblLayoutType.to_xml validates the value is "fixed" or "autofit". (The
%   CT_TblPr.autofit accessor -- OUT of scope, P6-2/P6-3 -- reads this .type and
%   treats absent/non-"fixed" as autofit; this leaf class stores only the attr.)
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on any <w:tblLayout> node inside a real table's <w:tblPr>,
%   so all positional args forward verbatim.
%
%   Example:
%       lay = mat2doc.oxml.OxmlElement("w:tblLayout");
%       lay.type = "fixed";   % <w:tblLayout w:type="fixed"/>
%
%   Ported from python-docx v1.2.0: src/docx/oxml/table.py::CT_TblLayoutType
%   (lines 285-294; registered for w:tblLayout)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        TYPE_ATTR    = "w:type"            % OptionalAttribute @ table.py:292-294
        TYPE_TYPE    = "ST_TblLayoutType"  % simple type (+oxml\+simpletypes)
        TYPE_DEFAULT = []                  % Python default: None
    end

    properties (Dependent)  % generated descriptor property
        type   % OptionalAttribute('w:type', ST_TblLayoutType) -> string or []
    end

    methods
        function obj = CT_TblLayoutType(varargin)
            % CT_TBLLAYOUTTYPE Construct a loose <w:tblLayout> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- type (OptionalAttribute, ST_TblLayoutType, default None) ----
        function value = get.type(obj)
            value = obj.getAttrTyped(obj.TYPE_ATTR, obj.TYPE_TYPE, obj.TYPE_DEFAULT);
        end
        function set.type(obj, value)
            obj.setAttrTyped(obj.TYPE_ATTR, obj.TYPE_TYPE, value, obj.TYPE_DEFAULT);
        end
    end
end

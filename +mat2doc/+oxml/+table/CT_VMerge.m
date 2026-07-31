classdef CT_VMerge < mat2doc.oxml.BaseOxmlElement
% CT_VMERGE `<w:vMerge>` element: vertical-merge behavior of a table cell.
%
%   One OptionalAttribute WITH A NON-None DEFAULT (table.py 975-977):
%     val OptionalAttribute("w:val", ST_Merge, default=ST_Merge.CONTINUE) -> string
%
%   H3 (tri-state -- the ONE non-None default in this WP): the Python default is
%   ST_Merge.CONTINUE, i.e. the STRING "continue" (NOT None/[]). So VAL_DEFAULT =
%   "continue". OptionalAttribute semantics (BaseOxmlElement getAttrTyped/
%   setAttrTyped, docx delta):
%     get: @w:val absent -> "continue" (the default); present -> from_xml (the
%          literal string, ST_Merge inherits identity convert_from_xml).
%     set: value == [] (None) OR value == "continue" (the default) -> REMOVE
%          @w:val; "restart" -> write @w:val="restart". So a <w:vMerge/> with no
%          @w:val reads as "continue", and assigning "continue" strips the attr.
%   H4/H3: this is exactly why setAttrTyped tests `isequal(value, default)` (the
%   docx `value == self._default` branch) -- "" is a real string, not the
%   default, and never removes.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on any <w:vMerge> node inside a real table cell's <w:tcPr>,
%   so all positional args forward verbatim.
%
%   Example:
%       vm = mat2doc.oxml.OxmlElement("w:vMerge");
%       vm.val                         % "continue"  (default, @w:val absent)
%       vm.val = "restart";            % <w:vMerge w:val="restart"/>
%       vm.val = "continue";           % strips @w:val -> <w:vMerge/>
%
%   Ported from python-docx v1.2.0: src/docx/oxml/table.py::CT_VMerge
%   (lines 972-977; registered for w:vMerge)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        VAL_ATTR    = "w:val"      % OptionalAttribute @ table.py:975-977
        VAL_TYPE    = "ST_Merge"   % simple type (+oxml\+simpletypes)
        VAL_DEFAULT = "continue"   % Python default: ST_Merge.CONTINUE == "continue"
    end

    properties (Dependent)  % generated descriptor property
        val   % OptionalAttribute('w:val', ST_Merge, default "continue") -> string
    end

    methods
        function obj = CT_VMerge(varargin)
            % CT_VMERGE Construct a loose <w:vMerge> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- val (OptionalAttribute, ST_Merge, default "continue") ----
        function value = get.val(obj)
            value = obj.getAttrTyped(obj.VAL_ATTR, obj.VAL_TYPE, obj.VAL_DEFAULT);
        end
        function set.val(obj, value)
            obj.setAttrTyped(obj.VAL_ATTR, obj.VAL_TYPE, value, obj.VAL_DEFAULT);
        end
    end
end

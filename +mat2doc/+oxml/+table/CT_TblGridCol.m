classdef CT_TblGridCol < mat2doc.oxml.BaseOxmlElement
% CT_TBLGRIDCOL `<w:gridCol>` element, child of `<w:tblGrid>`, one table column.
%
%   One OptionalAttribute (table.py 274-276):
%     w OptionalAttribute("w:w", ST_TwipsMeasure) -> Length, default None
%
%   H6 (EMU/Length): `w` is the column width in twips (ST_TwipsMeasure); the
%   getter returns a Length, the round-trip goes through Twips/Emu. H3
%   (tri-state): `w` has NO Python default (None -> []); the getter returns []
%   when @w:w is absent, the setter removes @w:w when assigned [] (None).
%
%   gridCol_idx (@property, table.py 278-282): the 0-BASED position of this
%   <w:gridCol> within its parent <w:tblGrid>'s gridCol_lst. Python:
%   `tblGrid.gridCol_lst.index(self)`. H1: Python list.index() is 0-based and
%   its VALUE is data, so the port subtracts 1 from the 1-based find() (% IDX).
%   H5: identity match (== on handles). ValueError if this element is not in its
%   parent's gridCol list (list.index() on absence) -- unreachable in practice
%   (a parented gridCol is always in its parent's list).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on each <w:gridCol> inside a real table's <w:tblGrid>, so
%   all positional args forward verbatim.
%
%   Example:
%       col = mat2doc.oxml.OxmlElement("w:gridCol");
%       col.w = mat2doc.shared.Twips(2880);   % <w:gridCol w:w="2880"/>
%
%   Ported from python-docx v1.2.0: src/docx/oxml/table.py::CT_TblGridCol
%   (lines 271-282; registered for w:gridCol)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        W_ATTR    = "w:w"              % OptionalAttribute @ table.py:274-276
        W_TYPE    = "ST_TwipsMeasure"  % simple type (+oxml\+simpletypes)
        W_DEFAULT = []                 % Python default: None
    end

    properties (Dependent)  % generated descriptor property + computed index
        w           % OptionalAttribute('w:w', ST_TwipsMeasure) -> Length or []
        gridCol_idx % @property: 0-based position within parent tblGrid's gridCol_lst
    end

    methods
        function obj = CT_TblGridCol(varargin)
            % CT_TBLGRIDCOL Construct a loose <w:gridCol> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- w (OptionalAttribute, ST_TwipsMeasure, default None) ----
        function value = get.w(obj)
            value = obj.getAttrTyped(obj.W_ATTR, obj.W_TYPE, obj.W_DEFAULT);
        end
        function set.w(obj, value)
            obj.setAttrTyped(obj.W_ATTR, obj.W_TYPE, value, obj.W_DEFAULT);
        end

        % ---- gridCol_idx (@property, table.py 278-282) ----
        function value = get.gridCol_idx(obj)
            % Python: tblGrid = cast(CT_TblGrid, self.getparent())
            %         return tblGrid.gridCol_lst.index(self)
            tblGrid = obj.getparent();
            lst = tblGrid.gridCol_lst;
            i = find(lst == obj, 1);          % H5 identity; 1-based find
            if isempty(i)                     % Python list.index() ValueError
                error("mat2doc:ValueError", ...
                    "%s is not in list", "gridCol");
            end
            value = i - 1;                    % IDX: 0-based to match Python list.index()
        end
    end
end

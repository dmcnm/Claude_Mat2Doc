classdef CT_TblGrid < mat2doc.oxml.BaseOxmlElement
% CT_TBLGRID `<w:tblGrid>` element, child of `<w:tbl>`.
%
%   Holds `<w:gridCol>` children that define the table's column count and widths.
%
%   DESCRIPTOR (table.py 268):
%     gridCol = ZeroOrMore("w:gridCol", successors=("w:tblGridChange",))
%
%   xmlchemy ZeroOrMore member generation (docx form): gridCol_lst, new_gridCol_,
%   insert_gridCol_, add_gridCol_, add_gridCol (PUBLIC). No bare `gridCol`
%   getter (delattr), no get_or_add, no remover. Underscore rotation:
%   _new_gridCol->new_gridCol_, _insert_gridCol->insert_gridCol_,
%   _add_gridCol->add_gridCol_ (the PUBLIC add_gridCol keeps its bare name --
%   docx generates a public adder for ZeroOrMore, D-delta-4).
%
%   H11 (child ordering): successors=("w:tblGridChange",) -> SUCCESSORS =
%   ["w:tblGridChange"]; a newly inserted gridCol goes before the FIRST present
%   <w:tblGridChange> (else appended), keeping the schema order. On a plain grid
%   (no tblGridChange) every add_gridCol APPENDS, so columns preserve add order.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on the single <w:tblGrid> of a real table, so all
%   positional args forward verbatim.
%
%   Example:
%       grid = mat2doc.oxml.OxmlElement("w:tblGrid");
%       c1 = grid.add_gridCol();  c1.w = mat2doc.shared.Twips(2880);
%       c2 = grid.add_gridCol();  c2.w = mat2doc.shared.Twips(2880);
%       numel(grid.gridCol_lst)   % 2
%
%   Ported from python-docx v1.2.0: src/docx/oxml/table.py::CT_TblGrid
%   (lines 259-268; registered for w:tblGrid)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        GRIDCOL_TAG = "w:gridCol"                 % ZeroOrMore @ table.py:268
        SUCCESSORS  = "w:tblGridChange"           % successors=("w:tblGridChange",)
    end

    properties (Dependent)  % ZeroOrMore list getter
        gridCol_lst   % list of <w:gridCol> children (document order)
    end

    methods
        function obj = CT_TblGrid(varargin)
            % CT_TBLGRID Construct a loose <w:tblGrid> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ============ gridCol (ZeroOrMore, successors=("w:tblGridChange",)) ============
        function lst = get.gridCol_lst(obj);          lst = obj.getChildList(obj.GRIDCOL_TAG); end
        function child = new_gridCol_(obj);           child = obj.newChild(obj.GRIDCOL_TAG); end
        function child = insert_gridCol_(obj, child); child = obj.insertChildInSequence(child, obj.SUCCESSORS); end
        function child = add_gridCol_(obj, varargin); child = obj.addChild(obj.GRIDCOL_TAG, obj.SUCCESSORS, varargin{:}); end
        function child = add_gridCol(obj);            child = obj.add_gridCol_(); end  % public adder (D-delta-4)
    end
end

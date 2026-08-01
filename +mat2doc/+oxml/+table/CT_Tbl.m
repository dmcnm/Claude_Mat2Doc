classdef CT_Tbl < mat2doc.oxml.BaseOxmlElement
% CT_TBL `<w:tbl>` element -- the table root. Registered for w:tbl
%   (oxml/__init__.py:173). Completes the table oxml layer.
%
%   Holds the OneAndOnlyOne <w:tblPr> (table properties) and <w:tblGrid> (column
%   grid) plus its ZeroOrMore <w:tr> rows (CT_Row, P6-2). Carries the whole-table
%   accessors (bidiVisual direction, tblStyle id, col_count, iter_tcs) and -- most
%   importantly -- the table CONSTRUCTOR new_tbl(rows, cols, width) that emits the
%   initial <w:tbl> XML the API tier (BlockItemContainer.add_table, P6-4a) calls.
%
%   ===================== DESCRIPTORS (table.py 152-154) =========================
%     tblPr   = OneAndOnlyOne("w:tblPr")   -- required; get.tblPr = getRequiredChild
%     tblGrid = OneAndOnlyOne("w:tblGrid") -- required; get.tblGrid = getRequiredChild
%     tr      = ZeroOrMore("w:tr")          -- successors=() -> append at end
%   The Callable annotations (table.py 149-150: add_tr / tr_lst) are type hints for
%   the generated <w:tr> members. OneAndOnlyOne generates ONLY the getter
%   (xmlchemy 499-505); ZeroOrMore generates (docx form, D-delta-4): tr_lst,
%   new_tr_, insert_tr_, add_tr_, add_tr (PUBLIC) -- NO bare `tr` getter, NO
%   get_or_add, NO remover. Underscore rotation: _new_tr->new_tr_,
%   _insert_tr->insert_tr_, _add_tr->add_tr_ (the PUBLIC add_tr keeps its bare
%   name). tr has no _new_tr override, so new_tr_ uses the default creator
%   (OxmlElement("w:tr") -> CT_Row via registry).
%
%   ===================== THE new_tbl CONSTRUCTOR (H11 byte-critical) ============
%   new_tbl(rows, cols, width) (table.py 191-197) = parse_xml(_tbl_xml(...)). The
%   four private builders _tbl_xml/_tblGrid_xml/_trs_xml/_tcs_xml (table.py
%   219-256) emit a SPECIFIC XML string that MUST be byte-identical to python-docx
%   before parse; they are transcribed VERBATIM below (tbl_xml_/tblGrid_xml_/
%   trs_xml_/tcs_xml_). The child order is tblPr, tblGrid, tr* (H11). The parser
%   drops the pretty-print whitespace (remove_blank_text), and the shared oxml
%   serializer is already byte-proven, so an identical pre-parse string yields
%   identical post-parse bytes. col_width = Emu(width // cols) is Python floor
%   division (H6: // -> floor), Emu(0) when cols == 0. col_width.twips is formatted
%   as a Python int via pyStr (H14; Python `'%d' % .twips` / f`{.twips}`).
%
%   ===================== bidiVisual_val (H10 / H4 / H3) =========================
%   getter (table.py 156-165): self.tblPr.bidiVisual is a CT_OnOff or [] (H3); .val
%   is a logical (default True). setter (table.py 167-173): value is a
%   WD_TABLE_DIRECTION member (or []). Python `bool(value)`: WD_TABLE_DIRECTION
%   subclasses int (enum/base.py BaseEnum(int, Enum)), so bool(LTR)=bool(0)=False,
%   bool(RTL)=bool(1)=True (H4). Ported as `logical(double(value.value) ~= 0)`
%   (H10 per-site double(.value)). Consequence via CT_OnOff.val's True default
%   (D-delta-1): RTL (True == default) removes @w:val -> bare <w:bidiVisual/>;
%   LTR (False) writes <w:bidiVisual w:val="0"/> -- faithful to python-docx.
%
%   tblStyle_val (table.py 199-217): get -> self.tblPr.tblStyle.val (a CT_String)
%   or [] (H3). set -> tblPr._remove_tblStyle() UNCONDITIONALLY, then if styleId
%   is [] return, else tblPr._add_tblStyle().val = styleId (the PRIVATE adder, not
%   get_or_add; add_tblStyle_ on CT_TblPr).
%
%   col_count (table.py 175-178): numel(self.tblGrid.gridCol_lst). iter_tcs
%   (table.py 180-189): a Python GENERATOR yielding each <w:tc> row-by-row,
%   left-to-right top-to-bottom; ported as a materialized 1xN heterogeneous
%   XmlElement array of CT_Tc (H9 -- no tree mutation during iteration by callers).
%
%   ===================== UN-DEFER SWEEP (registering w:tbl) =====================
%   Registering w:tbl -> CT_Tbl upgrades every previously-generic <w:tbl> node:
%     * CT_Tc merge (V2): self._tbl (tbl_) now yields a CT_Tbl, so the P6-3a
%       generic-ancestor shim `trLstOfTbl_` (tbl.xpath("./w:tr")) is REPLACED by
%       the faithful `self._tbl.tr_lst` (CT_Tbl.tr_lst == findall("w:tr"), the same
%       CT_Row handle list). Byte/identity-neutral -- the merge byte-matrix stays
%       byte-identical.
%     * CT_Body.tbl_lst / inner_content_elements, CT_HdrFtr.inner_content_elements,
%       CT_SectPr.iter_inner_content (SectBlockElementIterator_): all TAG-BASED
%       xpath, so the returned node-set is unchanged -- only the element CLASS of a
%       <w:tbl> match auto-upgrades generic->CT_Tbl. No code change; the RETURNED
%       class flips (documented for Gate-4).
%   The proxy Section.iter_inner_content / BlockItemContainer table members still
%   raise mat2doc:notYetPorted at the Table proxy (P6-4a) boundary -- keyed on the
%   ABSENCE of the Table proxy, NOT on the w:tbl element class, so they do NOT flip.
%
%   M1-NEUTRAL: default.docx has no <w:tbl> (styles.xml table STYLE defs use
%   tblPr/tcPr, not a w:tbl root), so nothing transits CT_Tbl on the M1 parse path;
%   Document().save() stays 17/17 byte-identical.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on the <w:tbl> root of a real table subtree.
%
%   Example:
%       tbl = mat2doc.oxml.table.CT_Tbl.new_tbl(2, 3, mat2doc.shared.Inches(6));
%       tbl.col_count            % 3
%       tbl.tblStyle_val = "TableGrid";   % <w:tblStyle w:val="TableGrid"/>
%       cells = tbl.iter_tcs;    % 1x6 CT_Tc array (row-major)
%
%   Ported from python-docx v1.2.0: src/docx/oxml/table.py::CT_Tbl
%   (lines 146-256; registered for w:tbl, oxml/__init__.py:173)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        NO_SUCCESSORS = string.empty(1, 0)   % tr: ZeroOrMore successors=() -> append at end
    end

    properties (Dependent)  % OneAndOnlyOne getters + ZeroOrMore list + @property members
        tblPr           % OneAndOnlyOne <w:tblPr> child (required; CT_TblPr)
        tblGrid         % OneAndOnlyOne <w:tblGrid> child (required; CT_TblGrid)
        tr_lst          % ZeroOrMore list of <w:tr> children (document order; CT_Row)
        bidiVisual_val  % @property (get+set): ./w:tblPr/w:bidiVisual/@w:val (logical) or []
        col_count       % @property: number of grid columns (len tblGrid.gridCol_lst)
        tblStyle_val    % @property (get+set): ./w:tblPr/w:tblStyle/@w:val (string) or []
    end

    methods
        function obj = CT_Tbl(varargin)
            % CT_TBL Construct a loose <w:tbl> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ============ tblPr / tblGrid (OneAndOnlyOne getters) ============
        function child = get.tblPr(obj);   child = obj.getRequiredChild("w:tblPr"); end
        function child = get.tblGrid(obj); child = obj.getRequiredChild("w:tblGrid"); end

        % ============ tr (ZeroOrMore, successors=() -> append at end) ============
        function lst = get.tr_lst(obj);          lst = obj.getChildList("w:tr"); end
        function child = new_tr_(obj);           child = obj.newChild("w:tr"); end
        function child = insert_tr_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_tr_(obj, varargin); child = obj.addChild("w:tr", obj.NO_SUCCESSORS, varargin{:}); end
        function child = add_tr(obj);            child = obj.add_tr_(); end   % public adder (D-delta-4)

        % ===================== @property members (table.py 156-217) ===============

        % ---- bidiVisual_val (get+set, table.py 156-173) ----
        function value = get.bidiVisual_val(obj)
            % Python: bidiVisual = self.tblPr.bidiVisual; if bidiVisual is None:
            %             return None
            %         return bidiVisual.val
            bidiVisual = obj.tblPr.bidiVisual;
            if isequal(bidiVisual, [])   % Python: if bidiVisual is None (H3)
                value = [];
                return
            end
            value = bidiVisual.val;
        end
        function set.bidiVisual_val(obj, value)
            % Python (table.py 168-173): tblPr = self.tblPr;
            %   if value is None: tblPr._remove_bidiVisual()
            %   else: tblPr.get_or_add_bidiVisual().val = bool(value)
            % value is a WD_TABLE_DIRECTION member (int subclass): bool(value) is
            % value.value != 0 (LTR=0 -> False, RTL=1 -> True). H10 per-site
            % double(.value); H4 truthiness. Via CT_OnOff.val's True default
            % (D-delta-1): True removes @val (bare <w:bidiVisual/>); False -> "0".
            tblPr = obj.tblPr;
            if isequal(value, [])   % Python: if value is None (H3)
                tblPr.remove_bidiVisual_();
            else
                bv = tblPr.get_or_add_bidiVisual();
                bv.val = logical(double(value.value) ~= 0);
            end
        end

        % ---- col_count (read-only, table.py 175-178) ----
        function value = get.col_count(obj)
            % Python: return len(self.tblGrid.gridCol_lst)
            value = numel(obj.tblGrid.gridCol_lst);
        end

        % ---- iter_tcs (table.py 180-189) ----
        function tcs = iter_tcs(obj)
            % ITER_TCS Each <w:tc> in this table, left-to-right and top-to-bottom
            %   (each cell of row 1, then row 2, ...), as a materialized 1xN
            %   heterogeneous XmlElement array of CT_Tc. Python is a GENERATOR (H9):
            %   callers iterate with no tree mutation, so laziness is unobservable.
            %   Python: for tr in self.tr_lst: for tc in tr.tc_lst: yield tc
            tcs = mat2doc.oxml.XmlElement.empty(1, 0);
            trs = obj.tr_lst;
            for i = 1:numel(trs)
                tcs = [tcs, trs(i).tc_lst]; %#ok<AGROW>
            end
        end

        % ---- tblStyle_val (get+set, table.py 199-217) ----
        function value = get.tblStyle_val(obj)
            % Python: tblStyle = self.tblPr.tblStyle; return None if None else tblStyle.val
            tblStyle = obj.tblPr.tblStyle;
            if isequal(tblStyle, [])   % Python: if tblStyle is None (H3)
                value = [];
                return
            end
            value = tblStyle.val;
        end
        function set.tblStyle_val(obj, styleId)
            % Python (table.py 213-217): tblPr = self.tblPr; tblPr._remove_tblStyle();
            %   if styleId is None: return; tblPr._add_tblStyle().val = styleId
            % The _remove_tblStyle() precedes the None short-circuit (order kept);
            % uses the PRIVATE _add adder (add_tblStyle_), not get_or_add.
            tblPr = obj.tblPr;
            tblPr.remove_tblStyle_();
            if isequal(styleId, [])   % Python: if styleId is None (H3)
                return
            end
            ts = tblPr.add_tblStyle_();
            ts.val = styleId;
        end
    end

    methods (Static)
        function elm = new_tbl(rows, cols, width)
            % NEW_TBL A new <w:tbl> with `rows` rows and `cols` columns, `width`
            %   distributed evenly across the columns. THE table constructor
            %   (table.py 191-197): return parse_xml(cls._tbl_xml(rows, cols, width)).
            %   Returns a CT_Tbl (w:tbl registered by this WP).
            elm = mat2doc.oxml.parse_xml( ...
                mat2doc.oxml.table.CT_Tbl.tbl_xml_(rows, cols, width));
        end
    end

    methods (Static, Access = private)
        function s = tbl_xml_(rows, cols, width)
            % TBL_XML_ VERBATIM port of CT_Tbl._tbl_xml (table.py 219-233). Emits
            %   the initial <w:tbl> XML string (byte-identical to python-docx before
            %   parse). col_width = Emu(width // cols) if cols > 0 else Emu(0) (H6:
            %   Python // -> floor).
            if cols > 0
                col_width = mat2doc.shared.Emu(floor(double(width) / cols));  % H6: // -> floor
            else
                col_width = mat2doc.shared.Emu(0);
            end
            LF = string(newline);
            s = "<w:tbl " + mat2doc.oxml.nsdecls("w") + ">" + LF + ...
                "  <w:tblPr>" + LF + ...
                "    <w:tblW w:type=""auto"" w:w=""0""/>" + LF + ...
                "    <w:tblLook w:firstColumn=""1"" w:firstRow=""1""" + LF + ...
                "               w:lastColumn=""0"" w:lastRow=""0"" w:noHBand=""0""" + LF + ...
                "               w:noVBand=""1"" w:val=""04A0""/>" + LF + ...
                "  </w:tblPr>" + LF + ...
                mat2doc.oxml.table.CT_Tbl.tblGrid_xml_(cols, col_width) + ...
                mat2doc.oxml.table.CT_Tbl.trs_xml_(rows, cols, col_width) + ...
                "</w:tbl>" + LF;
        end

        function s = tblGrid_xml_(col_count, col_width)
            % TBLGRID_XML_ VERBATIM port of CT_Tbl._tblGrid_xml (table.py 235-241).
            %   Python: xml = "  <w:tblGrid>\n"; for _ in range(col_count):
            %   xml += '    <w:gridCol w:w="%d"/>\n' % col_width.twips; ... .
            LF = string(newline);
            tw = mat2doc.shared.pyStr(col_width.twips, "int");   % H14 Python int '%d'
            s = "  <w:tblGrid>" + LF;
            for k = 1:col_count
                s = s + "    <w:gridCol w:w=""" + tw + """/>" + LF;
            end
            s = s + "  </w:tblGrid>" + LF;
        end

        function s = trs_xml_(row_count, col_count, col_width)
            % TRS_XML_ VERBATIM port of CT_Tbl._trs_xml (table.py 243-245):
            %   f"  <w:tr>\n{_tcs_xml(col_count, col_width)}  </w:tr>\n" * row_count.
            LF = string(newline);
            one = "  <w:tr>" + LF + ...
                mat2doc.oxml.table.CT_Tbl.tcs_xml_(col_count, col_width) + ...
                "  </w:tr>" + LF;
            s = "";
            for k = 1:row_count      % Python string * row_count
                s = s + one;
            end
        end

        function s = tcs_xml_(col_count, col_width)
            % TCS_XML_ VERBATIM port of CT_Tbl._tcs_xml (table.py 247-256): the
            %   <w:tc><w:tcPr><w:tcW .../></w:tcPr><w:p/></w:tc> block * col_count.
            LF = string(newline);
            tw = mat2doc.shared.pyStr(col_width.twips, "int");   % H14 Python int
            one = "    <w:tc>" + LF + ...
                "      <w:tcPr>" + LF + ...
                "        <w:tcW w:type=""dxa"" w:w=""" + tw + """/>" + LF + ...
                "      </w:tcPr>" + LF + ...
                "      <w:p/>" + LF + ...
                "    </w:tc>" + LF;
            s = "";
            for k = 1:col_count      % Python block * col_count
                s = s + one;
            end
        end
    end
end

classdef CT_Row < mat2doc.oxml.BaseOxmlElement
% CT_ROW `<w:tr>` element, a table row. Registered for w:tr (oxml/__init__.py:182).
%
%   Holds the row's optional <w:tblPrEx> (property exceptions) and <w:trPr> (row
%   properties) plus its <w:tc> cells. Row-level accessors delegate to <w:trPr>
%   (grid skip counts, row height); cell walking (tc_at_grid_offset) and new-cell
%   creation (_new_tc) depend on CT_Tc, which lands at P6-3a -- see the CT_Tc
%   BOUNDARY note below.
%
%   ===================== DESCRIPTORS (table.py 58-61) ===========================
%     tblPrEx = ZeroOrOne("w:tblPrEx")   # -- CUSTOM inserter _insert_tblPrEx --
%     trPr    = ZeroOrOne("w:trPr")       # -- CUSTOM inserter _insert_trPr --
%     tc      = ZeroOrMore("w:tc")
%   The Callable annotations (table.py 52-56: add_tc / get_or_add_trPr /
%   _add_trPr / tc_lst) are type hints for these generated members.
%
%   H11 -- CUSTOM INSERTERS (table.py 132-140). tblPrEx and trPr do NOT use the
%   generic successor-slice engine; each overrides _insert_x (like CT_P/CT_R
%   force pPr/rPr to the front). So get_or_add_x / add_x_ route through the OWN
%   insert_x_ override (NOT the generic addChild):
%     _insert_tblPrEx: self.insert(0, tblPrEx)                 -> obj.insert(1,...) (H1)
%     _insert_trPr:    if self.tblPrEx is not None:
%                          self.tblPrEx.addnext(trPr)
%                      else:
%                          self.insert(0, trPr)                -> obj.insert(1,...) (H1)
%   i.e. tblPrEx is forced to index 0; trPr goes immediately AFTER tblPrEx when
%   present, else to index 0. A wrong placement -> Word repair / byte divergence.
%   tc is ZeroOrMore with successors=() -> NO_SUCCESSORS (append at end) -- BUT
%   its creator is the _new_tc OVERRIDE (see CT_Tc BOUNDARY).
%
%   =============== CT_Tc BOUNDARY -- UN-STUBBED at P6-3a =========================
%   CT_Tc (w:tc) is now ported/registered (P6-3a), so <w:tc> children resolve to
%   CT_Tc and the two previously-deferred members run LIVE:
%     * _new_tc (table.py 142-143) = `return CT_Tc.new()`. new_tc_ now delegates to
%       mat2doc.oxml.table.CT_Tc.new(). The tc ZeroOrMore adders (add_tc_ / add_tc)
%       route through new_tc_ (the _new_tc override WINS over the generic creator,
%       xmlchemy _add_to_class), so they too now create real CT_Tc cells --
%       FAITHFUL: Python's add_tc also goes through _new_tc -> CT_Tc.new().
%     * tc_at_grid_offset (table.py 79-98) walks tc_lst summing `tc.grid_span`
%       (a CT_Tc accessor, now live). The P6-2 isa notYetPorted guard is removed;
%       the walk runs verbatim. (The four by-design P6-2 stub-battery throws in
%       Test_p6_2_table_props::test_ct_row_tc_boundary flip LIVE at this WP.)
%
%   ===================== @property members (table.py 63-130) ====================
%   grid_after / grid_before (read-only): 0 when trPr absent, else trPr.grid_after
%   / grid_before (H3). tr_idx: 0-based index of this <w:tr> among its parent's
%   <w:tr> siblings (H1: Python list.index() is 0-based DATA -> find()-1).
%   trHeight_hRule / trHeight_val (get+set): [] when trPr absent (get); set uses
%   get_or_add_trPr then delegates to CT_TrPr.
%
%   tr_idx CT_Tbl NOTE: Python is `tbl = cast(CT_Tbl, self.getparent());
%   tbl.tr_lst.index(self)`. CT_Tbl (P6-3) is not ported, so getparent() yields a
%   GENERIC XmlElement with no tr_lst. CT_Tbl.tr_lst is EXACTLY findall(qn("w:tr"))
%   (a ZeroOrMore list-getter == getChildList), so the port reads the parent's
%   w:tr children directly via findall -- behaviorally identical, tag-based, needs
%   NO CT_Tbl and NO stub (works now and after CT_Tbl registers). H5 identity.
%
%   M1-NEUTRAL: default.docx contains ZERO <w:tr> elements, so nothing transits
%   this class on the M1 parse path; registering w:tr -> CT_Row is byte-neutral.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on each <w:tr> inside a real table.
%
%   Example:
%       tr = mat2doc.oxml.OxmlElement("w:tr");   % a CT_Row
%       tr.grid_before                           % 0 (no <w:trPr>/<w:gridBefore>)
%       tr.trHeight_val = mat2doc.shared.Twips(360);  % <w:trPr><w:trHeight .../></w:trPr>
%
%   Ported from python-docx v1.2.0: src/docx/oxml/table.py::CT_Row
%   (lines 49-143; registered for w:tr)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        NO_SUCCESSORS = string.empty(1, 0)   % tc: successors=() -> append at end
    end

    properties (Dependent)  % generated ZeroOrOne getters + ZeroOrMore list + @property
        tblPrEx        % ZeroOrOne <w:tblPrEx> child or [] (CUSTOM inserter)
        trPr           % ZeroOrOne <w:trPr> child or [] (CUSTOM inserter)
        tc_lst         % ZeroOrMore list of <w:tc> children (document order; generic until P6-3a)
        grid_after     % @property: 0 when trPr absent, else trPr.grid_after
        grid_before    % @property: 0 when trPr absent, else trPr.grid_before
        tr_idx         % @property: 0-based index of this <w:tr> among sibling <w:tr>
        trHeight_hRule % @property: trPr.trHeight_hRule (WD_ROW_HEIGHT_RULE) or []
        trHeight_val   % @property: trPr.trHeight_val (Length) or []
    end

    methods
        function obj = CT_Row(varargin)
            % CT_ROW Construct a loose <w:tr> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ============ tblPrEx (ZeroOrOne, CUSTOM inserter: insert(0,...)) ============
        function child = get.tblPrEx(obj);            child = obj.getChild("w:tblPrEx"); end
        function child = get_or_add_tblPrEx(obj)
            % get_or_add (xmlchemy 557-562) routed through the OVERRIDE inserter.
            child = obj.tblPrEx;
            if isequal(child, [])   % Python: if child is None (H3)
                child = obj.add_tblPrEx_();
            end
        end
        function child = new_tblPrEx_(obj);           child = obj.newChild("w:tblPrEx"); end
        function child = insert_tblPrEx_(obj, tblPrEx)
            % OVERRIDE (table.py 132-133): self.insert(0, tblPrEx).
            obj.insert(1, tblPrEx);    % Python insert(0,...) -> 1-based (H1)
            child = tblPrEx;
        end
        function child = add_tblPrEx_(obj, varargin)
            % _add_tblPrEx (xmlchemy 284-291) routed through the OVERRIDE inserter.
            child = obj.new_tblPrEx_();
            for k = 1:2:numel(varargin)
                child.(varargin{k}) = varargin{k + 1};
            end
            child = obj.insert_tblPrEx_(child);
        end
        function remove_tblPrEx_(obj);                obj.removeChild("w:tblPrEx"); end

        % ============ trPr (ZeroOrOne, CUSTOM inserter: after tblPrEx else insert(0)) ============
        function child = get.trPr(obj);            child = obj.getChild("w:trPr"); end
        function child = get_or_add_trPr(obj)
            % get_or_add (xmlchemy 557-562) routed through the OVERRIDE inserter.
            child = obj.trPr;
            if isequal(child, [])   % Python: if child is None (H3)
                child = obj.add_trPr_();
            end
        end
        function child = new_trPr_(obj);           child = obj.newChild("w:trPr"); end
        function child = insert_trPr_(obj, trPr)
            % OVERRIDE (table.py 135-140): tblPrEx = self.tblPrEx;
            %   if tblPrEx is not None: tblPrEx.addnext(trPr)
            %   else: self.insert(0, trPr)
            tblPrEx = obj.tblPrEx;
            if ~isequal(tblPrEx, [])   % Python: if tblPrEx is not None (H3)
                tblPrEx.addnext(trPr);
            else
                obj.insert(1, trPr);   % Python insert(0,...) -> 1-based (H1)
            end
            child = trPr;
        end
        function child = add_trPr_(obj, varargin)
            % _add_trPr (xmlchemy 284-291) routed through the OVERRIDE inserter.
            child = obj.new_trPr_();
            for k = 1:2:numel(varargin)
                child.(varargin{k}) = varargin{k + 1};
            end
            child = obj.insert_trPr_(child);
        end
        function remove_trPr_(obj);                obj.removeChild("w:trPr"); end

        % ============ tc (ZeroOrMore, successors=(); creator is the _new_tc OVERRIDE) ============
        function lst = get.tc_lst(obj);           lst = obj.getChildList("w:tc"); end
        function child = new_tc_(obj) %#ok<MANU>
            % OVERRIDE (table.py 142-143): return CT_Tc.new(). UN-STUBBED at P6-3a
            % (CT_Tc + w:tc registration are now live). add_tc_ / add_tc route
            % through this override (the _new_tc creator WINS over the generic
            % adder, xmlchemy _add_to_class), exactly as Python's add_tc reaches
            % CT_Tc.new() via _new_tc.
            child = mat2doc.oxml.table.CT_Tc.new();
        end
        function child = insert_tc_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_tc_(obj, varargin)
            % _add_tc routed through the _new_tc OVERRIDE -> raises notYetPorted
            % (P6-3a) before any insert, exactly as Python's add_tc reaches
            % CT_Tc.new() via _new_tc.
            child = obj.new_tc_();
            for k = 1:2:numel(varargin)
                child.(varargin{k}) = varargin{k + 1}; %#ok<AGROW>
            end
            child = obj.insert_tc_(child);
        end
        function child = add_tc(obj);            child = obj.add_tc_(); end   % public adder (D-delta-4)

        % ===================== @property members (table.py 63-130) ====================

        % ---- grid_after (read-only, table.py 63-69) ----
        function value = get.grid_after(obj)
            % Python: trPr = self.trPr; if trPr is None: return 0; return trPr.grid_after
            trPr = obj.trPr;
            if isequal(trPr, [])   % Python: if trPr is None (H3)
                value = 0;
                return
            end
            value = trPr.grid_after;
        end

        % ---- grid_before (read-only, table.py 71-77) ----
        function value = get.grid_before(obj)
            % Python: trPr = self.trPr; if trPr is None: return 0; return trPr.grid_before
            trPr = obj.trPr;
            if isequal(trPr, [])   % Python: if trPr is None (H3)
                value = 0;
                return
            end
            value = trPr.grid_before;
        end

        % ---- tc_at_grid_offset (table.py 79-98) -- UN-STUBBED at P6-3a ----
        function tc = tc_at_grid_offset(obj, grid_offset)
            % TC_AT_GRID_OFFSET The `w:tc` in this row starting at exact grid_offset.
            %   Ported VERBATIM from table.py 79-98. UN-STUBBED at P6-3a: `tc` is now
            %   a CT_Tc (w:tc registered) so `tc.grid_span` runs -- the P6-2 isa
            %   notYetPorted guard is removed and the grid_span walk runs
            %   unchanged. H1: grid_offset is 0-based GRID data, kept RAW.
            arguments
                obj (1,1) mat2doc.oxml.table.CT_Row
                grid_offset (1,1) double
            end
            % Python: remaining_offset = grid_offset - self.grid_before
            remaining_offset = grid_offset - obj.grid_before;
            tcs = obj.tc_lst;
            for i = 1:numel(tcs)
                tc = tcs(i);
                if remaining_offset < 0   % Python: gone past grid_offset; stop
                    break
                end
                if remaining_offset == 0   % Python: arrived; this is the tc
                    return
                end
                % Python: remaining_offset -= tc.grid_span  (CT_Tc accessor, now live)
                remaining_offset = remaining_offset - tc.grid_span;
            end
            % Python: raise ValueError(f"no `tc` element at grid_offset={grid_offset}")
            error("mat2doc:ValueError", "no `tc` element at grid_offset=%s", ...
                mat2doc.shared.pyStr(grid_offset, "int"));
        end

        % ---- tr_idx (read-only, table.py 100-104) -- findall, no CT_Tbl dep ----
        function value = get.tr_idx(obj)
            % Python: tbl = cast(CT_Tbl, self.getparent()); return tbl.tr_lst.index(self)
            % CT_Tbl.tr_lst == findall(qn("w:tr")); read the parent's w:tr children
            % directly (tag-based, no CT_Tbl needed). H5 identity; H1 0-based.
            tbl = obj.getparent();
            trs = tbl.findall(mat2doc.oxml.qn("w:tr"));
            i = find(trs == obj, 1);
            if isempty(i)   % Python list.index() ValueError (unreachable for a parented row)
                error("mat2doc:ValueError", "%s is not in list", "tr");
            end
            value = i - 1;   % IDX: 0-based to match Python list.index()
        end

        % ---- trHeight_hRule (get+set, table.py 106-117) ----
        function value = get.trHeight_hRule(obj)
            % Python: trPr = self.trPr; if trPr is None: return None; return trPr.trHeight_hRule
            trPr = obj.trPr;
            if isequal(trPr, [])   % Python: if trPr is None (H3)
                value = [];
                return
            end
            value = trPr.trHeight_hRule;
        end
        function set.trHeight_hRule(obj, value)
            % Python (table.py 115-117): trPr = self.get_or_add_trPr();
            %   trPr.trHeight_hRule = value
            trPr = obj.get_or_add_trPr();
            trPr.trHeight_hRule = value;
        end

        % ---- trHeight_val (get+set, table.py 119-130) ----
        function value = get.trHeight_val(obj)
            % Python: trPr = self.trPr; if trPr is None: return None; return trPr.trHeight_val
            trPr = obj.trPr;
            if isequal(trPr, [])   % Python: if trPr is None (H3)
                value = [];
                return
            end
            value = trPr.trHeight_val;
        end
        function set.trHeight_val(obj, value)
            % Python (table.py 128-130): trPr = self.get_or_add_trPr();
            %   trPr.trHeight_val = value
            trPr = obj.get_or_add_trPr();
            trPr.trHeight_val = value;
        end
    end
end

classdef CT_Tc < mat2doc.oxml.BaseOxmlElement
% CT_TC `<w:tc>` table-cell element -- the table-cell MERGE engine.
%
%   Registered for w:tc (oxml/__init__.py:179). Holds the cell's optional
%   <w:tcPr> (cell properties) plus its block-level content (<w:p>/<w:tbl>/
%   <w:sdt>). This is the hardest WP of the Mat2Doc port: besides the tcPr
%   delegators (grid_span/vMerge/width) and the grid-geometry read tier
%   (grid_offset/left/right/top/bottom), it ports the destructive, byte-visible
%   MERGE ENGINE (merge/_span_dimensions/_grow_to/_span_to_width/_swallow_next_tc
%   + the content-move + row/cell navigation).
%
%   ============================ DESCRIPTORS (table.py 432-434) ==================
%     tcPr = ZeroOrOne("w:tcPr")   # -- CUSTOM inserter _insert_tcPr = insert(0,...)
%     p    = OneOrMore("w:p")
%     tbl  = OneOrMore("w:tbl")    # -- creator OVERRIDDEN by _new_tbl (raises)
%   The Callable annotations (table.py 424-429: add_p / get_or_add_tcPr / p_lst /
%   tbl_lst / _insert_tbl / _new_p) are type hints for these generated members.
%
%   H11 -- tcPr CUSTOM inserter (table.py 598-603): tcPr has MANY successors but
%   always comes FIRST, so _insert_tcPr just does `self.insert(0, tcPr)` rather
%   than spelling out successors. So get_or_add_tcPr / add_tcPr_ route through the
%   OWN insert_tcPr_ override (NOT the generic addChild): obj.insert(1, tcPr) --
%   Python insert(0,...) is 1-based here (H1). tcPr is thus ALWAYS the first child
%   of the tc, before all content. A wrong placement -> Word repair / byte
%   divergence. p / tbl are OneOrMore with successors=() (append at end); tbl's
%   creator is the _new_tbl OVERRIDE (raises NotImplementedError -- use
%   CT_Tbl.new_tbl()).
%
%   =============== CT_Tbl BOUNDARY -- the generic-ancestor tr_lst shim (C2) =====
%   CT_Tc needs `self._tbl.tr_lst[...]` (the enclosing tbl's rows: merge :515,
%   _tr_below :771-776, _tr_idx :779-781). CT_Tbl (w:tbl) is P6-3b -- NOT yet
%   ported/registered -- so `_tbl` (table.py 733-736, `./ancestor::w:tbl
%   [position()=1]`) resolves to a GENERIC XmlElement with NO tr_lst accessor.
%   RESOLUTION (planaudit_2026-07-31 condition C2): CT_Tbl.tr_lst is EXACTLY
%   findall("w:tr") in document order, and w:tr already dispatches to CT_Row
%   (P6-2). So the private shim `trLstOfTbl_` returns the tbl ancestor's
%   `xpath("./w:tr")` -- the identical CT_Row handle list. Byte-neutral,
%   identity-safe (persistent, parent_-linked handles). ADJUDICATED/UPGRADED at
%   P6-3b: when CT_Tbl registers, either swap to `tbl.tr_lst` or keep this shim
%   (decided at P6-3b Gate-2). This is the ONE Tc->Tbl seam; `_tbl` itself is the
%   plain generic ancestor element (used only to reach the rows).
%
%   ============================ H1 (0-based -> 1-based) landmines ===============
%   Every index below is 0-based Python data mapped explicitly to 1-based MATLAB;
%   each site is commented `IDX`:
%     * _tr_idx = tr_lst.index(self._tr) -> find()-1 (0-based row index).
%     * bottom  = _tr_idx + 1 (an EXCLUSIVE slice bound, NOT an index -- +1 ONCE).
%     * merge: tr_lst[top] -> trLst(top+1) (0-based top -> 1-based).
%     * _tr_below: tr_lst[tr_idx+1] with IndexError->None -> guard `i+1 <= numel`
%       then trLst(i+1) (the single most error-prone line: the +1 THEN the bound).
%     * _grow_to recursion height-1; _remove_trailing_empty_p block_items[-1]->end;
%       _is_empty block_items[0]->(1).
%     * grid_offset / left / right / tc_at_grid_offset(left): 0-based GRID data,
%       kept RAW (P6-2 grid convention is 0-based; never shift).
%
%   ============================ H5 (handle identity) ===========================
%   `tr_lst.index(self._tr)` (_tr_idx), `top_tc is not self` (_grow_to closure),
%   `other_tc is self` (_move_content_to), `next_tc is X` all use HANDLE identity
%   (`==`/`~=` / find-by-eq), NEVER isequal on content: on a UNIFORM grid (every
%   row/cell byte-identical) an isequal-based match silently returns row 1 and
%   corrupts the whole walk while passing every non-uniform fixture. Probe:
%   _tr_idx on a 3x3 uniform-row snippet -> rows 0/1/2 return 0/1/2.
%
%   ============================ H4 (truthiness) ================================
%   _add_width_of gates on `if self.width and other_tc.width` -- a 0-EMU width is
%   FALSY in Python (skips the add), and None is falsy. The port replicates
%   falsy-zero: skip when a width is [] (None) OR its EMU value == 0.
%
%   ============================ bare <w:vMerge/> (NEW-D risk) ==================
%   A vertical-merge CONTINUATION cell must serialize as a BARE <w:vMerge/> (no
%   @w:val), the top cell as <w:vMerge w:val="restart"/>. The port never writes
%   "continue": _grow_to's closure assigns vMerge="continue" to continuation cells
%   and CT_TcPr.vMerge_val -> CT_VMerge.val (OptionalAttribute default "continue",
%   D-delta-1) DELETES @w:val, producing the bare element. See CT_TcPr / CT_VMerge.
%
%   M1 NON-NEUTRAL (styles.xml byte-critical): registering w:tcPr puts the 595
%   <w:tcPr> nodes in word/styles.xml (+595 in stylesWithEffects.xml) on the parse
%   path. Byte-neutral EXPECTED (descriptors-only) but PROVEN at Gate-1 (M1 17/17,
%   styles.xml SHA). w:tc / w:gridSpan have 0 occurrences in default.docx -> those
%   two rows are M1-neutral.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on every <w:tc> inside a real table.
%
%   Example:
%       tc = mat2doc.oxml.table.CT_Tc.new();   % <w:tc><w:p/></w:tc>
%       tc.grid_span                            % 1  (no <w:tcPr>/<w:gridSpan>)
%       % merge: a.merge(b) returns the top-left cell of the new rectangular span
%
%   Ported from python-docx v1.2.0: src/docx/oxml/table.py::CT_Tc
%   (lines 421-783; registered for w:tc)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        NO_SUCCESSORS = string.empty(1, 0)   % p / tbl: OneOrMore successors=() -> append at end
    end

    properties (Dependent)  % generated ZeroOrOne/OneOrMore getters + @property members
        tcPr                    % ZeroOrOne <w:tcPr> child or [] (CUSTOM inserter insert(0,...))
        p_lst                   % OneOrMore list of <w:p> children (document order)
        tbl_lst                 % OneOrMore list of <w:tbl> children (document order)
        grid_offset             % @property: 0-based starting layout-grid column of this cell
        grid_span               % @property (get+set): columns this cell spans (1 default)
        inner_content_elements  % @property: ./w:p | ./w:tbl in document order
        left                    % @property: grid column index at which this cell starts (= grid_offset)
        right                   % @property: exclusive right-side grid column index (grid_offset+grid_span)
        vMerge                  % @property (get+set): ./w:tcPr/w:vMerge/@w:val (string) or []
        width                   % @property (get+set): EMU Length in ./w:tcPr/w:tcW or []
    end
    % NOTE: `top` and `bottom` are the two RECURSIVE @property members (top ->
    % _tc_above.top; bottom -> _tc_below.bottom). MATLAB forbids ANY textual
    % reference to a Dependent property's own name inside its getter (even on
    % another object of the class), so they are ported as ZERO-ARG METHODS
    % (`tc.top` / `tc.bottom` still read exactly like a property -- MATLAB allows
    % method-call-without-parens). All other @property members stay Dependent.

    methods
        function obj = CT_Tc(varargin)
            % CT_TC Construct a loose <w:tc> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ============ tcPr (ZeroOrOne, CUSTOM inserter: insert(0,...)) ============
        function child = get.tcPr(obj);            child = obj.getChild("w:tcPr"); end
        function child = get_or_add_tcPr(obj)
            % get_or_add (xmlchemy 557-562) routed through the OVERRIDE inserter.
            child = obj.tcPr;
            if isequal(child, [])   % Python: if child is None (H3)
                child = obj.add_tcPr_();
            end
        end
        function child = new_tcPr_(obj);           child = obj.newChild("w:tcPr"); end
        function child = insert_tcPr_(obj, tcPr)
            % OVERRIDE (table.py 598-603): self.insert(0, tcPr); return tcPr.
            obj.insert(1, tcPr);    % Python insert(0,...) -> 1-based (H1); tcPr FIRST child
            child = tcPr;
        end
        function child = add_tcPr_(obj, varargin)
            % _add_tcPr (xmlchemy 284-291) routed through the OVERRIDE inserter.
            child = obj.new_tcPr_();
            for k = 1:2:numel(varargin)
                child.(varargin{k}) = varargin{k + 1};
            end
            child = obj.insert_tcPr_(child);
        end
        function remove_tcPr_(obj);                obj.removeChild("w:tcPr"); end

        % ============ p (OneOrMore, successors=() -> append at end) ============
        function lst = get.p_lst(obj);            lst = obj.getChildList("w:p"); end
        function child = new_p_(obj);             child = obj.newChild("w:p"); end   % default creator (table.py 429)
        function child = insert_p_(obj, child);   child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_p_(obj, varargin);   child = obj.addChild("w:p", obj.NO_SUCCESSORS, varargin{:}); end
        function child = add_p(obj);              child = obj.add_p_(); end   % public adder (OneOrMore)

        % ============ tbl (OneOrMore, creator OVERRIDDEN by _new_tbl) ============
        function lst = get.tbl_lst(obj);          lst = obj.getChildList("w:tbl"); end
        function child = new_tbl_(obj) %#ok<STOUT,MANU>
            % OVERRIDE (table.py 632-635): raise NotImplementedError. FAITHFUL --
            % python-docx itself refuses to create a bare <w:tbl> here (a table
            % needs rows/cols); use CT_Tbl.new_tbl() instead.
            error("mat2doc:NotImplementedError", "%s", ...
                "use CT_Tbl.new_tbl() to add a new table, specifying rows and columns");
        end
        function child = insert_tbl_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_tbl_(obj, varargin)
            % _add_tbl routed through the _new_tbl OVERRIDE -> raises
            % NotImplementedError before any insert, exactly as Python's add_tbl
            % reaches _new_tbl.
            child = obj.new_tbl_();
            for k = 1:2:numel(varargin)
                child.(varargin{k}) = varargin{k + 1}; %#ok<AGROW>
            end
            child = obj.insert_tbl_(child);
        end
        function child = add_tbl(obj);            child = obj.add_tbl_(); end   % public adder (OneOrMore)

        % ===================== @property read tier (table.py 436-560) ===============

        % ---- bottom (read-only @property, table.py 436-447; MATLAB METHOD -- recursive) ----
        function value = bottom(obj)
            % Python: if self.vMerge is not None:
            %             tc_below = self._tc_below
            %             if tc_below is not None and tc_below.vMerge == ST_Merge.CONTINUE:
            %                 return tc_below.bottom
            %         return self._tr_idx + 1
            if ~isequal(obj.vMerge, [])   % self.vMerge is not None (H3)
                tc_below = obj.tc_below_();
                % tc_below is not None AND tc_below.vMerge == "continue" (H5 handle
                % nav; isequal on the STRING value -- None/"restart"/"continue")
                if ~isequal(tc_below, []) && isequal(tc_below.vMerge, "continue")
                    value = tc_below.bottom;   % recurse (method, not property)
                    return
                end
            end
            value = obj.tr_idx_() + 1;   % IDX: EXCLUSIVE bound (tr_idx is 0-based); +1 ONCE
        end

        % ---- clear_content (table.py 449-458) ----
        function clear_content(obj)
            % CLEAR_CONTENT Remove all content elements, preserving <w:tcPr>.
            %   Leaves the tc in an invalid state (no block-level child) -- the
            %   caller must add a <w:p>. Python: for e in self.xpath(
            %   "./*[not(self::w:tcPr)]"): self.remove(e).
            elems = obj.xpath("./*[not(self::w:tcPr)]");
            for k = 1:numel(elems)
                obj.remove(elems(k));
            end
        end

        % ---- grid_offset (read-only, table.py 460-470) ----
        function value = get.grid_offset(obj)
            % Python: grid_before = self._tr.grid_before
            %         preceding = sum(tc.grid_span for tc in
            %                         self.xpath("./preceding-sibling::w:tc"))
            %         return grid_before + preceding
            grid_before = obj.tr_().grid_before;
            preceding = obj.xpath("./preceding-sibling::w:tc");   % CT_Tc siblings
            s = 0;
            for k = 1:numel(preceding)
                s = s + preceding(k).grid_span;
            end
            value = grid_before + s;   % 0-based grid data, kept RAW (H1)
        end

        % ---- grid_span (get+set, table.py 472-484) ----
        function value = get.grid_span(obj)
            % Python: tcPr = self.tcPr; return 1 if tcPr is None else tcPr.grid_span
            tcPr = obj.tcPr;
            if isequal(tcPr, [])   % Python: if tcPr is None (H3)
                value = 1;
                return
            end
            value = tcPr.grid_span;
        end
        function set.grid_span(obj, value)
            % Python (table.py 482-484): tcPr = self.get_or_add_tcPr();
            %   tcPr.grid_span = value
            tcPr = obj.get_or_add_tcPr();
            tcPr.grid_span = value;
        end

        % ---- inner_content_elements (read-only, table.py 486-493) ----
        function value = get.inner_content_elements(obj)
            % Python: return self.xpath("./w:p | ./w:tbl")
            value = obj.xpath("./w:p | ./w:tbl");
        end

        % ---- iter_block_items (table.py 495-501) ----
        function items = iter_block_items(obj)
            % ITER_BLOCK_ITEMS Block-level content children (w:p / w:tbl / w:sdt),
            %   in document order. Python is a GENERATOR (H9): callers materialize
            %   with list(...); this returns a (1,N) snapshot array. lxml's child
            %   iterator captures the NEXT-sibling pointer BEFORE yielding, so a
            %   snapshot taken up front is faithful even when the caller MOVES each
            %   item out during the loop (_move_content_to) -- see that method.
            tags = [mat2doc.oxml.qn("w:p"), mat2doc.oxml.qn("w:tbl"), mat2doc.oxml.qn("w:sdt")];
            kids = obj.to_array();
            keep = false(1, numel(kids));
            for i = 1:numel(kids)
                keep(i) = any(kids(i).tag == tags);
            end
            items = kids(keep);
        end

        % ---- left (read-only, table.py 503-506) ----
        function value = get.left(obj)
            % Python: return self.grid_offset
            value = obj.grid_offset;
        end

        % ---- merge (table.py 508-517) -- THE public merge entry point ----
        function top_tc = merge(obj, other_tc)
            % MERGE Return the top-left <w:tc> of a new rectangular span formed by
            %   merging the region with this tc and other_tc as diagonal corners.
            %   Python: top, left, height, width = self._span_dimensions(other_tc)
            %           top_tc = self._tbl.tr_lst[top].tc_at_grid_offset(left)
            %           top_tc._grow_to(width, height); return top_tc
            arguments
                obj (1,1) mat2doc.oxml.table.CT_Tc
                other_tc (1,1) mat2doc.oxml.table.CT_Tc
            end
            [top, left, height, width] = obj.span_dimensions_(other_tc);
            trLst = obj.trLstOfTbl_();                         % CT_Tbl.tr_lst shim (C2)
            top_tc = trLst(top + 1).tc_at_grid_offset(left);   % IDX: tr_lst[top]->(top+1); left RAW 0-based
            top_tc.grow_to_(width, height);
        end

        % ---- right (read-only, table.py 524-532) ----
        function value = get.right(obj)
            % Python: return self.grid_offset + self.grid_span
            value = obj.grid_offset + obj.grid_span;
        end

        % ---- top (read-only @property, table.py 534-539; MATLAB METHOD -- recursive) ----
        function value = top(obj)
            % Python: if self.vMerge is None or self.vMerge == ST_Merge.RESTART:
            %             return self._tr_idx
            %         return self._tc_above.top
            vm = obj.vMerge;
            if isequal(vm, []) || isequal(vm, "restart")   % None or RESTART (H3)
                value = obj.tr_idx_();
                return
            end
            above = obj.tc_above_();
            value = above.top;   % recurse (method, not property)
        end

        % ---- vMerge (get+set, table.py 541-552) ----
        function value = get.vMerge(obj)
            % Python: tcPr = self.tcPr; return None if None else tcPr.vMerge_val
            tcPr = obj.tcPr;
            if isequal(tcPr, [])   % Python: if tcPr is None (H3)
                value = [];
                return
            end
            value = tcPr.vMerge_val;
        end
        function set.vMerge(obj, value)
            % Python (table.py 550-552): tcPr = self.get_or_add_tcPr();
            %   tcPr.vMerge_val = value
            tcPr = obj.get_or_add_tcPr();
            tcPr.vMerge_val = value;
        end

        % ---- width (get+set, table.py 554-565) ----
        function value = get.width(obj)
            % Python: tcPr = self.tcPr; return None if None else tcPr.width
            tcPr = obj.tcPr;
            if isequal(tcPr, [])   % Python: if tcPr is None (H3)
                value = [];
                return
            end
            value = tcPr.width;
        end
        function set.width(obj, value)
            % Python (table.py 563-565): tcPr = self.get_or_add_tcPr();
            %   tcPr.width = value
            tcPr = obj.get_or_add_tcPr();
            tcPr.width = value;
        end

        % ===================== MERGE ENGINE (write tier) =====================

        % ---- _add_width_of (table.py 567-573) ----
        function add_width_of_(obj, other_tc)
            % Python: if self.width and other_tc.width:
            %             self.width = Length(self.width + other_tc.width)
            % H4 falsy-zero: a width is truthy iff NOT None and its EMU value ~= 0
            % (Length subclasses int; Length(0) is falsy). Skip the add otherwise.
            sw = obj.width;
            ow = other_tc.width;
            if ~isequal(sw, []) && double(sw) ~= 0 && ~isequal(ow, []) && double(ow) ~= 0
                % self.width + other.width: Length(double subclass) arithmetic ->
                % plain double EMU sum; Length(...) re-wraps (H6). The width setter
                % re-emits it as dxa twips via CT_TblWidth.
                obj.width = mat2doc.shared.Length(double(sw) + double(ow));
            end
        end

        % ---- _grow_to (table.py 575-596) ----
        function grow_to_(obj, width, height, top_tc)
            % Python: def vMerge_val(top_tc):
            %             return (ST_Merge.CONTINUE if top_tc is not self
            %                     else None if height == 1 else ST_Merge.RESTART)
            %         top_tc = self if top_tc is None else top_tc
            %         self._span_to_width(width, top_tc, vMerge_val(top_tc))
            %         if height > 1:
            %             tc_below = self._tc_below; assert tc_below is not None
            %             tc_below._grow_to(width, height - 1, top_tc)
            arguments
                obj (1,1) mat2doc.oxml.table.CT_Tc
                width (1,1) double
                height (1,1) double
                top_tc = []   % Python default None
            end
            if isequal(top_tc, [])   % top_tc = self if top_tc is None else top_tc
                top_tc = obj;
            end
            % vMerge_val(top_tc) closure, evaluated on the RESOLVED top_tc (H5):
            if top_tc ~= obj          % top_tc is not self -> CONTINUE
                vm = "continue";
            elseif height == 1        % self, height 1 -> None (horizontal-only top)
                vm = [];
            else                      % self, height > 1 -> RESTART (vertical top)
                vm = "restart";
            end
            obj.span_to_width_(width, top_tc, vm);
            if height > 1
                tc_below = obj.tc_below_();
                assert(~isequal(tc_below, []));   % Python: assert tc_below is not None
                tc_below.grow_to_(width, height - 1, top_tc);   % IDX: height-1 recursion
            end
        end

        % ---- _is_empty (table.py 605-614) ----
        function tf = is_empty_(obj)
            % Python: block_items = list(self.iter_block_items())
            %         if len(block_items) > 1: return False
            %         only_item = block_items[0]
            %         return isinstance(only_item, CT_P) and len(only_item.r_lst) == 0
            block_items = obj.iter_block_items();
            if numel(block_items) > 1
                tf = false;
                return
            end
            only_item = block_items(1);   % IDX: block_items[0] -> (1)
            tf = isa(only_item, "mat2doc.oxml.text.CT_P") && numel(only_item.r_lst) == 0;
        end

        % ---- _move_content_to (table.py 616-630) ----
        function move_content_to_(obj, other_tc)
            % Python: if other_tc is self: return
            %         if self._is_empty: return
            %         other_tc._remove_trailing_empty_p()
            %         for block_element in self.iter_block_items():
            %             other_tc.append(block_element)     # append MOVES
            %         self.append(self._new_p())             # restore min <w:p>
            if other_tc == obj   % other_tc is self (H5 handle identity)
                return
            end
            if obj.is_empty_()
                return
            end
            other_tc.remove_trailing_empty_p_();
            % iter_block_items() is a snapshot taken BEFORE the moves; append MOVES
            % each element from self to other_tc (lxml move). Faithful to the
            % lazy Python generator (its cursor captured the next sibling before
            % each yield, so every original block item is moved) -- see
            % iter_block_items.
            items = obj.iter_block_items();
            for k = 1:numel(items)
                other_tc.append(items(k));
            end
            obj.append(obj.new_p_());   % add back the required minimum single empty <w:p>
        end

        % ---- _new_tbl (table.py 632-635): IS the tbl OneOrMore creator override
        % ---- new_tbl_ above (same symbol -- the generated creator for the `tbl`
        % ---- OneOrMore); no separate member. ----

        % ---- _next_tc (read-only, table.py 637-642) ----
        function nt = next_tc_(obj)
            % Python: following = self.xpath("./following-sibling::w:tc")
            %         return following[0] if following else None
            following = obj.xpath("./following-sibling::w:tc");
            if isempty(following)   % Python: if following else None (H3)
                nt = [];
            else
                nt = following(1);   % IDX: following[0] -> (1)
            end
        end

        % ---- _remove (table.py 644-648) ----
        function remove_(obj)
            % Python: parent = self.getparent(); assert parent is not None;
            %         parent.remove(self)
            parent_element = obj.getparent();
            assert(~isequal(parent_element, []));   % Python: assert parent is not None
            parent_element.remove(obj);
        end

        % ---- _remove_trailing_empty_p (table.py 650-659) ----
        function remove_trailing_empty_p_(obj)
            % Python: block_items = list(self.iter_block_items())
            %         last = block_items[-1]
            %         if not isinstance(last, CT_P): return
            %         if len(last.r_lst) > 0: return
            %         self.remove(last)
            block_items = obj.iter_block_items();
            last_content_elm = block_items(end);   % IDX: block_items[-1] -> (end)
            if ~isa(last_content_elm, "mat2doc.oxml.text.CT_P")
                return
            end
            p = last_content_elm;
            if numel(p.r_lst) > 0
                return
            end
            obj.remove(p);
        end

        % ---- _span_dimensions (table.py 661-689) ----
        function [top, left, height, width] = span_dimensions_(obj, other_tc)
            % Return (top, left, height, width) extents of the rectangular merged
            % cell using self and other_tc as opposite corners. All 0-based grid
            % math, kept RAW (H1). Raises mat2doc:InvalidSpanError for inverted-L
            % and tee-shaped (non-rectangular) requests (verbatim message).
            a = obj; b = other_tc;

            % raise_on_inverted_L(a, b) (table.py 666-670):
            if a.top == b.top && a.bottom ~= b.bottom
                mat2doc.exc.InvalidSpanError("requested span not rectangular");
            end
            if a.left == b.left && a.right ~= b.right
                mat2doc.exc.InvalidSpanError("requested span not rectangular");
            end

            % raise_on_tee_shaped(a, b) (table.py 672-679):
            if a.top < b.top      % top_most, other = (a, b) if a.top < b.top else (b, a)
                tm = a; ot = b;
            else
                tm = b; ot = a;
            end
            if tm.top < ot.top && tm.bottom > ot.bottom
                mat2doc.exc.InvalidSpanError("requested span not rectangular");
            end
            if a.left < b.left    % left_most, other = (a, b) if a.left < b.left else (b, a)
                lm = a; ot = b;
            else
                lm = b; ot = a;
            end
            if lm.left < ot.left && lm.right > ot.right
                mat2doc.exc.InvalidSpanError("requested span not rectangular");
            end

            % top,left,bottom,right = min/max of the two corners (table.py 684-689):
            top = min(a.top, b.top);
            left = min(a.left, b.left);
            bottom = max(a.bottom, b.bottom);
            right = max(a.right, b.right);
            height = bottom - top;   % 0-based extent math, RAW
            width = right - left;
        end

        % ---- _span_to_width (table.py 691-706) ----
        function span_to_width_(obj, grid_width, top_tc, vMerge)
            % Python: self._move_content_to(top_tc)
            %         while self.grid_span < grid_width:
            %             self._swallow_next_tc(grid_width, top_tc)
            %         self.vMerge = vMerge
            % Incorporates <w:tc> to the right until this cell spans grid_width;
            % swallowed cells are removed; content is appended to top_tc; the single
            % remaining cell's vMerge is set (None removes it).
            obj.move_content_to_(top_tc);
            while obj.grid_span < grid_width
                obj.swallow_next_tc_(grid_width, top_tc);
            end
            obj.vMerge = vMerge;
        end

        % ---- _swallow_next_tc (table.py 708-731) ----
        function swallow_next_tc_(obj, grid_width, top_tc)
            % Python: next_tc = self._next_tc
            %         raise_on_invalid_swallow(next_tc):
            %             if next_tc is None: raise InvalidSpanError("not enough grid columns")
            %             if self.grid_span + next_tc.grid_span > grid_width:
            %                 raise InvalidSpanError("span is not rectangular")
            %         assert next_tc is not None
            %         next_tc._move_content_to(top_tc)   # 1. move content
            %         self._add_width_of(next_tc)        # 2. add width
            %         self.grid_span += next_tc.grid_span# 3. grow span
            %         next_tc._remove()                  # 4. remove swallowed tc
            next_tc = obj.next_tc_();
            % raise_on_invalid_swallow (NOTE the TWO distinct non-rectangular texts):
            if isequal(next_tc, [])
                mat2doc.exc.InvalidSpanError("not enough grid columns");
            end
            if obj.grid_span + next_tc.grid_span > grid_width
                mat2doc.exc.InvalidSpanError("span is not rectangular");
            end
            assert(~isequal(next_tc, []));   % Python: assert next_tc is not None
            % EXACT swallow order (audit): move -> add width -> grow span -> remove.
            next_tc.move_content_to_(top_tc);
            obj.add_width_of_(next_tc);
            obj.grid_span = obj.grid_span + next_tc.grid_span;
            next_tc.remove_();
        end

        % ===================== NAVIGATION tier (private @property) ===============
        % Python private @property members (leading underscore rotated to trailing).

        % ---- _tbl (table.py 733-736) ----
        function tbl = tbl_(obj)
            % Python: return cast(CT_Tbl, self.xpath(
            %             "./ancestor::w:tbl[position()=1]")[0])
            % CT_Tbl is P6-3b (unregistered) -> the nearest w:tbl ancestor is a
            % GENERIC XmlElement, used only to reach the rows (see trLstOfTbl_).
            res = obj.xpath("./ancestor::w:tbl[position()=1]");
            tbl = res(1);   % IDX: xpath(...)[0] -> (1)
        end

        % ---- CT_Tbl.tr_lst SHIM (planaudit_2026-07-31 condition C2) ----
        function trLst = trLstOfTbl_(obj)
            % The `self._tbl.tr_lst` used by merge/_tr_below/_tr_idx. CT_Tbl.tr_lst
            % (OneOrMore w:tr) is EXACTLY findall("w:tr") in document order; since
            % CT_Tbl is not yet ported (P6-3b), reach the rows via the tbl
            % ancestor's `xpath("./w:tr")` -- w:tr already dispatches CT_Row (P6-2),
            % so this returns the identical CT_Row handle list (H5 persistent
            % handles). Byte-neutral; adjudicated/upgraded at P6-3b when CT_Tbl
            % registers (swap to tbl.tr_lst, or keep -- decided at P6-3b Gate-2).
            tbl = obj.tbl_();
            trLst = tbl.xpath("./w:tr");
        end

        % ---- _tc_above (table.py 738-741) ----
        function tc = tc_above_(obj)
            % Python: return self._tr_above.tc_at_grid_offset(self.grid_offset)
            % grid_offset is 0-based GRID data -> passed RAW (P6-2 convention).
            tc = obj.tr_above_().tc_at_grid_offset(obj.grid_offset);
        end

        % ---- _tc_below (table.py 743-749) ----
        function tc = tc_below_(obj)
            % Python: tr_below = self._tr_below
            %         if tr_below is None: return None
            %         return tr_below.tc_at_grid_offset(self.grid_offset)
            tr_below = obj.tr_below_();
            if isequal(tr_below, [])   % Python: if tr_below is None (H3)
                tc = [];
                return
            end
            tc = tr_below.tc_at_grid_offset(obj.grid_offset);   % grid_offset RAW 0-based
        end

        % ---- _tr (table.py 751-754) ----
        function tr = tr_(obj)
            % Python: return cast(CT_Row, self.xpath(
            %             "./ancestor::w:tr[position()=1]")[0])
            res = obj.xpath("./ancestor::w:tr[position()=1]");
            tr = res(1);   % IDX: xpath(...)[0] -> (1); a CT_Row (P6-2)
        end

        % ---- _tr_above (table.py 756-765) ----
        function tr = tr_above_(obj)
            % Python: tr_aboves = self.xpath(
            %             "./ancestor::w:tr[position()=1]/preceding-sibling::w:tr[1]")
            %         if not tr_aboves: raise ValueError("no tr above topmost tr in w:tbl")
            %         return tr_aboves[0]
            tr_aboves = obj.xpath("./ancestor::w:tr[position()=1]/preceding-sibling::w:tr[1]");
            if isempty(tr_aboves)   % Python: if not tr_aboves (H3)
                error("mat2doc:ValueError", "%s", "no tr above topmost tr in w:tbl");
            end
            tr = tr_aboves(1);   % IDX: tr_aboves[0] -> (1)
        end

        % ---- _tr_below (table.py 767-776) ----
        function tr = tr_below_(obj)
            % Python: tr_lst = self._tbl.tr_lst
            %         tr_idx = tr_lst.index(self._tr)
            %         try: return tr_lst[tr_idx + 1]
            %         except IndexError: return None
            trLst = obj.trLstOfTbl_();
            myTr = obj.tr_();
            i = find(trLst == myTr, 1);   % H5 handle identity; i-1 == 0-based tr_idx
            if isempty(i)   % list.index() ValueError (unreachable for a parented cell)
                error("mat2doc:ValueError", "%s", "tr is not in list");
            end
            % Python tr_lst[tr_idx+1] = 0-based index i -> 1-based trLst(i+1);
            % IndexError (i.e. tr_idx+1 >= len) -> None (IDX: the +1 THEN the bound).
            if (i + 1) <= numel(trLst)
                tr = trLst(i + 1);
            else
                tr = [];   % Python: except IndexError: return None
            end
        end

        % ---- _tr_idx (table.py 778-781) ----
        function value = tr_idx_(obj)
            % Python: return self._tbl.tr_lst.index(self._tr)
            % 0-based row index of this cell's row within the tbl (H1: find()-1).
            % H5: HANDLE identity -- on a uniform 3x3 grid a content-compare would
            % silently return row 1 and corrupt the whole walk. Probe: rows 0/1/2.
            trLst = obj.trLstOfTbl_();
            myTr = obj.tr_();
            i = find(trLst == myTr, 1);   % handle identity
            if isempty(i)   % list.index() ValueError (unreachable for a parented cell)
                error("mat2doc:ValueError", "%s", "tr is not in list");
            end
            value = i - 1;   % IDX: 0-based to match Python list.index()
        end
    end

    methods (Static)
        function elm = new()
            % NEW A new <w:tc> containing an empty <w:p> (the required EG_BlockLevelElt).
            %   Python (table.py 519-522): return cast(CT_Tc, parse_xml(
            %       "<w:tc %s><w:p/></w:tc>" % nsdecls("w"))). Returns a CT_Tc
            %   (w:tc registered by this WP).
            elm = mat2doc.oxml.parse_xml( ...
                "<w:tc " + mat2doc.oxml.nsdecls("w") + "><w:p/></w:tc>");
        end
    end
end

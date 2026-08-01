classdef CT_TcPr < mat2doc.oxml.BaseOxmlElement
% CT_TCPR `<w:tcPr>` element, defining table-cell properties.
%
%   Holds a cell's width (<w:tcW>), horizontal grid span (<w:gridSpan>),
%   vertical-merge behavior (<w:vMerge>) and vertical alignment (<w:vAlign>).
%   Registered for w:tcPr (oxml/__init__.py:180). The audit confirmed this class
%   is ALREADY MINIMAL upstream (table.py 784-891): exactly 4 ZeroOrOne
%   descriptors + 4 accessor pairs + the 18-tag _tag_seq -- NOTHING to defer
%   (there are no borders/shading/margins ACCESSORS in python-docx v1.2.0).
%
%   ============================ H11 (child ordering) ============================
%   _tag_seq (table.py 795-814, VERBATIM, 18 tags) is stored as Constant TAG_SEQ.
%   Each ZeroOrOne descriptor's successors = Python `_tag_seq[s0:]` is expressed
%   as `TAG_SEQ(s0+1:end)` (H1: 0-based Python slice start s0 -> 1-based MATLAB
%   start s0+1). Per-descriptor slice (own 1-based idx / Python slice / MATLAB
%   slice / first successor tag):
%     tcW      = 2 : successors=_tag_seq[2:]  -> TAG_SEQ(3:end)   first "w:gridSpan"
%     gridSpan = 3 : successors=_tag_seq[3:]  -> TAG_SEQ(4:end)   first "w:hMerge"
%     vMerge   = 5 : successors=_tag_seq[5:]  -> TAG_SEQ(6:end)   first "w:tcBorders"
%     vAlign   = 12: successors=_tag_seq[12:] -> TAG_SEQ(13:end)  first "w:hideMark"
%   Upstream swallow fixture `(w:tcW..., w:gridSpan...)` byte-pins tcW-before-
%   gridSpan; a wrong slice would scramble the tcPr children -> Word repair /
%   byte divergence.
%
%   ===================== GENERATED DESCRIPTOR FAMILIES ==========================
%   tcW / gridSpan / vMerge / vAlign are ZeroOrOne (table.py 815-826), generic
%   engine (no _new_x/_insert_x override on CT_TcPr): get.x, get_or_add_x, new_x_,
%   insert_x_, add_x_, remove_x_ (underscore rotation of _new_x/_insert_x/_add_x/
%   _remove_x; get_or_add_x public). The Callable annotations (table.py 787-793:
%   get_or_add_gridSpan / get_or_add_tcW / get_or_add_vAlign / _add_vMerge /
%   _remove_gridSpan / _remove_vAlign / _remove_vMerge) are type hints only.
%
%   CHILD-CLASS REGISTRATION (this WP registers w:gridSpan; the others are live):
%     w:gridSpan -> CT_DecimalNumber (oxml/__init__.py:172, registered by P6-3a)
%     w:tcW      -> CT_TblWidth       (P6-1)
%     w:vMerge   -> CT_VMerge         (P6-1)
%     w:vAlign   -> CT_VerticalJc     (P6-1)
%   grid_span reads .val on the CT_DecimalNumber gridSpan child (must resolve to
%   CT_DecimalNumber -- registered by this WP); width/vMerge_val/vAlign_val read
%   the P6-1 leaf accessors.
%
%   ===================== H3 / H4 (@property accessors) ==========================
%   grid_span (get+set, table.py 829-842): get -> 1 when w:gridSpan absent (H3),
%   else gridSpan.val (H6 int). set -> ALWAYS _remove_gridSpan() first, then write
%   @w:val ONLY `if value > 1` (H4 strict): span==1 (or shrinking to 1) emits NO
%   <w:gridSpan> element at all.
%   vMerge_val (get+set, table.py 863-876): get -> [] (None) when w:vMerge absent,
%   else vMerge.val (the CT_VMerge OptionalAttribute, default "continue"). set ->
%   ALWAYS _remove_vMerge() first, then `if value is not None: _add_vMerge().val =
%   value`. The bare-<w:vMerge/> continuation cell (sharpest NEW-D risk) falls out
%   of this: assigning "continue" (the ST_Merge default) makes CT_VMerge's
%   OptionalAttribute setter DELETE @w:val (D-delta-1), serializing a BARE
%   <w:vMerge/>; assigning "restart" writes @w:val="restart"; assigning None
%   removes the element. No special-casing here -- the P6-1 CT_VMerge setter is
%   the byte mechanism.
%   width (get+set, table.py 878-889): get -> [] (None) when w:tcW absent, else
%   tcW.width (the CT_TblWidth dxa union, [] for non-dxa). set -> get_or_add_tcW +
%   assign (EMU->dxa twips inside CT_TblWidth).
%   vAlign_val (get+set, table.py 844-861): get -> [] (None) when w:vAlign absent,
%   else vAlign.val (CT_VerticalJc RequiredAttribute). set -> `if value is None:
%   _remove_vAlign(); return` else get_or_add_vAlign().val = value.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on every <w:tcPr> inside a real table cell AND on the 595
%   <w:tcPr> nodes in word/styles.xml + 595 in stylesWithEffects.xml (inside the
%   <w:tblStylePr> table-style overrides). Registering w:tcPr thus puts styles.xml
%   on the live parse path -> P6-3a is NON-M1-neutral (byte-neutral EXPECTED --
%   descriptors-only, no parse-time behavior, the CT_TblPr precedent -- but must be
%   PROVEN: M1 17/17, styles.xml SHA re-derived).
%
%   Example:
%       tcPr = mat2doc.oxml.OxmlElement("w:tcPr");   % a CT_TcPr
%       tcPr.grid_span                               % 1 (no <w:gridSpan>)
%       tcPr.grid_span = 2;                          % <w:gridSpan w:val="2"/>
%       tcPr.vMerge_val = "continue";                % <w:vMerge/> (bare, D-delta-1)
%       tcPr.vMerge_val = "restart";                 % <w:vMerge w:val="restart"/>
%
%   Ported from python-docx v1.2.0: src/docx/oxml/table.py::CT_TcPr
%   (lines 784-891; registered for w:tcPr)

    properties (Constant, Hidden)  % _tag_seq VERBATIM (table.py 795-814; 18 tags)
        TAG_SEQ = [ ...
            "w:cnfStyle", "w:tcW", "w:gridSpan", ...            %  1- 3  (tcW own @2, gridSpan own @3)
            "w:hMerge", "w:vMerge", "w:tcBorders", ...          %  4- 6  (vMerge own @5)
            "w:shd", "w:noWrap", "w:tcMar", ...                 %  7- 9
            "w:textDirection", "w:tcFitText", "w:vAlign", ...   % 10-12  (vAlign own @12)
            "w:hideMark", "w:headers", "w:cellIns", ...         % 13-15
            "w:cellDel", "w:cellMerge", "w:tcPrChange" ]        % 16-18
    end

    properties (Dependent)  % generated ZeroOrOne getters + @property accessors
        tcW        % ZeroOrOne <w:tcW> child or []
        gridSpan   % ZeroOrOne <w:gridSpan> child or []
        vMerge     % ZeroOrOne <w:vMerge> child or []
        vAlign     % ZeroOrOne <w:vAlign> child or []
        grid_span  % @property: 1 when w:gridSpan absent, else its .val (int); set writes @w:val only if >1
        vMerge_val % @property: w:vMerge/@w:val (string) or []; set removes then re-adds if not None
        width      % @property: EMU Length in w:tcW or [] (None / non-dxa)
        vAlign_val % @property: w:vAlign/@w:val (WD_CELL_VERTICAL_ALIGNMENT) or []
    end

    methods
        function obj = CT_TcPr(varargin)
            % CT_TCPR Construct a loose <w:tcPr> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ============ tcW (ZeroOrOne, successors=_tag_seq[2:] -> TAG_SEQ(3:end)) ============
        function child = get.tcW(obj);            child = obj.getChild("w:tcW"); end
        function child = get_or_add_tcW(obj);     child = obj.getOrAddChild("w:tcW", obj.TAG_SEQ(3:end)); end
        function child = new_tcW_(obj);           child = obj.newChild("w:tcW"); end
        function child = insert_tcW_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(3:end)); end
        function child = add_tcW_(obj, varargin); child = obj.addChild("w:tcW", obj.TAG_SEQ(3:end), varargin{:}); end
        function remove_tcW_(obj);                obj.removeChild("w:tcW"); end

        % ============ gridSpan (ZeroOrOne, successors=_tag_seq[3:] -> TAG_SEQ(4:end)) ============
        function child = get.gridSpan(obj);            child = obj.getChild("w:gridSpan"); end
        function child = get_or_add_gridSpan(obj);     child = obj.getOrAddChild("w:gridSpan", obj.TAG_SEQ(4:end)); end
        function child = new_gridSpan_(obj);           child = obj.newChild("w:gridSpan"); end
        function child = insert_gridSpan_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(4:end)); end
        function child = add_gridSpan_(obj, varargin); child = obj.addChild("w:gridSpan", obj.TAG_SEQ(4:end), varargin{:}); end
        function remove_gridSpan_(obj);                obj.removeChild("w:gridSpan"); end

        % ============ vMerge (ZeroOrOne, successors=_tag_seq[5:] -> TAG_SEQ(6:end)) ============
        function child = get.vMerge(obj);            child = obj.getChild("w:vMerge"); end
        function child = get_or_add_vMerge(obj);     child = obj.getOrAddChild("w:vMerge", obj.TAG_SEQ(6:end)); end
        function child = new_vMerge_(obj);           child = obj.newChild("w:vMerge"); end
        function child = insert_vMerge_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(6:end)); end
        function child = add_vMerge_(obj, varargin); child = obj.addChild("w:vMerge", obj.TAG_SEQ(6:end), varargin{:}); end
        function remove_vMerge_(obj);                obj.removeChild("w:vMerge"); end

        % ============ vAlign (ZeroOrOne, successors=_tag_seq[12:] -> TAG_SEQ(13:end)) ============
        function child = get.vAlign(obj);            child = obj.getChild("w:vAlign"); end
        function child = get_or_add_vAlign(obj);     child = obj.getOrAddChild("w:vAlign", obj.TAG_SEQ(13:end)); end
        function child = new_vAlign_(obj);           child = obj.newChild("w:vAlign"); end
        function child = insert_vAlign_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(13:end)); end
        function child = add_vAlign_(obj, varargin); child = obj.addChild("w:vAlign", obj.TAG_SEQ(13:end), varargin{:}); end
        function remove_vAlign_(obj);                obj.removeChild("w:vAlign"); end

        % ===================== @property accessors (table.py 829-889) =====================

        % ---- grid_span (get+set, table.py 829-842) ----
        function value = get.grid_span(obj)
            % Python: gridSpan = self.gridSpan; return 1 if None else gridSpan.val
            gridSpan = obj.gridSpan;
            if isequal(gridSpan, [])   % Python: if gridSpan is None (H3)
                value = 1;
                return
            end
            value = gridSpan.val;
        end
        function set.grid_span(obj, value)
            % Python (table.py 838-842): self._remove_gridSpan();
            %   if value > 1: self.get_or_add_gridSpan().val = value
            obj.remove_gridSpan_();
            if value > 1   % H4: STRICT > 1 -- span==1 emits NO <w:gridSpan>
                child = obj.get_or_add_gridSpan();
                child.val = value;
            end
        end

        % ---- vAlign_val (get+set, table.py 844-861) ----
        function value = get.vAlign_val(obj)
            % Python: vAlign = self.vAlign; return None if None else vAlign.val
            vAlign = obj.vAlign;
            if isequal(vAlign, [])   % Python: if vAlign is None (H3)
                value = [];
                return
            end
            value = vAlign.val;
        end
        function set.vAlign_val(obj, value)
            % Python (table.py 857-861): if value is None: self._remove_vAlign();
            %   return; self.get_or_add_vAlign().val = value
            if isequal(value, [])   % Python: if value is None (H3)
                obj.remove_vAlign_();
                return
            end
            child = obj.get_or_add_vAlign();
            child.val = value;
        end

        % ---- vMerge_val (get+set, table.py 863-876) ----
        function value = get.vMerge_val(obj)
            % Python: vMerge = self.vMerge; return None if None else vMerge.val
            vMerge = obj.vMerge;
            if isequal(vMerge, [])   % Python: if vMerge is None (H3)
                value = [];
                return
            end
            value = vMerge.val;
        end
        function set.vMerge_val(obj, value)
            % Python (table.py 873-876): self._remove_vMerge();
            %   if value is not None: self._add_vMerge().val = value
            % The bare-<w:vMerge/> byte behavior is delegated ENTIRELY to CT_VMerge's
            % OptionalAttribute setter (default "continue" -> deletes @w:val); no
            % special-casing here (see class header H3/H4).
            obj.remove_vMerge_();
            if ~isequal(value, [])   % Python: if value is not None (H3)
                child = obj.add_vMerge_();
                child.val = value;
            end
        end

        % ---- width (get+set, table.py 878-889) ----
        function value = get.width(obj)
            % Python: tcW = self.tcW; return None if None else tcW.width
            tcW = obj.tcW;
            if isequal(tcW, [])   % Python: if tcW is None (H3)
                value = [];
                return
            end
            value = tcW.width;
        end
        function set.width(obj, value)
            % Python (table.py 887-889): tcW = self.get_or_add_tcW(); tcW.width = value
            tcW = obj.get_or_add_tcW();
            tcW.width = value;
        end
    end
end

classdef CT_PPr < mat2doc.oxml.BaseOxmlElement
% CT_PPR `<w:pPr>` element: the paragraph-properties container for a paragraph.
%
%   THE M2 BYTE-CRITICAL paragraph-properties element. Its 12 ZeroOrOne child
%   descriptors carry NON-CONTIGUOUS H11 successor slices of a single 36-element
%   _tag_seq; the slices drive insert_element_before, which re-sorts scrambled
%   child adds into canonical OOXML schema order so document.xml/styles.xml stay
%   byte-identical. (Same discipline as CT_RPr.) The M2 add_heading path inserts
%   pStyle, which sorts before EVERYTHING (successors=_tag_seq[1:]).
%
%   _tag_seq (parfmt.py 64-101, VERBATIM, 36 tags) is stored as the Constant
%   TAG_SEQ. Each descriptor's successors = Python _tag_seq[s0:] is expressed as
%   TAG_SEQ(s0+1:end) (H1: 0-based Python slice start s0 -> 1-based MATLAB start
%   s0+1). The own tag of a descriptor sits at TAG_SEQ(s0) (1-based).
%
%   H11 SUCCESSOR-SLICE TABLE (prop : own 1-based idx : Python slice : MATLAB slice):
%     pStyle          =  1 : successors=_tag_seq[1:]  -> TAG_SEQ(2:end)
%     keepNext        =  2 : successors=_tag_seq[2:]  -> TAG_SEQ(3:end)
%     keepLines       =  3 : successors=_tag_seq[3:]  -> TAG_SEQ(4:end)
%     pageBreakBefore =  4 : successors=_tag_seq[4:]  -> TAG_SEQ(5:end)
%     widowControl    =  6 : successors=_tag_seq[6:]  -> TAG_SEQ(7:end)   (framePr@5 skipped)
%     numPr           =  7 : successors=_tag_seq[7:]  -> TAG_SEQ(8:end)
%     tabs            = 11 : successors=_tag_seq[11:] -> TAG_SEQ(12:end)
%     spacing         = 22 : successors=_tag_seq[22:] -> TAG_SEQ(23:end)
%     ind             = 23 : successors=_tag_seq[23:] -> TAG_SEQ(24:end)
%     jc              = 27 : successors=_tag_seq[27:] -> TAG_SEQ(28:end)
%     outlineLvl      = 31 : successors=_tag_seq[31:] -> TAG_SEQ(32:end)
%     sectPr          = 35 : successors=_tag_seq[35:] -> TAG_SEQ(36:end)
%
%   GENERATED DESCRIPTOR FAMILY (per ZeroOrOne, xmlchemy docx form): get.x,
%   get_or_add_x, new_x_, insert_x_, add_x_, remove_x_ (underscore rotation:
%   Python _new_x/_insert_x/_add_x/_remove_x -> new_x_/insert_x_/add_x_/
%   remove_x_; get_or_add_x public). The pyright Callable annotations
%   (parfmt.py 57-62: get_or_add_ind/get_or_add_pStyle/get_or_add_sectPr/
%   _insert_sectPr/_remove_pStyle/_remove_sectPr) are type hints only and add no
%   members beyond this family. ALL 12 descriptors use the generic
%   BaseOxmlElement engine (no _new_x/_insert_x override on CT_PPr).
%
%   CHILD-CLASS REGISTRATION (parse-time class of each descriptor's child):
%     pStyle -> CT_String, keepNext/keepLines/pageBreakBefore/widowControl ->
%     CT_OnOff, tabs -> CT_TabStops, spacing -> CT_Spacing, ind -> CT_Ind,
%     jc -> CT_Jc (all registered by THIS WP or earlier). numPr -> CT_NumPr (P8),
%     outlineLvl -> CT_DecimalNumber (shared.py, numbering/shared WP), sectPr ->
%     CT_SectPr (P5) are NOT yet ported -> resolve to generic XmlElement. This is
%     BYTE-NEUTRAL and behavior-neutral: no ported CT_PPr accessor reads .val on
%     numPr/outlineLvl/sectPr, and the descriptor family (get_or_add/insert/
%     remove) uses the generic engine which does not depend on the child class.
%
%   @property members (parfmt.py 122-339): first_line_indent, ind_left, ind_right,
%   jc_val, keepLines_val, keepNext_val, pageBreakBefore_val, spacing_after,
%   spacing_before, spacing_line, spacing_lineRule, style, widowControl_val --
%   all ported LIVE (deps present). spacing_lineRule maps a present @w:line with
%   absent @w:lineRule to WD_LINE_SPACING.MULTIPLE (parfmt.py 296-298).
%
%   H3 (None): inline isequal(x, []) (established Mat2Doc oxml convention).
%   H4 (truthiness): first_line_indent uses `value < 0` (value is a Length/[]).
%   H6 (Length): Length(-hanging) and -value negate a Length -> double EMU count,
%   accepted by the CT_Ind ST_TwipsMeasure setter (Emu(count).twips).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on the many <w:pPr> nodes inside styles.xml/document.xml.
%
%   Example:
%       pPr = mat2doc.oxml.OxmlElement("w:pPr");
%       pPr.style = "Heading1";          % <w:pStyle w:val="Heading1"/> (M2 path)
%       pPr.get_or_add_jc().val = mat2doc.enum.text.WD_ALIGN_PARAGRAPH.CENTER;
%
%   Ported from python-docx v1.2.0: src/docx/oxml/text/parfmt.py::CT_PPr
%   (lines 54-339; registered for w:pPr)

    properties (Constant, Hidden)  % _tag_seq VERBATIM (parfmt.py 64-101; 36 tags)
        TAG_SEQ = [ ...
            "w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", ...
            "w:framePr", "w:widowControl", "w:numPr", "w:suppressLineNumbers", ...
            "w:pBdr", "w:shd", "w:tabs", "w:suppressAutoHyphens", ...
            "w:kinsoku", "w:wordWrap", "w:overflowPunct", "w:topLinePunct", ...
            "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd", ...
            "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing", ...
            "w:mirrorIndents", "w:suppressOverlap", "w:jc", "w:textDirection", ...
            "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl", "w:divId", ...
            "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange" ]
    end

    properties (Dependent)  % generated ZeroOrOne getters + @property members
        % -- 12 ZeroOrOne child getters (read-only; use get_or_add_x/remove_x_) --
        pStyle
        keepNext
        keepLines
        pageBreakBefore
        widowControl
        numPr
        tabs
        spacing
        ind
        jc
        outlineLvl
        sectPr
        % -- @property members (parfmt.py 122-339) --
        first_line_indent   % Length from w:ind/@w:firstLine|@w:hanging, or []
        ind_left            % w:ind/@w:left or []
        ind_right           % w:ind/@w:right or []
        jc_val              % ./w:jc/@val (WD_ALIGN_PARAGRAPH) or []
        keepLines_val       % keepLines/@val or []
        keepNext_val        % keepNext/@val or []
        pageBreakBefore_val % pageBreakBefore/@val or []
        spacing_after       % w:spacing/@w:after or []
        spacing_before      % w:spacing/@w:before or []
        spacing_line        % w:spacing/@w:line or []
        spacing_lineRule    % w:spacing/@w:lineRule (WD_LINE_SPACING) or [] (MULTIPLE if line present)
        style               % ./w:pStyle/@val or []
        widowControl_val    % widowControl/@val or []
    end

    methods
        function obj = CT_PPr(varargin)
            % CT_PPR Construct a loose <w:pPr> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- pStyle (ZeroOrOne, successors=_tag_seq[1:] -> TAG_SEQ(2:end)) ----
        function child = get.pStyle(obj);            child = obj.getChild("w:pStyle"); end
        function child = get_or_add_pStyle(obj);     child = obj.getOrAddChild("w:pStyle", obj.TAG_SEQ(2:end)); end
        function child = new_pStyle_(obj);           child = obj.newChild("w:pStyle"); end
        function child = insert_pStyle_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(2:end)); end
        function child = add_pStyle_(obj, varargin); child = obj.addChild("w:pStyle", obj.TAG_SEQ(2:end), varargin{:}); end
        function remove_pStyle_(obj);                obj.removeChild("w:pStyle"); end

        % ---- keepNext (ZeroOrOne, successors=_tag_seq[2:] -> TAG_SEQ(3:end)) ----
        function child = get.keepNext(obj);            child = obj.getChild("w:keepNext"); end
        function child = get_or_add_keepNext(obj);     child = obj.getOrAddChild("w:keepNext", obj.TAG_SEQ(3:end)); end
        function child = new_keepNext_(obj);           child = obj.newChild("w:keepNext"); end
        function child = insert_keepNext_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(3:end)); end
        function child = add_keepNext_(obj, varargin); child = obj.addChild("w:keepNext", obj.TAG_SEQ(3:end), varargin{:}); end
        function remove_keepNext_(obj);                obj.removeChild("w:keepNext"); end

        % ---- keepLines (ZeroOrOne, successors=_tag_seq[3:] -> TAG_SEQ(4:end)) ----
        function child = get.keepLines(obj);            child = obj.getChild("w:keepLines"); end
        function child = get_or_add_keepLines(obj);     child = obj.getOrAddChild("w:keepLines", obj.TAG_SEQ(4:end)); end
        function child = new_keepLines_(obj);           child = obj.newChild("w:keepLines"); end
        function child = insert_keepLines_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(4:end)); end
        function child = add_keepLines_(obj, varargin); child = obj.addChild("w:keepLines", obj.TAG_SEQ(4:end), varargin{:}); end
        function remove_keepLines_(obj);                obj.removeChild("w:keepLines"); end

        % ---- pageBreakBefore (ZeroOrOne, successors=_tag_seq[4:] -> TAG_SEQ(5:end)) ----
        function child = get.pageBreakBefore(obj);            child = obj.getChild("w:pageBreakBefore"); end
        function child = get_or_add_pageBreakBefore(obj);     child = obj.getOrAddChild("w:pageBreakBefore", obj.TAG_SEQ(5:end)); end
        function child = new_pageBreakBefore_(obj);           child = obj.newChild("w:pageBreakBefore"); end
        function child = insert_pageBreakBefore_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(5:end)); end
        function child = add_pageBreakBefore_(obj, varargin); child = obj.addChild("w:pageBreakBefore", obj.TAG_SEQ(5:end), varargin{:}); end
        function remove_pageBreakBefore_(obj);                obj.removeChild("w:pageBreakBefore"); end

        % ---- widowControl (ZeroOrOne, successors=_tag_seq[6:] -> TAG_SEQ(7:end)) ----
        function child = get.widowControl(obj);            child = obj.getChild("w:widowControl"); end
        function child = get_or_add_widowControl(obj);     child = obj.getOrAddChild("w:widowControl", obj.TAG_SEQ(7:end)); end
        function child = new_widowControl_(obj);           child = obj.newChild("w:widowControl"); end
        function child = insert_widowControl_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(7:end)); end
        function child = add_widowControl_(obj, varargin); child = obj.addChild("w:widowControl", obj.TAG_SEQ(7:end), varargin{:}); end
        function remove_widowControl_(obj);                obj.removeChild("w:widowControl"); end

        % ---- numPr (ZeroOrOne, successors=_tag_seq[7:] -> TAG_SEQ(8:end)) ----
        function child = get.numPr(obj);            child = obj.getChild("w:numPr"); end
        function child = get_or_add_numPr(obj);     child = obj.getOrAddChild("w:numPr", obj.TAG_SEQ(8:end)); end
        function child = new_numPr_(obj);           child = obj.newChild("w:numPr"); end
        function child = insert_numPr_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(8:end)); end
        function child = add_numPr_(obj, varargin); child = obj.addChild("w:numPr", obj.TAG_SEQ(8:end), varargin{:}); end
        function remove_numPr_(obj);                obj.removeChild("w:numPr"); end

        % ---- tabs (ZeroOrOne, successors=_tag_seq[11:] -> TAG_SEQ(12:end)) ----
        function child = get.tabs(obj);            child = obj.getChild("w:tabs"); end
        function child = get_or_add_tabs(obj);     child = obj.getOrAddChild("w:tabs", obj.TAG_SEQ(12:end)); end
        function child = new_tabs_(obj);           child = obj.newChild("w:tabs"); end
        function child = insert_tabs_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(12:end)); end
        function child = add_tabs_(obj, varargin); child = obj.addChild("w:tabs", obj.TAG_SEQ(12:end), varargin{:}); end
        function remove_tabs_(obj);                obj.removeChild("w:tabs"); end

        % ---- spacing (ZeroOrOne, successors=_tag_seq[22:] -> TAG_SEQ(23:end)) ----
        function child = get.spacing(obj);            child = obj.getChild("w:spacing"); end
        function child = get_or_add_spacing(obj);     child = obj.getOrAddChild("w:spacing", obj.TAG_SEQ(23:end)); end
        function child = new_spacing_(obj);           child = obj.newChild("w:spacing"); end
        function child = insert_spacing_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(23:end)); end
        function child = add_spacing_(obj, varargin); child = obj.addChild("w:spacing", obj.TAG_SEQ(23:end), varargin{:}); end
        function remove_spacing_(obj);                obj.removeChild("w:spacing"); end

        % ---- ind (ZeroOrOne, successors=_tag_seq[23:] -> TAG_SEQ(24:end)) ----
        function child = get.ind(obj);            child = obj.getChild("w:ind"); end
        function child = get_or_add_ind(obj);     child = obj.getOrAddChild("w:ind", obj.TAG_SEQ(24:end)); end
        function child = new_ind_(obj);           child = obj.newChild("w:ind"); end
        function child = insert_ind_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(24:end)); end
        function child = add_ind_(obj, varargin); child = obj.addChild("w:ind", obj.TAG_SEQ(24:end), varargin{:}); end
        function remove_ind_(obj);                obj.removeChild("w:ind"); end

        % ---- jc (ZeroOrOne, successors=_tag_seq[27:] -> TAG_SEQ(28:end)) ----
        function child = get.jc(obj);            child = obj.getChild("w:jc"); end
        function child = get_or_add_jc(obj);     child = obj.getOrAddChild("w:jc", obj.TAG_SEQ(28:end)); end
        function child = new_jc_(obj);           child = obj.newChild("w:jc"); end
        function child = insert_jc_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(28:end)); end
        function child = add_jc_(obj, varargin); child = obj.addChild("w:jc", obj.TAG_SEQ(28:end), varargin{:}); end
        function remove_jc_(obj);                obj.removeChild("w:jc"); end

        % ---- outlineLvl (ZeroOrOne, successors=_tag_seq[31:] -> TAG_SEQ(32:end)) ----
        function child = get.outlineLvl(obj);            child = obj.getChild("w:outlineLvl"); end
        function child = get_or_add_outlineLvl(obj);     child = obj.getOrAddChild("w:outlineLvl", obj.TAG_SEQ(32:end)); end
        function child = new_outlineLvl_(obj);           child = obj.newChild("w:outlineLvl"); end
        function child = insert_outlineLvl_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(32:end)); end
        function child = add_outlineLvl_(obj, varargin); child = obj.addChild("w:outlineLvl", obj.TAG_SEQ(32:end), varargin{:}); end
        function remove_outlineLvl_(obj);                obj.removeChild("w:outlineLvl"); end

        % ---- sectPr (ZeroOrOne, successors=_tag_seq[35:] -> TAG_SEQ(36:end)) ----
        function child = get.sectPr(obj);            child = obj.getChild("w:sectPr"); end
        function child = get_or_add_sectPr(obj);     child = obj.getOrAddChild("w:sectPr", obj.TAG_SEQ(36:end)); end
        function child = new_sectPr_(obj);           child = obj.newChild("w:sectPr"); end
        function child = insert_sectPr_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(36:end)); end
        function child = add_sectPr_(obj, varargin); child = obj.addChild("w:sectPr", obj.TAG_SEQ(36:end), varargin{:}); end
        function remove_sectPr_(obj);                obj.removeChild("w:sectPr"); end

        % =================== @property members (parfmt.py 122-339) ===================

        % ---- first_line_indent (parfmt.py 122-151) ----
        function value = get.first_line_indent(obj)
            ind = obj.ind;
            if isequal(ind, [])          % Python: if ind is None
                value = [];
                return
            end
            hanging = ind.hanging;
            if ~isequal(hanging, [])     % Python: if hanging is not None
                value = mat2doc.shared.Length(-hanging);   % Python: Length(-hanging)
                return
            end
            firstLine = ind.firstLine;
            if isequal(firstLine, [])    % Python: if firstLine is None
                value = [];
                return
            end
            value = firstLine;
        end
        function set.first_line_indent(obj, value)
            % Python (parfmt.py 141-151)
            if isequal(obj.ind, []) && isequal(value, [])   % if self.ind is None and value is None
                return
            end
            ind = obj.get_or_add_ind();
            ind.firstLine = [];     % Python: ind.firstLine = ind.hanging = None
            ind.hanging = [];
            if isequal(value, [])   % Python: if value is None
                return
            elseif value < 0        % Python: elif value < 0 (H4/H6)
                ind.hanging = -value;
            else
                ind.firstLine = value;
            end
        end

        % ---- ind_left (parfmt.py 153-166) ----
        function value = get.ind_left(obj)
            ind = obj.ind;
            if isequal(ind, [])   % Python: if ind is None
                value = [];
                return
            end
            value = ind.left;
        end
        function set.ind_left(obj, value)
            % Python: if value is None and self.ind is None: return
            if isequal(value, []) && isequal(obj.ind, [])
                return
            end
            ind = obj.get_or_add_ind();
            ind.left = value;
        end

        % ---- ind_right (parfmt.py 168-181) ----
        function value = get.ind_right(obj)
            ind = obj.ind;
            if isequal(ind, [])   % Python: if ind is None
                value = [];
                return
            end
            value = ind.right;
        end
        function set.ind_right(obj, value)
            % Python: if value is None and self.ind is None: return
            if isequal(value, []) && isequal(obj.ind, [])
                return
            end
            ind = obj.get_or_add_ind();
            ind.right = value;
        end

        % ---- jc_val (parfmt.py 183-193) ----
        function value = get.jc_val(obj)
            % Python: return self.jc.val if self.jc is not None else None
            if ~isequal(obj.jc, [])
                value = obj.jc.val;
            else
                value = [];
            end
        end
        function set.jc_val(obj, value)
            if isequal(value, [])   % Python: if value is None
                obj.remove_jc_();
                return
            end
            j = obj.get_or_add_jc();
            j.val = value;
        end

        % ---- keepLines_val (parfmt.py 195-208) ----
        function value = get.keepLines_val(obj)
            keepLines = obj.keepLines;
            if isequal(keepLines, [])   % Python: if keepLines is None
                value = [];
                return
            end
            value = keepLines.val;
        end
        function set.keepLines_val(obj, value)
            if isequal(value, [])       % Python: if value is None
                obj.remove_keepLines_();
            else
                el = obj.get_or_add_keepLines();
                el.val = value;
            end
        end

        % ---- keepNext_val (parfmt.py 210-223) ----
        function value = get.keepNext_val(obj)
            keepNext = obj.keepNext;
            if isequal(keepNext, [])    % Python: if keepNext is None
                value = [];
                return
            end
            value = keepNext.val;
        end
        function set.keepNext_val(obj, value)
            if isequal(value, [])       % Python: if value is None
                obj.remove_keepNext_();
            else
                el = obj.get_or_add_keepNext();
                el.val = value;
            end
        end

        % ---- pageBreakBefore_val (parfmt.py 225-238) ----
        function value = get.pageBreakBefore_val(obj)
            pageBreakBefore = obj.pageBreakBefore;
            if isequal(pageBreakBefore, [])   % Python: if pageBreakBefore is None
                value = [];
                return
            end
            value = pageBreakBefore.val;
        end
        function set.pageBreakBefore_val(obj, value)
            if isequal(value, [])       % Python: if value is None
                obj.remove_pageBreakBefore_();
            else
                el = obj.get_or_add_pageBreakBefore();
                el.val = value;
            end
        end

        % ---- spacing_after (parfmt.py 240-252) ----
        function value = get.spacing_after(obj)
            spacing = obj.spacing;
            if isequal(spacing, [])   % Python: if spacing is None
                value = [];
                return
            end
            value = spacing.after;
        end
        function set.spacing_after(obj, value)
            % Python: if value is None and self.spacing is None: return
            if isequal(value, []) && isequal(obj.spacing, [])
                return
            end
            sp = obj.get_or_add_spacing();
            sp.after = value;
        end

        % ---- spacing_before (parfmt.py 254-266) ----
        function value = get.spacing_before(obj)
            spacing = obj.spacing;
            if isequal(spacing, [])   % Python: if spacing is None
                value = [];
                return
            end
            value = spacing.before;
        end
        function set.spacing_before(obj, value)
            % Python: if value is None and self.spacing is None: return
            if isequal(value, []) && isequal(obj.spacing, [])
                return
            end
            sp = obj.get_or_add_spacing();
            sp.before = value;
        end

        % ---- spacing_line (parfmt.py 268-280) ----
        function value = get.spacing_line(obj)
            spacing = obj.spacing;
            if isequal(spacing, [])   % Python: if spacing is None
                value = [];
                return
            end
            value = spacing.line;
        end
        function set.spacing_line(obj, value)
            % Python: if value is None and self.spacing is None: return
            if isequal(value, []) && isequal(obj.spacing, [])
                return
            end
            sp = obj.get_or_add_spacing();
            sp.line = value;
        end

        % ---- spacing_lineRule (parfmt.py 282-304) ----
        function value = get.spacing_lineRule(obj)
            spacing = obj.spacing;
            if isequal(spacing, [])   % Python: if spacing is None
                value = [];
                return
            end
            lineRule = spacing.lineRule;
            % Python: if lineRule is None and spacing.line is not None:
            %             return WD_LINE_SPACING.MULTIPLE
            if isequal(lineRule, []) && ~isequal(spacing.line, [])
                value = mat2doc.enum.text.WD_LINE_SPACING.MULTIPLE;
                return
            end
            value = lineRule;
        end
        function set.spacing_lineRule(obj, value)
            % Python: if value is None and self.spacing is None: return
            if isequal(value, []) && isequal(obj.spacing, [])
                return
            end
            sp = obj.get_or_add_spacing();
            sp.lineRule = value;
        end

        % ---- style (parfmt.py 306-324) ----
        function value = get.style(obj)
            pStyle = obj.pStyle;
            if isequal(pStyle, [])   % Python: if pStyle is None
                value = [];
                return
            end
            value = pStyle.val;
        end
        function set.style(obj, style)
            if isequal(style, [])    % Python: if style is None
                obj.remove_pStyle_();
                return
            end
            pStyle = obj.get_or_add_pStyle();
            pStyle.val = style;
        end

        % ---- widowControl_val (parfmt.py 326-339) ----
        function value = get.widowControl_val(obj)
            widowControl = obj.widowControl;
            if isequal(widowControl, [])   % Python: if widowControl is None
                value = [];
                return
            end
            value = widowControl.val;
        end
        function set.widowControl_val(obj, value)
            if isequal(value, [])       % Python: if value is None
                obj.remove_widowControl_();
            else
                el = obj.get_or_add_widowControl();
                el.val = value;
            end
        end
    end
end

classdef ParagraphFormat < mat2doc.shared.ElementProxy
% PARAGRAPHFORMAT Paragraph formatting: alignment, indentation, spacing, tabs.
%
%   Provides access to paragraph formatting such as justification,
%   indentation, line spacing, space before and after, and widow/orphan
%   control. An ElementProxy subclass: reference semantics (handle) and H5
%   element-identity eq/ne are inherited unchanged.
%
%   WRAPPED ELEMENT (parfmt.py 20, 27, 237): unlike Font (which wraps the
%   <w:r>), ParagraphFormat wraps the pPr's OWNER -- a <w:p> (CT_P) or a style
%   pPr owner -- and reaches its properties through `self._element.pPr` /
%   `self._element.get_or_add_pPr()`. Every accessor delegates to the P4-2
%   CT_PPr @property members (jc_val, first_line_indent, ind_left/ind_right,
%   keepLines_val/keepNext_val, pageBreakBefore_val, widowControl_val,
%   spacing_after/before/line/lineRule). ParagraphFormat adds NO oxml logic,
%   NO registry rows and NO serialization code (API/proxy tier).
%
%   TRI-STATE (H3): every property is None-vs-value; None (inherited from the
%   style hierarchy) is the sentinel [], tested via inline isequal(x, []) (the
%   established Mat2Doc None idiom).
%
%   LINE SPACING (parfmt.py 103-160, 256-286) -- the subtlest logic here:
%     * line_spacing GET returns a FLOAT number of lines (D-STYPE-1) when the
%       rule is MULTIPLE (spacing_line / Pt(12)); otherwise the absolute Length.
%     * line_spacing SET dispatches on the ARGUMENT TYPE: a Length -> absolute
%       (EXACTLY unless already AT_LEAST); any non-Length numeric (a float
%       multiple) -> Emu(value * Twips(240)) + MULTIPLE.
%     * line_spacing_rule GET maps the MULTIPLE special members
%       (Twips 240 -> SINGLE, 360 -> ONE_POINT_FIVE, 480 -> DOUBLE); SET of
%       those members writes the corresponding Twips line + MULTIPLE.
%   Note Pt(12) == Twips(240) == 152400 EMU (single-space unit); parfmt.py uses
%   Pt(12) on the get side and Twips(240) on the set side -- both ported verbatim.
%
%   tab_stops (parfmt.py 233-238) is a @lazyproperty: a TabStops view of the
%   pPr, computed (and CACHED) on first access. get_or_add_pPr() runs once and
%   mutates by adding a <w:pPr> if absent (Python lazyproperty body semantics).
%   Realized via the cache + computed-flag idiom (mat2doc.shared.lazyproperty);
%   read-only (no set.tab_stops), mirroring Python's AttributeError.
%
%   Example:
%       p  = mat2doc.oxml.OxmlElement("w:p");
%       pf = mat2doc.text.ParagraphFormat(p);
%       pf.alignment = mat2doc.enum.text.WD_ALIGN_PARAGRAPH.CENTER;
%       pf.left_indent = mat2doc.shared.Pt(36);
%       pf.line_spacing = 1.5;   % -> <w:spacing w:line="360" w:lineRule="auto"/>
%       pf.line_spacing_rule     % WD_LINE_SPACING.ONE_POINT_FIVE
%       pf.space_after = mat2doc.shared.Pt(6);
%       ts = pf.tab_stops;       % a TabStops view
%
%   Ported from python-docx v1.2.0: src/docx/text/parfmt.py::ParagraphFormat

    properties (Dependent)
        alignment          % WD_ALIGN_PARAGRAPH|[] -- ./w:jc/@val (via CT_PPr.jc_val)
        first_line_indent  % Length|[] -- w:ind firstLine(+)/hanging(-)
        keep_together      % bool|[] -- w:keepLines
        keep_with_next     % bool|[] -- w:keepNext
        left_indent        % Length|[] -- w:ind/@w:left
        line_spacing       % float|Length|[] -- w:spacing @w:line/@w:lineRule
        line_spacing_rule  % WD_LINE_SPACING|[] -- w:spacing @w:lineRule
        page_break_before  % bool|[] -- w:pageBreakBefore
        right_indent       % Length|[] -- w:ind/@w:right
        space_after        % Length|[] -- w:spacing/@w:after
        space_before       % Length|[] -- w:spacing/@w:before
        tab_stops          % TabStops (read-only, lazy) -- w:tabs
        widow_control      % bool|[] -- w:widowControl
    end

    properties (Access = private)
        % @lazyproperty tab_stops cache (mat2doc.shared.lazyproperty idiom):
        % NEVER use isempty on the cache as the sentinel (H3) -- the computed
        % flag is the sole sentinel.
        tab_stops_cache_
        tab_stops_isComputed_ (1, 1) logical = false
    end

    methods
        function obj = ParagraphFormat(element, parent)
            % PARAGRAPHFORMAT Wrap a pPr-owner element (parfmt.py; via ElementProxy).
            %
            %   Inputs:  element - the oxml element carrying <w:pPr> (a
            %                      mat2doc.oxml.text.CT_P or a style pPr owner).
            %            parent  - (optional) an object providing `part`;
            %                      default [] (None).
            %   Outputs: obj     - a scalar ParagraphFormat handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/shared.py::ElementProxy.__init__
            arguments
                element
                parent = []   % None sentinel (H3)
            end
            obj@mat2doc.shared.ElementProxy(element, parent);
        end

        % ---- alignment (parfmt.py 12-28) ----
        function value = get.alignment(obj)
            % Python (parfmt.py 20-23): pPr None-guard; else pPr.jc_val.
            pPr = obj.element_.pPr;
            if isequal(pPr, [])          % Python: if pPr is None
                value = [];
                return
            end
            value = pPr.jc_val;
        end
        function set.alignment(obj, value)
            % Python (parfmt.py 26-28): pPr = get_or_add_pPr(); pPr.jc_val = value.
            pPr = obj.element_.get_or_add_pPr();
            pPr.jc_val = value;
        end

        % ---- first_line_indent (parfmt.py 30-47) ----
        function value = get.first_line_indent(obj)
            % Python (parfmt.py 39-42): pPr None-guard; else pPr.first_line_indent.
            pPr = obj.element_.pPr;
            if isequal(pPr, [])          % Python: if pPr is None
                value = [];
                return
            end
            value = pPr.first_line_indent;
        end
        function set.first_line_indent(obj, value)
            % Python (parfmt.py 45-47): pPr.first_line_indent = value.
            pPr = obj.element_.get_or_add_pPr();
            pPr.first_line_indent = value;
        end

        % ---- keep_together (parfmt.py 49-63) ----
        function value = get.keep_together(obj)
            % Python (parfmt.py 56-59): pPr None-guard; else pPr.keepLines_val.
            pPr = obj.element_.pPr;
            if isequal(pPr, [])          % Python: if pPr is None
                value = [];
                return
            end
            value = pPr.keepLines_val;
        end
        function set.keep_together(obj, value)
            % Python (parfmt.py 62-63): get_or_add_pPr().keepLines_val = value.
            obj.element_.get_or_add_pPr().keepLines_val = value;
        end

        % ---- keep_with_next (parfmt.py 65-81) ----
        function value = get.keep_with_next(obj)
            % Python (parfmt.py 74-77): pPr None-guard; else pPr.keepNext_val.
            pPr = obj.element_.pPr;
            if isequal(pPr, [])          % Python: if pPr is None
                value = [];
                return
            end
            value = pPr.keepNext_val;
        end
        function set.keep_with_next(obj, value)
            % Python (parfmt.py 80-81): get_or_add_pPr().keepNext_val = value.
            obj.element_.get_or_add_pPr().keepNext_val = value;
        end

        % ---- left_indent (parfmt.py 83-100) ----
        function value = get.left_indent(obj)
            % Python (parfmt.py 92-95): pPr None-guard; else pPr.ind_left.
            pPr = obj.element_.pPr;
            if isequal(pPr, [])          % Python: if pPr is None
                value = [];
                return
            end
            value = pPr.ind_left;
        end
        function set.left_indent(obj, value)
            % Python (parfmt.py 98-100): pPr.ind_left = value.
            pPr = obj.element_.get_or_add_pPr();
            pPr.ind_left = value;
        end

        % ---- line_spacing (parfmt.py 102-131) ----
        function value = get.line_spacing(obj)
            % Python (parfmt.py 114-117): pPr None-guard; else
            %   _line_spacing(pPr.spacing_line, pPr.spacing_lineRule).
            pPr = obj.element_.pPr;
            if isequal(pPr, [])          % Python: if pPr is None
                value = [];
                return
            end
            value = mat2doc.text.ParagraphFormat.line_spacing_( ...
                pPr.spacing_line, pPr.spacing_lineRule);
        end
        function set.line_spacing(obj, value)
            % Python (parfmt.py 119-131): dispatch on value type.
            pPr = obj.element_.get_or_add_pPr();
            if isequal(value, [])                            % Python: if value is None
                pPr.spacing_line = [];
                pPr.spacing_lineRule = [];
            elseif isa(value, "mat2doc.shared.Length")       % Python: elif isinstance(value, Length)
                pPr.spacing_line = value;
                % Python: if pPr.spacing_lineRule != WD_LINE_SPACING.AT_LEAST
                %   (None != AT_LEAST -> True). [] guard replicates None !=.
                lineRule = pPr.spacing_lineRule;
                if isequal(lineRule, []) || ...
                        lineRule ~= mat2doc.enum.text.WD_LINE_SPACING.AT_LEAST
                    pPr.spacing_lineRule = mat2doc.enum.text.WD_LINE_SPACING.EXACTLY;
                end
            else                                             % Python: else (a float multiple)
                % D-STYPE-1: a non-Length numeric is a float number of lines.
                % Emu(value * Twips(240)) truncates toward zero (Emu int()).
                pPr.spacing_line = mat2doc.shared.Emu(value * mat2doc.shared.Twips(240));
                pPr.spacing_lineRule = mat2doc.enum.text.WD_LINE_SPACING.MULTIPLE;
            end
        end

        % ---- line_spacing_rule (parfmt.py 133-160) ----
        function value = get.line_spacing_rule(obj)
            % Python (parfmt.py 142-145): pPr None-guard; else
            %   _line_spacing_rule(pPr.spacing_line, pPr.spacing_lineRule).
            pPr = obj.element_.pPr;
            if isequal(pPr, [])          % Python: if pPr is None
                value = [];
                return
            end
            value = mat2doc.text.ParagraphFormat.line_spacing_rule_( ...
                pPr.spacing_line, pPr.spacing_lineRule);
        end
        function set.line_spacing_rule(obj, value)
            % Python (parfmt.py 147-160): SINGLE/ONE_POINT_FIVE/DOUBLE write the
            %   matching Twips line + MULTIPLE; anything else assigns lineRule.
            pPr = obj.element_.get_or_add_pPr();
            if isequal(value, [])                                     % Python: None -> falls to else
                pPr.spacing_lineRule = [];
            elseif value == mat2doc.enum.text.WD_LINE_SPACING.SINGLE
                pPr.spacing_line = mat2doc.shared.Twips(240);
                pPr.spacing_lineRule = mat2doc.enum.text.WD_LINE_SPACING.MULTIPLE;
            elseif value == mat2doc.enum.text.WD_LINE_SPACING.ONE_POINT_FIVE
                pPr.spacing_line = mat2doc.shared.Twips(360);
                pPr.spacing_lineRule = mat2doc.enum.text.WD_LINE_SPACING.MULTIPLE;
            elseif value == mat2doc.enum.text.WD_LINE_SPACING.DOUBLE
                pPr.spacing_line = mat2doc.shared.Twips(480);
                pPr.spacing_lineRule = mat2doc.enum.text.WD_LINE_SPACING.MULTIPLE;
            else                                                     % Python: else (AT_LEAST/EXACTLY/MULTIPLE)
                pPr.spacing_lineRule = value;
            end
        end

        % ---- page_break_before (parfmt.py 162-176) ----
        function value = get.page_break_before(obj)
            % Python (parfmt.py 169-172): pPr None-guard; else pPr.pageBreakBefore_val.
            pPr = obj.element_.pPr;
            if isequal(pPr, [])          % Python: if pPr is None
                value = [];
                return
            end
            value = pPr.pageBreakBefore_val;
        end
        function set.page_break_before(obj, value)
            % Python (parfmt.py 175-176): get_or_add_pPr().pageBreakBefore_val = value.
            obj.element_.get_or_add_pPr().pageBreakBefore_val = value;
        end

        % ---- right_indent (parfmt.py 178-195) ----
        function value = get.right_indent(obj)
            % Python (parfmt.py 187-190): pPr None-guard; else pPr.ind_right.
            pPr = obj.element_.pPr;
            if isequal(pPr, [])          % Python: if pPr is None
                value = [];
                return
            end
            value = pPr.ind_right;
        end
        function set.right_indent(obj, value)
            % Python (parfmt.py 193-195): pPr.ind_right = value.
            pPr = obj.element_.get_or_add_pPr();
            pPr.ind_right = value;
        end

        % ---- space_after (parfmt.py 197-213) ----
        function value = get.space_after(obj)
            % Python (parfmt.py 206-209): pPr None-guard; else pPr.spacing_after.
            pPr = obj.element_.pPr;
            if isequal(pPr, [])          % Python: if pPr is None
                value = [];
                return
            end
            value = pPr.spacing_after;
        end
        function set.space_after(obj, value)
            % Python (parfmt.py 212-213): get_or_add_pPr().spacing_after = value.
            obj.element_.get_or_add_pPr().spacing_after = value;
        end

        % ---- space_before (parfmt.py 215-231) ----
        function value = get.space_before(obj)
            % Python (parfmt.py 224-227): pPr None-guard; else pPr.spacing_before.
            pPr = obj.element_.pPr;
            if isequal(pPr, [])          % Python: if pPr is None
                value = [];
                return
            end
            value = pPr.spacing_before;
        end
        function set.space_before(obj, value)
            % Python (parfmt.py 230-231): get_or_add_pPr().spacing_before = value.
            obj.element_.get_or_add_pPr().spacing_before = value;
        end

        % ---- tab_stops (read-only @lazyproperty; parfmt.py 233-238) ----
        function value = get.tab_stops(obj)
            % Python (parfmt.py 236-238): pPr = get_or_add_pPr(); return TabStops(pPr).
            %   @lazyproperty -> cache + computed-flag idiom (side effect of
            %   get_or_add_pPr runs exactly once on first access).
            if ~obj.tab_stops_isComputed_
                pPr = obj.element_.get_or_add_pPr();
                obj.tab_stops_cache_ = mat2doc.text.TabStops(pPr);
                obj.tab_stops_isComputed_ = true;
            end
            value = obj.tab_stops_cache_;
        end

        % ---- widow_control (parfmt.py 240-254) ----
        function value = get.widow_control(obj)
            % Python (parfmt.py 247-250): pPr None-guard; else pPr.widowControl_val.
            pPr = obj.element_.pPr;
            if isequal(pPr, [])          % Python: if pPr is None
                value = [];
                return
            end
            value = pPr.widowControl_val;
        end
        function set.widow_control(obj, value)
            % Python (parfmt.py 253-254): get_or_add_pPr().widowControl_val = value.
            obj.element_.get_or_add_pPr().widowControl_val = value;
        end
    end

    methods (Static, Access = private)
        function value = line_spacing_(spacing_line, spacing_lineRule)
            % LINE_SPACING_ Line spacing from (line, lineRule) (parfmt.py 256-269).
            %   Returns a FLOAT number of lines (D-STYPE-1) when lineRule is
            %   MULTIPLE, else the absolute Length; [] when spacing_line is None.
            %
            %   Ported from python-docx v1.2.0: src/docx/text/parfmt.py::ParagraphFormat._line_spacing
            if isequal(spacing_line, [])         % Python: if spacing_line is None
                value = [];
                return
            end
            % Python: if spacing_lineRule == WD_LINE_SPACING.MULTIPLE
            %   ([] guard replicates Python None == member -> False).
            if ~isequal(spacing_lineRule, []) && ...
                    spacing_lineRule == mat2doc.enum.text.WD_LINE_SPACING.MULTIPLE
                % Python: return spacing_line / Pt(12) -- int/int true division
                % -> float (D-STYPE-1). Explicit double() guarantees the float.
                value = double(spacing_line) / double(mat2doc.shared.Pt(12));
                return
            end
            value = spacing_line;                % absolute Length
        end

        function value = line_spacing_rule_(line, lineRule)
            % LINE_SPACING_RULE_ Rule from (line, lineRule) (parfmt.py 271-286).
            %   Returns the SINGLE/ONE_POINT_FIVE/DOUBLE special members when the
            %   rule is MULTIPLE and line is exactly Twips 240/360/480; else
            %   returns lineRule unchanged.
            %
            %   Ported from python-docx v1.2.0: src/docx/text/parfmt.py::ParagraphFormat._line_spacing_rule
            % Python: if lineRule == WD_LINE_SPACING.MULTIPLE
            %   ([] guard replicates Python None == member -> False).
            if ~isequal(lineRule, []) && ...
                    lineRule == mat2doc.enum.text.WD_LINE_SPACING.MULTIPLE
                % Python: if line == Twips(240) -> SINGLE (Length == Length by EMU).
                %   ([] guard replicates Python None == Twips(...) -> False.)
                if ~isequal(line, []) && line == mat2doc.shared.Twips(240)
                    value = mat2doc.enum.text.WD_LINE_SPACING.SINGLE;
                    return
                end
                if ~isequal(line, []) && line == mat2doc.shared.Twips(360)
                    value = mat2doc.enum.text.WD_LINE_SPACING.ONE_POINT_FIVE;
                    return
                end
                if ~isequal(line, []) && line == mat2doc.shared.Twips(480)
                    value = mat2doc.enum.text.WD_LINE_SPACING.DOUBLE;
                    return
                end
            end
            value = lineRule;
        end
    end
end

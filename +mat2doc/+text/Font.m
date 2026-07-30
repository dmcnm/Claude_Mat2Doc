classdef Font < mat2doc.shared.ElementProxy
% FONT Character properties: font name, size, bold, italic, subscript, color.
%
%   Ported from python-docx v1.2.0: src/docx/text/font.py::Font (lines 17-428).
%   Proxy for the parent of a `<w:rPr>` element (a `w:r` / CT_R). An
%   ElementProxy subclass: reference semantics (handle) and H5 element-identity
%   eq/ne are inherited unchanged. Every accessor reaches through the wrapped
%   run to its run-properties (`self._element.rPr` / `get_or_add_rPr`) and
%   delegates to the CT_RPr helpers ported at P4-1a -- Font adds NO oxml logic,
%   NO registry rows and NO serialization code (API/proxy tier).
%
%   CONSTRUCTOR (font.py 21-24): `Font(r, parent=None)`. Python
%   `super().__init__(r, parent); self._element = r; self._r = r`. The
%   `_element` re-assign is redundant with the super call; `_r` (-> r_) is set
%   but NEVER read anywhere in font.py -- both ported verbatim for fidelity.
%
%   TRI-STATE (H3): all ~27 properties are None-vs-value. None (the setting is
%   inherited from the style hierarchy) is the sentinel [] on get and on set,
%   via inline isequal(x, []) (the established Mat2Doc None idiom).
%
%   BOOLEAN PROPERTIES (font.py 26-416): 20 read/write tri-state booleans, each
%   delegating to _get_bool_prop / _set_bool_prop (-> get_bool_prop_ /
%   set_bool_prop_), which in turn call CT_RPr.get_bool_val_ / set_bool_val_
%   (P4-1a). The `name` argument each passes is the CT_RPr *property* name
%   (the w: local tag minus prefix), which CT_RPr.get_bool_val_ resolves via
%   obj.(name). The Font-property -> CT_RPr-property map (verified vs font.py):
%       all_caps->caps  bold->b  complex_script->cs  cs_bold->bCs
%       cs_italic->iCs  double_strike->dstrike  emboss->emboss  hidden->vanish
%       imprint->imprint  italic->i  math->oMath  no_proof->noProof
%       outline->outline  rtl->rtl  shadow->shadow  small_caps->smallCaps
%       snap_to_grid->snapToGrid  spec_vanish->specVanish  strike->strike
%       web_hidden->webHidden
%
%   NON-BOOLEAN PROPERTIES:
%     * color (read-only, font.py 50-54): returns a FRESH ColorFormat wrapping
%       the run each access (Python `return ColorFormat(self._element)`; not
%       cached).
%     * highlight_color (font.py 133-144): get/set WD_COLOR_INDEX via
%       CT_RPr.highlight_val.
%     * name (font.py 184-200): get reads rFonts_ascii; set writes BOTH
%       rFonts_ascii AND rFonts_hAnsi (the ascii/hAnsi asymmetry -- and the
%       hAnsi-doesn't-remove-rFonts quirk -- lives inside CT_RPr, P4-1a).
%     * size (font.py 254-278): get returns CT_RPr.sz_val (a Length in EMU,
%       converted from the sz half-point value by ST_HpsMeasure at P4-1a); set
%       wraps the arg in Emu() and assigns sz_val ([] -> remove). D-STYPE-1
%       (int/float indistinguishability on Length) applies -- adopt-only.
%     * subscript / superscript (font.py 334-367): get/set CT_RPr.subscript /
%       .superscript (the w:vertAlign truth table lives in CT_RPr, P4-1a).
%     * underline (font.py 369-403): the WD_UNDERLINE tri-state-plus-enum. See
%       the get/set bodies for the INHERITED/SINGLE/NONE special-casing.
%
%   Example:
%       r = mat2doc.oxml.OxmlElement("w:r");
%       f = mat2doc.text.Font(r);
%       f.bold = true;                                       % <w:rPr><w:b/></w:rPr>
%       f.size = mat2doc.shared.Pt(12);                      % <w:sz w:val="24"/>
%       f.name = "Calibri";                                  % rFonts @w:ascii/@w:hAnsi
%       f.underline = mat2doc.enum.text.WD_UNDERLINE.DOUBLE; % <w:u w:val="double"/>
%       f.color.rgb = mat2doc.shared.RGBColor(255, 0, 0);    % <w:color w:val="FF0000"/>
%
%   Ported from python-docx v1.2.0: src/docx/text/font.py::Font

    properties (Access = private)
        r_   % Python self._r (font.py 24): set from the run, NEVER read in font.py
    end

    properties (Dependent)
        all_caps        % bool|[] -- w:caps
        bold            % bool|[] -- w:b
        color           % ColorFormat (read-only) -- font.color
        complex_script  % bool|[] -- w:cs
        cs_bold         % bool|[] -- w:bCs
        cs_italic       % bool|[] -- w:iCs
        double_strike   % bool|[] -- w:dstrike
        emboss          % bool|[] -- w:emboss
        hidden          % bool|[] -- w:vanish
        highlight_color % WD_COLOR_INDEX|[] -- w:highlight
        italic          % bool|[] -- w:i
        imprint         % bool|[] -- w:imprint
        math            % bool|[] -- w:oMath
        name            % string|[] -- w:rFonts @w:ascii/@w:hAnsi
        no_proof        % bool|[] -- w:noProof
        outline         % bool|[] -- w:outline
        rtl             % bool|[] -- w:rtl
        shadow          % bool|[] -- w:shadow
        size            % Length|[] -- w:sz (EMU)
        small_caps      % bool|[] -- w:smallCaps
        snap_to_grid    % bool|[] -- w:snapToGrid
        spec_vanish     % bool|[] -- w:specVanish
        strike          % bool|[] -- w:strike
        subscript       % bool|[] -- w:vertAlign == subscript
        superscript     % bool|[] -- w:vertAlign == superscript
        underline       % bool|WD_UNDERLINE|[] -- w:u
        web_hidden      % bool|[] -- w:webHidden
    end

    methods
        function obj = Font(r, parent)
            % FONT Wrap a run element `r` (font.py 21-24).
            %
            %   Inputs:  r      - a mat2doc.oxml.text.CT_R (the `w:r` whose
            %                     `w:rPr` carries the character properties).
            %            parent - (optional) an object providing `part`;
            %                     default [] (None).
            %   Outputs: obj    - a scalar Font handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/text/font.py::Font.__init__
            arguments
                r
                parent = []   % None sentinel (H3)
            end
            obj@mat2doc.shared.ElementProxy(r, parent);  % Python: super().__init__(r, parent)
            obj.element_ = r;   % Python: self._element = r (redundant, faithful)
            obj.r_ = r;         % Python: self._r = r (set but unused in font.py)
        end

        % ---- all_caps (font.py 26-36) ----
        function value = get.all_caps(obj);        value = obj.get_bool_prop_("caps"); end
        function set.all_caps(obj, value);         obj.set_bool_prop_("caps", value); end

        % ---- bold (font.py 38-48) ----
        function value = get.bold(obj);            value = obj.get_bool_prop_("b"); end
        function set.bold(obj, value);             obj.set_bool_prop_("b", value); end

        % ---- color (read-only; font.py 50-54) ----
        function value = get.color(obj)
            % Python: return ColorFormat(self._element) -- a fresh proxy each access.
            value = mat2doc.dml.ColorFormat(obj.element_);
        end

        % ---- complex_script (font.py 56-67) ----
        function value = get.complex_script(obj);  value = obj.get_bool_prop_("cs"); end
        function set.complex_script(obj, value);   obj.set_bool_prop_("cs", value); end

        % ---- cs_bold (font.py 69-80) ----
        function value = get.cs_bold(obj);         value = obj.get_bool_prop_("bCs"); end
        function set.cs_bold(obj, value);          obj.set_bool_prop_("bCs", value); end

        % ---- cs_italic (font.py 82-93) ----
        function value = get.cs_italic(obj);       value = obj.get_bool_prop_("iCs"); end
        function set.cs_italic(obj, value);        obj.set_bool_prop_("iCs", value); end

        % ---- double_strike (font.py 95-105) ----
        function value = get.double_strike(obj);   value = obj.get_bool_prop_("dstrike"); end
        function set.double_strike(obj, value);    obj.set_bool_prop_("dstrike", value); end

        % ---- emboss (font.py 107-118) ----
        function value = get.emboss(obj);          value = obj.get_bool_prop_("emboss"); end
        function set.emboss(obj, value);           obj.set_bool_prop_("emboss", value); end

        % ---- hidden (font.py 120-131) ----
        function value = get.hidden(obj);          value = obj.get_bool_prop_("vanish"); end
        function set.hidden(obj, value);           obj.set_bool_prop_("vanish", value); end

        % ---- highlight_color (font.py 133-144) ----
        function value = get.highlight_color(obj)
            % Python (font.py 136-139): rPr None-guard; else rPr.highlight_val.
            rPr = obj.element_.rPr;
            if isequal(rPr, [])                 % Python: if rPr is None
                value = [];
                return
            end
            value = rPr.highlight_val;
        end
        function set.highlight_color(obj, value)
            % Python (font.py 142-144): rPr = get_or_add_rPr(); rPr.highlight_val = value.
            rPr = obj.element_.get_or_add_rPr();
            rPr.highlight_val = value;
        end

        % ---- italic (font.py 146-157) ----
        function value = get.italic(obj);          value = obj.get_bool_prop_("i"); end
        function set.italic(obj, value);           obj.set_bool_prop_("i", value); end

        % ---- imprint (font.py 159-169) ----
        function value = get.imprint(obj);         value = obj.get_bool_prop_("imprint"); end
        function set.imprint(obj, value);          obj.set_bool_prop_("imprint", value); end

        % ---- math (font.py 171-182) ----
        function value = get.math(obj);            value = obj.get_bool_prop_("oMath"); end
        function set.math(obj, value);             obj.set_bool_prop_("oMath", value); end

        % ---- name (font.py 184-200) ----
        function value = get.name(obj)
            % Python (font.py 191-194): rPr None-guard; else rPr.rFonts_ascii.
            rPr = obj.element_.rPr;
            if isequal(rPr, [])                 % Python: if rPr is None
                value = [];
                return
            end
            value = rPr.rFonts_ascii;
        end
        function set.name(obj, value)
            % Python (font.py 197-200): set BOTH ascii and hAnsi to `value`.
            rPr = obj.element_.get_or_add_rPr();
            rPr.rFonts_ascii = value;
            rPr.rFonts_hAnsi = value;
        end

        % ---- no_proof (font.py 202-213) ----
        function value = get.no_proof(obj);        value = obj.get_bool_prop_("noProof"); end
        function set.no_proof(obj, value);         obj.set_bool_prop_("noProof", value); end

        % ---- outline (font.py 215-227) ----
        function value = get.outline(obj);         value = obj.get_bool_prop_("outline"); end
        function set.outline(obj, value);          obj.set_bool_prop_("outline", value); end

        % ---- rtl (font.py 229-239) ----
        function value = get.rtl(obj);             value = obj.get_bool_prop_("rtl"); end
        function set.rtl(obj, value);              obj.set_bool_prop_("rtl", value); end

        % ---- shadow (font.py 241-252) ----
        function value = get.shadow(obj);          value = obj.get_bool_prop_("shadow"); end
        function set.shadow(obj, value);           obj.set_bool_prop_("shadow", value); end

        % ---- size (font.py 254-278) ----
        function value = get.size(obj)
            % Python (font.py 270-273): rPr None-guard; else rPr.sz_val.
            rPr = obj.element_.rPr;
            if isequal(rPr, [])                 % Python: if rPr is None
                value = [];
                return
            end
            value = rPr.sz_val;
        end
        function set.size(obj, emu)
            % Python (font.py 276-278): rPr.sz_val = None if emu is None else Emu(emu).
            rPr = obj.element_.get_or_add_rPr();
            if isequal(emu, [])                 % Python: if emu is None
                rPr.sz_val = [];
            else
                rPr.sz_val = mat2doc.shared.Emu(emu);
            end
        end

        % ---- small_caps (font.py 280-291) ----
        function value = get.small_caps(obj);      value = obj.get_bool_prop_("smallCaps"); end
        function set.small_caps(obj, value);       obj.set_bool_prop_("smallCaps", value); end

        % ---- snap_to_grid (font.py 293-304) ----
        function value = get.snap_to_grid(obj);    value = obj.get_bool_prop_("snapToGrid"); end
        function set.snap_to_grid(obj, value);     obj.set_bool_prop_("snapToGrid", value); end

        % ---- spec_vanish (font.py 306-319) ----
        function value = get.spec_vanish(obj);     value = obj.get_bool_prop_("specVanish"); end
        function set.spec_vanish(obj, value);      obj.set_bool_prop_("specVanish", value); end

        % ---- strike (font.py 321-332) ----
        function value = get.strike(obj);          value = obj.get_bool_prop_("strike"); end
        function set.strike(obj, value);           obj.set_bool_prop_("strike", value); end

        % ---- subscript (font.py 334-349) ----
        function value = get.subscript(obj)
            % Python (font.py 341-344): rPr None-guard; else rPr.subscript.
            rPr = obj.element_.rPr;
            if isequal(rPr, [])                 % Python: if rPr is None
                value = [];
                return
            end
            value = rPr.subscript;
        end
        function set.subscript(obj, value)
            % Python (font.py 347-349): rPr = get_or_add_rPr(); rPr.subscript = value.
            rPr = obj.element_.get_or_add_rPr();
            rPr.subscript = value;
        end

        % ---- superscript (font.py 351-367) ----
        function value = get.superscript(obj)
            % Python (font.py 359-362): rPr None-guard; else rPr.superscript.
            rPr = obj.element_.rPr;
            if isequal(rPr, [])                 % Python: if rPr is None
                value = [];
                return
            end
            value = rPr.superscript;
        end
        function set.superscript(obj, value)
            % Python (font.py 364-367): rPr = get_or_add_rPr(); rPr.superscript = value.
            rPr = obj.element_.get_or_add_rPr();
            rPr.superscript = value;
        end

        % ---- underline (font.py 369-403) ----
        function value = get.underline(obj)
            % Python (font.py 380-392): rPr None-guard; then map rPr.u_val:
            %   None  (u absent, or INHERITED)  -> []   (inherited)
            %   SINGLE                          -> true
            %   NONE                            -> false
            %   any other WD_UNDERLINE member   -> the member itself
            rPr = obj.element_.rPr;
            if isequal(rPr, [])                 % Python: if rPr is None
                value = [];
                return
            end
            val = rPr.u_val;                    % [] (None) or a WD_UNDERLINE member
            if isequal(val, [])
                % Python: val is None -> every `==` below is False -> returns val (None).
                value = [];
            elseif val == mat2doc.enum.text.WD_UNDERLINE.INHERITED
                value = [];                     % Python: None if val == INHERITED
            elseif val == mat2doc.enum.text.WD_UNDERLINE.SINGLE
                value = true;                   % Python: True if val == SINGLE
            elseif val == mat2doc.enum.text.WD_UNDERLINE.NONE
                value = false;                  % Python: False if val == NONE
            else
                value = val;                    % Python: else val
            end
        end
        function set.underline(obj, value)
            % Python (font.py 395-403): map True->SINGLE, False->NONE (strict
            %   `is` identity -- NOT truthiness; a WD_UNDERLINE member or []
            %   passes through unchanged), then assign rPr.u_val.
            rPr = obj.element_.get_or_add_rPr();
            if islogical(value) && isscalar(value) && value          % Python: value is True
                val = mat2doc.enum.text.WD_UNDERLINE.SINGLE;
            elseif islogical(value) && isscalar(value) && ~value     % Python: value is False
                val = mat2doc.enum.text.WD_UNDERLINE.NONE;
            else
                val = value;                    % [] (None) or a WD_UNDERLINE member
            end
            rPr.u_val = val;                    % u_val setter: [] -> remove; member -> set
        end

        % ---- web_hidden (font.py 405-416) ----
        function value = get.web_hidden(obj);      value = obj.get_bool_prop_("webHidden"); end
        function set.web_hidden(obj, value);       obj.set_bool_prop_("webHidden", value); end
    end

    methods (Access = private)
        function value = get_bool_prop_(obj, name)
            % GET_BOOL_PROP_ Value of the boolean `w:rPr` child named `name`
            %   (font.py 418-423). Python `_get_bool_prop`: rPr None-guard,
            %   else rPr._get_bool_val(name). `name` is the CT_RPr property name.
            rPr = obj.element_.rPr;
            if isequal(rPr, [])                 % Python: if rPr is None
                value = [];
                return
            end
            value = rPr.get_bool_val_(name);    % Python: rPr._get_bool_val(name)
        end
        function set_bool_prop_(obj, name, value)
            % SET_BOOL_PROP_ Assign `value` to the boolean `w:rPr` child named
            %   `name` (font.py 425-428). Python `_set_bool_prop`:
            %   rPr = get_or_add_rPr(); rPr._set_bool_val(name, value).
            rPr = obj.element_.get_or_add_rPr();
            rPr.set_bool_val_(name, value);     % Python: rPr._set_bool_val(name, value)
        end
    end
end

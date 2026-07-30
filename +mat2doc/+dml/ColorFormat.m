classdef ColorFormat < mat2doc.shared.ElementProxy
% COLORFORMAT Access to color settings: RGB color, theme color, and type.
%
%   Ported from python-docx v1.2.0: src/docx/dml/color.py::ColorFormat
%   (lines 22-112). An ElementProxy over the `<w:rPr>` PARENT element (a
%   `w:r` / CT_R), NOT over the `<w:color>` element itself: every accessor
%   reaches the color via `self._element.rPr.color` (the private `_color`
%   helper). Reference semantics (handle) inherited from ElementProxy; H5
%   element-identity eq/ne inherited unchanged.
%
%   CONSTRUCTOR (color.py 25-27): `ColorFormat(rPr_parent)` -- a SINGLE
%   positional arg. Python `super().__init__(rPr_parent)` leaves the
%   ElementProxy parent at its None default, then re-assigns
%   `self._element = rPr_parent`. Ported verbatim (the re-assign is redundant
%   with the super call but faithful).
%
%   TYPE-DETECTION (type, color.py 86-101): reads the `w:color` child once
%   (`_color`) and classifies -- THEME if @w:themeColor present, else AUTO if
%   @w:val == "auto", else RGB; [] (None) when there is no color at all.
%   The order matters: themeColor is checked BEFORE val, so a color carrying
%   BOTH a themeColor and an RGB @w:val reports THEME (Word writes the RGB as
%   a "good guess" but the theme color wins) -- ported in exactly that order.
%
%   H3 (None tri-state): every absent-child / absent-attr path returns [] and
%   every set-to-[] path removes, via inline isequal(x, []) (the established
%   Mat2Doc oxml/proxy None idiom -- no shared isNone helper).
%
%   H4/mixed-type compare: Python `color.val == ST_HexColorAuto.AUTO`
%   compares a value that is EITHER an RGBColor OR the string "auto" against
%   the string "auto". `color.val` here comes from CT_Color.val
%   (ST_HexColor.convert_from_xml -> RGBColor object, or the "auto" string).
%   Ported as isequal(color.val, ST_HexColorAuto.AUTO): isequal(RGBColor,
%   "auto") is false (different class, no eq invoked) and isequal("auto",
%   "auto") is true -- reproducing the Python `==` result for both branches.
%
%   REQUIRED-ATTR EDGE (rgb on a themeColor-only color): `rgb` reads
%   `color.val` directly, and `CT_Color.val` is a REQUIRED @w:val attribute.
%   On a parsed `<w:color w:themeColor="accent1"/>` (a themeColor present but
%   NO @w:val), reading `rgb` RAISES mat2doc:InvalidXmlError -- byte-identical
%   to python-docx, which raises the same required-attribute error. `type` and
%   `theme_color` on that same element are SAFE (type checks themeColor before
%   val), so only `rgb` trips the required-attr path.
%
%   Example:
%       r  = mat2doc.oxml.OxmlElement("w:r");
%       cf = mat2doc.dml.ColorFormat(r);
%       cf.rgb = mat2doc.shared.RGBColor(60, 47, 128);   % <w:color w:val="3C2F80"/>
%       cf.type == mat2doc.enum.dml.MSO_COLOR_TYPE.RGB   % true
%       cf.theme_color = mat2doc.enum.dml.MSO_THEME_COLOR.ACCENT_1;
%       cf.type == mat2doc.enum.dml.MSO_COLOR_TYPE.THEME % true
%
%   Ported from python-docx v1.2.0: src/docx/dml/color.py::ColorFormat

    properties (Dependent)
        rgb          % RGBColor value, or [] (None) -- see color.py 29-49
        theme_color  % MSO_THEME_COLOR member, or [] (None) -- see color.py 60-76
        type         % MSO_COLOR_TYPE member (read-only), or [] (None) -- color.py 86-101
    end

    methods
        function obj = ColorFormat(rPr_parent)
            % COLORFORMAT Wrap the `w:rPr` parent element (color.py 25-27).
            %
            %   Inputs:  rPr_parent - the oxml element that owns (or will own)
            %                         a `<w:rPr>`, currently always a `w:r`
            %                         (CT_R). The color is reached through it.
            %   Outputs: obj        - a scalar ColorFormat handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/dml/color.py::ColorFormat.__init__
            obj@mat2doc.shared.ElementProxy(rPr_parent);  % parent defaults to [] (None)
            obj.element_ = rPr_parent;                    % Python: self._element = rPr_parent
        end

        % ---- rgb (color.py 29-58) ----
        function value = get.rgb(obj)
            % Python (color.py 44-49): color = self._color; None-guard; return
            %   None when @w:val == "auto"; else the RGBColor.
            color = obj.color_();
            if isequal(color, [])                       % Python: if color is None
                value = [];
                return
            end
            if isequal(color.val, mat2doc.oxml.simpletypes.ST_HexColorAuto.AUTO)
                value = [];                             % Python: color.val == AUTO -> None
                return
            end
            value = color.val;                          % Python: cast(RGBColor, color.val)
        end
        function set.rgb(obj, value)
            % Python (color.py 51-58): no-op when clearing an already-absent
            %   color; otherwise remove any existing color, then (unless
            %   clearing) add a fresh w:color and set its @w:val.
            if isequal(value, []) && isequal(obj.color_(), [])  % value is None and self._color is None
                return
            end
            rPr = obj.element_.get_or_add_rPr();
            rPr.remove_color_();                        % Python: rPr._remove_color()
            if ~isequal(value, [])                      % Python: if value is not None
                c = rPr.get_or_add_color();
                c.val = value;                          % Python: ....val = value
            end
        end

        % ---- theme_color (color.py 60-84) ----
        function value = get.theme_color(obj)
            % Python (color.py 73-76): color = self._color; None-guard; return
            %   color.themeColor.
            color = obj.color_();
            if isequal(color, [])                       % Python: if color is None
                value = [];
                return
            end
            value = color.themeColor;
        end
        function set.theme_color(obj, value)
            % Python (color.py 78-84): clearing removes any color (only when a
            %   color AND an rPr exist); otherwise force a color and set its
            %   @w:themeColor.
            if isequal(value, [])                       % Python: if value is None
                if ~isequal(obj.color_(), []) && ~isequal(obj.element_.rPr, [])
                    obj.element_.rPr.remove_color_();   % Python: ...rPr._remove_color()
                end
                return
            end
            % Python: self._element.get_or_add_rPr().get_or_add_color().themeColor = value
            c = obj.element_.get_or_add_rPr().get_or_add_color();
            c.themeColor = value;
        end

        % ---- type (read-only; color.py 86-101) ----
        function value = get.type(obj)
            % Python (color.py 94-101): None when no color; THEME when a theme
            %   color is set (checked FIRST); AUTO when @w:val == "auto"; else
            %   RGB.
            color = obj.color_();
            if isequal(color, [])                       % Python: if color is None
                value = [];
                return
            end
            if ~isequal(color.themeColor, [])           % Python: if color.themeColor is not None
                value = mat2doc.enum.dml.MSO_COLOR_TYPE.THEME;
                return
            end
            if isequal(color.val, mat2doc.oxml.simpletypes.ST_HexColorAuto.AUTO)
                value = mat2doc.enum.dml.MSO_COLOR_TYPE.AUTO;  % Python: color.val == AUTO
                return
            end
            value = mat2doc.enum.dml.MSO_COLOR_TYPE.RGB;
        end
    end

    methods (Access = private)
        function c = color_(obj)
            % COLOR_ Return `w:rPr/w:color` or [] (None) if not present.
            %   Python `_color` @property (color.py 103-112); leading-underscore
            %   rotation -> trailing (design.md section 2).
            %
            %   Ported from python-docx v1.2.0: src/docx/dml/color.py::ColorFormat._color
            rPr = obj.element_.rPr;
            if isequal(rPr, [])                         % Python: if rPr is None
                c = [];
                return
            end
            c = rPr.color;
        end
    end
end

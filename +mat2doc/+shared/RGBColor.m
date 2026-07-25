classdef RGBColor
% RGBCOLOR Immutable value object defining a particular RGB color.
%
%   Holds a red/green/blue triplet of integers in the range 0-255.
%
%   DOCX DELTAS vs the Mat2Ppt design guide (python-docx v1.2.0 is the
%   source of truth):
%     - Location: python-docx defines RGBColor in shared.py (python-pptx
%       had it in dml/color.py), so the port lives in +mat2doc\+shared\.
%     - Error split: docx __new__ raises TypeError for a non-int component
%       and ValueError for an out-of-range component (pptx conflated both to
%       ValueError). Both are honored here as mat2doc:TypeError /
%       mat2doc:ValueError (D-004 namespace).
%     - Repr: docx defines __repr__ = "RGBColor(0x%02x, 0x%02x, 0x%02x)"
%       (pptx had none); ported as repr_().
%
%   VALUE SEMANTICS (H2/H5-adjacent): the Python original is a `tuple`
%   subclass (``class RGBColor(Tuple[int, int, int])``), an immutable value
%   with tuple equality-by-value. The port is a MATLAB VALUE class (not a
%   handle) holding the three components in immutable properties; eq/ne
%   compare by value so two RGBColor instances with the same triplet are
%   equal, and a comparison against a non-RGBColor is false (the MATLAB
%   analogue of tuple ==).
%
%   INT/FLOAT INDISTINGUISHABILITY (D-STYPE-1): a MATLAB double cannot
%   distinguish 255 int from 255.0 float, so an integer-VALUED real numeric
%   scalar is accepted as the Python int; a non-integral value is treated as
%   "not an int" and raises TypeError (mirroring isinstance(val, int)).
%
%   HEX FORMATTING: str_ mirrors Python __str__ = "%02X%02X%02X" % self --
%   a fixed-width UPPERCASE hex string (e.g. "3C2F80"). This is a direct
%   printf-style hex format, NOT a decimal serialization, so it does not
%   route through a pyStr numeric helper (H14 applies to str()/repr of
%   numbers, not to %X hex). char returns the same value as a char row
%   vector for display.
%
%   Example:
%       c = mat2doc.shared.RGBColor(60, 47, 128);
%       c.str_()                                % "3C2F80"
%       c.repr_()                               % "RGBColor(0x3c, 0x2f, 0x80)"
%       d = mat2doc.shared.RGBColor.from_string("3C2F80");
%       c == d                                  % true (value equality)
%
%   Ported from python-docx v1.2.0: src/docx/shared.py::RGBColor

    properties (SetAccess = immutable)
        r   % red component, integer 0-255
        g   % green component, integer 0-255
        b   % blue component, integer 0-255
    end

    methods
        function obj = RGBColor(r, g, b)
            % RGBCOLOR Construct from three integer values 0-255 (shared.py 127-134).
            %
            %   Python __new__ iterates (r, g, b) and for each val raises
            %   TypeError if not isinstance(val, int), else ValueError if
            %   val < 0 or val > 255. MATLAB has no distinct int literal, so
            %   an integer-VALUED real numeric scalar in range is accepted
            %   (D-STYPE-1); a non-integral / non-numeric value is "not an
            %   int" -> TypeError.
            %
            %   Inputs:  r, g, b - integer-valued numeric scalars, each 0-255.
            %   Outputs: obj     - a scalar RGBColor value.
            %
            %   Ported from python-docx v1.2.0: src/docx/shared.py::RGBColor.__new__
            if nargin == 0
                % No-arg path so MATLAB can build class prototypes / empties.
                return
            end
            msg = "RGBColor() takes three integer values 0-255";
            comps = {r, g, b};
            for i = 1:3
                v = comps{i};
                % Python: if not isinstance(val, int): raise TypeError(msg)
                if ~(isnumeric(v) && isscalar(v) && isreal(v) && isfinite(v) ...
                        && v == floor(v))
                    error("mat2doc:TypeError", "%s", msg);
                end
                % Python: if val < 0 or val > 255: raise ValueError(msg)
                if v < 0 || v > 255
                    error("mat2doc:ValueError", "%s", msg);
                end
            end
            obj.r = double(r);
            obj.g = double(g);
            obj.b = double(b);
        end

        function s = str_(obj)
            % STR_ Hex string rgb value, like "3C2F80" (Python __str__, shared.py 139-141).
            %
            %   Python: "%02X%02X%02X" % self -- fixed two-digit UPPERCASE
            %   hex for each component.
            %
            %   Ported from python-docx v1.2.0: src/docx/shared.py::RGBColor.__str__
            s = string(sprintf("%02X%02X%02X", obj.r, obj.g, obj.b));
        end

        function s = repr_(obj)
            % REPR_ Python repr, like "RGBColor(0x3c, 0x2f, 0x80)" (shared.py 136-137).
            %
            %   Python: "RGBColor(0x%02x, 0x%02x, 0x%02x)" % self -- fixed
            %   two-digit LOWERCASE hex with a 0x prefix per component.
            %
            %   Ported from python-docx v1.2.0: src/docx/shared.py::RGBColor.__repr__
            s = string(sprintf("RGBColor(0x%02x, 0x%02x, 0x%02x)", ...
                obj.r, obj.g, obj.b));
        end

        function c = char(obj)
            % CHAR Char row-vector form of the hex string (for display/convenience).
            c = char(obj.str_());
        end

        function tf = eq(a, b)
            % EQ True iff a and b are RGBColor values with the same triplet.
            %
            %   Mirrors Python tuple equality: RGBColor compares by value; a
            %   comparison against a non-RGBColor is false.
            if ~isa(a, "mat2doc.shared.RGBColor") || ~isa(b, "mat2doc.shared.RGBColor")
                tf = false;
                return
            end
            tf = isequal(a.r, b.r) && isequal(a.g, b.g) && isequal(a.b, b.b);
        end

        function tf = ne(a, b)
            % NE Complement of eq.
            tf = ~eq(a, b);
        end
    end

    methods (Static)
        function obj = from_string(rgb_hex_str)
            % FROM_STRING New instance from an RGB hex string like "3C2F80" (shared.py 143-149).
            %
            %   Python: r = int(rgb_hex_str[:2], 16); g = int(rgb_hex_str[2:4], 16);
            %   b = int(rgb_hex_str[4:], 16); return cls(r, g, b) -- slices the
            %   6-char hex string into three base-16 byte values.
            %
            %   Ported from python-docx v1.2.0: src/docx/shared.py::RGBColor.from_string
            s = char(rgb_hex_str);
            r = mat2doc.shared.RGBColor.parseHexByte_(s(1:2));    % rgb_hex_str[:2]
            g = mat2doc.shared.RGBColor.parseHexByte_(s(3:4));    % rgb_hex_str[2:4]
            b = mat2doc.shared.RGBColor.parseHexByte_(s(5:end));  % rgb_hex_str[4:]
            obj = mat2doc.shared.RGBColor(r, g, b);
        end
    end

    methods (Static, Access = private)
        function b = parseHexByte_(slice)
            % Faithful int(slice, 16): a non-hex slice raises the same
            % ValueError CPython does (hex2dec otherwise leaks
            % MATLAB:hex2dec:InvalidCharacters). Message byte-matches Python
            % int('<slice>', 16). hex2dec is base MATLAB (no toolbox) and is
            % case-insensitive, matching int(_, 16).
            if isempty(regexp(slice, '^[0-9A-Fa-f]+$', 'once'))
                error("mat2doc:ValueError", ...
                    "invalid literal for int() with base 16: '%s'", slice);
            end
            b = hex2dec(slice);
        end
    end
end

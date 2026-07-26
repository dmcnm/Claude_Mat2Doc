function s = pyStr(value, kind)
% PYSTR Format a value exactly as Python str() would (H14).
%
%   s = MAT2DOC.SHARED.PYSTR(value) returns a string scalar matching what
%   Python 3 str() produces for the corresponding Python value. This is the
%   ONLY permitted numeric->text conversion at XML serialization sites; raw
%   num2str / sprintf('%g') there is a defect (H14, design.md section 8).
%
%   s = MAT2DOC.SHARED.PYSTR(value, kind) forces the Python-side numeric
%   type, kind = "auto" (default) | "int" | "float".
%
%   Behavior by class of value:
%     string/char       - passed through unchanged (Python str(s) is s)
%     logical           - "True" / "False" (Python str(bool))
%     mat2doc.shared.Length and subclasses
%                       - always Python int formatting, regardless of kind
%                         (Length subclasses int in Python, so str() uses
%                         int.__str__)
%     int8..uint64      - Python int formatting: digits, no decimal point,
%                         '-' sign only ("914400", "-914")
%     double            - kind "int":   Python int formatting (value must be
%                                       integral and finite)
%                         kind "float": Python float repr - shortest string
%                                       of significant digits that round-trips
%                                       to the exact same IEEE double,
%                                       positional for decimal exponent
%                                       -4..15 ("0.5", "2.0", "0.0001"),
%                                       scientific otherwise ("1e-05",
%                                       "1.5e+16"); "inf"/"-inf"/"nan";
%                                       "-0.0" for negative zero
%                         kind "auto":  integral finite values format as int,
%                                       all others as float
%
%   DOCUMENTED LIMITS:
%     - MATLAB cannot distinguish the double 2.0 from the integer 2, so
%       "auto" formats integral doubles as Python int str ("2"). Where the
%       Python source value is a float (e.g. str(2.0) -> '2.0'), the call
%       site MUST pass kind = "float".
%     - Only double-precision floats are supported (Python floats are
%       doubles); single input is an error rather than a silent
%       approximation.
%     - Shortest-repr search uses correctly-rounded sprintf('%.*e') /
%       str2double round-trips over 1..17 significant digits, the same
%       shortest-round-trip contract as CPython's float_repr.
%
%   Inputs:  value - string/char/logical/integer/double/Length
%            kind  - "auto" (default) | "int" | "float"
%   Outputs: s - string scalar
%
%   Example:
%       mat2doc.shared.pyStr(mat2doc.shared.Emu(914400))   % "914400"
%       mat2doc.shared.pyStr(0.5, "float")                 % "0.5"
%       mat2doc.shared.pyStr(2.0, "float")                 % "2.0"
%
%   Mat2Doc infrastructure (no python-docx counterpart). Mandated by
%   design.md section 8 (numeric->XML formatting) and hazard H14. Re-ported
%   faithfully from Mat2Ppt +util\pyStr.m (design.md section 7: shared idioms
%   are re-ported into each toolbox's home package, no shared code). First
%   established for Mat2Doc by P1-7 (coreprops), the first numeric-serialization
%   site in the docx port.

arguments
    value
    kind (1, 1) string {mustBeMember(kind, ["auto", "int", "float"])} = "auto"
end

% --- str passthrough -------------------------------------------------
if isstring(value) || ischar(value)
    s = string(value);
    return
end

% --- bool ------------------------------------------------------------
if islogical(value)
    if ~isscalar(value)
        error("mat2doc:pyStr:notScalar", "pyStr requires a scalar value.");
    end
    if value
        s = "True";
    else
        s = "False";
    end
    return
end

if ~isnumeric(value) || ~isscalar(value)
    error("mat2doc:pyStr:unsupportedType", ...
        "pyStr supports string/char, logical, integer, double, and Length values; got %s.", ...
        class(value));
end

% --- Length: Python Length subclasses int, str() -> int digits --------
if isa(value, "mat2doc.shared.Length")
    s = intStr(double(value));
    return
end

% --- native integer classes -------------------------------------------
if isinteger(value)
    s = string(sprintf("%d", value));
    return
end

if isa(value, "single")
    error("mat2doc:pyStr:unsupportedType", ...
        "pyStr does not accept single; convert deliberately to double first.");
end

% --- double: dispatch on kind -----------------------------------------
v = double(value);
switch kind
    case "int"
        if ~isfinite(v) || v ~= fix(v)
            error("mat2doc:pyStr:notIntegral", ...
                "kind=""int"" requires an integral finite value; got %.17g.", v);
        end
        s = intStr(v);
    case "float"
        s = floatRepr(v);
    otherwise % "auto"
        if isfinite(v) && v == fix(v)
            s = intStr(v);
        else
            s = floatRepr(v);
        end
end
end

function s = intStr(v)
% Python str(int): digits with '-' sign only, no decimal point.
if v == 0
    v = 0; % normalize -0.0 -> '0' (Python int has no negative zero)
end
s = string(sprintf("%.0f", v));
end

function s = floatRepr(x)
% Python repr(float) / str(float): shortest round-tripping representation.
if isnan(x)
    s = "nan"; % Python str(float('nan'))
    return
end
if isinf(x)
    if x > 0
        s = "inf";
    else
        s = "-inf";
    end
    return
end
if x == 0
    if 1 / x < 0
        s = "-0.0"; % negative zero: Python repr(-0.0)
    else
        s = "0.0";
    end
    return
end

neg = x < 0;
ax = abs(x);

% Shortest significant-digit count 1..17 whose decimal round-trips to the
% exact same double (CPython float_repr contract).
digits = '';
E = 0;
for p = 1:17
    s0 = sprintf('%.*e', p - 1, ax);
    if str2double(s0) == ax
        % group 2 always participates (may capture ''), so tok is 1x3
        tok = regexp(s0, '^(\d)\.?(\d*)e([+-]\d+)$', 'tokens', 'once');
        digits = [tok{1}, tok{2}];
        E = str2double(tok{3}); % decimal exponent of leading digit
        break
    end
end
assert(~isempty(digits), "mat2doc:pyStr:internal", "shortest-repr search failed");

nd = numel(digits);
if E >= -4 && E < 16
    % Positional notation (CPython: -4 <= exponent < 16).
    if E >= nd - 1
        body = [digits, repmat('0', 1, E - nd + 1), '.0'];
    elseif E >= 0
        body = [digits(1:E + 1), '.', digits(E + 2:end)];
    else
        body = ['0.', repmat('0', 1, -E - 1), digits];
    end
else
    % Scientific notation: mantissa digit[.digits], exponent sign + >=2 digits.
    if nd == 1
        mant = digits;
    else
        mant = [digits(1), '.', digits(2:end)];
    end
    body = [mant, sprintf('e%+03d', E)];
end

if neg
    s = "-" + string(body);
else
    s = string(body);
end
end

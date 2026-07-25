function d = pyIntArg(value, strMode)
% PYINTARG Python int()-argument domain conversion for Length constructors.
%
%   d = PYINTARG(value, strMode) accepts value exactly where CPython 3.13
%   int(x) accepts x and errors where int(x) raises, for the Length-family
%   constructors (Length, Emu, Inches, Cm, Mm, Pt, Twips):
%
%       numeric scalar   -> double(value), untruncated (the constructor's
%                           int() site applies fix()); non-finite raises
%                           like Python int() does (identifier
%                           MATLAB:validators:mustBeFinite; Python raises
%                           ValueError on nan / OverflowError on inf - the
%                           raising decision is faithful, D-004)
%       logical scalar   -> 0/1 (Python bool is an int subclass:
%                           int(True) == 1, int(False) == 0)
%       char row/string  -> strMode "parse" (Length, Emu -- upstream applies
%                           int() directly to the raw argument; LIVE path is
%                           the docx oxml simpletypes ST_ length lexer which
%                           returns Emu(str_value)):
%                           base-10 int-literal PARSE per int(str) --
%                           optional surrounding whitespace, optional single
%                           +/- sign, digits with single underscores strictly
%                           between digits. int(str) PARSES, it never
%                           truncates: '2.5' raises mat2doc:ValueError
%                           (contrast int(2.5) -> 2), message replicating
%                           Python's "invalid literal for int() with base
%                           10: '...'".
%                           strMode "reject" (Inches, Cm, Mm, Pt, Twips --
%                           upstream computes int(str * K), string
%                           REPETITION, which raises ValueError for every
%                           realistic string in CPython 3.13): raises
%                           mat2doc:ValueError.
%       anything else    -> mat2doc:TypeError (Python: TypeError)
%
%   Inputs:  value   - constructor argument (see domain above)
%            strMode - "parse" or "reject" (string handling, see above)
%   Outputs: d       - double scalar, the value int() would operate on
%
%   Example:
%       d = pyIntArg("914400", "parse");   % 914400 (Python int('914400'))
%
%   Mat2Doc infrastructure (shared package-private helper), no python-docx
%   counterpart; replicates CPython 3.13 int() construction semantics.
%   Design carried from the Mat2Ppt +util/private/pyIntArg (no shared code;
%   re-implemented, mat2doc:-namespaced). Adopted rulings D-002/D-003/D-004
%   per decision_2026-07-25_mat2doc_deviation_preadoption.md.

if ischar(value) || isstring(value)
    if strMode == "reject"
        % Python: int(str * K) -- str*int repetition; every string raises
        % ValueError in CPython 3.13 (digit strings once the repeat exceeds
        % the 4300-digit limit -- Twips K=635 keeps a <=6-digit string under
        % that, where CPython builds a garbage repunit int; anything else:
        % "invalid literal..." on the repeated string). Exception type
        % replicated; message is port-specific (D-003).
        error("mat2doc:ValueError", ...
            "string input to this Length constructor is invalid: " + ...
            "Python evaluates int(str * K) (string repetition), which " + ...
            "raises ValueError for every realistic string in CPython 3.13");
    end
    d = parseIntLiteral(value);
elseif islogical(value)
    if ~isscalar(value)
        throwTypeError(value);
    end
    d = double(value);   % Python: int(True) == 1, int(False) == 0
elseif isnumeric(value)
    if ~isscalar(value)
        throwTypeError(value);
    end
    % Python int() raises ValueError on nan, OverflowError on inf; keep the
    % established raising behavior/identifier (D-004).
    mustBeFinite(value);
    d = double(value);
else
    throwTypeError(value);
end
end

function d = parseIntLiteral(value)
% CPython 3.13 int(str) base-10 grammar (ASCII subset -- Python also accepts
% Unicode decimal digits and Unicode whitespace; recorded divergence D-002:
% the XML simpletype lexical space is ASCII-only).
if isstring(value)
    if ~isscalar(value)
        throwTypeError(value);
    end
    if ismissing(value)
        throwTypeError(value);
    end
    s = char(value);
else
    if ~isrow(value) && ~isempty(value)
        throwTypeError(value);   % char matrix has no Python analogue
    end
    s = value;
end
% Strip surrounding ASCII whitespace (Python strips Unicode whitespace).
t = regexprep(s, '^[ \t\n\v\f\r]+', '');
t = regexprep(t, '[ \t\n\v\f\r]+$', '');
% Grammar: optional single sign, digits, single underscores BETWEEN digits
% only ('1_000' ok; '_1', '1_', '1__0', '+_1' all raise, probed).
if isempty(regexp(t, '^[+-]?[0-9](_?[0-9])*$', 'once'))
    error("mat2doc:ValueError", ...
        "invalid literal for int() with base 10: %s", reprStr(s));
end
digits = regexprep(strrep(t, '_', ''), '^[+-]', '');
if strlength(digits) > 4300
    % CPython 3.11+ int-from-str digit limit, replicated verbatim.
    error("mat2doc:ValueError", ...
        "Exceeds the limit (4300 digits) for integer string conversion: " + ...
        "value has %d digits; use sys.set_int_max_str_digits() to " + ...
        "increase the limit", strlength(digits));
end
d = str2double(strrep(t, '_', ''));
if ~isfinite(d)
    % <= 4300 digits but > realmax: Python's arbitrary-precision int
    % succeeds; a double cannot hold it. Raise rather than silently
    % saturate (Length = exact-integer double, design.md section 8;
    % recorded divergence D-002 -- far outside any ST_ range).
    error("mat2doc:ValueError", ...
        "integer string %s exceeds the range representable in a MATLAB " + ...
        "double (Length holds exact-integer doubles, design.md " + ...
        "section 8)", reprStr(s));
end
d = d + 0;   % normalize IEEE -0 from '-0': Python int('-0') is exactly 0
end

function throwTypeError(value)
% Python 3.13 message: "int() argument must be a string, a bytes-like
% object or a real number, not 'list'" -- format replicated with the MATLAB
% class name (non-scalar arrays correspond to Python sequence input).
error("mat2doc:TypeError", ...
    "int() argument must be a string, a bytes-like object or a real " + ...
    "number, not '%s'", class(value));
end

function r = reprStr(s)
% Python str repr, subset sufficient for error messages: single quotes
% unless the string contains ' and no " (then double quotes, probed:
% int("a'b") -> ... base 10: "a'b"); escape backslash, quote, and control
% characters tab/newline/CR/VT/FF as Python repr does.
s = char(s);
if any(s == '''') && ~any(s == '"')
    q = '"';
else
    q = '''';
end
b = strrep(s, '\', '\\');
b = strrep(b, sprintf('\t'), '\t');
b = strrep(b, newline, '\n');
b = strrep(b, sprintf('\r'), '\r');
b = strrep(b, sprintf('\x0B'), '\x0b');
b = strrep(b, sprintf('\f'), '\x0c');
if q == ''''
    b = strrep(b, '''', '\''');
end
r = [q b q];
end

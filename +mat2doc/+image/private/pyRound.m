function y = pyRound(x)
% PYROUND Python 3 built-in round(x) with a single argument (H6).
%
%   y = PYROUND(x) returns x rounded to the nearest integer, ties rounded to
%   the nearest EVEN integer (banker's rounding), exactly as CPython 3's
%   one-argument round() does. MATLAB's built-in round() rounds half AWAY from
%   zero, so it is NOT a faithful substitute at the .5 boundary and is a defect
%   at every ported round() site (H6).
%
%   Used by Image.scaled_dimensions (round(length * scaling_factor)) and by the
%   +image format parsers arriving at P7-1b/P7-2 (Png._dpi / Bmp._dpi compute
%   int(round(px_per_unit * 0.0254))); all route their round() through here so
%   the half-to-even tie behavior matches Python exactly (fix()/int() is then
%   applied by the caller to the already-integral result).
%
%   Inputs:  x - real finite double scalar
%   Outputs: y - integral double, nearest integer with ties to even
%
%   Example:
%       pyRound(2.5)   % 2  (ties to even; MATLAB round(2.5) is 3)
%       pyRound(3.5)   % 4
%       pyRound(-0.5)  % 0
%
%   Mat2Doc infrastructure (package-private helper for +image), no python-docx
%   counterpart; replicates CPython 3 round(x). Mirrors the identical private
%   helper in +oxml/+simpletypes/private/pyRound.m (each package keeps its own
%   private copy; no shared code path). Mandated by H6.

r = floor(x);
d = x - r;
if d < 0.5
    y = r;
elseif d > 0.5
    y = r + 1;
else
    % Exact tie (x is an integer + 0.5): round to the even neighbour.
    if mod(r, 2) == 0
        y = r;
    else
        y = r + 1;
    end
end
% Normalize IEEE -0.0 to +0.0.
y = y + 0;
end

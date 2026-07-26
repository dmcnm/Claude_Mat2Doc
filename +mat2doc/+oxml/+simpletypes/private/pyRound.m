function y = pyRound(x)
% PYROUND Python 3 built-in round(x) with a single argument (H6).
%
%   y = PYROUND(x) returns x rounded to the nearest integer, ties rounded
%   to the nearest EVEN integer (banker's rounding), exactly as CPython 3's
%   one-argument round() does. MATLAB's built-in round() rounds half AWAY
%   from zero, so it is NOT a faithful substitute at the .5 boundary and is
%   a defect at every ported round() site (H6).
%
%   Two docx simpletypes sites route here:
%     * ST_UniversalMeasure.convert_from_xml -> Emu(int(round(quantity *
%       multiplier))) (simpletypes.py 424)
%     * ST_SignedTwipsMeasure.convert_from_xml -> Twips(int(round(float(
%       str_value)))) (simpletypes.py 366)
%
%   The result is an integral double (Python round(x) with no ndigits
%   returns an int); callers apply int()/fix() or wrap in Emu/Twips, for
%   which this integral value is already exact. The rounding replicates
%   Python's round-half-to-even for the finite, in-range domain these simple
%   types operate over (all magnitudes far below 2^53, where every double is
%   an exact integer and the tie test is exact).
%
%   Inputs:  x - real finite double scalar
%   Outputs: y - integral double, nearest integer with ties to even
%
%   Example:
%       pyRound(2.5)   % 2  (ties to even; MATLAB round(2.5) is 3)
%       pyRound(3.5)   % 4
%       pyRound(-0.5)  % 0
%
%   Mat2Doc infrastructure (package-private helper), no python-docx
%   counterpart; replicates CPython 3 round(x) semantics. Mandated by H6.
%   Re-ported faithfully from Mat2Ppt +oxml\+simpletypes\private\pyRound.m
%   (design.md section 7: no shared code between toolboxes).

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
% Normalize IEEE -0.0 (guard the -0.0 input path).
y = y + 0;
end

classdef Length < double
% LENGTH Base class for length classes Inches, Emu, Cm, Mm, Pt, and Twips.
%
%   Provides properties for converting length values to convenient units.
%
%   len = MAT2DOC.SHARED.LENGTH(emu) constructs a length of emu English
%   Metric Units (914400 EMU per inch, 635 EMU per twip).
%
%   Python Length subclasses int: arithmetic on Length values returns a
%   plain int (the Length type is lost); only the constructors re-wrap.
%   This class subclasses double to replicate that exactly: because the
%   class defines properties, every built-in numeric operation (+, -, *,
%   /, floor, ...) returns a plain double (verified on R2024b), while
%   isa(len, 'double') remains true and double(len) is the EMU value.
%   Values are exact integers held in doubles (EMU magnitudes are far
%   below 2^53; design.md section 8).
%
%   NOTE (R2024b): concatenation of Length objects ([a b]) errors because
%   the subclass defines properties. Hold Length values in cell arrays or
%   convert with double() first. Python code never relies on Length
%   surviving aggregation arithmetic, so this is not API-visible.
%
%   Construction domain (Python int() applied to the raw argument): numeric
%   is truncated toward zero (int() -> fix()); logical maps to 0/1 (Python
%   int(True) == 1); a base-10 integer string/char is PARSED like int(str)
%   ('2.5' errors, it is never truncated). The live string path is the docx
%   oxml simpletypes ST_ length lexer, which constructs Emu(str_value) from
%   a raw XML attribute string on essentially every opened document.
%
%   Divergences from the Python original are all on DEAD paths (unreachable
%   through any ported call site, API-invisible); adopted rulings carried
%   from Mat2Ppt per validation\summary\decision_2026-07-25_mat2doc_
%   deviation_preadoption.md (mat2doc:-namespaced):
%     - D-002 (string-input dead paths): string inputs where CPython int()
%       yields exact big/odd integers; the multiplier constructors
%       (Inches/Cm/Mm/Pt/Twips) evaluate int(str * K) (string repetition):
%       for Twips (K=635) a <=6-digit numeric string builds a CPython
%       repunit int the port rejects; the live Emu(str) path carries only
%       bounded ASCII EMU strings, exact there.
%     - D-003 (multiplier-constructor message wording): the string-rejection
%       error MESSAGES of Inches/Cm/Mm/Pt/Twips are port-authored; the
%       exception CLASS (mat2doc:ValueError) is faithful to CPython.
%     - D-004 (error-identifier namespace): non-finite / wrong-type inputs
%       raise MATLAB-native identifiers (MATLAB:validators:mustBeFinite,
%       mat2doc:TypeError) where CPython raises ValueError / OverflowError /
%       TypeError; the RAISING decision is faithful and no upstream except
%       observes the identifier.
%     - D-STYPE-1..4 (int/float indistinguishability): a MATLAB double
%       cannot distinguish 914400 int from 914400.0 float, exactly as in
%       Mat2Ppt; construction accepts an integer-valued double as the int.
%
%   Inputs:  emu - length in EMU: numeric scalar (truncated toward zero
%                  like Python int()), logical scalar (true -> 1), or a
%                  base-10 integer string/char (parsed like Python
%                  int(str); '2.5' errors, it is never truncated)
%   Outputs: len - Length instance
%
%   Dependent properties (read-only), per the Python @property bodies:
%       cm     - floating point length in centimeters
%       emu    - integer length in English Metric Units (returns self)
%       inches - floating point length in inches
%       mm     - floating point length in millimeters
%       pt     - floating point length in points
%       twips  - integer length in twips (a twip is a twentieth of a point)
%
%   Example:
%       len = mat2doc.shared.Length(914400);
%       len.inches   % 1
%
%   Ported from python-docx v1.2.0: src/docx/shared.py::Length

    properties (Constant)
        % Python private class attributes; leading underscore rotated to
        % trailing per design.md section 2 (RATIFIED underscore rotation).
        EMUS_PER_INCH_ = 914400        % _EMUS_PER_INCH
        EMUS_PER_CM_ = 360000          % _EMUS_PER_CM
        EMUS_PER_MM_ = 36000           % _EMUS_PER_MM
        EMUS_PER_PT_ = 12700           % _EMUS_PER_PT
        EMUS_PER_TWIP_ = 635           % _EMUS_PER_TWIP
    end

    properties (Dependent)
        cm       % Floating point length in centimeters.
        emu      % Integer length in English Metric Units.
        inches   % Floating point length in inches.
        mm       % Floating point length in millimeters.
        pt       % Floating point length in points.
        twips    % Integer length in twips (1/20 point, 635 EMU).
    end

    methods
        function obj = Length(emu)
            % Python: int.__new__(cls, emu) - the full int() construction
            % domain:
            %   numeric -> truncate toward zero (H6: int() -> fix());
            %              non-finite raises (Python ValueError/
            %              OverflowError; identifier differs - D-004)
            %   string  -> base-10 int-literal PARSE, never truncation
            %              ('2.5' raises mat2doc:ValueError). LIVE path:
            %              the docx oxml simpletypes ST_ length lexer
            %              constructs Emu(str_value) from a raw XML string.
            %   logical -> 0/1 (Python int(True) == 1)
            % Domain conversion via the shared package-private pyIntArg.
            emu = pyIntArg(emu, "parse");
            % + 0 normalizes IEEE -0.0 -> +0.0: fix() of a negative value in
            % (-1, 0) (e.g. Length(-0.5), or Inches(-1e-7) whose product
            % truncates into that band) yields -0.0, which would surface as
            % "-0.0" through the float getters (.inches/.cm/.mm/.pt), whereas
            % Python int() gives 0 (str() "0.0"). Mirrors the string-path
            % normalization in pyIntArg.
            obj = obj@double(fix(emu) + 0);
        end

        function value = get.cm(obj)
            % Floating point length in centimeters.
            % Python: self / float(self._EMUS_PER_CM) - true division.
            value = double(obj) / mat2doc.shared.Length.EMUS_PER_CM_;
        end

        function value = get.emu(obj)
            % Integer length in English Metric Units.
            % Python: return self - returns the Length instance itself.
            value = obj;
        end

        function value = get.inches(obj)
            % Floating point length in inches.
            % Python: self / float(self._EMUS_PER_INCH) - true division.
            value = double(obj) / mat2doc.shared.Length.EMUS_PER_INCH_;
        end

        function value = get.mm(obj)
            % Floating point length in millimeters.
            % Python: self / float(self._EMUS_PER_MM) - true division.
            value = double(obj) / mat2doc.shared.Length.EMUS_PER_MM_;
        end

        function value = get.pt(obj)
            % Floating point length in points.
            % Python: self / float(self._EMUS_PER_PT) - true division.
            value = double(obj) / mat2doc.shared.Length.EMUS_PER_PT_;
        end

        function value = get.twips(obj)
            % Integer length in twips (a twip is a twentieth of a point).
            % Python: int(round(self / float(self._EMUS_PER_TWIP)))
            % Python 3 round() with no ndigits rounds half-to-EVEN and
            % returns an int; int() of that int is identity. MATLAB round()
            % rounds half AWAY from zero, so the tie case is corrected to
            % even (H6/H14: match Python's rounding exactly). For integer
            % EMU / 635 an exact .5 tie never arises mathematically (635 is
            % odd), but the CPython algorithm is replicated verbatim so the
            % result is provably identical for every input double.
            value = mat2doc.shared.Length.pyRoundHalfToEven_( ...
                double(obj) / mat2doc.shared.Length.EMUS_PER_TWIP_);
        end
    end

    methods (Static, Access = private)
        function n = pyRoundHalfToEven_(x)
            % Replicates CPython float.__round__(x) with no ndigits
            % (floatobject.c): r = C round(x) [ties away from zero]; if the
            % distance is exactly 0.5, snap to the nearest even integer via
            % 2*round(x/2). MATLAB round() == C round() (half away from zero).
            r = round(x);
            if abs(x - r) == 0.5
                r = 2 * round(x / 2);
            end
            n = r;
        end
    end
end

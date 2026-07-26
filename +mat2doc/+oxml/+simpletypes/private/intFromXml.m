function n = intFromXml(str_value)
% INTFROMXML Python int(str_value) base-10 parse, returning a plain double.
%
%   n = INTFROMXML(str_value) parses an XML integer literal exactly as
%   CPython int(str) does and returns the value as a plain (non-Length)
%   double. This is the bare-integer analogue of BaseIntType.convert_from_xml
%   (`int(str_value)`, simpletypes.py 68-69) for the simple types that return
%   a plain int (XsdInt / XsdLong / XsdUnsignedInt / XsdUnsignedLong /
%   ST_DecimalNumber / ST_DrawingElementId / ST_CoordinateUnqualified) and
%   for the measure ST_ types that apply int(str_value) BEFORE a Twips()/Pt()
%   constructor (ST_TwipsMeasure Twips(int(str)), simpletypes.py 402).
%
%   The parse is delegated to the audited Length string-parse path
%   (shared.private.pyIntArg "parse", reached through the Emu constructor):
%   double(Emu(str_value)) equals int(str_value) for every valid XML integer
%   literal. Emu applies int() -> fix(), a no-op on an already-integral
%   parsed value, and normalizes IEEE -0.0; an invalid literal such as '2.5'
%   raises mat2doc:ValueError with Python's "invalid literal for int() with
%   base 10: ..." message, matching int('2.5'). Reusing Emu keeps the
%   int()-domain grammar (ASCII [0-9], D-002) identical to the +shared P1-1
%   implementation (H6) rather than duplicating the grammar.
%
%   Inputs:  str_value - char row or string scalar, an XML integer literal
%   Outputs: n - double scalar (exact integer value of int(str_value))
%
%   Example:
%       intFromXml("914400")   % 914400
%       intFromXml("-42")      % -42
%
%   Mat2Doc infrastructure (package-private helper), no python-docx
%   counterpart; the BaseIntType `int(str_value)` conversion for bare-int and
%   pre-int()ed measure simple types. Mandated by H6/D-002.

n = double(mat2doc.shared.Emu(str_value));
end

classdef BaseIntEnum
% BASEINTENUM Shared root giving BaseEnum / BaseXmlEnum value-based == / ~=.
%
%   In python-docx v1.2.0 both ``BaseEnum(int, enum.Enum)`` and
%   ``BaseXmlEnum(int, enum.Enum)`` are INT SUBCLASSES (base.py lines 15, 33:
%   ``int.__new__(cls, ms_api_value)``). Their members therefore ARE integers,
%   so Python compares them by their MS-API integer value:
%
%       WD_PARAGRAPH_ALIGNMENT.CENTER == WD_TABLE_ALIGNMENT.CENTER   # int==int -> True
%       member == 1                                                 # int==int -> True
%       member == "CENTER"                                          # int==str -> False
%
%   MATLAB ``enumeration`` classes compare ``==`` by MEMBER IDENTITY by default,
%   which makes a cross-class comparison silently False and a member-vs-int
%   comparison an error/False. This root class restores the python-docx
%   int-subclass semantics by overriding ``eq``/``ne`` so that EVERY
%   BaseEnum/BaseXmlEnum member compares by ``double(value)``. Because a single
%   inherited ``eq`` lives here, two members of ANY two int-enum classes (same
%   base or the two different bases) dispatch to the same method with no
%   ambiguity (verified: cross-class and cross-base ``==`` resolve here).
%
%   Operand matrix (element-wise, native ``==`` broadcasting):
%     - member vs member (any BaseEnum/BaseXmlEnum) -> double(a.value)==double(b.value)
%     - member vs numeric (double/int/logical)      -> double(a.value)==double(b)
%     - member vs string/char                       -> false (Python int != str)
%     - member vs [] (None) / missing / other type  -> false
%   Non-numeric and None operands map to NaN in the comparison vector; NaN is
%   unequal to everything (NaN==NaN is false), which is exactly the Python
%   "int is never equal to a str/None" result. An operand's numeric vector keeps
%   its SHAPE, so scalar<->array broadcasting is delegated to MATLAB ``==``; a
%   genuinely empty enum/numeric array yields an empty result exactly as native
%   ``==`` does, while the ``[]`` None sentinel is treated as a scalar (Python
%   ``member == None`` -> scalar False).
%
%   NOT applied to the plain-``enum.Enum`` ports WD_BREAK_TYPE and
%   WD_INLINE_SHAPE_TYPE: those are NOT int subclasses in python-docx (they
%   compare by identity and are NOT equal to their int), so they do NOT derive
%   from this class - they derive from the sibling root ``BasePlainEnum`` which
%   gives them identity ``==`` (same member true; string/int/other-class/None
%   false), matching Python plain-``enum.Enum`` semantics.
%
%   Byte-neutral: this class changes ONLY comparison operators; ``value``,
%   ``xml_value``, ``from_xml``, ``to_xml`` and every serialization path are
%   untouched, so saved ``.docx`` output is unchanged.
%
%   Ported from python-docx v1.2.0: src/docx/enum/base.py::BaseEnum /
%   BaseXmlEnum (the ``int`` subclass equality of ``int.__new__(cls, ...)``)

    methods
        function tf = eq(a, b)
            % EQ Value-based ``==`` for int-enum members (Python int equality).
            tf = mat2doc.enum.base.BaseIntEnum.valVec_(a) == ...
                 mat2doc.enum.base.BaseIntEnum.valVec_(b);
        end

        function tf = ne(a, b)
            % NE Value-based ``~=`` (the negation of eq), element-wise.
            tf = mat2doc.enum.base.BaseIntEnum.valVec_(a) ~= ...
                 mat2doc.enum.base.BaseIntEnum.valVec_(b);
        end
    end

    methods (Static, Access = private)
        function v = valVec_(x)
            % Map an operand to a numeric comparison vector matching MATLAB ``==``
            % broadcasting. An int-enum member yields its double(value); a numeric
            % or logical operand yields its double; a string/char/None/other
            % operand yields NaN (unequal to everything -> Python int != str/None).
            if isa(x, "mat2doc.enum.base.BaseIntEnum")
                % .value is a scalar int32 on each concrete member; keep shape.
                v = reshape(double([x.value]), size(x));
            elseif ischar(x)
                v = NaN;                       % a char row is ONE string scalar
            elseif isstring(x)
                v = NaN(size(x));              % each string element: no int value
            elseif islogical(x)
                v = double(x);                 % Python bool is an int subclass
            elseif isnumeric(x)
                if isempty(x)
                    v = NaN;                   % [] None sentinel -> scalar False
                else
                    v = double(x);
                end
            else
                v = NaN;                       % any other type -> never equal
            end
        end
    end
end

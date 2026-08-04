classdef BasePlainEnum
% BASEPLAINENUM Shared root giving the plain-``enum.Enum`` ports IDENTITY == / ~=.
%
%   In python-docx v1.2.0 two enumerations are declared as PLAIN
%   ``class X(enum.Enum)`` (NOT int subclasses): ``WD_BREAK_TYPE``
%   (enum/text.py:70) and ``WD_INLINE_SHAPE_TYPE`` (enum/shape.py:6). A plain
%   Enum member compares by IDENTITY, never by value and never against a
%   non-member:
%
%       WD_BREAK_TYPE.LINE == WD_BREAK_TYPE.LINE    # same member   -> True
%       WD_BREAK_TYPE.LINE == WD_BREAK_TYPE.PAGE    # other member  -> False
%       WD_BREAK_TYPE.LINE == "LINE"                # str           -> False
%       WD_BREAK_TYPE.LINE == 6                     # int (value)   -> False
%       WD_BREAK_TYPE.LINE == WD_INLINE_SHAPE_TYPE.X  # other class -> False
%       WD_BREAK_TYPE.LINE == None                  #               -> False
%
%   These enums were (correctly) EXCLUDED from ``BaseIntEnum`` (value-eq) because
%   they are not int subclasses. But MATLAB's built-in ``enumeration`` ``==``
%   compares a member to a string by NAME, so ``WD_BREAK_TYPE.LINE == "LINE"``
%   returned TRUE -- diverging from Python's False. This root class restores plain
%   identity semantics by overriding ``eq``/``ne`` so a member is equal ONLY to a
%   member of the SAME plain-enum class with the SAME name (Python member
%   identity), and is NEVER equal to a string, number, logical, other-class
%   member, ``[]`` (None), or ``missing``.
%
%   Operand matrix (element-wise, native ``==`` broadcasting):
%     - member vs member, SAME plain-enum class -> string(name)==string(name)
%     - member vs member, DIFFERENT class       -> false (even if values collide)
%     - member vs string / char                 -> false (THE FIX; was true)
%     - member vs numeric / logical             -> false (plain enum != its int)
%     - member vs [] (None) / missing / other   -> false
%   Within one class, member names are unique (an alias such as
%   WD_BREAK_TYPE.TEXT_WRAPPING resolves to the canonical LINE_CLEAR_ALL member,
%   name "LINE_CLEAR_ALL"), so name equality IS member identity. The cross-class
%   and non-member cases map both operands to NaN comparison vectors; NaN is
%   unequal to everything (NaN==NaN is false), which is exactly the Python "a
%   plain Enum member is never equal to a non-member" result. Each operand's
%   vector keeps its SHAPE so scalar<->array broadcasting is delegated to MATLAB
%   ``==``; a char row is treated as ONE string scalar and the ``[]`` None
%   sentinel as a scalar (Python ``member == None`` -> scalar False).
%
%   The value-based int-enum root is the SIBLING ``BaseIntEnum`` (BaseEnum /
%   BaseXmlEnum). No int-enum derives from this class, and neither plain enum
%   derives from ``BaseIntEnum`` -- the two families are disjoint by construction.
%
%   The name idiom ``string(member) == "NAME"`` is UNAFFECTED (string<->string,
%   not routed through this eq).
%
%   Byte-neutral: this class changes ONLY comparison operators; the ``value``
%   property, member declarations and constructors are untouched, and plain enums
%   carry no ``xml_value``/``from_xml``/``to_xml`` serialization path, so saved
%   ``.docx`` output is unchanged.
%
%   Ported from python-docx v1.2.0: src/docx/enum/text.py::WD_BREAK_TYPE /
%   src/docx/enum/shape.py::WD_INLINE_SHAPE_TYPE (plain ``enum.Enum`` identity ==)

    methods
        function tf = eq(a, b)
            % EQ Identity ``==`` for plain-enum members (Python plain-Enum identity).
            import mat2doc.enum.base.BasePlainEnum
            sameClass = isa(a, 'mat2doc.enum.base.BasePlainEnum') && ...
                        isa(b, 'mat2doc.enum.base.BasePlainEnum') && ...
                        strcmp(class(a), class(b));
            if sameClass
                % Same plain-enum class: identity == name equality, element-wise.
                tf = BasePlainEnum.names_(a) == BasePlainEnum.names_(b);
            else
                % Different class or non-member operand -> never equal. Map both to
                % NaN vectors (NaN==NaN is false) so native == owns the broadcast.
                tf = BasePlainEnum.neverVec_(a) == BasePlainEnum.neverVec_(b);
            end
        end

        function tf = ne(a, b)
            % NE Identity ``~=`` (the exact negation of eq), element-wise.
            tf = ~eq(a, b);
        end
    end

    methods (Static, Access = private)
        function s = names_(x)
            % Member names of a plain-enum array, shape-preserved. `string` on an
            % enumeration member returns its name and does NOT route through eq.
            s = reshape(string(x), size(x));
        end

        function v = neverVec_(x)
            % Map any operand to an all-NaN comparison vector (so it equals
            % nothing), preserving shape for MATLAB == broadcasting. Mirrors the
            % sibling BaseIntEnum.valVec_ branch structure, but a plain-enum member
            % (even the SAME value in another class) maps to NaN -> never equal.
            if isa(x, 'mat2doc.enum.base.BasePlainEnum')
                v = NaN(size(x));              % other-class member: never matches
            elseif ischar(x)
                v = NaN;                       % a char row is ONE string scalar
            elseif isstring(x)
                v = NaN(size(x));              % each string element: no member
            elseif isnumeric(x) || islogical(x)
                if isempty(x)
                    v = NaN;                   % [] None sentinel -> scalar False
                else
                    v = NaN(size(x));
                end
            else
                v = NaN;                       % missing / other type -> never equal
            end
        end
    end
end

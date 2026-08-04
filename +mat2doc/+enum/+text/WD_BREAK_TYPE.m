classdef WD_BREAK_TYPE < mat2doc.enum.base.BasePlainEnum
% WD_BREAK_TYPE Corresponds to the MS WdBreakType enumeration.
%
%   http://msdn.microsoft.com/en-us/library/office/ff195905.aspx
%
%   Alias: WD_BREAK (mat2doc.enum.text.WD_BREAK).
%
%   Design realization. This is NOT a BaseEnum/BaseXmlEnum: in python-docx it is
%   a PLAIN ``enum.Enum`` whose members carry only a bare integer value (no
%   xml_value, no per-member docstring, no from_xml/to_xml). The port is a plain
%   value classdef with an ``enumeration`` block and a single immutable ``value``
%   (int32) property, mirroring the PROG_ID plain-enum precedent in Mat2Ppt. It
%   derives from ``mat2doc.enum.base.BasePlainEnum`` SOLELY for Python plain-Enum
%   IDENTITY ``==``/``~=`` (a member equals only the same member of the same
%   class; never a string/int/other-class member/None) -- see that class. This is
%   the identity sibling of the int-enum value-eq root ``BaseIntEnum``; the plain
%   enums do NOT derive from ``BaseIntEnum`` (they are not int subclasses).
%
%   MEMBER ALIAS (closes the P3-1 WD_BREAK_TYPE carry-forward). python-docx
%   declares ``LINE_CLEAR_ALL = 11`` and ``TEXT_WRAPPING = 11`` (text.py lines
%   80, 86). Python enum makes the second same-valued member an ALIAS of the
%   first: ``WD_BREAK_TYPE.TEXT_WRAPPING is WD_BREAK_TYPE.LINE_CLEAR_ALL`` is
%   True, its ``.name`` is "LINE_CLEAR_ALL", and iteration/``__members__`` list
%   only the canonical LINE_CLEAR_ALL (verified against the oracle).
%
%   MATLAB enumeration members compare ``==`` by MEMBER IDENTITY, NOT by the
%   ``value`` property (empirically verified: two distinct members that both
%   carry value 11 are NOT ``==``). Declaring TEXT_WRAPPING as a SECOND
%   enumeration member would therefore make it a distinct object that is neither
%   ``==`` to nor named the same as LINE_CLEAR_ALL - diverging from Python. So
%   the alias is realized the same way Python realizes it: as a class attribute
%   (a Constant property) pointing at the canonical LINE_CLEAR_ALL member. Then
%   ``WD_BREAK_TYPE.TEXT_WRAPPING`` returns the LINE_CLEAR_ALL member itself:
%   ``string(...TEXT_WRAPPING)`` is "LINE_CLEAR_ALL", ``TEXT_WRAPPING ==
%   LINE_CLEAR_ALL`` is true, its value is 11, and ``enumeration('...')`` lists
%   only the 10 canonical members - matching Python exactly.
%
%   Consumer: Run.add_break (docx/text/run.py), ported at P4-4b; it maps a member
%   to (w:br type, clear) attribute values via an identity lookup, so member
%   identity (not the int) is what matters at the call site.
%
%   Example:
%       tw = mat2doc.enum.text.WD_BREAK_TYPE.TEXT_WRAPPING;
%       string(tw)                                             % "LINE_CLEAR_ALL"
%       tw == mat2doc.enum.text.WD_BREAK_TYPE.LINE_CLEAR_ALL   % true
%       double(tw.value)                                       % 11
%
%   Ported from python-docx v1.2.0: src/docx/enum/text.py::WD_BREAK_TYPE

    properties (SetAccess = immutable)
        % The MS API integer value assigned to this member (Python `_value_`).
        value (1, 1) int32
    end

    properties (Constant)
        % Alias member: Python `TEXT_WRAPPING = 11`, an alias of LINE_CLEAR_ALL
        % (text.py line 86; "-- added for consistency, not in MS version --").
        % Resolves to the canonical LINE_CLEAR_ALL member (same object).
        TEXT_WRAPPING = mat2doc.enum.text.WD_BREAK_TYPE.LINE_CLEAR_ALL
    end

    methods
        function obj = WD_BREAK_TYPE(v)
            % Ported from the bare `NAME = <int>` member declarations. There is
            % no shared enum base (Python subclasses only enum.Enum), so the
            % constructor stores the integer value directly.
            if nargin == 0
                % No-arg path so MATLAB can construct the class prototype.
                return
            end
            obj.value = v;
        end
    end

    enumeration
        COLUMN             (8)
        LINE               (6)
        LINE_CLEAR_LEFT    (9)
        LINE_CLEAR_RIGHT   (10)
        LINE_CLEAR_ALL     (11)
        PAGE               (7)
        SECTION_CONTINUOUS (3)
        SECTION_EVEN_PAGE  (4)
        SECTION_NEXT_PAGE  (2)
        SECTION_ODD_PAGE   (5)
    end
end

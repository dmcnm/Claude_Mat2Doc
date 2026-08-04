classdef WD_INLINE_SHAPE_TYPE < mat2doc.enum.base.BasePlainEnum
% WD_INLINE_SHAPE_TYPE Corresponds to the MS WdInlineShapeType enumeration.
%
%   http://msdn.microsoft.com/en-us/library/office/ff192587.aspx.
%
%   Alias: WD_INLINE_SHAPE (mat2doc.enum.shape.WD_INLINE_SHAPE).
%
%   Design realization. This is NOT a BaseEnum/BaseXmlEnum: in python-docx it is
%   a PLAIN ``enum.Enum`` whose members carry only a bare integer value (no
%   xml_value, no per-member docstring, no from_xml/to_xml). The port is a plain
%   value classdef with an ``enumeration`` block and a single immutable ``value``
%   (int32) property, mirroring the WD_BREAK_TYPE plain-enum precedent (P3-3). It
%   derives from ``mat2doc.enum.base.BasePlainEnum`` SOLELY for Python plain-Enum
%   IDENTITY ``==``/``~=`` (a member equals only the same member of the same
%   class; never a string/int/other-class member/None) -- see that class. The
%   plain enums do NOT derive from the int-enum value-eq root ``BaseIntEnum``.
%
%   NO internal member alias. Verified against the python-docx v1.2.0 oracle: the
%   5 members carry 5 DISTINCT integer values (CHART 12, LINKED_PICTURE 4,
%   PICTURE 3, SMART_ART 15, NOT_IMPLEMENTED -6), so unlike WD_BREAK_TYPE (which
%   needed a TEXT_WRAPPING Constant aliasing LINE_CLEAR_ALL) there is no
%   duplicate-value alias to realize. NOT_IMPLEMENTED is negative (-6), stored
%   as int32.
%
%   Example:
%       t = mat2doc.enum.shape.WD_INLINE_SHAPE_TYPE.PICTURE;
%       double(t.value)   % 3
%       string(t)         % "PICTURE"  (member name)
%
%   Ported from python-docx v1.2.0: src/docx/enum/shape.py::WD_INLINE_SHAPE_TYPE

    properties (SetAccess = immutable)
        % The MS API integer value assigned to this member (Python `_value_`).
        value (1, 1) int32
    end

    methods
        function obj = WD_INLINE_SHAPE_TYPE(v)
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
        CHART           (12)
        LINKED_PICTURE  (4)
        PICTURE         (3)
        SMART_ART       (15)
        NOT_IMPLEMENTED (-6)
    end
end

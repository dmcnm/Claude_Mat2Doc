classdef BaseEnum
% BASEENUM Base class for enumerations that do NOT map XML attribute values.
%
%   A concrete enumeration derives from this class and declares an
%   ``enumeration`` block whose members pass ``(ms_api_value, docstr)`` to the
%   constructor. Each member's integer ``value`` corresponds to the integer
%   assigned the same-named member in the Microsoft API enum of the same name.
%
%   Design realization (design.md section 2; audit_P3-1_enum_base.md):
%   Python's ``BaseEnum(int, enum.Enum)`` is an int subclass whose members
%   carry an integer ``value`` (``_value_ = ms_api_value``) and a ``__doc__``
%   string. MATLAB enumeration classes that subclass a built-in numeric type
%   cannot add properties, so the port makes BaseEnum a plain VALUE class (not
%   handle - enum members are value objects) that owns the ``value``/``doc``
%   properties; the concrete enum subclasses it with an ``enumeration`` block.
%   Python ``int(member)`` sites port as explicit ``double(member.value)``
%   (members are NOT numeric in MATLAB, so bare ``double(member)`` does not
%   work; a member-vs-int comparison such as Python ``member == ms_api_value``
%   likewise ports as ``double(member.value) == ms_api_value`` at each site).
%
%   Inputs (constructor):
%       ms_api_value - scalar, the MS API integer constant (stored as int32)
%       docstr       - string, the member docstring (Python __doc__, .strip()ed)
%   Outputs: obj - a BaseEnum value (concrete-subclass instance)
%
%   Example:
%       % Exercised through a concrete BaseEnum subclass (never instantiated
%       % directly). If WD_SECTION_START were such a subclass:
%       a = mat2doc.enum.WD_SECTION_START.NEW_PAGE;
%       double(a.value)   % 2  (Python int(WD_SECTION_START.NEW_PAGE))
%       string(a)         % "NEW_PAGE"  (member name, Python self.name)
%       a.str_()          % "NEW_PAGE (2)"  (Python str(member))
%
%   Ported from python-docx v1.2.0: src/docx/enum/base.py::BaseEnum

    properties (SetAccess = immutable)
        % The MS API integer value assigned to this member (Python `_value_`).
        value (1, 1) int32
        % The member docstring (Python `__doc__`, stripped). Non-behavioral:
        % only DocsPageFormatter reads it, and that is not ported (see audit).
        doc (1, 1) string
    end

    methods
        function obj = BaseEnum(ms_api_value, docstr)
            % Ported from BaseEnum.__new__ (base.py lines 22-26).
            if nargin == 0
                % No-arg path so MATLAB can construct the class prototype.
                return
            end
            obj.value = ms_api_value;
            obj.doc = strip(string(docstr));   % Python docstr.strip()
        end

        function s = str_(obj)
            % STR_ Symbolic name and value, e.g. "NEW_PAGE (2)" (Python __str__).
            %
            %   Ported from BaseEnum.__str__ (base.py lines 28-30):
            %   f"{self.name} ({self.value})". `string(obj)` yields the MATLAB
            %   enumeration member name (Python `self.name`); `char`/`string`
            %   are deliberately NOT overridden so the member name stays
            %   directly accessible.
            s = string(obj) + " (" + string(double(obj.value)) + ")";
        end
    end
end

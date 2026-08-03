classdef BaseXmlEnum < mat2doc.enum.base.BaseIntEnum
% BASEXMLENUM Base class for enumerations that also map XML attribute values.
%
%   A concrete enumeration derives from this class and declares an
%   ``enumeration`` block whose members pass ``(ms_api_value, xml_value,
%   docstr)`` to the constructor. Each member carries the MS API integer
%   ``value`` (Python ``_value_``) PLUS the XML attribute value ``xml_value``
%   used to (de)serialize it. In python-docx v1.2.0 ``xml_value`` is typed
%   ``str | None``: a member may legally carry ``None`` (stored here as a
%   ``<missing>`` string) or the empty string ``""`` or an ordinary token. The
%   xmlchemy attribute descriptors call the concrete class's static
%   ``from_xml`` / ``to_xml`` to translate between the XML string form and the
%   enumeration member.
%
%   Design realization (design.md section 2; audit_P3-1_enum_base.md):
%   Python's ``BaseXmlEnum(int, enum.Enum)`` adds an ``xml_value`` attribute to
%   each member and defines classmethods ``from_xml``/``to_xml`` (NO
%   ``validate`` in docx v1.2.0). MATLAB static methods have no ``cls``
%   binding, so the shared logic lives here as Hidden static helpers taking the
%   concrete class name; each concrete subclass exposes thin static
%   ``from_xml``/``to_xml`` that forward with their own class name hardcoded
%   (the explicit, no-metaclass pattern).
%
%   DOCX-vs-PPTX SEMANTIC DELTAS (do NOT copy the Mat2Ppt BaseXmlEnum verbatim -
%   the boundary audit found it differs). Ported to docx v1.2.0 semantics:
%     1. ``from_xml`` has NO None/empty-string short-circuit (pptx added one;
%        docx does NOT). It is a straight equality scan: the member whose
%        ``xml_value`` equals the query, with None-tolerant comparison
%        (``None == None`` -> match). Therefore ``from_xml(None)`` returns the
%        member whose ``xml_value is None`` IF one exists, else raises - and a
%        member with ``xml_value == ""`` IS reachable via ``from_xml("")``.
%     2. ``to_xml`` guard is ``if not xml_value`` (Python falsy) - true for
%        BOTH None and "" - so a resolved member with a <missing> OR empty-
%        string ``xml_value`` raises "has no XML representation".
%     3. NO ``validate`` classmethod (pptx has one; docx v1.2.0 does not).
%     4. ValueError message forms are docx-specific (below), emitted under
%        mat2doc:ValueError (D-005 identifier convention).
%
%   H3 tri-state (None vs "" vs missing). Python None ports to a <missing>
%   string; Python "" ports to the real empty string "". The two are distinct:
%   ``from_xml(None)`` (query normalized to <missing>) matches only a <missing>
%   member; ``from_xml("")`` matches only a "" member. Both a <missing> and a
%   "" member are rejected by ``to_xml`` ("has no XML representation").
%
%   Example:
%       % Exercised through a concrete BaseXmlEnum subclass (the machinery is
%       % never instantiated directly). If WD_TAB_ALIGNMENT were such a class:
%       m = mat2doc.enum.WD_TAB_ALIGNMENT.from_xml("center");
%       string(m)                                     % "CENTER"
%       double(m.value)                               % 1
%       mat2doc.enum.WD_TAB_ALIGNMENT.to_xml(m)       % "center"
%
%   Ported from python-docx v1.2.0: src/docx/enum/base.py::BaseXmlEnum

    properties (SetAccess = immutable)
        % The MS API integer value assigned to this member (Python `_value_`).
        value (1, 1) int32
        % The XML attribute value for this member (Python `xml_value: str |
        % None`). A real string for a mapped member, "" for an empty-string
        % member, or a <missing> string for a Python-None member. H3: <missing>
        % is the None sentinel; "" is a distinct real value.
        xml_value (1, 1) string
        % The member docstring (Python `__doc__`, stripped). Non-behavioral.
        doc (1, 1) string
    end

    methods
        function obj = BaseXmlEnum(ms_api_value, xml_value, docstr)
            % Ported from BaseXmlEnum.__new__ (base.py lines 42-47).
            if nargin == 0
                return
            end
            obj.value = ms_api_value;
            obj.xml_value = mat2doc.enum.base.BaseXmlEnum.asXmlVal_(xml_value);
            obj.doc = strip(string(docstr));   % Python docstr.strip()
        end

        function s = str_(obj)
            % STR_ Symbolic name and value, e.g. "CENTER (1)" (Python __str__).
            %   Ported from BaseXmlEnum.__str__ (base.py lines 49-51).
            s = string(obj) + " (" + string(double(obj.value)) + ")";
        end
    end

    % --- Shared machinery. Hidden keeps it off the public doc surface; the
    % --- concrete subclasses forward to these with their own class name. -----
    methods (Static, Hidden)
        function member = from_xml_(fullClassName, xml_value)
            % FROM_XML_ Member whose xml_value == `xml_value` (None-tolerant).
            %   Ported from BaseXmlEnum.from_xml (base.py lines 53-66):
            %       member = next((m for m in cls if m.xml_value == xml_value), None)
            %       if member is None:
            %           raise ValueError(f"{cls.__name__} has no XML mapping for '{xml_value}'")
            %       return member
            %   NO None/empty short-circuit (the docx delta): the query is
            %   compared for equality against every member, with None (<missing>)
            %   matching only a None member. mat2doc:ValueError on no match.
            query = mat2doc.enum.base.BaseXmlEnum.asXmlVal_(xml_value);
            members = enumeration(fullClassName);
            found = false;
            member = [];
            for k = 1:numel(members)
                if mat2doc.enum.base.BaseXmlEnum.xmlEq_(members(k).xml_value, query)
                    member = members(k);
                    found = true;
                    break
                end
            end
            if ~found
                error("mat2doc:ValueError", "%s has no XML mapping for '%s'", ...
                    mat2doc.enum.base.BaseXmlEnum.shortName_(fullClassName), ...
                    mat2doc.enum.base.BaseXmlEnum.queryStr_(query));
            end
        end

        function s = to_xml_(fullClassName, value)
            % TO_XML_ XML attribute value for `value` (a member, its int, or None).
            %   Ported from BaseXmlEnum.to_xml (base.py lines 68-77):
            %       member = cls(value)
            %       xml_value = member.xml_value
            %       if not xml_value:
            %           raise ValueError(f"{cls.__name__}.{member.name} has no XML representation")
            %       return xml_value
            %   `cls(value)` resolves a member by identity (already a member) or
            %   by its int value; a bad value raises mat2doc:ValueError
            %   ("<v> is not a valid <Cls>", the Python stdlib enum form). The
            %   `if not xml_value` guard is Python-falsy: BOTH a <missing> and an
            %   empty-string xml_value raise "has no XML representation".
            member = mat2doc.enum.base.BaseXmlEnum.resolveMember_(fullClassName, value);
            xv = member.xml_value;
            if ismissing(xv) || strlength(xv) == 0
                error("mat2doc:ValueError", "%s.%s has no XML representation", ...
                    mat2doc.enum.base.BaseXmlEnum.shortName_(fullClassName), ...
                    string(member));
            end
            s = xv;
        end
    end

    % --- Private helpers ----------------------------------------------------
    methods (Static, Access = private)
        function member = resolveMember_(fullClassName, value)
            % Python cls(value): identity for a member, value-lookup for an int,
            % ValueError otherwise (including None -> "None is not a valid Cls").
            if isa(value, fullClassName)
                member = value;
                return
            end
            v = mat2doc.enum.base.BaseXmlEnum.intOf_(value);   % [] if not numeric-like
            if ~isempty(v)
                members = enumeration(fullClassName);
                idx = find(double([members.value]) == v, 1);
                if ~isempty(idx)
                    member = members(idx);
                    return
                end
            end
            error("mat2doc:ValueError", "%s is not a valid %s", ...
                mat2doc.enum.base.BaseXmlEnum.valueStr_(value), ...
                mat2doc.enum.base.BaseXmlEnum.shortName_(fullClassName));
        end

        function v = intOf_(value)
            % The integer value of `value` for the cls(value) lookup, or [] if
            % none (None sentinel [] / <missing> / string / non-scalar -> []).
            if isnumeric(value) && isscalar(value) && ~isempty(value)
                v = double(value);
            elseif (isa(value, "mat2doc.enum.base.BaseXmlEnum") || ...
                    isa(value, "mat2doc.enum.base.BaseEnum")) && isscalar(value)
                v = double(value.value);
            else
                v = [];
            end
        end

        function s = valueStr_(value)
            % Python repr-ish rendering of `value` for the "is not a valid"
            % error message (CPython stdlib enum formats it with %r, so a str
            % value renders single-quoted; int/None render as their str form).
            if (isa(value, "mat2doc.enum.base.BaseXmlEnum") || ...
                    isa(value, "mat2doc.enum.base.BaseEnum")) && isscalar(value)
                s = value.str_();                        % "NAME (value)"
            elseif isnumeric(value) && isscalar(value) && ~isempty(value)
                s = mat2doc.shared.pyStr(double(value));  % Python str(int)
            elseif isnumeric(value) && isempty(value)
                s = "None";                               % [] None sentinel
            elseif (isstring(value) || ischar(value)) && isscalar(string(value))
                sv = string(value);
                if ismissing(sv)
                    s = "None";                           % <missing> = None
                else
                    s = "'" + sv + "'";                   % Python repr(str)
                end
            else
                s = string(class(value));
            end
        end

        function nrm = asXmlVal_(v)
            % Normalize an xml_value (member store or query) to the canonical
            % form: a <missing> string for Python None, else a real string.
            % H3: the [] None sentinel and a <missing> string both mean None;
            % a char/string '' maps to the real empty string "" (NOT None).
            if isnumeric(v) && isempty(v)
                nrm = string(missing);                    % [] None sentinel
            elseif isstring(v) && isscalar(v) && ismissing(v)
                nrm = string(missing);                    % <missing> = None
            else
                nrm = string(v);                          % char/string/num -> string
            end
        end

        function tf = xmlEq_(a, b)
            % Python `a == b` for two xml_values (each <missing> for None or a
            % real string): None==None -> true; None vs string -> false; else
            % string equality ("" == "" -> true).
            if ismissing(a) && ismissing(b)
                tf = true;
            elseif ismissing(a) || ismissing(b)
                tf = false;
            else
                tf = (a == b);
            end
        end

        function s = queryStr_(query)
            % Python str(xml_value) inside the from_xml message's literal
            % single quotes: None -> "None", else the string ("" -> "").
            if ismissing(query)
                s = "None";
            else
                s = query;
            end
        end

        function s = shortName_(fullClassName)
            % "mat2doc.enum.WD_TAB_ALIGNMENT" -> "WD_TAB_ALIGNMENT"
            % (Python cls.__name__, unqualified).
            parts = split(string(fullClassName), ".");
            s = parts(end);
        end
    end
end

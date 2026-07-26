classdef BaseStringEnumerationType < mat2doc.oxml.simpletypes.BaseStringType
% BASESTRINGENUMERATIONTYPE Simple-type base for enumerated xsd:string values.
%
%   Python BaseStringEnumerationType.validate (simpletypes.py 97-101) is
%   `validate_string(value); if value not in cls._members: raise ValueError`,
%   where cls._members LATE-BINDS to the concrete subclass's tuple. MATLAB
%   static methods have no such late binding (H10) and cannot read a subclass
%   constant from a shared base method, so this base cannot host a working
%   validate. Instead it exposes validate_enum(value, members): each concrete
%   enumeration simple type (ST_HexColorAuto, ST_Merge, ST_VerticalAlignRun)
%   declares its own validate that calls validate_enum with its members, and
%   its own to_xml routed to that validate (BaseStringType.to_xml would call
%   BaseStringType.validate and skip the membership check).
%
%   validate_enum is ALSO reused by the plain XsdString subclasses that carry
%   an inline valid-values tuple with the identical ValueError message
%   ("must be one of %s, got '%s'"): ST_BrClear, ST_BrType, ST_TblLayoutType,
%   ST_TblWidth (simpletypes.py 181-196, 379-394). Those are not
%   BaseStringEnumerationType subclasses in Python but produce the same check,
%   so they call this static directly.
%
%   from_xml / convert_from_xml / convert_to_xml are the identity string
%   transforms inherited from BaseStringType unchanged (Python's enumeration
%   simple types add only validation).
%
%   Ported from python-docx v1.2.0:
%   src/docx/oxml/simpletypes.py::BaseStringEnumerationType (lines 94-101)

    methods (Static)
        function validate_enum(value, members)
            % VALIDATE_ENUM Ported BaseStringEnumerationType.validate body.
            %   lines 98-101: validate_string(value); if value not in members
            %   raise ValueError "must be one of %s, got '%s'" % (members, value)
            %   -- the %s of a Python tuple renders as ('a', 'b'), reproduced
            %   here (a singleton keeps the trailing comma: ('a',)). `members`
            %   is a string array in the source declaration order.
            v = mat2doc.oxml.simpletypes.BaseSimpleType.validate_string(value);
            if ~any(v == members)
                error("mat2doc:ValueError", "must be one of %s, got '%s'", ...
                    mat2doc.oxml.simpletypes.BaseStringEnumerationType.tupleRepr_(members), v);
            end
        end
    end

    methods (Static, Access = private)
        function s = tupleRepr_(members)
            % TUPLEREPR_ Python repr of a tuple of str, e.g. ('continue', 'restart').
            parts = "'" + members + "'";
            if numel(members) == 1
                s = "(" + parts(1) + ",)";
            else
                s = "(" + strjoin(parts, ", ") + ")";
            end
        end
    end
end

classdef BaseSimpleType
% BASESIMPLETYPE Root of the simple-type (scalar XML attribute value) hierarchy.
%
%   A "simple-type" transforms values to and from the XML string form of an
%   attribute (validation + format translation). The xmlchemy attribute
%   descriptors hold a simple-type CLASS and call its from_xml / to_xml; the
%   Mat2Doc engine (BaseOxmlElement.getAttrTyped / setAttrTyped) dispatches to
%   the named simple-type's static from_xml / to_xml instead.
%
%   CLASSMETHOD LATE-BINDING HAZARD (H10): Python's from_xml / to_xml are
%   @classmethod and call `cls.convert_from_xml` / `cls.validate`, which bind
%   to the MOST-DERIVED class at call time. MATLAB static methods have NO such
%   late binding, so from_xml / to_xml are defined on the concrete branch base
%   (BaseIntType / BaseStringType, or the direct-BaseSimpleType leaf itself:
%   XsdBoolean / ST_DateTime / ST_UniversalMeasure / ST_Coordinate /
%   ST_HexColor) that owns the convert/validate they use, rather than here.
%   Any subclass that overrides a convert_from_xml re-declares from_xml, and
%   one that overrides validate or convert_to_xml re-declares to_xml, so the
%   routed convert/validate is that subclass's own (see each ST_* header).
%
%   Python BaseSimpleType additionally defines an ABSTRACT-ish
%   convert_from_xml default (`return int(str_value)`, simpletypes.py 36-37),
%   convert_to_xml / validate stubs, and the from_xml / to_xml templates
%   (simpletypes.py 25-33). Those templates and the int default are ported on
%   the branch bases per H10; the int default is never reached (every concrete
%   leaf either overrides convert_from_xml or inherits a branch base that
%   does), so it is intentionally omitted here.
%
%   docx BaseSimpleType has NO float validators (validate_float /
%   validate_float_in_range) -- python-pptx did, python-docx does not
%   (D-STYPE-2 re-home: docx's only float parse lives inside
%   ST_UniversalMeasure.convert_from_xml). Only the integer + string
%   validators are ported here.
%
%   INTEGRAL-TYPE HAZARD (D-STYPE-1): Python validate_int rejects a float
%   whose value is integral (isinstance(2.0, int) is False), but MATLAB
%   cannot distinguish the double 2.0 from an integer 2 (the same ambiguity
%   documented for pyStr "auto" and RGBColor). validate_int here accepts any
%   finite integral-VALUED real numeric (and Length, which subclasses int in
%   Python, and logical, an int subclass). It therefore accepts 2.0 where
%   CPython raises TypeError -- a divergence only for a caller programming
%   error, recorded as adopted deviation D-STYPE-1. Non-integral doubles
%   (2.5), [] (the None analogue), non-scalars, and non-numerics are still
%   rejected, matching Python isinstance(value, int) -> False -> TypeError.
%
%   TYPE-TOKEN DIVERGENCE (D-005): the validate_int / validate_string error
%   messages report the MATLAB class token (e.g. 'double') where CPython
%   reports type(value) (e.g. "<class 'int'>"). Exception CLASS and message
%   template are otherwise faithful; adopted ruling D-005.
%
%   Example:
%       mat2doc.oxml.simpletypes.BaseSimpleType.validate_string("rId7")  % "rId7"
%       try
%           mat2doc.oxml.simpletypes.BaseSimpleType.validate_int(3.5);
%       catch e
%           disp(e.message)   % "value must be <type 'int'>, got double"
%       end
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::BaseSimpleType
%   (lines 22-63; from_xml/to_xml templates ported on the branch bases)

    methods (Static)
        function validate_int(value)
            % VALIDATE_INT Raise TypeError unless value is an int (lines 45-48).
            %   Python: if not isinstance(value, int): raise TypeError(
            %   "value must be <type 'int'>, got %s" % type(value)). See the
            %   INTEGRAL-TYPE HAZARD (D-STYPE-1) and TYPE-TOKEN (D-005) notes:
            %   MATLAB accepts a finite integral-VALUED real numeric, a Length,
            %   or a logical; [] / non-scalar / non-integral / non-numeric
            %   raise. Message template + class faithful; class token per D-005.
            if isscalar(value) && mat2doc.oxml.simpletypes.BaseSimpleType.isIntegral_(value)
                return
            end
            error("mat2doc:TypeError", ...
                "value must be <type 'int'>, got %s", class(value));
        end

        function validate_int_in_range(value, min_inclusive, max_inclusive)
            % VALIDATE_INT_IN_RANGE Int type check then inclusive range (lines 50-57).
            %   Python: validate_int(value); if value < min or value > max:
            %   raise ValueError("value must be in range %d to %d inclusive,
            %   got %d" % (min, max, value)). validate_int runs first, so value
            %   is integral here; the %d fields render via pyStr "int".
            mat2doc.oxml.simpletypes.BaseSimpleType.validate_int(value);
            if value < min_inclusive || value > max_inclusive
                error("mat2doc:ValueError", ...
                    "value must be in range %s to %s inclusive, got %s", ...
                    mat2doc.shared.pyStr(min_inclusive, "int"), ...
                    mat2doc.shared.pyStr(max_inclusive, "int"), ...
                    mat2doc.shared.pyStr(value, "int"));
            end
        end

        function v = validate_string(value)
            % VALIDATE_STRING Return value if it is a string, else TypeError (lines 59-63).
            %   Python: if not isinstance(value, str): raise TypeError(
            %   "value must be a string, got %s" % type(value)); return value.
            %   MATLAB has no `str`; a string scalar or a char row IS the str
            %   case. Class token per D-005; template + class faithful.
            if (isstring(value) && isscalar(value) && ~ismissing(value)) || ...
                    (ischar(value) && (isrow(value) || isempty(value)))
                v = string(value);
                return
            end
            error("mat2doc:TypeError", "value must be a string, got %s", class(value));
        end
    end

    methods (Static, Access = private)
        function tf = isIntegral_(value)
            % ISINTEGRAL_ Integral-value test (see D-STYPE-1 in the header).
            %   True for a Length (Python int subclass), a MATLAB integer-class
            %   scalar, a logical scalar (bool is an int subclass), or a finite
            %   integral-VALUED real double scalar.
            if isa(value, "mat2doc.shared.Length")
                tf = true;
                return
            end
            if ~isscalar(value) || ~isreal(value)
                tf = false;
                return
            end
            if islogical(value) || isinteger(value)
                tf = true;
                return
            end
            tf = isnumeric(value) && isfinite(value) && value == fix(value);
        end
    end
end

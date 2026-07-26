classdef XsdLong < mat2doc.oxml.simpletypes.BaseIntType
% XSDLONG The xsd:long simple type: signed 64-bit integer.
%
%   Overrides validate with the signed 64-bit range; re-declares to_xml so
%   the template routes to THIS validate (H10). from_xml / convert_* are
%   inherited from BaseIntType.
%
%   RANGE-PRECISION HAZARD (D-STYPE-3): the bounds -9223372036854775808 and
%   9223372036854775807 (+/- 2^63) exceed 2^53, so a MATLAB double cannot
%   hold 2^63-1 exactly (it rounds to 9223372036854775808). The range test
%   and the error message therefore use the rounded upper bound. This is a
%   DEAD path in docx: every concrete subclass in scope
%   (ST_CoordinateUnqualified, ST_PositiveCoordinate) overrides validate with
%   a narrower range well below 2^53, so XsdLong.validate is never reached on
%   any live docx call path. Recorded as adopted deviation D-STYPE-3.
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::XsdLong
%   (lines 148-151)

    methods (Static)
        function s = to_xml(value)
            % TO_XML BaseSimpleType.to_xml (lines 29-33), bound to THIS validate.
            mat2doc.oxml.simpletypes.XsdLong.validate(value);
            s = mat2doc.oxml.simpletypes.XsdLong.convert_to_xml(value);
        end

        function validate(value)
            % VALIDATE lines 149-151: int in signed 64-bit range (D-STYPE-3).
            mat2doc.oxml.simpletypes.BaseSimpleType.validate_int_in_range( ...
                value, -9223372036854775808, 9223372036854775807);
        end
    end
end

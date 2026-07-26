classdef XsdUnsignedLong < mat2doc.oxml.simpletypes.BaseIntType
% XSDUNSIGNEDLONG The xsd:unsignedLong simple type: 0..18446744073709551615.
%
%   Overrides validate; re-declares to_xml so the template routes to THIS
%   validate (H10). from_xml / convert_* inherited from BaseIntType.
%
%   NO python-pptx analogue -- this class is docx-specific (base for
%   ST_HpsMeasure and ST_TwipsMeasure).
%
%   RANGE-PRECISION HAZARD (D-STYPE-3): the upper bound 18446744073709551615
%   (2^64 - 1) exceeds 2^53, so a MATLAB double cannot hold it exactly (it
%   rounds to 18446744073709551616). The range test and the error message
%   therefore use the rounded upper bound. This is a DEAD upper edge: both
%   concrete subclasses override convert_to_xml to re-emit a re-derived twip /
%   half-point count (ST_TwipsMeasure, ST_HpsMeasure), and no realistic
%   measure reaches 2^64; XsdUnsignedLong.validate on the raw far-upper bound
%   is never exercised on a live docx path. Same divergence CLASS as XsdLong,
%   recorded under adopted deviation D-STYPE-3 (no new D-number).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::XsdUnsignedLong
%   (lines 175-178)

    methods (Static)
        function s = to_xml(value)
            % TO_XML BaseSimpleType.to_xml (lines 29-33), bound to THIS validate.
            mat2doc.oxml.simpletypes.XsdUnsignedLong.validate(value);
            s = mat2doc.oxml.simpletypes.XsdUnsignedLong.convert_to_xml(value);
        end

        function validate(value)
            % VALIDATE lines 176-178: int in range 0..18446744073709551615 (D-STYPE-3).
            mat2doc.oxml.simpletypes.BaseSimpleType.validate_int_in_range( ...
                value, 0, 18446744073709551615);
        end
    end
end

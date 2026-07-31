classdef CT_DecimalNumber < mat2doc.oxml.BaseOxmlElement
% CT_DECIMALNUMBER Used for `w:numId`, `w:ilvl`, `w:abstractNumId` and others.
%
%   Contains a text representation of a decimal number (e.g. 42) in its REQUIRED
%   `val` attribute (ST_DecimalNumber, an xsd:int). Direct sibling of the shared
%   CT_OnOff / CT_String simple-child classes.
%
%   `val` is a RequiredAttribute (ST_DecimalNumber): getAttrRequired raises
%   mat2doc:InvalidXmlError when @w:val is absent; the setter always writes it
%   (never removes) and range-validates via ST_DecimalNumber.to_xml. from_xml
%   parses the XML integer literal exactly as Python int(str_value) (H6),
%   returning a double holding an exact int.
%
%   FIRST REGISTERED for this WP at w:uiPriority (styles) and the deferred
%   w:outlineLvl (parfmt, now closed); the numbering/table CT_DecimalNumber
%   registrations (w:gridSpan/w:numId/... ) are DEFERRED to their P6/P8 WPs
%   (design.md section 4: register ONLY tags that appear in the WP's parts).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this class as feval(cls, name, ownDecls[, resolvedUri]) on the
%   <w:uiPriority>/<w:outlineLvl>/... nodes inside a real styles.xml, so all
%   positional args forward verbatim.
%
%   Example:
%       up = mat2doc.oxml.shared.CT_DecimalNumber.new("w:uiPriority", 9);
%       up.val                                 % 9  (double)
%
%   Ported from python-docx v1.2.0: src/docx/oxml/shared.py::CT_DecimalNumber
%   (lines 13-24)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        VAL_ATTR = "w:val"             % RequiredAttribute @ shared.py:18
        VAL_TYPE = "ST_DecimalNumber"  % simple type (+oxml\+simpletypes)
    end

    properties (Dependent)  % generated descriptor property
        val   % RequiredAttribute('w:val', ST_DecimalNumber) -> double (int); InvalidXmlError if absent
    end

    methods
        function obj = CT_DecimalNumber(varargin)
            % CT_DECIMALNUMBER Construct a loose decimal-number element -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        function value = get.val(obj)
            value = obj.getAttrRequired(obj.VAL_ATTR, obj.VAL_TYPE);
        end
        function set.val(obj, value)
            obj.setAttrRequired(obj.VAL_ATTR, obj.VAL_TYPE, value);
        end
    end

    methods (Static)
        function elm = new(nsptagname, val)
            % NEW A new CT_DecimalNumber with tag `nsptagname` and @w:val set to `val`.
            %   Ported from python-docx v1.2.0: src/docx/oxml/shared.py::
            %   CT_DecimalNumber.new (lines 20-24):
            %       return OxmlElement(nsptagname, attrs={qn("w:val"): str(val)})
            %   Unlike CT_String.new (which routes through the descriptor setter),
            %   this passes the attribute DIRECTLY as str(val) -- so it does NOT
            %   range-validate (Python str() never validates). H14: str(int) is
            %   pyStr kind "int" (digits, no decimal point, '-' sign only). The
            %   attrs currency is an Nx2 string [name,value] (OxmlElement).
            arguments
                nsptagname (1,1) string
                val (1,1) {mustBeNumeric}
            end
            elm = mat2doc.oxml.OxmlElement(nsptagname, ...
                [mat2doc.oxml.qn("w:val"), mat2doc.shared.pyStr(val, "int")]);
        end
    end
end

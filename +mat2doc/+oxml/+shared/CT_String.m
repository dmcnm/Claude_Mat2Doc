classdef CT_String < mat2doc.oxml.BaseOxmlElement
% CT_STRING Used for `w:rStyle`, `w:pStyle`, `w:tblStyle` and others.
%
%   Contains a string (e.g. a style name) in its REQUIRED `val` attribute. This
%   is the target of the w:rStyle font-block registry row (and w:pStyle /
%   w:basedOn / w:name / w:next / w:tblStyle elsewhere).
%
%   `val` is a RequiredAttribute (ST_String): getAttrRequired raises
%   mat2doc:InvalidXmlError when @w:val is absent (a malformed CT_String); the
%   setter always writes it (never removes). ST_String is the identity string
%   transform.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this class as feval(cls, name, ownDecls[, resolvedUri]) on the
%   <w:rStyle>/<w:pStyle>/... nodes inside a real styles.xml/document.xml.
%
%   Example:
%       s = mat2doc.oxml.shared.CT_String.new("w:rStyle", "Emphasis");
%       s.val                                  % "Emphasis"
%
%   Ported from python-docx v1.2.0: src/docx/oxml/shared.py::CT_String
%   (lines 39-52; registered for w:rStyle plus w:pStyle/w:basedOn/w:name/
%   w:next/w:tblStyle)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        VAL_ATTR = "w:val"      % RequiredAttribute @ shared.py:45
        VAL_TYPE = "ST_String"  % simple type (+oxml\+simpletypes)
    end

    properties (Dependent)  % generated descriptor property
        val   % RequiredAttribute('w:val', ST_String) -> string; InvalidXmlError if absent
    end

    methods
        function obj = CT_String(varargin)
            % CT_STRING Construct a loose value-string element -- TRANSPARENT PASS-THROUGH.
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
            % NEW A new CT_String with tag `nsptagname` and @w:val set to `val`.
            %   Ported from python-docx v1.2.0: src/docx/oxml/shared.py::
            %   CT_String.new (lines 47-52): elm = OxmlElement(nsptagname);
            %   elm.val = val; return elm. (The Python `cast(CT_String, ...)` is a
            %   type hint only -- OxmlElement already returns a CT_String for a
            %   registered nsptagname.)
            arguments
                nsptagname (1,1) string
                val (1,1) string
            end
            elm = mat2doc.oxml.OxmlElement(nsptagname);
            elm.val = val;
        end
    end
end

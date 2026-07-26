classdef BaseOxmlElement < mat2doc.oxml.XmlElement
% BASEOXMLELEMENT OPC-local base class for the [Content_Types].xml / .rels custom elements.
%
%   Base class for the OPC custom element classes (CT_Default, CT_Override,
%   CT_Types, CT_Relationship, CT_Relationships), adding standardized behavior in
%   one place. In docx this is `class BaseOxmlElement(etree.ElementBase)` in
%   docx/opc/oxml.py:75-85 -- a SEPARATE, minimal base DISTINCT from the
%   WordprocessingML xmlchemy BaseOxmlElement (mat2doc.oxml.BaseOxmlElement). It
%   is NOT the xmlchemy descriptor base: it adds only the `.xml` pretty-print
%   test-helper property; the CT_* subclasses are RAW etree elements that use
%   plain get/set/find/findall/append directly (NOT the ZeroOrOne/OptionalAttribute
%   descriptor machinery). This port therefore extends mat2doc.oxml.XmlElement
%   (the ElementBase analogue), NOT mat2doc.oxml.BaseOxmlElement -- so the CT_*
%   classes deliberately do NOT inherit xpath/getChild/getAttrTyped/etc.
%
%   `.xml` (Dependent property; Python @property, opc/oxml.py:79-85): pretty-
%   printed XML without a declaration, for tests -- routes through the single
%   pretty-print engine via mat2doc.opc.oxml.serialize_for_reading. Test-only
%   surface (see the engine's VERIFY note). NOTE: CT_Relationships SHADOWS this
%   `.xml` in Python to return declaration BYTES; MATLAB cannot override an
%   inherited property, so that byte member is rotated to `xml_file_bytes` there
%   (plan-audit condition B4) and CT_Relationships inherits this pretty `.xml`
%   unchanged.
%
%   TRANSPARENT PASS-THROUGH constructor (design.md section 2 "CT_* constructor
%   contract (INTEGRATION-CRITICAL)", INT-1): all positional args forwarded via
%   varargin to the XmlElement constructor with NO re-validation of the nsmap
%   argument, so the parser can instantiate registered CT_* classes via
%   feval(cls, name, ownDecls) where ownDecls is the Nx2 string decl-pair.
%
%   Example:
%       t = mat2doc.opc.oxml.CT_Types.new();
%       disp(t.xml)     % "<Types xmlns=...>\n</Types>\n"  (pretty, no declaration)
%
%   Ported from python-docx v1.2.0: src/docx/opc/oxml.py::BaseOxmlElement
%   (lines 75-85)

    properties (Dependent)
        xml   % pretty-printed XML string, no declaration (test helper)
    end

    methods
        function obj = BaseOxmlElement(varargin)
            obj = obj@mat2doc.oxml.XmlElement(varargin{:});
        end

        function value = get.xml(obj)
            % @property xml (opc/oxml.py 79-85): serialize_for_reading(self).
            value = mat2doc.opc.oxml.serialize_for_reading(obj);
        end
    end
end

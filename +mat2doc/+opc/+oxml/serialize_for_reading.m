function xml = serialize_for_reading(element)
% SERIALIZE_FOR_READING Human-readable (pretty-printed) XML for `element`, no declaration.
%
%   xml = MAT2DOC.OPC.OXML.SERIALIZE_FOR_READING(element) serializes `element` to
%   human-readable XML suitable for tests, with no XML declaration -- the MATLAB
%   replacement for docx's OPC-local
%   `etree.tostring(element, encoding="unicode", pretty_print=True)`
%   (opc/oxml.py:62-67).
%
%   ONE ENGINE, BOTH CALL SITES (plan-audit condition B1): docx defines an
%   IDENTICAL serialize_for_reading in both docx/oxml/xmlchemy.py:22-27 and
%   docx/opc/oxml.py:62-67 (same one-line etree.tostring call). This file is a
%   thin delegator to the single ported pretty-print engine
%   mat2doc.oxml.serialize_for_reading, provided as a distinct file for citation
%   fidelity to the docx module boundary. It is test-only surface -- see the
%   VERIFY note on the engine.
%
%   Inputs:  element - (1,1) mat2doc.oxml.XmlElement
%   Outputs: xml     - (1,1) string, pretty-printed, no declaration, trailing LF
%
%   Ported from python-docx v1.2.0: src/docx/opc/oxml.py::serialize_for_reading
%   (lines 62-67; implementation in mat2doc.oxml.serialize_for_reading)

xml = mat2doc.oxml.serialize_for_reading(element);
end

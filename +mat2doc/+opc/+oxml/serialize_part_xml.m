function xml_bytes = serialize_part_xml(part_elm)
% SERIALIZE_PART_XML Serialize `part_elm` to XML-part bytes (no insignificant whitespace, UTF-8 declaration).
%
%   xml_bytes = MAT2DOC.OPC.OXML.SERIALIZE_PART_XML(part_elm) serializes the
%   element tree rooted at part_elm to UTF-8 file bytes with the XML declaration
%   header -- the MATLAB replacement for
%   `etree.tostring(part_elm, encoding="UTF-8", standalone=True)`.
%
%   DELEGATION (plan-audit condition B5): docx DEFINES serialize_part_xml only
%   here, in src/docx/opc/oxml.py:53-59. P1-2 pre-ported it into
%   +oxml\serialize_part_xml.m (the byte-critical lxml serializer, validated at
%   P1-2 against harness\common\golden\*_docx.bin). P1-4 does NOT re-port it;
%   this file is a thin delegator provided for citation fidelity to the docx
%   module boundary (the symbol's home is opc/oxml.py). All byte conventions --
%   single-quote+LF declaration, insertion-order attributes, H7 escaping table,
%   H8/D-serializer-nsdecl verbatim-until-moved namespace handling -- are those
%   of mat2doc.oxml.serialize_part_xml.
%
%   Inputs:  part_elm  - (1,1) mat2doc.oxml.XmlElement, the part root element
%   Outputs: xml_bytes - 1xN uint8, exact file bytes (Python `bytes`)
%
%   Ported from python-docx v1.2.0: src/docx/opc/oxml.py::serialize_part_xml
%   (lines 53-59; implementation in mat2doc.oxml.serialize_part_xml, P1-2)

xml_bytes = mat2doc.oxml.serialize_part_xml(part_elm);
end

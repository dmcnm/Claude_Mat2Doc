function root = parse_xml(xml)
% PARSE_XML `etree.fromstring()` replacement that uses the OPC oxml parser.
%
%   root = MAT2DOC.OPC.OXML.PARSE_XML(xml) parses xml (UTF-8 bytes or a decoded
%   char/string XML literal) and returns the root XmlElement of an
%   order-preserving tree, applying the element-class registry so a parsed
%   <Types>/<Relationships>/<Default>/<Override>/<Relationship> instantiates the
%   corresponding CT_* subclass. This is the MATLAB replacement for docx's
%   OPC-local `parse_xml(text)` = `etree.fromstring(text, oxml_parser)`, where
%   `oxml_parser` carries the OPC element_class_lookup (opc/oxml.py:20-22).
%
%   DELEGATION + REGISTRY UNIFICATION (plan-audit condition B2, option A):
%   in docx there are TWO parsers -- the main WordprocessingML parser
%   (docx/oxml/parser.py) and this OPC parser (docx/opc/oxml.py) -- each with its
%   OWN element_class_lookup (the main one holds the 120 w:* classes; the OPC one
%   holds only the 5 CT_* classes bound via ct_namespace/pr_namespace,
%   opc/oxml.py:240-247). This port keeps ONE parser (mat2doc.oxml.parse_xml)
%   and ONE registry (mat2doc.oxml.registry) into which the 5 OPC classes are
%   merged BY RAW CLARK NAME (see registry.m). Unifying is behavior-preserving
%   because the key spaces are DISJOINT namespaces: an OPC part
%   ([Content_Types].xml / .rels / the CT_*.new literals) contains only ct:/pr:
%   tags, and a WordprocessingML part contains only the main-map tags, so no tag
%   ever resolves differently under the merged table than under docx's two
%   separate lookups. Hence this function is a thin delegator to
%   mat2doc.oxml.parse_xml; it exists as a distinct file for citation fidelity
%   to the docx module boundary (opc/oxml.py defines its own parse_xml).
%
%   Inputs:  xml  - 1xN uint8 (UTF-8 bytes) OR a (1,1) string / char row
%   Outputs: root - scalar mat2doc.oxml.XmlElement (or registered CT_* subclass)
%
%   Example:
%       t = mat2doc.opc.oxml.parse_xml("<Types xmlns=""" + ...
%           mat2doc.opc.NAMESPACE.OPC_CONTENT_TYPES + """/>");
%       class(t)   % "mat2doc.opc.oxml.CT_Types"
%
%   Ported from python-docx v1.2.0: src/docx/opc/oxml.py::parse_xml (lines 36-38)

root = mat2doc.oxml.parse_xml(xml);
end

function root = parse_xml(xml)
% PARSE_XML Return the root XmlElement obtained by parsing the XML in xml.
%
%   root = MAT2DOC.OXML.PARSE_XML(xml) parses xml -- either UTF-8 bytes (uint8
%   row vector, the on-disk part blob) or an already-decoded char row / string
%   scalar (a loose-element XML literal) -- and returns the root
%   mat2doc.oxml.XmlElement of an order-preserving tree, applying the class
%   registry at each element (registered CT_* classes instantiate; plain
%   XmlElement otherwise). This is the MATLAB replacement for
%   `etree.fromstring(xml, oxml_parser)`.
%
%   The parser is the read side of the own OOXML-subset XML layer (D-001):
%   Spike S1 proved matlab.io.xml.dom sorts attributes, so document order is
%   preserved by this purpose-built parser. It matches lxml 5.3.0 under the
%   docx config `etree.XMLParser(remove_blank_text=True,
%   resolve_entities=False)` (docx/oxml/parser.py:19). See the engine class
%   mat2doc.oxml.XmlParser for the parsed subset and behavior.
%
%   Both call currencies python-docx uses are supported: bytes (the on-disk
%   part blob) and str literals (loose-element XML fragments, e.g.
%   docx/opc/oxml.py CT_* .new() factories).
%
%   Inputs:  xml  - 1xN uint8 (UTF-8 bytes) OR a (1,1) string / char row
%   Outputs: root - scalar mat2doc.oxml.XmlElement (or registered subclass)
%
%   Example:
%       xml  = "<w:p xmlns:w=""http://schemas.openxmlformats.org/" + ...
%              "wordprocessingml/2006/main""><w:r><w:t>hi</w:t></w:r></w:p>";
%       root = mat2doc.oxml.parse_xml(xml);
%       root.tag                      % "{...wordprocessingml/2006/main}p"
%       % Bytes currency (an on-disk part blob) round-trips byte-for-byte:
%       blob  = uint8(unicode2native(xml, "UTF-8"));
%       bytes = mat2doc.oxml.serialize_part_xml(mat2doc.oxml.parse_xml(blob));
%
%   Ported from python-docx v1.2.0: src/docx/oxml/parser.py::parse_xml
%   (lines 23-29; design-realization of the read side, D-001)

% -- decode to a character string (H2): bytes via native2unicode UTF-8; an
%    already-decoded char/string is used verbatim (that is how python-docx
%    passes loose-element XML literals to parse_xml).
if isa(xml, "uint8")
    if ~isrow(xml) && ~isempty(xml)
        error("mat2doc:TypeError", "parse_xml expects a uint8 row vector of bytes");
    end
    text = string(native2unicode(xml, "UTF-8"));
elseif ischar(xml) && (isrow(xml) || isempty(xml))
    text = string(xml);
elseif isstring(xml) && isscalar(xml) && ~ismissing(xml)
    text = xml;
else
    error("mat2doc:TypeError", ...
        "parse_xml expects uint8 bytes, a char row, or a string scalar");
end

parser = mat2doc.oxml.XmlParser(text);
root = parser.parse();
end

classdef XmlPart < mat2doc.opc.Part
% XMLPART Base class for package parts carrying an XML payload (most of them).
%
%   An XmlPart parses its blob into an XmlElement tree on load and RE-SERIALIZES
%   that tree on save (the `blob` override) -- the byte-order-critical path: the
%   part is written as serialize_part_xml(element), byte-matched to lxml (P1-4).
%   This is WHY python-docx's open+save regenerates content parts rather than
%   copying their bytes: parsing with remove_blank_text=True strips the template's
%   pretty-print indentation, so re-serializing collapses the whitespace. It is
%   the mechanism behind the M1 finding -- registering exactly python-docx's XML
%   content types to XmlPart (vs. the generic passthrough Part) reproduces its
%   output part-for-part (see PartFactory).
%
%   H9 (lazy/eager currency): the element is parsed ONCE at load() time and held;
%   `element` returns that live tree, and `blob` re-serializes it on each access.
%   This matches python-docx: XmlPart.load parses eagerly (part.py 229-232), the
%   `element` property is a stored attribute, and `blob` re-serializes on demand
%   (part.py 220-222). The M1 reserialize path depends on this exact currency.
%
%   ARG ORDER (docx): XmlPart.__init__(partname, content_type, element, package)
%   -- element in the blob slot, package last (part.py 214-218). The base ctor is
%   called with NO blob (blob defaults to None); the element is stored separately.
%
%   Example:
%       % An XmlPart RE-SERIALIZES its parsed element on blob (contrast the
%       % verbatim base Part). Loading parses the XML; blob emits it afresh
%       % through serialize_part_xml, prepending the XML declaration.
%       xml = "<w:p xmlns:w='" + ...
%           "http://schemas.openxmlformats.org/wordprocessingml/2006/main'/>";
%       xp  = mat2doc.opc.XmlPart.load( ...
%           mat2doc.opc.PackURI("/word/document.xml"), ...
%           mat2doc.opc.CONTENT_TYPE.WML_DOCUMENT_MAIN, ...
%           uint8(unicode2native(xml, "UTF-8")), []);
%       disp(class(xp))                              % "mat2doc.opc.XmlPart"
%       disp(startsWith(char(xp.blob()), "<?xml"))   % true (re-serialized)
%
%   Ported from python-docx v1.2.0: src/docx/opc/part.py::XmlPart
%   (lines 207-247)

    properties (Access = private)
        element_
    end

    methods
        function obj = XmlPart(partname, content_type, element, package)
            % part.py 214-218: base ctor WITHOUT a blob (blob=None); store the
            %   parsed element.
            obj@mat2doc.opc.Part(partname, content_type, [], package);
            obj.element_ = element;
        end

        function b = blob(obj)
            % blob property OVERRIDE (part.py 220-222): the XML serialization of
            %   this part -- serialize_part_xml(self._element). The re-serialize
            %   path that must be byte-identical to lxml. Imported (docx part.py:9)
            %   from docx.opc.oxml -> mat2doc.opc.oxml.serialize_part_xml.
            b = mat2doc.opc.oxml.serialize_part_xml(obj.element_);
        end

        function e = element(obj)
            % element property (part.py 224-227): the root XML element of this
            %   XML part (the live parsed tree, H9).
            e = obj.element_;
        end

        function p = part(obj)
            % part property (part.py 234-241): this part -- the end of the
            %   parent-delegation chain for child objects.
            p = obj;
        end
    end

    methods (Static)
        function obj = load(partname, content_type, blob, package)
            % LOAD Return an XmlPart with the XML in `blob` parsed
            %   (part.py 229-232): element = parse_xml(blob). Imported
            %   (docx part.py:13) from docx.oxml.parser -> mat2doc.oxml.parse_xml.
            element = mat2doc.oxml.parse_xml(blob);
            obj = mat2doc.opc.XmlPart(partname, content_type, element, package);
        end
    end

    methods (Access = protected)
        function n = rel_ref_count_(obj, rId)
            % _rel_ref_count OVERRIDE (part.py 243-247): count of references to
            %   `rId` in this part's XML -- len([r for r in xpath("//@r:id")
            %   if r == rId]). xpath returns a (1,N) string array (empty if none).
            rIds = obj.element_.xpath("//@r:id");
            n = sum(rIds == string(rId));
        end
    end
end

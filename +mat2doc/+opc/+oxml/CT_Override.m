classdef CT_Override < mat2doc.opc.oxml.BaseOxmlElement
% CT_OVERRIDE `<Override>` element in `[Content_Types].xml`.
%
%   Specifies the content type applied to the part with the given partname. RAW
%   etree element (see CT_Default): attributes read via plain get, written by the
%   `new` factory via plain set; no validation or simple-type conversion.
%   Registered for `ct:Override` (Clark {..content-types}Override).
%
%   The content_type / partname @property getters return the plain attribute
%   string, or [] (None, H3) when absent.
%
%   TRANSPARENT PASS-THROUGH constructor (design.md section 2, INT-1).
%
%   Example:
%       o = mat2doc.opc.oxml.CT_Override.new("/word/document.xml", ...
%           mat2doc.opc.CONTENT_TYPE.WML_DOCUMENT_MAIN);
%       disp(o.partname)       % "/word/document.xml"
%       disp(o.content_type)   % "...wordprocessingml.document.main+xml"
%
%   Ported from python-docx v1.2.0: src/docx/opc/oxml.py::CT_Override (lines
%   115-138; registered for `ct:Override` via ct_namespace at line 242)

    properties (Dependent)
        content_type   % ContentType attribute (opc/oxml.py 119-123)
        partname       % PartName attribute (opc/oxml.py 134-138)
    end

    methods
        function obj = CT_Override(varargin)
            obj = obj@mat2doc.opc.oxml.BaseOxmlElement(varargin{:});
        end

        function value = get.content_type(obj)
            % @property content_type (opc/oxml.py 119-123): self.get("ContentType").
            value = obj.get("ContentType");   % [] (None) when absent (H3)
        end

        function value = get.partname(obj)
            % @property partname (opc/oxml.py 134-138): self.get("PartName").
            value = obj.get("PartName");      % [] (None) when absent (H3)
        end
    end

    methods (Static)
        function override = new(partname, content_type)
            % NEW Return a new `<Override>` element with attributes set (opc/oxml.py
            %   125-132). Attribute set order is PartName THEN ContentType (H11),
            %   e.g. <Override PartName="/word/document.xml" ContentType="..."/>.
            %   partname is a PackURI (a str subclass) at the P1-6a call site;
            %   coerced with string() so its partname text is stored (XmlElement.set
            %   requires text), mirroring lxml accepting the str-subclass directly.
            arguments
                partname
                content_type (1,1) string
            end
            % Python: parse_xml('<Override xmlns="%s"/>' % nsmap["ct"])
            xml = "<Override xmlns=""" + ...
                mat2doc.opc.oxml.nsmap().ct + """/>";
            override = mat2doc.opc.oxml.parse_xml(xml);
            override.set("PartName", string(partname));
            override.set("ContentType", content_type);
        end
    end
end

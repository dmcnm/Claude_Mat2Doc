classdef CT_Default < mat2doc.opc.oxml.BaseOxmlElement
% CT_DEFAULT `<Default>` element in `[Content_Types].xml`.
%
%   Specifies the default content type applied to any part with the given
%   extension. RAW etree element (docx opc CT_* do NOT use the xmlchemy
%   descriptor engine): the two attributes are read via plain get and written by
%   the `new` factory via plain set -- no RequiredAttribute validation, no
%   simple-type conversion (contrast the Mat2Ppt pptx CT_Default, which is a
%   descriptor class; docx is deliberately rawer, so NO validation is added here
%   -- that would be behavior-adding). Registered for `ct:Default` (Clark
%   {..content-types}Default) so the parser instantiates it for every <Default>.
%
%   The two @property getters return the plain attribute string, or [] (None,
%   H3) when the attribute is absent -- docx does NOT raise on a missing
%   attribute here (unlike a RequiredAttribute).
%
%   TRANSPARENT PASS-THROUGH constructor (design.md section 2, INT-1): all
%   positional args via varargin, NO struct typing of the nsmap arg, so the
%   parser can construct it via feval(cls, name, ownDecls).
%
%   Example:
%       d = mat2doc.opc.oxml.CT_Default.new("png", mat2doc.opc.CONTENT_TYPE.PNG);
%       disp(d.extension)      % "png"
%       disp(d.content_type)   % "image/png"
%
%   Ported from python-docx v1.2.0: src/docx/opc/oxml.py::CT_Default (lines
%   88-112; registered for `ct:Default` via ct_namespace at line 241)

    properties (Dependent)
        content_type   % ContentType attribute (opc/oxml.py 94-98)
        extension      % Extension attribute (opc/oxml.py 100-103)
    end

    methods
        function obj = CT_Default(varargin)
            obj = obj@mat2doc.opc.oxml.BaseOxmlElement(varargin{:});
        end

        function value = get.content_type(obj)
            % @property content_type (opc/oxml.py 94-98): self.get("ContentType").
            value = obj.get("ContentType");   % [] (None) when absent (H3)
        end

        function value = get.extension(obj)
            % @property extension (opc/oxml.py 100-103): self.get("Extension").
            value = obj.get("Extension");     % [] (None) when absent (H3)
        end
    end

    methods (Static)
        function default = new(ext, content_type)
            % NEW Return a new `<Default>` element with attributes set (opc/oxml.py
            %   105-112). Attribute set order is Extension THEN ContentType --
            %   that insertion order IS the serialized attribute order (H11), e.g.
            %   <Default Extension="jpeg" ContentType="image/jpeg"/>.
            arguments
                ext (1,1) string
                content_type (1,1) string
            end
            % Python: parse_xml('<Default xmlns="%s"/>' % nsmap["ct"])
            xml = "<Default xmlns=""" + ...
                mat2doc.opc.oxml.nsmap().ct + """/>";
            default = mat2doc.opc.oxml.parse_xml(xml);
            default.set("Extension", ext);
            default.set("ContentType", content_type);
        end
    end
end

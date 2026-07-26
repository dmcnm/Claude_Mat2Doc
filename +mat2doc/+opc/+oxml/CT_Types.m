classdef CT_Types < mat2doc.opc.oxml.BaseOxmlElement
% CT_TYPES `<Types>` element, the container/root of `[Content_Types].xml`.
%
%   Container for <Default> and <Override> elements. RAW etree container (docx
%   opc CT_* do NOT use the xmlchemy child-descriptor engine): add_default /
%   add_override build a CT_Default / CT_Override and APPEND it (plain lxml
%   append, NOT a ZeroOrMore inserter). Appends VERBATIM in caller order and
%   NEVER reorders (trap 6): the sort of defaults-by-extension and
%   overrides-by-partname, and the <Default>-before-<Override> grouping, are the
%   CALLER's responsibility (pkgwriter._ContentTypesItem._element, P1-6a), not
%   this element's. Registered for `ct:Types` (Clark {..content-types}Types).
%
%   TRANSPARENT PASS-THROUGH constructor (design.md section 2, INT-1).
%
%   Example:
%       t = mat2doc.opc.oxml.CT_Types.new();
%       t.add_default("xml", mat2doc.opc.CONTENT_TYPE.XML);
%       t.add_override("/word/document.xml", ...
%           mat2doc.opc.CONTENT_TYPE.WML_DOCUMENT_MAIN);
%       disp(numel(t.defaults))    % 1
%       disp(numel(t.overrides))   % 1
%
%   Ported from python-docx v1.2.0: src/docx/opc/oxml.py::CT_Types (lines
%   209-237; registered for `ct:Types` via ct_namespace at line 243)

    properties (Dependent)
        defaults    % list of <ct:Default> children, document order (opc/oxml.py 224-226)
        overrides   % list of <ct:Override> children, document order (opc/oxml.py 235-237)
    end

    methods
        function obj = CT_Types(varargin)
            obj = obj@mat2doc.opc.oxml.BaseOxmlElement(varargin{:});
        end

        function add_default(obj, ext, content_type)
            % ADD_DEFAULT Add a <Default> child (opc/oxml.py 213-216). Like
            %   Python, returns NOTHING -- builds and appends.
            arguments
                obj (1,1) mat2doc.opc.oxml.CT_Types
                ext (1,1) string
                content_type (1,1) string
            end
            default = mat2doc.opc.oxml.CT_Default.new(ext, content_type);
            obj.append(default);   % plain lxml append (H8: MOVED -> D-serializer-nsdecl)
        end

        function add_override(obj, partname, content_type)
            % ADD_OVERRIDE Add an <Override> child (opc/oxml.py 218-222). Like
            %   Python, returns NOTHING. partname is a PackURI at the P1-6a call
            %   site (coerced to string inside CT_Override.new).
            arguments
                obj (1,1) mat2doc.opc.oxml.CT_Types
                partname
                content_type (1,1) string
            end
            override = mat2doc.opc.oxml.CT_Override.new(partname, content_type);
            obj.append(override);
        end

        function list = get.defaults(obj)
            % @property defaults (opc/oxml.py 224-226):
            %   self.findall(qn("ct:Default")) -- OPC-LOCAL qn (ct: prefix).
            list = obj.findall(mat2doc.opc.oxml.qn("ct:Default"));
        end

        function list = get.overrides(obj)
            % @property overrides (opc/oxml.py 235-237):
            %   self.findall(qn("ct:Override")) -- OPC-LOCAL qn (ct: prefix).
            list = obj.findall(mat2doc.opc.oxml.qn("ct:Override"));
        end
    end

    methods (Static)
        function types = new()
            % NEW Return a new `<Types>` element (opc/oxml.py 228-233).
            xml = "<Types xmlns=""" + ...
                mat2doc.opc.oxml.nsmap().ct + """/>";
            types = mat2doc.opc.oxml.parse_xml(xml);
        end
    end
end

classdef CT_Relationships < mat2doc.opc.oxml.BaseOxmlElement
% CT_RELATIONSHIPS `<Relationships>` element, the root element of a `.rels` part.
%
%   RAW etree container (docx opc CT_* do NOT use the xmlchemy child-descriptor
%   engine): add_rel builds a CT_Relationship and APPENDS it (plain lxml append,
%   NOT a ZeroOrMore inserter -- contrast the Mat2Ppt pptx CT_Relationships,
%   which routes through insert_relationship_). Registered for `pr:Relationships`
%   (Clark {..package/2006/relationships}Relationships) so the parser
%   instantiates it as the root of a parsed .rels.
%
%   `.xml` BYTE-MEMBER ROTATION (plan-audit condition B4): Python CT_Relationships
%   OVERRIDES the base `.xml` @property (opc/oxml.py 202-206) to return the .rels
%   FILE BYTES (serialize_part_xml, declaration + no pretty-print), shadowing
%   BaseOxmlElement's pretty `.xml`. MATLAB cannot override an inherited property,
%   so that byte member is rotated to `xml_file_bytes` (the exact python-pptx
%   v1.0.2 / Mat2Ppt precedent, +mat2ppt/+opc/+oxml/CT_Relationships.m). This
%   class therefore INHERITS BaseOxmlElement's pretty `.xml` unchanged and adds
%   `xml_file_bytes` for the byte path. HAND-OFF: the sole byte consumer is
%   Relationships.xml (rel.py, P1-5) -- the P1-5 brief MUST call the rotated name
%   `xml_file_bytes`, and the P1-5/phys_pkg hand-off carries this (phys_pkg is NOT
%   implemented in P1-4).
%
%   TRANSPARENT PASS-THROUGH constructor (design.md section 2, INT-1).
%
%   Example:
%       rels = mat2doc.opc.oxml.CT_Relationships.new();
%       rels.add_rel("rId1", mat2doc.opc.RELATIONSHIP_TYPE.OFFICE_DOCUMENT, ...
%           "word/document.xml", false);
%       bytes = rels.xml_file_bytes;             % uint8 .rels bytes + declaration
%       disp(numel(rels.Relationship_lst))       % 1
%
%   Ported from python-docx v1.2.0: src/docx/opc/oxml.py::CT_Relationships (lines
%   181-206; registered for `pr:Relationships` via pr_namespace at line 247)

    properties (Dependent)
        Relationship_lst   % list of <pr:Relationship> children, document order
        xml_file_bytes     % .rels file bytes (declaration + no whitespace); Python `.xml` override
    end

    methods
        function obj = CT_Relationships(varargin)
            obj = obj@mat2doc.opc.oxml.BaseOxmlElement(varargin{:});
        end

        function add_rel(obj, rId, reltype, target, is_external)
            % ADD_REL Add a child <Relationship> with attributes set (opc/oxml.py
            %   184-189). target_mode = EXTERNAL if is_external else INTERNAL. Like
            %   Python, returns NOTHING -- it appends and stops.
            arguments
                obj (1,1) mat2doc.opc.oxml.CT_Relationships
                rId (1,1) string
                reltype (1,1) string
                target (1,1) string
                is_external (1,1) logical = false   % Python default is_external=False
            end
            if is_external
                target_mode = mat2doc.opc.RELATIONSHIP_TARGET_MODE.EXTERNAL;
            else
                target_mode = mat2doc.opc.RELATIONSHIP_TARGET_MODE.INTERNAL;
            end
            relationship = mat2doc.opc.oxml.CT_Relationship.new(rId, reltype, ...
                target, target_mode);
            obj.append(relationship);   % plain lxml append (H8: MOVED -> D-serializer-nsdecl)
        end

        function list = get.Relationship_lst(obj)
            % @property Relationship_lst (opc/oxml.py 197-200):
            %   self.findall(qn("pr:Relationship")) -- OPC-LOCAL qn (pr: prefix).
            list = obj.findall(mat2doc.opc.oxml.qn("pr:Relationship"));
        end

        function value = get.xml_file_bytes(obj)
            % Python `.xml` override (opc/oxml.py 202-206): serialize_part_xml(self)
            %   -- .rels bytes with declaration. Rotated name per condition B4.
            value = mat2doc.opc.oxml.serialize_part_xml(obj);
        end
    end

    methods (Static)
        function elm = new()
            % NEW Return a new `<Relationships>` element (opc/oxml.py 191-195).
            xml = "<Relationships xmlns=""" + ...
                mat2doc.opc.oxml.nsmap().pr + """/>";
            elm = mat2doc.opc.oxml.parse_xml(xml);
        end
    end
end

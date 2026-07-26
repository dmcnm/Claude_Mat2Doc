classdef CT_Relationship < mat2doc.opc.oxml.BaseOxmlElement
% CT_RELATIONSHIP `<Relationship>` element in a `.rels` part.
%
%   Represents a single relationship from a source to a target part. RAW etree
%   element (see CT_Default): four attributes read via plain get, written by the
%   `new` factory via plain set; NO validation, NO simple-type conversion
%   (contrast the Mat2Ppt pptx CT_Relationship, a descriptor class with
%   RequiredAttribute/OptionalAttribute -- docx is deliberately rawer, so none of
%   that is added here). Registered for `pr:Relationship` (Clark
%   {..package/2006/relationships}Relationship).
%
%   TargetMode tri-state (trap 2, H3): the `new` factory writes the TargetMode
%   attribute ONLY when target_mode == "External"; INTERNAL relationships omit
%   the attribute entirely. The target_mode @property getter returns the stored
%   attribute or the DEFAULT "Internal" when absent (two-arg get with default).
%
%   TRANSPARENT PASS-THROUGH constructor (design.md section 2, INT-1).
%
%   Example:
%       r = mat2doc.opc.oxml.CT_Relationship.new("rId1", ...
%           mat2doc.opc.RELATIONSHIP_TYPE.OFFICE_DOCUMENT, "word/document.xml");
%       disp(r.rId)           % "rId1"
%       disp(r.target_mode)   % "Internal"  (default; no TargetMode attr written)
%
%   Ported from python-docx v1.2.0: src/docx/opc/oxml.py::CT_Relationship (lines
%   140-178; registered for `pr:Relationship` via pr_namespace at line 246)

    properties (Dependent)
        rId           % Id attribute (opc/oxml.py 155-158)
        reltype       % Type attribute (opc/oxml.py 160-163)
        target_ref    % Target attribute (opc/oxml.py 165-169)
        target_mode   % TargetMode attribute, default "Internal" (opc/oxml.py 171-178)
    end

    methods
        function obj = CT_Relationship(varargin)
            obj = obj@mat2doc.opc.oxml.BaseOxmlElement(varargin{:});
        end

        function value = get.rId(obj)
            % @property rId (opc/oxml.py 155-158): self.get("Id").
            value = obj.get("Id");     % [] (None) when absent (H3)
        end

        function value = get.reltype(obj)
            % @property reltype (opc/oxml.py 160-163): self.get("Type").
            value = obj.get("Type");   % [] (None) when absent (H3)
        end

        function value = get.target_ref(obj)
            % @property target_ref (opc/oxml.py 165-169): self.get("Target").
            value = obj.get("Target"); % [] (None) when absent (H3)
        end

        function value = get.target_mode(obj)
            % @property target_mode (opc/oxml.py 171-178): self.get("TargetMode",
            %   RTM.INTERNAL) -- default "Internal" when the attribute is absent.
            value = obj.get("TargetMode", ...
                mat2doc.opc.RELATIONSHIP_TARGET_MODE.INTERNAL);
        end
    end

    methods (Static)
        function relationship = new(rId, reltype, target, target_mode)
            % NEW Return a new `<Relationship>` element (opc/oxml.py 143-153).
            %   Attribute set order is Id, Type, Target, [TargetMode] (H11), e.g.
            %   <Relationship Id="rId3" Type="..." Target="..."/>. TargetMode is
            %   written ONLY when target_mode == EXTERNAL (trap 2); target_mode
            %   defaults to RTM.INTERNAL, matching the Python keyword default.
            arguments
                rId (1,1) string
                reltype (1,1) string
                target (1,1) string
                target_mode (1,1) string = mat2doc.opc.RELATIONSHIP_TARGET_MODE.INTERNAL
            end
            % Python: parse_xml('<Relationship xmlns="%s"/>' % nsmap["pr"])
            xml = "<Relationship xmlns=""" + ...
                mat2doc.opc.oxml.nsmap().pr + """/>";
            relationship = mat2doc.opc.oxml.parse_xml(xml);
            relationship.set("Id", rId);
            relationship.set("Type", reltype);
            relationship.set("Target", target);
            % Python: if target_mode == RTM.EXTERNAL: set("TargetMode", RTM.EXTERNAL)
            if target_mode == mat2doc.opc.RELATIONSHIP_TARGET_MODE.EXTERNAL
                relationship.set("TargetMode", ...
                    mat2doc.opc.RELATIONSHIP_TARGET_MODE.EXTERNAL);
            end
        end
    end
end

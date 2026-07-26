classdef SerializedRelationship_ < handle
% SERIALIZEDRELATIONSHIP_ Value object for a serialized relationship in a package.
%
%   Serialized means the target part is referred to via its partname rather than a
%   direct link to an in-memory Part object. Constructed from a baseURI (the
%   source part's baseURI) and a parsed <Relationship> element (CT_Relationship).
%   A handle class: the Python original is a plain object AND `target_partname`
%   lazy-caches its computed PackURI (see below), which requires reference
%   semantics.
%
%   V-3 LAZY-CACHE (hand-off from P1-5 audit FLAG-1 ruling): in python-docx
%   v1.2.0 the ONLY relationship class that caches its target partname is THIS one
%   (`_SerializedRelationship.target_partname`, pkgreader.py 211-227), guarded by
%   `hasattr(self, "_target_partname")`. The P1-5 `_Relationship` (rel.py) does
%   NOT cache and has no `target_partname` -- that caching belongs HERE. Ported
%   faithfully with a private cache property + a logical computed-flag (design.md
%   section 2 lazyproperty pattern; NEVER isempty as the sentinel).
%
%   UNDERSCORE ROTATION (design.md section 2): `_SerializedRelationship` ->
%   SerializedRelationship_.
%
%   Ported from python-docx v1.2.0: src/docx/opc/pkgreader.py::_SerializedRelationship
%   (lines 167-227)

    properties (Access = private)
        baseURI_               % string, the source part's baseURI
        rId_                   % string (Id attribute)
        reltype_               % string (Type attribute)
        target_mode_           % string (TargetMode attribute; default "Internal")
        target_ref_            % string (Target attribute)
        target_partname_ = []  % cached PackURI (V-3 lazy-load)
        has_target_partname_ (1,1) logical = false  % hasattr(self,"_target_partname")
    end

    properties (Dependent, SetAccess = private)
        is_external
        reltype
        rId
        target_mode
        target_ref
        target_partname
    end

    methods
        function obj = SerializedRelationship_(baseURI, rel_elm)
            % pkgreader.py 174-180: capture baseURI and the four relationship
            %   attributes off the parsed <Relationship> element (CT_Relationship).
            obj.baseURI_ = baseURI;
            obj.rId_ = rel_elm.rId;
            obj.reltype_ = rel_elm.reltype;
            obj.target_mode_ = rel_elm.target_mode;
            obj.target_ref_ = rel_elm.target_ref;
        end

        function value = get.is_external(obj)
            % @property is_external (pkgreader.py 182-185): target_mode == EXTERNAL.
            value = obj.target_mode_ == mat2doc.opc.RELATIONSHIP_TARGET_MODE.EXTERNAL;
        end

        function value = get.reltype(obj)
            % @property reltype (pkgreader.py 187-190).
            value = obj.reltype_;
        end

        function value = get.rId(obj)
            % @property rId (pkgreader.py 192-196).
            value = obj.rId_;
        end

        function value = get.target_mode(obj)
            % @property target_mode (pkgreader.py 198-202).
            value = obj.target_mode_;
        end

        function value = get.target_ref(obj)
            % @property target_ref (pkgreader.py 204-209).
            value = obj.target_ref_;
        end

        function value = get.target_partname(obj)
            % @property target_partname (pkgreader.py 211-227): PackURI targeted by
            %   this relationship. Raises ValueError when target_mode == External.
            %   Lazy-loads and caches _target_partname on first internal access.
            if obj.is_external
                % Python message is split across two source lines (218-223);
                % concatenated here to the single string it denotes.
                error("mat2doc:ValueError", ...
                    "target_partname attribute on Relationship is undefined " + ...
                    "where TargetMode == ""External""");
            end
            % lazy-load _target_partname attribute (hasattr guard)
            if ~obj.has_target_partname_
                obj.target_partname_ = mat2doc.opc.PackURI.from_rel_ref( ...
                    obj.baseURI_, obj.target_ref);
                obj.has_target_partname_ = true;
            end
            value = obj.target_partname_;
        end
    end
end

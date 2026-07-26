classdef Relationship_ < handle
% RELATIONSHIP_ Value object for a relationship to a part.
%
%   Holds rId / reltype / target / baseURI / is_external. `target` is either a
%   Part (internal relationship) or a string URI (external). is_external, reltype,
%   rId, target_part and target_ref are read-only accessors. For an internal
%   relationship target_ref is the target partname made relative to baseURI (the
%   serialization form); for external it is the URI verbatim; target_part raises
%   for an external relationship.
%
%   NO CACHING (IMPORTANT -- deviates from the P1-5 brief's hazard #3 premise).
%   python-docx v1.2.0 rel.py declares EVERY _Relationship accessor as a PLAIN
%   `@property` (rel.py:127-153, catalog-confirmed) -- there is NO `@lazyproperty`
%   and NO `target_partname` member at all. target_ref is RECOMPUTED on each
%   access. The Mat2Ppt caching precedent (#11) is for pptx `package.py`'s
%   `@lazyproperty`-based `_Relationship`, a DIFFERENT (refactored) design; docx's
%   older rel.py does not cache. Adding a cache here would be behavior-adding
%   (design.md section 7), so this port mirrors the plain @property form exactly.
%   Serialized bytes are identical either way (target_ref is computed once per
%   emit). See the audit record hazard #3 / VERIFY note. Class is `handle` for
%   reference semantics (design.md section 2); it is never mutated after
%   construction.
%
%   UNDERSCORE ROTATION (design.md section 2): Python `_Relationship` ->
%   Relationship_; `_rId`/`_reltype`/`_target`/`_baseURI`/`_is_external` ->
%   rId_/reltype_/target_/base_uri_/is_external_.
%
%   Example:
%       rel = mat2doc.opc.Relationship_("rId1", ...
%           mat2doc.opc.RELATIONSHIP_TYPE.HYPERLINK, "http://example", "/", true);
%       disp(rel.is_external)    % true
%       disp(rel.rId)            % "rId1"
%       disp(rel.target_ref)     % "http://example"  (URI verbatim)
%
%   Ported from python-docx v1.2.0: src/docx/opc/rel.py::_Relationship
%   (lines 114-153)

    properties (Access = private)
        rId_
        reltype_
        target_        % Part (internal) OR string URI (external)
        base_uri_
        is_external_ (1,1) logical = false
    end

    properties (Dependent, SetAccess = private)
        is_external
        reltype
        rId
        target_part
        target_ref
    end

    methods
        function obj = Relationship_(rId, reltype, target, baseURI, external)
            % rel.py 117-125. Argument order rId, reltype, target, baseURI,
            %   external (external default False). is_external = bool(external).
            arguments
                rId
                reltype
                target
                baseURI
                external = false
            end
            obj.rId_ = rId;
            obj.reltype_ = reltype;
            obj.target_ = target;
            obj.base_uri_ = baseURI;
            obj.is_external_ = logical(external);   % Python: bool(external)
        end

        function value = get.is_external(obj)
            % @property is_external (rel.py 127-129).
            value = obj.is_external_;
        end

        function value = get.reltype(obj)
            % @property reltype (rel.py 131-133).
            value = obj.reltype_;
        end

        function value = get.rId(obj)
            % @property rId (rel.py 135-137).
            value = obj.rId_;
        end

        function value = get.target_part(obj)
            % @property target_part (rel.py 139-145): undefined (raises) when the
            %   relationship is external.
            if obj.is_external_
                error("mat2doc:ValueError", ...
                    "target_part property on _Relationship is undefined " + ...
                    "when target mode is External");
            end
            value = obj.target_;
        end

        function value = get.target_ref(obj)
            % @property target_ref (rel.py 147-153): external -> the URI verbatim;
            %   internal -> the target partname relative to baseURI. RECOMPUTED
            %   each access (plain @property; see NO CACHING note).
            if obj.is_external_
                value = obj.target_;   % the URI string
            else
                value = obj.target_.partname.relative_ref(obj.base_uri_);
            end
        end
    end
end

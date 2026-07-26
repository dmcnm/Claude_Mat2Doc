classdef Relationships < handle
% RELATIONSHIPS Collection of |Relationship_| objects with dict semantics.
%
%   python-docx `Relationships` is PUBLIC and subclasses `Dict[str, _Relationship]`
%   (rel.py 13): relationships are keyed by rId, iterated in dict INSERTION order.
%   This collection adds relationships (add_relationship / get_or_add /
%   get_or_add_ext_rel), looks a part up by relationship type (part_with_reltype),
%   exposes the internal-target map (related_parts), and regenerates the .rels
%   bytes (the `xml` accessor).
%
%   ORDERING (H11, CRITICAL docx-vs-pptx divergence): the `xml` accessor iterates
%   `self.values()` in INSERTION order and does NOT sort by rId -- contrast the
%   Mat2Ppt pptx `_Relationships.xml`, which sorts <Relationship> elements by
%   numeric rId. Ported the docx insertion-order behavior exactly: relationships
%   are held in insertion-ordered parallel arrays (rIds_ / rels_), and the emitted
%   .rels element order is the order relationships were added.
%
%   DICT BASE SURFACE: docx code and its consumers use the inherited dict surface
%   -- `rels[rId]` (getitem), `rId in rels` (contains), `len(rels)`, `rels.values()`.
%   These are ported as getitem/contains/len/keys/values (faithful realization of
%   the Dict base, not added features).
%
%   `related_parts` currency (H3/H11): returns the {rId -> target Part} map for
%   INTERNAL relationships as a MATLAB dictionary(string -> cell); a consumer
%   reads a part with `rp(rId){1}` (the Part handle is cell-wrapped, the
%   established Mat2Doc map currency). Built incrementally by add_relationship for
%   internal targets only, matching `self._target_parts_by_rId`.
%
%   `xml` HAND-OFF (P1-4 condition B4 -> P1-6a): Python `Relationships.xml`
%   (rel.py 61-68) returns `CT_Relationships.new()...xml` -- the .rels FILE BYTES.
%   Because MATLAB cannot override the inherited pretty `.xml`, P1-4 rotated that
%   byte member to `CT_Relationships.xml_file_bytes`; this accessor calls the
%   rotated name and returns uint8 bytes. The sole consumer is P1-6a's
%   PackageWriter, which does `phys_writer.write(uri, part.rels.xml)` /
%   `_write_pkg_rels` (docx pkgwriter.py 54, 59) -> ZipPkgWriter_.write.
%
%   UNDERSCORE ROTATION (design.md section 2): the private `_add_relationship`
%   stays PUBLIC in docx (add_relationship, NO underscore); `_get_matching` /
%   `_get_rel_of_type` / `_next_rId` -> get_matching_ / get_rel_of_type_ /
%   next_rId_.
%
%   Example:
%       r  = mat2doc.opc.Relationships("/");
%       RT = mat2doc.opc.RELATIONSHIP_TYPE;
%       id1 = r.get_or_add_ext_rel(RT.HYPERLINK, "http://a.example");
%       id2 = r.get_or_add_ext_rel(RT.HYPERLINK, "http://b.example");
%       disp(id1 + " " + id2)                    % "rId1 rId2"
%       disp(startsWith(char(r.xml), "<?xml"))   % true (.rels bytes)
%
%   Ported from python-docx v1.2.0: src/docx/opc/rel.py::Relationships
%   (lines 13-111)

    properties (Access = private)
        base_uri_ (1,1) string
        rIds_ = string.empty(1, 0)                       % dict keys, insertion order
        rels_ = mat2doc.opc.Relationship_.empty(1, 0)    % parallel to rIds_
        target_parts_by_rId_                             % dictionary(string -> cell)
    end

    properties (Dependent, SetAccess = private)
        related_parts   % {rId -> target Part} for internal rels (rel.py 55-59)
        xml             % .rels file bytes (uint8) (rel.py 61-68)
    end

    methods
        function obj = Relationships(baseURI)
            % rel.py 16-19.
            arguments
                baseURI (1,1) string
            end
            obj.base_uri_ = baseURI;
            obj.target_parts_by_rId_ = configureDictionary("string", "cell");
        end

        % ---- inherited dict surface (Dict[str, _Relationship]) ----
        function tf = contains(obj, rId)
            % __contains__ / `rId in rels`.
            tf = any(obj.rIds_ == string(rId));
        end

        function rel = getitem(obj, rId)
            % __getitem__ / rels[rId]; KeyError if absent.
            idx = find(obj.rIds_ == string(rId), 1);
            if isempty(idx)
                error("mat2doc:KeyError", "no relationship with key '%s'", ...
                    string(rId));
            end
            rel = obj.rels_(idx);
        end

        function k = keys(obj)
            % Mapping keys (rIds), insertion order.
            k = obj.rIds_;
        end

        function v = values(obj)
            % Mapping values (the Relationship_ objects), insertion order.
            v = obj.rels_;
        end

        function n = len(obj)
            % __len__.
            n = numel(obj.rIds_);
        end

        % ---- rel.py public API ----
        function rel = add_relationship(obj, reltype, target, rId, is_external)
            % ADD_RELATIONSHIP Add and return a Relationship_ (rel.py 21-29).
            %   Mirrors dict setitem `self[rId] = rel` (replace value in place if
            %   the key already exists; else append). Internal targets are also
            %   recorded in the related_parts map.
            arguments
                obj (1,1) mat2doc.opc.Relationships
                reltype (1,1) string
                target
                rId (1,1) string
                is_external (1,1) logical = false   % Python default is_external=False
            end
            rel = mat2doc.opc.Relationship_(rId, reltype, target, ...
                obj.base_uri_, is_external);
            % self[rId] = rel
            idx = find(obj.rIds_ == rId, 1);
            if isempty(idx)
                obj.rIds_(end + 1) = rId;
                obj.rels_(end + 1) = rel;
            else
                obj.rels_(idx) = rel;   % dict: replace value, keep position
            end
            if ~is_external
                obj.target_parts_by_rId_(rId) = {target};
            end
        end

        function rel = get_or_add(obj, reltype, target_part)
            % GET_OR_ADD Relationship of `reltype` to `target_part`, added if not
            %   already present (rel.py 31-38). Returns the Relationship_.
            rel = obj.get_matching_(reltype, target_part, false);
            if isempty(rel)   % Python: if rel is None (H3)
                rId = obj.next_rId_();
                rel = obj.add_relationship(reltype, target_part, rId, false);
            end
        end

        function rId = get_or_add_ext_rel(obj, reltype, target_ref)
            % GET_OR_ADD_EXT_REL rId of an external `reltype` to `target_ref`,
            %   added if not already present (rel.py 40-47). Returns the rId.
            rel = obj.get_matching_(reltype, string(target_ref), true);
            if isempty(rel)   % Python: if rel is None (H3)
                rId = obj.next_rId_();
                rel = obj.add_relationship(reltype, string(target_ref), rId, true);
            end
            rId = rel.rId;   % Python: return rel.rId
        end

        function part = part_with_reltype(obj, reltype)
            % PART_WITH_RELTYPE Target part of the single rel of `reltype`
            %   (rel.py 49-53). KeyError if none, ValueError if more than one.
            rel = obj.get_rel_of_type_(reltype);
            part = rel.target_part;
        end

        function value = get.related_parts(obj)
            % @property related_parts (rel.py 55-59): the {rId -> target Part} map
            %   for internal relationships (self._target_parts_by_rId).
            value = obj.target_parts_by_rId_;
        end

        function value = get.xml(obj)
            % @property xml (rel.py 61-68): serialize the collection as .rels
            %   bytes. <Relationship> elements are emitted in `self.values()`
            %   INSERTION order (NO rId sort -- see ORDERING note). Returns uint8
            %   .rels bytes with an XML declaration.
            rels_elm = mat2doc.opc.oxml.CT_Relationships.new();
            for i = 1:numel(obj.rels_)
                rel = obj.rels_(i);
                rels_elm.add_rel(rel.rId, rel.reltype, rel.target_ref, ...
                    rel.is_external);
            end
            value = rels_elm.xml_file_bytes;   % Python: rels_elm.xml (P1-4 B4 rotation)
        end
    end

    methods (Access = private)
        function rel = get_matching_(obj, reltype, target, is_external)
            % _get_matching (rel.py 70-87): first rel matching `reltype`,
            %   `target`, and `is_external`, or [] (None) if none. For external
            %   the target compared is target_ref (string equality); for internal
            %   it is target_part (handle identity, H5).
            rel = mat2doc.opc.Relationship_.empty(1, 0);   % None
            for i = 1:numel(obj.rels_)
                r = obj.rels_(i);
                if r.reltype ~= reltype
                    continue
                end
                if r.is_external ~= is_external
                    continue
                end
                if r.is_external
                    rel_target = r.target_ref;
                else
                    rel_target = r.target_part;
                end
                if rel_target == target   % string == or handle-identity ==
                    rel = r;
                    return
                end
            end
        end

        function rel = get_rel_of_type_(obj, reltype)
            % _get_rel_of_type (rel.py 89-102): the single rel of `reltype`.
            %   KeyError if none; ValueError if more than one.
            matching = mat2doc.opc.Relationship_.empty(1, 0);
            for i = 1:numel(obj.rels_)
                if obj.rels_(i).reltype == reltype
                    matching(end + 1) = obj.rels_(i); %#ok<AGROW>
                end
            end
            if numel(matching) == 0
                error("mat2doc:KeyError", ...
                    "no relationship of type '%s' in collection", reltype);
            end
            if numel(matching) > 1
                error("mat2doc:ValueError", ...
                    "multiple relationships of type '%s' in collection", reltype);
            end
            rel = matching(1);
        end

        function rId = next_rId_(obj)
            % _next_rId (rel.py 104-111): first unused "rIdN" starting at rId1,
            %   reusing gaps (ascending scan of range(1, len+2)).
            rId = "";   % unreachable per pigeonhole; Python falls through to None
            for n = 1:(numel(obj.rIds_) + 1)
                candidate = "rId" + n;              % "rId%d" % n
                if ~any(obj.rIds_ == candidate)     % candidate not in self
                    rId = candidate;
                    return
                end
            end
        end
    end
end

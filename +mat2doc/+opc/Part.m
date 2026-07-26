classdef Part < handle & matlab.mixin.Heterogeneous
% PART Base class for package parts (and the default, verbatim-blob part class).
%
%   A part has a partname (PackURI), a content_type (string), an optional blob
%   (its serialized bytes) and a back-reference to its package. The base class
%   returns the blob VERBATIM -- the bytes loaded during OpcPackage.open -- so
%   binary parts (images, thumbnails) and any part whose content-type is not
%   registered for an XML subclass round-trip byte-for-byte. XmlPart OVERRIDES
%   `blob` to re-serialize its parsed element (the whitespace-collapse path).
%
%   HETEROGENEOUS ROOT (design.md section 2, VERIFY-1b): Part derives
%   matlab.mixin.Heterogeneous so a mixed Part / XmlPart vector forms one array
%   (OpcPackage.iter_parts must return an OBJECT ARRAY, never a cell -- the
%   already-merged P1-6a PackageWriter.write_parts_ indexes `parts(k)` and reads
%   named members). The identity operators `eq`/`ne` are Sealed here (forwarding
%   to `handle`), exactly as the heterogeneous-Sealed contract requires: MATLAB
%   refuses to dispatch an unsealed method on a heterogeneous array, and
%   iter_parts / iter_rels compare parts by identity (`any(visited == part)`).
%
%   PROPERTY-AS-METHOD (design.md section 2): Python's @property / @lazyproperty
%   members (`blob`, `content_type`, `package`, `partname`, `related_parts`,
%   `rels`) are ported as zero-argument methods so XmlPart can OVERRIDE `blob`
%   (MATLAB cannot override a superclass property accessor, but can override a
%   method). `part.blob`, `part.partname`, `part.rels` etc. still read naturally
%   with no parentheses (MATLAB invokes a zero-arg method on field access), which
%   is exactly how the P1-6a writer consumes them.
%
%   ARG ORDER (docx-vs-pptx divergence): python-docx Part.__init__ is
%   (partname, content_type, blob=None, package=None) -- blob BEFORE package,
%   both defaulting to None. (python-pptx used (partname, content_type, package,
%   blob).) The docx order is preserved here, and in `load` and PartFactory.
%
%   UNDERSCORE ROTATION (design.md section 2): private `_rels` lazyproperty ->
%   the rels_cache_/rels_computed_ pair; `_rel_ref_count` -> rel_ref_count_.
%   NOTE: docx Part.blob is a getter-only @property (NO setter, unlike pptx),
%   so there is deliberately no set_blob here.
%
%   Example:
%       % A base Part returns its blob VERBATIM -- no re-serialization (the
%       % path that keeps thumbnails and the passthrough XML parts byte-stable).
%       p = mat2doc.opc.Part( ...
%           mat2doc.opc.PackURI("/word/theme/theme1.xml"), ...
%           mat2doc.opc.CONTENT_TYPE.OFC_THEME, uint8('<a:theme/>'), []);
%       disp(char(p.blob()))    % "<a:theme/>"  (identical to the stored bytes)
%
%   Ported from python-docx v1.2.0: src/docx/opc/part.py::Part (lines 21-162)

    properties (Access = protected)
        partname_
        content_type_
        blob_
        package_
        rels_cache_                          % _rels lazyproperty cache
        rels_computed_ (1,1) logical = false
    end

    methods
        function obj = Part(partname, content_type, blob, package)
            % part.py 28-39: store partname/content_type/blob/package. blob and
            %   package default to None ([]); XmlPart passes blob = [] and its
            %   element separately.
            arguments
                partname (1,1) mat2doc.opc.PackURI
                content_type (1,1) string
                blob = []          % Python bytes | None (H3); [] is None
                package = []       % Python Package | None (H3)
            end
            obj.partname_ = partname;
            obj.content_type_ = content_type;
            obj.blob_ = blob;
            obj.package_ = package;
        end

        function after_unmarshal(obj) %#ok<MANU>
            % AFTER_UNMARSHAL Entry point for post-unmarshaling processing
            %   (part.py 41-49). Base is intentionally a no-op; subclasses may
            %   override WITHOUT forwarding to super.
        end

        function before_marshal(obj) %#ok<MANU>
            % BEFORE_MARSHAL Entry point for pre-serialization processing
            %   (part.py 51-59). Base is intentionally a no-op; subclasses may
            %   override WITHOUT forwarding to super.
        end

        function b = blob(obj)
            % blob property (part.py 61-68): the bytes loaded during
            %   OpcPackage.open. Python `return self._blob or b""` -- the blob, or
            %   empty bytes when it is None/empty (H4 truthiness: bytes are falsy
            %   only when empty, so isempty() is exact for the byte currency).
            %   Overridden by XmlPart.
            if isempty(obj.blob_)
                b = uint8.empty(1, 0);      % Python b""
            else
                b = obj.blob_;
            end
        end

        function ct = content_type(obj)
            % content_type property (part.py 70-73).
            ct = obj.content_type_;
        end

        function drop_rel(obj, rId)
            % DROP_REL Remove relationship `rId` if its reference count < 2
            %   (part.py 75-82). A ref-count of 0 is an implicit relationship.
            %   LIVE at P2-2 (was a notYetPorted stub at P1-6b): rel_ref_count_
            %   is XmlPart's `element.xpath("//@r:id")` count (0 for the base
            %   Part), and the delete is Relationships.delitem (`del self.rels
            %   [rId]`), both ported in this WP. Real callers land in P5 (Section
            %   footer/header handling) and P5-3b (DocumentPart.drop_header_part);
            %   proven now by a direct unit check.
            arguments
                obj
                rId (1,1) string
            end
            if obj.rel_ref_count_(rId) < 2
                obj.rels().delitem(rId);   % Python: del self.rels[rId]
            end
        end

        function rel = load_rel(obj, reltype, target, rId, is_external)
            % LOAD_REL Add and return a Relationship_ of `reltype` relating
            %   `target` to this part with key `rId` (part.py 88-97). Target mode
            %   is EXTERNAL iff is_external. For use during load, where rId is
            %   well-known.
            arguments
                obj
                reltype (1,1) string
                target
                rId (1,1) string
                is_external (1,1) logical = false   % Python default is_external=False
            end
            rel = obj.rels().add_relationship(reltype, target, rId, is_external);
        end

        function p = package(obj)
            % package property (part.py 99-102).
            p = obj.package_;
        end

        function pn = partname(obj)
            % partname property (part.py 104-108).
            pn = obj.partname_;
        end

        function set_partname(obj, partname)
            % SET_PARTNAME partname setter (part.py 110-115). `partname` is a
            %   zero-arg method here (property-as-method), so the Python
            %   attribute-set `part.partname = PackURI(...)` ports as
            %   `part.set_partname(...)`. Faithful to the isinstance guard: a
            %   non-PackURI raises mat2doc:TypeError.
            if ~isa(partname, "mat2doc.opc.PackURI")
                % Python: type(partname).__name__ (bare class name)
                nm = split(string(class(partname)), ".");
                error("mat2doc:TypeError", ...
                    "partname must be instance of PackURI, got '%s'", nm(end));
            end
            obj.partname_ = partname;
        end

        function part = part_related_by(obj, reltype)
            % PART_RELATED_BY Target part of the single rel of `reltype`
            %   (part.py 117-124). KeyError if none, ValueError if more than one.
            part = obj.rels().part_with_reltype(reltype);
        end

        function rId = relate_to(obj, target, reltype, is_external)
            % RELATE_TO rId of an existing-or-new relationship of `reltype` to
            %   `target` (part.py 126-136). External targets are referred to by
            %   ref string; internal targets by Part handle.
            arguments
                obj
                target
                reltype (1,1) string
                is_external (1,1) logical = false   % Python default is_external=False
            end
            if is_external
                rId = obj.rels().get_or_add_ext_rel(reltype, target);
            else
                rel = obj.rels().get_or_add(reltype, target);
                rId = rel.rId;
            end
        end

        function rp = related_parts(obj)
            % related_parts property (part.py 138-143): {rId -> target Part} map.
            rp = obj.rels().related_parts;
        end

        function r = rels(obj)
            % rels lazyproperty (part.py 145-150): |Relationships| over this
            %   part's base URI. Cached via a logical flag (design.md
            %   @lazyproperty rule; NEVER isempty as the cache sentinel). The
            %   legacy `self._rels` attribute python-docx also assigns (for
            %   python-docx-template) is not part of the used surface.
            if ~obj.rels_computed_
                obj.rels_cache_ = mat2doc.opc.Relationships(obj.partname_.baseURI);
                obj.rels_computed_ = true;
            end
            r = obj.rels_cache_;
        end

        function ref = target_ref(obj, rId)
            % TARGET_REF URL in the target ref of relationship `rId`
            %   (part.py 152-155).
            rel = obj.rels().getitem(rId);   % Python: self.rels[rId]
            ref = rel.target_ref;
        end
    end

    methods (Static)
        function part = load(partname, content_type, blob, package)
            % LOAD Return a base Part loaded from arguments (part.py 84-86): a
            %   straight pass-through construction. XmlPart overrides with a parse
            %   pre-processing step.
            part = mat2doc.opc.Part(partname, content_type, blob, package);
        end
    end

    methods (Sealed)
        function tf = eq(a, b)
            % EQ Identity comparison, Sealed so `==` dispatches on a heterogeneous
            %   Part/XmlPart array (design.md section 2 heterogeneous-Sealed
            %   contract). Forwards to the handle builtin. H5 element identity.
            tf = eq@handle(a, b);
        end

        function tf = ne(a, b)
            % NE Sealed identity `~=`, forwarding to the handle builtin.
            tf = ne@handle(a, b);
        end
    end

    methods (Access = protected)
        function n = rel_ref_count_(obj, rId) %#ok<INUSD>
            % _rel_ref_count (part.py 157-162): only an XML part can contain
            %   references, so this is 0 for the base Part. XmlPart overrides.
            n = 0;
        end
    end
end

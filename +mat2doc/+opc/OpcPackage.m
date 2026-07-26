classdef OpcPackage < handle
% OPCPACKAGE Main API class for the OPC package (open / save / traverse).
%
%   pkg = MAT2DOC.OPC.OPCPACKAGE.OPEN(pkg_file) loads a package from `pkg_file`
%   (a path to a .docx, or a uint8 whole-zip byte vector). pkg.SAVE(out_file)
%   writes it back out. The load builds every part (via PartFactory) complete
%   with its relationships (via the Unmarshaller), and populates the
%   package-level Relationships collection.
%
%   Reference semantics (design.md section 2; H5 element identity): a handle
%   class -- the package and its parts are a live object graph shared by
%   reference, exactly like the Python proxy model. A part reached through two
%   relationships is the SAME handle, and the DFS traversals dedup on that
%   identity (`any(visited == part)`).
%
%   Traversal:
%       iter_parts()  - exactly one reference to each part, depth-first over the
%                       relationship graph (external targets excluded). Returns a
%                       HETEROGENEOUS mat2doc.opc.Part OBJECT ARRAY (Part / XmlPart
%                       mix), never a cell (VERIFY-1b) -- the P1-6a PackageWriter
%                       indexes parts(k) and reads named members.
%       iter_rels()   - exactly one reference to each relationship, depth-first
%                       (external rels included; the visited set gates recursion
%                       only). Returns a Relationship_ handle array.
%   main_document_part is the part related to the package by OFFICE_DOCUMENT.
%
%   SAVE REGENERATES (the M1 byte-identity path): save() does not copy the
%   original zip. It re-emits the package from the live logical model via
%   PackageWriter -- XmlPart bodies are re-serialized (whitespace-collapse),
%   base-Part bodies are kept VERBATIM (see PartFactory).
%
%   ARG/CTOR (docx): OpcPackage() takes NO arguments (Python `cls()`); the reader
%   is opened inside the static `open`. (python-pptx stored pkg_file on the
%   instance; python-docx does not.)
%
%   UNDERSCORE ROTATION (design.md section 2): private `_rels` lazyproperty ->
%   the rels_cache_/rels_computed_ pair; `_core_properties_part` ->
%   core_properties_part_.
%
%   Example:
%       tpl = fullfile(fileparts(fileparts(which('mat2doc.opc.OpcPackage'))), ...
%           "templates", "default.docx");
%       pkg = mat2doc.opc.OpcPackage.open(tpl);
%       disp(numel(pkg.iter_parts()))     % part count
%       pkg.save("out.docx");
%
%   Ported from python-docx v1.2.0: src/docx/opc/package.py::OpcPackage
%   (lines 24-179)

    properties (Access = protected)
        rels_cache_                          % _rels lazyproperty cache
        rels_computed_ (1,1) logical = false
    end

    methods
        function obj = OpcPackage()
            % package.py: OpcPackage has no __init__; `open` builds it with cls().
        end

        function after_unmarshal(obj) %#ok<MANU>
            % AFTER_UNMARSHAL Entry point for post-unmarshaling processing
            %   (package.py 31-38). Base is intentionally a no-op; subclasses (the
            %   P1-8 docx Package, which gathers image parts) may override WITHOUT
            %   forwarding to super.
        end

        function cp = core_properties(obj)
            % core_properties property (package.py 40-44): read/write access to
            %   the Dublin Core properties. Delegates to the core-properties part.
            %   (The CoreProperties object is ported in P2; not exercised at M1.)
            cp = obj.core_properties_part_().core_properties;
        end

        function rels = iter_rels(obj)
            % ITER_RELS Exactly one reference to each relationship, depth-first
            %   over the rels graph (package.py 46-67). EVERY relationship is
            %   yielded (external ones too); the `visited` set of parts gates only
            %   RECURSION, so a part reached by several rels still has each of
            %   those rels yielded once. Generator -> precomputed Relationship_
            %   handle array (H9); `visited` is the shared mutable list, realized
            %   as a nested-function closure variable (fresh per call).
            rels = mat2doc.opc.Relationship_.empty(1, 0);
            visited = mat2doc.opc.Part.empty(1, 0);
            walk_rels(obj);

            function walk_rels(source)
                % Python walk_rels(source, visited): yield each rel, then recurse
                %   into non-external, not-yet-visited target parts.
                vals = source.rels().values();
                for i = 1:numel(vals)
                    rel = vals(i);
                    rels(end + 1) = rel;                    %#ok<AGROW>
                    if rel.is_external   % external items have no rels to descend
                        continue
                    end
                    part = rel.target_part;
                    if ~isempty(visited) && any(visited == part)   % part in visited
                        continue
                    end
                    visited(end + 1) = part;                %#ok<AGROW>
                    walk_rels(part);
                end
            end
        end

        function parts = iter_parts(obj)
            % ITER_PARTS Exactly one reference to each part, depth-first over the
            %   rels graph (package.py 69-87). External targets excluded, each
            %   part yielded once (visited-dedup by handle identity, H5). Returns
            %   a HETEROGENEOUS Part object array in traversal order (VERIFY-1b).
            %   Generator -> precomputed array (H9); `visited` is the shared
            %   mutable list as a closure variable.
            parts = mat2doc.opc.Part.empty(1, 0);
            visited = mat2doc.opc.Part.empty(1, 0);
            walk_parts(obj);

            function walk_parts(source)
                % Python walk_parts(source, visited): for each non-external rel to
                %   a not-yet-visited part, record and yield it, then recurse.
                vals = source.rels().values();
                for i = 1:numel(vals)
                    rel = vals(i);
                    if rel.is_external
                        continue
                    end
                    part = rel.target_part;
                    if ~isempty(visited) && any(visited == part)   % part in visited
                        continue
                    end
                    visited(end + 1) = part;                %#ok<AGROW>
                    parts(end + 1) = part;                  %#ok<AGROW>
                    walk_parts(part);
                end
            end
        end

        function rel = load_rel(obj, reltype, target, rId, is_external)
            % LOAD_REL Add and return a Relationship_ of `reltype` between this
            %   package and `target` with key `rId` (package.py 89-97). For use
            %   during load, where rId is well-known.
            arguments
                obj
                reltype (1,1) string
                target
                rId (1,1) string
                is_external (1,1) logical = false   % Python default is_external=False
            end
            rel = obj.rels().add_relationship(reltype, target, rId, is_external);
        end

        function part = main_document_part(obj)
            % MAIN_DOCUMENT_PART Reference to the main document part
            %   (package.py 99-107): part_related_by(RT.OFFICE_DOCUMENT).
            part = obj.part_related_by( ...
                mat2doc.opc.RELATIONSHIP_TYPE.OFFICE_DOCUMENT);
        end

        function uri = next_partname(obj, template)
            % NEXT_PARTNAME PackURI whose numeric suffix is the next available for
            %   `template`, a printf %d template (package.py 109-121), e.g.
            %   "/word/header%d.xml". Scans n = 1 .. len(partnames)+1.
            arguments
                obj
                template (1,1) string
            end
            % partnames = {part.partname for part in self.iter_parts()}
            parts = obj.iter_parts();
            partnames = string.empty(1, 0);
            for i = 1:numel(parts)
                partnames(end + 1) = string(parts(i).partname()); %#ok<AGROW>
            end
            % Python: for n in range(1, len(partnames) + 2)  ->  n = 1 .. len+1
            for n = 1:(numel(partnames) + 1)
                % candidate_partname = template % n
                candidate = string(sprintf(char(template), n));
                if ~any(partnames == candidate)   % candidate not in partnames
                    uri = mat2doc.opc.PackURI(candidate);
                    return
                end
            end
            % Pigeonhole guarantees a return above (Python implicitly returns None).
            uri = mat2doc.opc.PackURI.empty(1, 0);
        end

        function part = part_related_by(obj, reltype)
            % PART_RELATED_BY Part to which this package has a relationship of
            %   `reltype` (package.py 131-137). KeyError if none, ValueError if
            %   more than one.
            part = obj.rels().part_with_reltype(reltype);
        end

        function p = parts(obj)
            % parts property (package.py 139-142): list(self.iter_parts()) -- a
            %   Part object array (one reference to each part).
            p = obj.iter_parts();
        end

        function rId = relate_to(obj, part, reltype)
            % RELATE_TO rId of an existing-or-new relationship of `reltype` to
            %   `part` (package.py 144-151).
            rel = obj.rels().get_or_add(reltype, part);
            rId = rel.rId;
        end

        function r = rels(obj)
            % rels lazyproperty (package.py 153-157): the package's Relationships
            %   over the package base URI ("/"). Cached via a logical flag
            %   (design.md @lazyproperty rule; NEVER isempty as the sentinel).
            if ~obj.rels_computed_
                obj.rels_cache_ = mat2doc.opc.Relationships( ...
                    mat2doc.opc.PACKAGE_URI().baseURI);
                obj.rels_computed_ = true;
            end
            r = obj.rels_cache_;
        end

        function save(obj, pkg_file)
            % SAVE Write this package to `pkg_file` (package.py 159-166): fire
            %   before_marshal on every part, then PackageWriter.write with the
            %   package rels and parts. (Python evaluates self.parts twice -- once
            %   for the loop, once as the write arg; both DFS walks return the same
            %   handles in the same order, so the array is captured once here.)
            parts = obj.parts();
            for i = 1:numel(parts)
                parts(i).before_marshal();
            end
            mat2doc.opc.PackageWriter.write(pkg_file, obj.rels(), parts);
        end
    end

    methods (Static)
        function package = open(pkg_file)
            % OPEN Return an OpcPackage loaded with the contents of `pkg_file`
            %   (package.py 123-129): read the file, build the (sub)class instance,
            %   unmarshal parts + relationships into it.
            %   NOTE (classmethod cls()): Python `cls()` constructs the calling
            %   subclass; MATLAB static methods carry no cls, so this constructs
            %   OpcPackage directly. The P1-8 docx Package provides its OWN `open`
            %   so parts receive a Package (not OpcPackage) back-reference. The
            %   Python PartFactory class object is passed here as the create handle.
            pkg_reader = mat2doc.opc.PackageReader.from_file(pkg_file);
            package = mat2doc.opc.OpcPackage();
            mat2doc.opc.Unmarshaller.unmarshal( ...
                pkg_reader, package, @mat2doc.opc.PartFactory.create);
        end
    end

    methods (Access = protected)
        function part = core_properties_part_(obj)
            % _core_properties_part (package.py 168-179): the CorePropertiesPart
            %   related to this package, creating a default one if absent (not
            %   common). Not exercised at M1 (default.docx ships core.xml, so
            %   part_related_by succeeds and returns it).
            try
                % Python: cast(CorePropertiesPart, self.part_related_by(...))
                part = obj.part_related_by( ...
                    mat2doc.opc.RELATIONSHIP_TYPE.CORE_PROPERTIES);
            catch e
                if e.identifier == "mat2doc:KeyError"
                    % Python: CorePropertiesPart.default(self); self.relate_to(...)
                    % CorePropertiesPart is ported in P2 (core properties).
                    error("mat2doc:notYetPorted", "%s", ...
                        "mat2doc.opc.parts.CorePropertiesPart.default (owning " + ...
                        "WP: P2 core properties) required by " + ...
                        "mat2doc.opc.OpcPackage._core_properties_part");
                else
                    rethrow(e)   % ValueError (more than one) propagates
                end
            end
        end
    end
end

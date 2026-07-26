classdef PackageReader < handle
% PACKAGEREADER Low-level, read-only access to a serialized OPC package.
%
%   Provides access to the contents of a zip-format OPC package via its serialized
%   parts and package relationships. Constructed from a physical package file (a
%   .docx path or a uint8 whole-zip stream) with the static factory from_file.
%
%   ITERATOR CURRENCY (H9): the Python iter_sparts / iter_srels are generators.
%   MATLAB precomputes them into struct arrays (design.md section 2 -- generators
%   with no mutation-during-iteration become arrays):
%     * iter_sparts() -> 1xN struct with fields
%           partname (PackURI), content_type (string), reltype (string), blob (uint8)
%     * iter_srels()  -> 1xM struct with fields
%           source_uri (PackURI), srel (SerializedRelationship_)
%   HAND-OFF (P1-6b/P1-8 unmarshaller): consume `for s = reader.iter_sparts()` /
%   `for s = reader.iter_srels()` and read the named fields (mirrors Python tuple
%   unpacking `for partname, content_type, reltype, blob in ...`).
%
%   Ported from python-docx v1.2.0: src/docx/opc/pkgreader.py::PackageReader
%   (lines 10-83)

    properties (Access = private)
        pkg_srels_   % SerializedRelationships_ (package-level rels)
        sparts_ = mat2doc.opc.SerializedPart_.empty(1, 0)  % 1xN SerializedPart_
    end

    methods
        function obj = PackageReader(content_types, pkg_srels, sparts) %#ok<INUSA>
            % pkgreader.py 14-17: store pkg_srels and sparts. NOTE the Python
            %   __init__ receives content_types but does NOT store it (it is used
            %   only during from_file); ported faithfully -- content_types is
            %   accepted and ignored.
            obj.pkg_srels_ = pkg_srels;
            obj.sparts_ = sparts;
        end

        function tuples = iter_sparts(obj)
            % ITER_SPARTS 4-tuple (partname, content_type, reltype, blob) for each
            %   serialized part (pkgreader.py 29-33). Precomputed struct array (H9).
            tuples = struct('partname', {}, 'content_type', {}, ...
                'reltype', {}, 'blob', {});
            for k = 1:numel(obj.sparts_)
                s = obj.sparts_(k);
                t = struct();
                t.partname = s.partname;
                t.content_type = s.content_type;
                t.reltype = s.reltype;
                t.blob = s.blob;
                tuples(end + 1) = t; %#ok<AGROW>
            end
        end

        function tuples = iter_srels(obj)
            % ITER_SRELS 2-tuple (source_uri, srel) for each relationship in the
            %   package (pkgreader.py 35-42): first the package rels (source =
            %   PACKAGE_URI), then each part's rels (source = part.partname).
            %   Precomputed struct array (H9).
            tuples = struct('source_uri', {}, 'srel', {});
            pkg_uri = mat2doc.opc.PACKAGE_URI();
            pkg_rel_arr = obj.pkg_srels_.to_array();
            for i = 1:numel(pkg_rel_arr)
                t = struct();
                t.source_uri = pkg_uri;
                t.srel = pkg_rel_arr(i);
                tuples(end + 1) = t; %#ok<AGROW>
            end
            for k = 1:numel(obj.sparts_)
                spart = obj.sparts_(k);
                part_rel_arr = spart.srels.to_array();
                for i = 1:numel(part_rel_arr)
                    t = struct();
                    t.source_uri = spart.partname;
                    t.srel = part_rel_arr(i);
                    tuples(end + 1) = t; %#ok<AGROW>
                end
            end
        end
    end

    methods (Static)
        function reader = from_file(pkg_file)
            % FROM_FILE Build a PackageReader loaded with the contents of pkg_file
            %   (pkgreader.py 19-27).
            phys_reader = mat2doc.opc.PhysPkgReader.factory(pkg_file);
            content_types = mat2doc.opc.ContentTypeMap_.from_xml( ...
                phys_reader.content_types_xml);
            pkg_srels = mat2doc.opc.PackageReader.srels_for_( ...
                phys_reader, mat2doc.opc.PACKAGE_URI());
            sparts = mat2doc.opc.PackageReader.load_serialized_parts_( ...
                phys_reader, pkg_srels, content_types);
            phys_reader.close();
            reader = mat2doc.opc.PackageReader(content_types, pkg_srels, sparts);
        end
    end

    methods (Static, Access = private)
        function sparts = load_serialized_parts_(phys_reader, pkg_srels, content_types)
            % _load_serialized_parts (pkgreader.py 44-55): walk the relationship
            %   graph from pkg_srels, building a SerializedPart_ per reached part.
            sparts = mat2doc.opc.SerializedPart_.empty(1, 0);
            part_walker = mat2doc.opc.PackageReader.walk_phys_parts_( ...
                phys_reader, pkg_srels, strings(1, 0));
            for k = 1:numel(part_walker)
                w = part_walker(k);
                content_type = content_types.getitem(w.partname);
                spart = mat2doc.opc.SerializedPart_(w.partname, content_type, ...
                    w.reltype, w.blob, w.srels);
                sparts(end + 1) = spart; %#ok<AGROW>
            end
        end

        function srels = srels_for_(phys_reader, source_uri)
            % _srels_for (pkgreader.py 57-62): SerializedRelationships_ for the
            %   source identified by source_uri.
            rels_xml = phys_reader.rels_xml_for(source_uri);
            srels = mat2doc.opc.SerializedRelationships_.load_from_xml( ...
                source_uri.baseURI, rels_xml);
        end

        function [results, visited] = walk_phys_parts_(phys_reader, srels, visited)
            % _walk_phys_parts (pkgreader.py 64-83): DFS the relationship graph
            %   rooted at srels, yielding a 4-tuple (partname, blob, reltype,
            %   srels) per part. Generator -> precomputed struct array (H9). The
            %   `visited` partname list is THREADED through the recursion (returned
            %   and re-passed) to replicate Python's SHARED mutable list -- a part
            %   reached by two rel paths is emitted once, globally across the walk.
            results = struct('partname', {}, 'blob', {}, 'reltype', {}, 'srels', {});
            srel_arr = srels.to_array();
            for j = 1:numel(srel_arr)
                srel = srel_arr(j);
                if srel.is_external
                    continue   % Python: if srel.is_external: continue
                end
                partname = srel.target_partname;   % PackURI
                % Python: `if partname in visited_partnames: continue` -- membership
                %   by PackURI __eq__ (string identity); visited stored as strings.
                if any(visited == string(partname))
                    continue
                end
                visited(end + 1) = string(partname); %#ok<AGROW>
                reltype = srel.reltype;
                part_srels = mat2doc.opc.PackageReader.srels_for_( ...
                    phys_reader, partname);
                blob = phys_reader.blob_for(partname);
                r = struct();
                r.partname = partname;
                r.blob = blob;
                r.reltype = reltype;
                r.srels = part_srels;
                results(end + 1) = r; %#ok<AGROW>
                % Recurse with the CURRENT (mutated) visited list; append its
                % results after this part -- exactly Python's yield-then-recurse.
                [child, visited] = mat2doc.opc.PackageReader.walk_phys_parts_( ...
                    phys_reader, part_srels, visited);
                results = [results, child]; %#ok<AGROW>
            end
        end
    end
end

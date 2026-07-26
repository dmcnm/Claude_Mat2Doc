classdef Unmarshaller
% UNMARSHALLER Builds a package's part graph and relationships from a reader.
%
%   Hosts static methods for unmarshalling a package from a PackageReader. The
%   entry point Unmarshaller.UNMARSHAL(pkg_reader, package, part_factory)
%   constructs every part (delegating to `part_factory`), realizes the
%   relationship graph (resolving each serialized target to the actual Part), then
%   fires after_unmarshal on every part and on the package.
%
%   CONSUMES P1-6a READER CURRENCY (VERIFY-1 hand-off): pkg_reader.iter_sparts()
%   is a 1xN struct array with fields (partname, content_type, reltype, blob) and
%   iter_srels() a 1xM struct array with fields (source_uri, srel) -- the
%   precomputed-generator shape P1-6a froze (H9). The Python tuple-unpacking
%   `for partname, content_type, reltype, blob in pkg_reader.iter_sparts()` and
%   `for source_uri, srel in pkg_reader.iter_srels()` port to reads of those
%   named struct fields.
%
%   PARTS MAP CURRENCY: Python `parts` is a dict {partname -> Part} iterated in
%   INSERTION order (`for part in parts.values()`). containers.Map sorts keys
%   (H11, forbidden), so `parts` is realized as an insertion-ordered struct with
%   parallel fields:
%       parts.keys - 1xK string, the partname text keys in insertion order
%       parts.vals - 1xK heterogeneous mat2doc.opc.Part array (Part / XmlPart mix)
%   Lookups `parts[uri]` are find(parts.keys == string(uri)); `parts.values()` is
%   parts.vals in order. partnames are unique, so a key is appended once.
%
%   Ported from python-docx v1.2.0: src/docx/opc/package.py::Unmarshaller
%   (lines 182-219)

    methods (Static)
        function unmarshal(pkg_reader, package, part_factory)
            % UNMARSHAL Construct the graph of parts and realized relationships
            %   from `pkg_reader`, delegating part construction to `part_factory`
            %   (package.py 185-196). `part_factory` is a function handle
            %   (@mat2doc.opc.PartFactory.create) -- the MATLAB stand-in for the
            %   Python PartFactory class object called as part_factory(...).
            parts = mat2doc.opc.Unmarshaller.unmarshal_parts_( ...
                pkg_reader, package, part_factory);
            mat2doc.opc.Unmarshaller.unmarshal_relationships_( ...
                pkg_reader, package, parts);
            % for part in parts.values(): part.after_unmarshal()
            for k = 1:numel(parts.vals)
                parts.vals(k).after_unmarshal();
            end
            package.after_unmarshal();
        end
    end

    methods (Static, Access = private)
        function parts = unmarshal_parts_(pkg_reader, package, part_factory)
            % _unmarshal_parts (package.py 198-209): a {partname -> Part} map,
            %   each part constructed via `part_factory`. Insertion-ordered
            %   parallel-array struct (see the PARTS MAP CURRENCY note).
            parts = struct("keys", string.empty(1, 0), ...
                "vals", mat2doc.opc.Part.empty(1, 0));
            sparts = pkg_reader.iter_sparts();
            for k = 1:numel(sparts)
                s = sparts(k);
                % parts[partname] = part_factory(partname, content_type, reltype,
                %                                 blob, package)
                part = part_factory(s.partname, s.content_type, s.reltype, ...
                    s.blob, package);
                key = string(s.partname);
                idx = find(parts.keys == key, 1);
                if isempty(idx)
                    parts.keys(end + 1) = key;         %#ok<AGROW>
                    parts.vals(end + 1) = part;        %#ok<AGROW>
                else
                    parts.vals(idx) = part;   % dict setitem: replace in place
                end
            end
        end

        function unmarshal_relationships_(pkg_reader, package, parts)
            % _unmarshal_relationships (package.py 211-219): add each serialized
            %   relationship to its source (the package for source_uri "/",
            %   otherwise the source part), with target resolved to the actual
            %   target Part (or the ref string for external rels).
            srels = pkg_reader.iter_srels();
            for k = 1:numel(srels)
                source_uri = srels(k).source_uri;   % PackURI
                srel = srels(k).srel;               % SerializedRelationship_
                % source = package if source_uri == "/" else parts[source_uri]
                if source_uri == "/"   % PackURI eq "/" (string identity)
                    source = package;
                else
                    idx = find(parts.keys == string(source_uri), 1);
                    source = parts.vals(idx);
                end
                % target = srel.target_ref if srel.is_external
                %          else parts[srel.target_partname]
                if srel.is_external
                    target = srel.target_ref;
                else
                    tidx = find(parts.keys == string(srel.target_partname), 1);
                    target = parts.vals(tidx);
                end
                source.load_rel(srel.reltype, target, srel.rId, srel.is_external);
            end
        end
    end
end

classdef IfdEntries_
% IFDENTRIES_ Image File Directory for a TIFF image, with mapping semantics.
%
%   Provides tag-code -> value lookup (Python `_IfdEntries`, a dict-like). The
%   Python `{e.tag: e.value for e in ...}` comprehension keeps the LAST value for
%   any duplicate tag; from_stream replicates that by overwriting on rebuild.
%   Lookup is by key only (never ordered iteration), so storage order is inert
%   (H11). Values are heterogeneous (short/long int as double, ASCII as string,
%   rational as double), so they are held in a cell array.
%
%   MATLAB has no operator overloading for `in`/`[]`, so the Python dunder methods
%   map to explicit methods: __contains__ -> contains_, __getitem__ -> getitem_,
%   get -> get.
%
%   UNDERSCORE ROTATION (design.md section 2): python-docx _IfdEntries ->
%   IfdEntries_.
%
%   Ported from python-docx v1.2.0: src/docx/image/tiff.py::_IfdEntries
%   (lines 118-145)

    properties (Access = private)
        tags_     % 1xK double: entry tag codes (deduplicated, last-wins)
        values_   % 1xK cell:   corresponding values
    end

    methods
        function obj = IfdEntries_(tags, values)
            % Ported from _IfdEntries.__init__ (tiff.py 122-124). `entries` (a
            %   dict) is represented here as the deduplicated (tags, values) pair.
            obj.tags_ = tags;
            obj.values_ = values;
        end

        function tf = contains_(obj, key)
            % __contains__ (tiff.py 126-128): `key in ifd_entries`.
            tf = any(obj.tags_ == key);
        end

        function v = getitem_(obj, key)
            % __getitem__ (tiff.py 130-132): `ifd_entries[key]`. KeyError when
            %   absent (matches Python dict indexing).
            idx = find(obj.tags_ == key, 1);
            if isempty(idx)
                error("mat2doc:KeyError", "%s", string(key));
            end
            v = obj.values_{idx};
        end

        function v = get(obj, tag_code, default)
            % get (tiff.py 142-145): value for `tag_code`, or `default` (None ->
            %   [] when omitted, H3) if no matching tag is present.
            arguments
                obj
                tag_code
                default = []     % None (H3)
            end
            idx = find(obj.tags_ == tag_code, 1);
            if isempty(idx)
                v = default;
            else
                v = obj.values_{idx};
            end
        end
    end

    methods (Static)
        function obj = from_stream(stream, offset)
            % from_stream (tiff.py 134-140): parse the IFD at `offset`, building
            %   {tag: value} with last-wins duplicate resolution (the dict
            %   comprehension semantics).
            ifd_parser = mat2doc.image.IfdParser_(stream, offset);
            entries = ifd_parser.iter_entries();      % 1xN cell of IfdEntry_
            tags = zeros(1, 0);
            values = cell(1, 0);
            for i = 1:numel(entries)
                e = entries{i};
                tag = e.tag;
                j = find(tags == tag, 1);
                if isempty(j)
                    tags(end + 1) = tag;              %#ok<AGROW>
                    values{end + 1} = e.value;        %#ok<AGROW>
                else
                    values{j} = e.value;              % last-wins overwrite
                end
            end
            obj = mat2doc.image.IfdEntries_(tags, values);
        end
    end
end

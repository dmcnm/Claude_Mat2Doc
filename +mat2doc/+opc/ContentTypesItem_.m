classdef ContentTypesItem_ < handle
% CONTENTTYPESITEM_ Composes [Content_Types].xml from a list of parts.
%
%   Service class; the single public entry is the static from_parts(parts), whose
%   result exposes `blob` (the serialized [Content_Types].xml bytes). A handle
%   class (holds the CaseInsensitiveDict of defaults and the dict of overrides
%   that _add_content_type mutates).
%
%   DEFAULT vs OVERRIDE (pkgwriter.py 92-99, H11/H15): a part is classified as a
%   <Default> iff (ext.lower(), content_type) is a ROW of default_content_types
%   (a duplicate-key PAIR-LIST, so BOTH fields must match -- see
%   default_content_types.m); otherwise it is an <Override>. `_defaults` is a
%   CaseInsensitiveDict keyed by (lowercased) extension; `_overrides` is a plain
%   dict keyed by the raw (case-preserved) partname. Element order in the emitted
%   part is defaults-sorted-by-extension THEN overrides-sorted-by-partname
%   (pkgwriter.py 110-115): the SORT is applied at emit, not at insertion.
%
%   UNDERSCORE ROTATION (design.md section 2): `_ContentTypesItem` ->
%   ContentTypesItem_; `_add_content_type` -> add_content_type_; `_element` ->
%   element_ (private method).
%
%   Ported from python-docx v1.2.0: src/docx/opc/pkgwriter.py::_ContentTypesItem
%   (lines 62-115)

    properties (Access = private)
        defaults_    % CaseInsensitiveDict (extension -> content_type)
        overrides_   % dictionary(string partname -> string content_type)
    end

    properties (Dependent, SetAccess = private)
        blob
    end

    methods
        function obj = ContentTypesItem_()
            % pkgwriter.py 70-72: empty CaseInsensitiveDict + empty dict.
            obj.defaults_ = mat2doc.opc.CaseInsensitiveDict();
            obj.overrides_ = dictionary(string.empty(0, 1), string.empty(0, 1));
        end

        function value = get.blob(obj)
            % @property blob (pkgwriter.py 74-78): serialize_part_xml(self._element)
            %   -- the [Content_Types].xml file bytes (uint8, declaration included).
            value = mat2doc.opc.oxml.serialize_part_xml(obj.element_());
        end

        function add_content_type_(obj, partname, content_type)
            % _add_content_type (pkgwriter.py 92-99): default if the (ext.lower(),
            %   content_type) pair is a row of default_content_types, else override.
            ext = partname.ext;   % partname is a PackURI
            pairs = mat2doc.opc.default_content_types();
            % Python: `if (ext.lower(), content_type) in default_content_types`
            %   -- a ROW membership test over the Nx2 pair-list (both columns).
            is_default = any(pairs(:, 1) == lower(ext) & pairs(:, 2) == content_type);
            if is_default
                obj.defaults_.set(ext, content_type);   % CID folds key to lowercase
            else
                obj.overrides_(string(partname)) = content_type;
            end
        end
    end

    methods (Static)
        function cti = from_parts(parts)
            % FROM_PARTS Build the content-types item for `parts` (pkgwriter.py
            %   80-90). Pre-seeds the `rels` and `xml` defaults, then classifies
            %   each part.
            cti = mat2doc.opc.ContentTypesItem_();
            cti.defaults_.set("rels", mat2doc.opc.CONTENT_TYPE.OPC_RELATIONSHIPS);
            cti.defaults_.set("xml", mat2doc.opc.CONTENT_TYPE.XML);
            for k = 1:numel(parts)
                cti.add_content_type_(parts(k).partname, parts(k).content_type);
            end
        end
    end

    methods (Access = private)
        function types_elm = element_(obj)
            % _element (pkgwriter.py 101-115): a fresh <Types> element with
            %   <Default> children sorted by extension THEN <Override> children
            %   sorted by partname. The sort keys mirror Python `sorted()` exactly
            %   (ascending string order; ASCII partnames/exts -> identical to
            %   Python code-point order, H11).
            types_elm = mat2doc.opc.oxml.CT_Types.new();
            ext_keys = sort(obj.defaults_.keys());
            for i = 1:numel(ext_keys)
                types_elm.add_default(ext_keys(i), obj.defaults_.get(ext_keys(i)));
            end
            part_keys = sort(reshape(keys(obj.overrides_), 1, []));
            for i = 1:numel(part_keys)
                types_elm.add_override(part_keys(i), obj.overrides_(part_keys(i)));
            end
        end
    end
end

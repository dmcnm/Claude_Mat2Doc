classdef ContentTypeMap_ < handle
% CONTENTTYPEMAP_ Dictionary of content type by partname, from [Content_Types].xml.
%
%   Value type providing dictionary semantics for looking up a content type by
%   part name, e.g. `content_type = cti.getitem(PackURI('/word/document.xml'))`.
%   A handle class (holds the two CaseInsensitiveDict handles the Python
%   __init__ creates and that _add_default/_add_override mutate).
%
%   PRECEDENCE (pkgreader.py 95-105, H15): Override (exact partname) beats Default
%   (by extension). BOTH `_overrides` and `_defaults` are CaseInsensitiveDict in
%   python-docx's reader, so the override partname match AND the default extension
%   match are case-INSENSITIVE (contrast the writer side, where partnames are the
%   raw dict keys). A miss on both raises KeyError (mat2doc:KeyError).
%
%   UNDERSCORE ROTATION (design.md section 2): `_ContentTypeMap` -> ContentTypeMap_;
%   `__getitem__` -> getitem; `_add_default`/`_add_override` ->
%   add_default_/add_override_.
%
%   Ported from python-docx v1.2.0: src/docx/opc/pkgreader.py::_ContentTypeMap
%   (lines 86-127)

    properties (Access = private)
        overrides_   % CaseInsensitiveDict (partname -> content_type)
        defaults_    % CaseInsensitiveDict (extension -> content_type)
    end

    methods
        function obj = ContentTypeMap_()
            % pkgreader.py 90-93: two CaseInsensitiveDicts.
            obj.overrides_ = mat2doc.opc.CaseInsensitiveDict();
            obj.defaults_ = mat2doc.opc.CaseInsensitiveDict();
        end

        function content_type = getitem(obj, partname)
            % __getitem__ (pkgreader.py 95-105): content type for partname.
            %   Override (exact partname) beats Default (by extension); miss ->
            %   KeyError. Key MUST be a PackURI (Python isinstance guard).
            if ~isa(partname, "mat2doc.opc.PackURI")
                % Python: KeyError("_ContentTypeMap key must be <type 'PackURI'>,
                %   got %s" % type(partname)). MATLAB reports class(partname) in
                %   place of Python's type() repr (diagnostic-only, unreachable in
                %   the library flow where callers always pass a PackURI).
                error("mat2doc:KeyError", ...
                    "_ContentTypeMap key must be <type 'PackURI'>, got %s", ...
                    class(partname));
            end
            if obj.overrides_.isKey(partname)
                content_type = obj.overrides_.get(partname);
                return
            end
            if obj.defaults_.isKey(partname.ext)
                content_type = obj.defaults_.get(partname.ext);
                return
            end
            error("mat2doc:KeyError", ...
                "no content type for partname '%s' in [Content_Types].xml", ...
                string(partname));
        end

        function add_default_(obj, extension, content_type)
            % _add_default (pkgreader.py 119-122): map extension -> content_type
            %   (CaseInsensitiveDict folds the extension key to lowercase).
            obj.defaults_.set(extension, content_type);
        end

        function add_override_(obj, partname, content_type)
            % _add_override (pkgreader.py 124-127): map partname -> content_type.
            obj.overrides_.set(partname, content_type);
        end
    end

    methods (Static)
        function ct_map = from_xml(content_types_xml)
            % FROM_XML Build a ContentTypeMap_ from [Content_Types].xml bytes
            %   (pkgreader.py 107-117). Populates overrides then defaults in
            %   document order (order is immaterial -- lookups are keyed).
            types_elm = mat2doc.opc.oxml.parse_xml(content_types_xml);
            ct_map = mat2doc.opc.ContentTypeMap_();
            overrides = types_elm.overrides;   % CT_Override array (document order)
            for k = 1:numel(overrides)
                ct_map.add_override_(overrides(k).partname, ...
                    overrides(k).content_type);
            end
            defaults = types_elm.defaults;     % CT_Default array (document order)
            for k = 1:numel(defaults)
                ct_map.add_default_(defaults(k).extension, ...
                    defaults(k).content_type);
            end
        end
    end
end

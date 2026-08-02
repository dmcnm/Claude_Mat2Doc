classdef Package < mat2doc.opc.OpcPackage
% PACKAGE Customizations specific to a WordprocessingML (WML) package.
%
%   Extends the OPC-level mat2doc.opc.OpcPackage with the WordprocessingML-level
%   surface: image-part gathering after load, and the image-part helpers. The
%   api-level factory mat2doc.Document opens documents through Package.open so
%   that every part's package back-reference is a Package and
%   document_part.core_properties resolves.
%
%   REFERENCE SEMANTICS (design.md section 2): a handle class (via OpcPackage <
%   handle) -- one live object graph shared by reference, exactly like the Python
%   proxy model.
%
%   OWN STATIC `open` (the inherited-static trap): MATLAB static methods are not
%   polymorphic in `cls`, so the inherited OpcPackage.open (which constructs a
%   base OpcPackage) is OVERRIDDEN here to construct a mat2doc.package.Package,
%   faithfully realizing Python `Package.open` == inherited OpcPackage.open with
%   cls=Package (opc/package.py 123-129 with cls bound to docx.package.Package).
%   Without this override, parts would receive a base OpcPackage back-reference.
%
%   IMAGE GATHERING LIVE (benign at M1): after_unmarshal -> _gather_image_parts
%   runs on every open. default.docx has NO internal RT.IMAGE relationship (its
%   thumbnail is a THUMBNAIL reltype), so the reltype guard short-circuits before
%   any image part is appended -- the walk is a no-op and the empty ImageParts is
%   never grown, so the empty-doc save is byte-unchanged (M1 17/17 re-proven).
%
%   P7-4 UN-STUB (image tier now ported): get_or_add_image_part and the
%   image_parts lazyproperty are LIVE -- image_parts returns a
%   mat2doc.package.ImageParts collection; get_or_add_image_part delegates to it
%   (SHA1 dedupe). _gather_image_parts appends every existing internal IMAGE
%   target part (an ImagePart) into that collection.
%
%   UNDERSCORE ROTATION (design.md section 2): private `_gather_image_parts` ->
%   gather_image_parts_; the image_parts @lazyproperty cache ->
%   image_parts_cache_/image_parts_computed_.
%
%   Example:
%       tpl = fullfile(fileparts(fileparts(which( ...
%           "mat2doc.package.Package"))), "templates", "default.docx");
%       pkg = mat2doc.package.Package.open(tpl);
%       disp(class(pkg))                % "mat2doc.package.Package"
%       disp(class(pkg.main_document_part()))  % "mat2doc.parts.DocumentPart"
%       pkg.save("out.docx");
%
%   Ported from python-docx v1.2.0: src/docx/package.py::Package (lines 15-47)

    properties (Access = private)
        image_parts_cache_                   % image_parts @lazyproperty cache
        image_parts_computed_ (1,1) logical = false
    end

    methods
        function obj = Package()
            % package.py: Package(OpcPackage) declares no __init__; it is built
            %   by the inherited cls() in `open`. MATLAB does not inherit
            %   constructors, so this pass-through forwards to the (no-arg)
            %   OpcPackage constructor.
            obj@mat2doc.opc.OpcPackage();
        end

        function after_unmarshal(obj)
            % AFTER_UNMARSHAL OVERRIDE (package.py 18-23): post-load processing --
            %   gather the package's image parts. Overrides the base no-op WITHOUT
            %   forwarding to super (matching Python, which does not call
            %   super().after_unmarshal). LIVE at M1; a no-op on default.docx.
            obj.gather_image_parts_();
        end

        function image_part = get_or_add_image_part(obj, image_descriptor)
            % GET_OR_ADD_IMAGE_PART (package.py 25-31): the ImagePart for
            %   `image_descriptor`, newly created if a matching one is not already
            %   present. Python:
            %     return self.image_parts.get_or_add_image_part(image_descriptor)
            image_part = obj.image_parts().get_or_add_image_part(image_descriptor);
        end

        function ip = image_parts(obj)
            % image_parts @lazyproperty (package.py 33-36): the ImageParts
            %   collection for this package. Python: `return ImageParts()`.
            %   Cached via a logical flag (design.md @lazyproperty rule; NEVER
            %   isempty as the sentinel) so the collection is a single shared
            %   handle across gather / get_or_add.
            if ~obj.image_parts_computed_
                obj.image_parts_cache_ = mat2doc.package.ImageParts();
                obj.image_parts_computed_ = true;
            end
            ip = obj.image_parts_cache_;
        end
    end

    methods (Static)
        function package = open(pkg_file)
            % OPEN Return a Package loaded with the contents of `pkg_file` -- the
            %   OWN static override of OpcPackage.open constructing a
            %   mat2doc.package.Package (opc/package.py 123-129, cls=Package):
            %   read the file, build the Package instance, unmarshal parts +
            %   relationships into it. Passes the PartFactory create handle, as the
            %   base does (docx has a single PartFactory registered in
            %   docx/__init__.py; the same handle serves the subclass).
            pkg_reader = mat2doc.opc.PackageReader.from_file(pkg_file);
            package = mat2doc.package.Package();
            mat2doc.opc.Unmarshaller.unmarshal( ...
                pkg_reader, package, @mat2doc.opc.PartFactory.create);
        end
    end

    methods (Access = protected)
        function gather_image_parts_(obj)
            % _GATHER_IMAGE_PARTS (package.py 38-47): load the image-part
            %   collection with every internal image part in the package. Iterates
            %   iter_rels (H9 -> precomputed Relationship_ array), skipping
            %   external rels and non-IMAGE reltypes. On default.docx there are no
            %   internal IMAGE rels, so the collection stays empty and this is a
            %   no-op (the M1 benign path). LIVE at P7-4: for a document opened WITH
            %   images, each IMAGE target part (an ImagePart) is appended once.
            rels = obj.iter_rels();
            RT = mat2doc.opc.RELATIONSHIP_TYPE;
            for i = 1:numel(rels)
                rel = rels(i);
                if rel.is_external          % Python: if rel.is_external: continue
                    continue
                end
                if rel.reltype ~= RT.IMAGE  % Python: if rel.reltype != RT.IMAGE
                    continue
                end
                % -- reached only for an internal IMAGE rel (none at M1) --
                ip = obj.image_parts();     % LIVE (P7-4): shared ImageParts
                if ip.contains(rel.target_part)   % Python: in self.image_parts
                    continue
                end
                ip.append(rel.target_part);
            end
        end
    end
end

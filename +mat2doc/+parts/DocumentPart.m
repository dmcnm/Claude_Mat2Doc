classdef DocumentPart < mat2doc.opc.XmlPart
% DOCUMENTPART Main document part of a WordprocessingML (WML) package (.docx).
%
%   Acts as broker to other parts such as image, core properties, and style
%   parts. It also acts as a convenient delegate when a mid-document object needs
%   a service involving a remote ancestor.
%
%   THIN M1 SLICE: this is the walking-skeleton entry part. It subclasses
%   mat2doc.opc.XmlPart -- so it PARSES on load and RE-SERIALIZES on save through
%   serialize_part_xml (byte-matched to lxml) -- and exposes only the members the
%   open->save path traverses: `document`, `save`, and `core_properties`. Every
%   feature accessor (styles/settings/inline_shapes/numbering_part/comments,
%   header & footer parts, style lookup) is a mat2doc:notYetPorted stub; NONE is
%   on the open/save path.
%
%   VERIFY-M1-DOCPART-BASE: in python-docx DocumentPart extends StoryPart <
%   BaseStoryPart < XmlPart. The StoryPart / BaseStoryPart tier is not ported at
%   M1 (it carries the story-authoring surface -- new_run, paragraphs, tables,
%   embed helpers -- all feature-only). P2-2 inserts BaseStoryPart/StoryPart
%   ABOVE this class; because those bases add methods only (blob/load/element
%   inherit unchanged from XmlPart), the reparented DocumentPart emits identical
%   bytes -- the superclass insertion is byte-neutral. Recorded for the Auditor.
%
%   BYTE-IDENTITY (M1 round-trip): DocumentPart inherits XmlPart.blob unchanged
%   (parse + serialize_part_xml) and does not mutate the parsed tree on a plain
%   open->save, so the PartFactory flip WML_DOCUMENT_MAIN -> DocumentPart (from
%   the base XmlPart stand-in) produces bytes IDENTICAL to the previous
%   base-XmlPart dispatch -- only the reloaded part's TYPE changes.
%
%   OWN CONSTRUCTOR + OWN STATIC `load` (the inherited-static trap): MATLAB does
%   not inherit constructors or dispatch inherited static methods to the
%   subclass, so this class declares its own pass-through constructor and its own
%   `load` (the PartFactory entry point) constructing a DocumentPart -- the
%   faithful realization of Python's inherited-but-cls-bound XmlPart.load
%   (opc/part.py 229-232, where cls is DocumentPart). Without its own `load`, the
%   factory would silently build a base XmlPart and `.document` would be missing.
%
%   ARG ORDER (docx): DocumentPart(partname, content_type, element, package) --
%   element third, package last (XmlPart.__init__ order, opc/part.py 214-218).
%
%   `document` is a PLAIN @property (parts/document.py 58-61) -- it constructs a
%   FRESH mat2doc.document.Document on EACH access (NOT cached / not a
%   lazyproperty), so two reads are distinct proxies, exactly as in python-docx.
%
%   Example:
%       CT  = mat2doc.opc.CONTENT_TYPE;
%       xml = "<w:document xmlns:w='" + ...
%           "http://schemas.openxmlformats.org/wordprocessingml/2006/main'>" + ...
%           "<w:body/></w:document>";
%       dp  = mat2doc.parts.DocumentPart.load( ...
%           mat2doc.opc.PackURI("/word/document.xml"), ...
%           CT.WML_DOCUMENT_MAIN, uint8(unicode2native(xml, "UTF-8")), []);
%       d   = dp.document();            % a mat2doc.document.Document
%
%   Ported from python-docx v1.2.0: src/docx/parts/document.py::DocumentPart
%   (the M1 slice: __init__/blob/element via XmlPart; core_properties 52-56,
%   document 58-61, save 111-114. P2-2 un-stubs the remaining members.)

    methods
        function obj = DocumentPart(partname, content_type, element, package)
            % Pass-through to the XmlPart constructor (design.md CT_* / part
            %   constructor contract): forward ALL args, no re-validation. ARG
            %   ORDER element-before-package matches docx XmlPart.__init__.
            obj@mat2doc.opc.XmlPart(partname, content_type, element, package);
        end

        function cp = core_properties(obj)
            % CORE_PROPERTIES (parts/document.py 52-56): a CoreProperties object
            %   with read/write access to the core properties of this document.
            %   Python: return self.package.core_properties. LIVE at M1 (P1-7
            %   made the core-properties path real).
            cp = obj.package().core_properties();
        end

        function d = document(obj)
            % DOCUMENT (parts/document.py 58-61, plain @property): a Document
            %   object providing access to the content of this document. Python:
            %   return Document(self._element, self). Constructs a FRESH proxy on
            %   each access (NOT cached). `element()` is the inherited XmlPart
            %   accessor of the parsed w:document root.
            d = mat2doc.document.Document(obj.element(), obj);
        end

        function save(obj, path_or_stream)
            % SAVE (parts/document.py 111-114): save this document to
            %   `path_or_stream`. Python: self.package.save(path_or_stream).
            %   `package()` is a mat2doc.package.Package (Package.open set the
            %   back-reference); its save delegates to the P1-6a PackageWriter.
            obj.package().save(path_or_stream);
        end

        % ------------------------------------------------------------------
        % Feature stubs (mat2doc:notYetPorted) -- P2-2 un-stubs these. NONE is
        % reached on the open->save path.
        % ------------------------------------------------------------------

        function [footer_part, rId] = add_footer_part(obj) %#ok<MANU,STOUT>
            % ADD_FOOTER_PART STUB (parts/document.py 35-39). P2 header/footer tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.parts.FooterPart (owning WP: P2 header/footer tier) " + ...
                "required by mat2doc.parts.DocumentPart.add_footer_part");
        end

        function [header_part, rId] = add_header_part(obj) %#ok<MANU,STOUT>
            % ADD_HEADER_PART STUB (parts/document.py 41-45). P2 header/footer tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.parts.HeaderPart (owning WP: P2 header/footer tier) " + ...
                "required by mat2doc.parts.DocumentPart.add_header_part");
        end

        function c = comments(obj) %#ok<MANU,STOUT>
            % COMMENTS STUB (parts/document.py 47-50). P2 comments tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.parts.CommentsPart / mat2doc.comments.Comments (owning " + ...
                "WP: P2 comments tier) required by mat2doc.parts.DocumentPart.comments");
        end

        function drop_header_part(obj, rId) %#ok<INUSD,MANU>
            % DROP_HEADER_PART STUB (parts/document.py 63-65). P2 header/footer tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.parts.DocumentPart.drop_header_part (owning WP: P2 " + ...
                "header/footer tier) is not yet ported");
        end

        function fp = footer_part(obj, rId) %#ok<INUSD,MANU,STOUT>
            % FOOTER_PART STUB (parts/document.py 67-69). P2 header/footer tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.parts.FooterPart (owning WP: P2 header/footer tier) " + ...
                "required by mat2doc.parts.DocumentPart.footer_part");
        end

        function style = get_style(obj, style_id, style_type) %#ok<INUSD,MANU,STOUT>
            % GET_STYLE STUB (parts/document.py 71-77). P2 styles tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.styles.Styles.get_by_id (owning WP: P2 styles tier) " + ...
                "required by mat2doc.parts.DocumentPart.get_style");
        end

        function style_id = get_style_id(obj, style_or_name, style_type) %#ok<INUSD,MANU,STOUT>
            % GET_STYLE_ID STUB (parts/document.py 79-87). P2 styles tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.styles.Styles.get_style_id (owning WP: P2 styles tier) " + ...
                "required by mat2doc.parts.DocumentPart.get_style_id");
        end

        function hp = header_part(obj, rId) %#ok<INUSD,MANU,STOUT>
            % HEADER_PART STUB (parts/document.py 89-91). P2 header/footer tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.parts.HeaderPart (owning WP: P2 header/footer tier) " + ...
                "required by mat2doc.parts.DocumentPart.header_part");
        end

        function s = inline_shapes(obj) %#ok<MANU,STOUT>
            % INLINE_SHAPES STUB (parts/document.py 93-96, @lazyproperty). P2 shape tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.shape.InlineShapes (owning WP: P2 shape tier) required " + ...
                "by mat2doc.parts.DocumentPart.inline_shapes");
        end

        function np = numbering_part(obj) %#ok<MANU,STOUT>
            % NUMBERING_PART STUB (parts/document.py 98-109, @lazyproperty). P2-2 numbering tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.parts.NumberingPart (owning WP: P2-2 numbering tier) " + ...
                "required by mat2doc.parts.DocumentPart.numbering_part");
        end

        function s = settings(obj) %#ok<MANU,STOUT>
            % SETTINGS STUB (parts/document.py 116-120). P2-2 settings tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.settings.Settings (owning WP: P2-2 settings tier) " + ...
                "required by mat2doc.parts.DocumentPart.settings");
        end

        function s = styles(obj) %#ok<MANU,STOUT>
            % STYLES STUB (parts/document.py 122-126). P2-2 styles tier.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.styles.Styles (owning WP: P2-2 styles tier) required " + ...
                "by mat2doc.parts.DocumentPart.styles");
        end
    end

    methods (Static)
        function obj = load(partname, content_type, blob, package)
            % LOAD OWN static override (the inherited-static trap) -- the faithful
            %   MATLAB realization of the inherited XmlPart.load with
            %   cls=DocumentPart (opc/part.py 229-232): parse the blob into an
            %   element tree and construct a DocumentPart. ARG ORDER
            %   blob-before-package matches PartFactory.create. WITHOUT this
            %   override, DocumentPart.load would resolve to the inherited
            %   XmlPart.load and silently construct a base XmlPart -- so the
            %   PartFactory flip WML_DOCUMENT_MAIN -> DocumentPart would be inert.
            element = mat2doc.oxml.parse_xml(blob);
            obj = mat2doc.parts.DocumentPart( ...
                partname, content_type, element, package);
        end
    end
end

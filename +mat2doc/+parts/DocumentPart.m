classdef DocumentPart < mat2doc.parts.StoryPart
% DOCUMENTPART Main document part of a WordprocessingML (WML) package (.docx).
%
%   Acts as broker to other parts such as image, core properties, and style
%   parts. It also acts as a convenient delegate when a mid-document object needs
%   a service involving a remote ancestor (the Parented.part chain ends here).
%
%   TIER (P2-2): DocumentPart < mat2doc.parts.StoryPart < mat2doc.opc.XmlPart
%   (parts/document.py:26 `class DocumentPart(StoryPart)`). The StoryPart tier
%   was inserted at P2-2 (discharges VERIFY-M1-DOCPART-BASE); because StoryPart
%   adds methods only (blob/load/element inherit unchanged from XmlPart), the
%   reparenting is BYTE-NEUTRAL -- the reparented DocumentPart emits identical
%   bytes. (Earlier: DocumentPart < XmlPart directly, the M1 stand-in.)
%
%   P2-2 UN-STUB THE OBJECT GRAPH (not the FEATURES): the PART-level plumbing is
%   made live; the FEATURE surface still raises notYetPorted at its real-phase
%   owner. Split:
%     LIVE (part level / M1):
%       document / save / core_properties  (M1)
%       _styles_part   -> real mat2doc.parts.StylesPart   (part-level plumbing)
%       _settings_part -> real mat2doc.parts.SettingsPart (part-level plumbing)
%     STILL STUB (feature surface, correct real-phase owner):
%       styles              -> P4-7 (StylesProxy / StyleFactory)
%       settings            -> P5-1 (Settings proxy)
%       numbering_part      -> P8-1 (NumberingPart.new / numbering definitions)
%       comments / _comments_part -> P8-2 (CommentsPart + Comments)
%       inline_shapes       -> P7   (InlineShapes)
%       add/drop/footer/header_part -> P5-3b (Header/FooterPart)
%       get_style / get_style_id -> P4-7 (styles resolution)
%
%   IDENTITY (H5/H9): _styles_part/_settings_part are PLAIN @property (docx
%   document.py 156-169), NOT lazyproperty -- each call runs part_related_by; on
%   a package that HAS the part (default.docx does), it returns the SAME loaded
%   part handle every time (identity via the live rels). Only when the part is
%   ABSENT do they create a default and relate it, after which subsequent reads
%   find that same handle -- "materialized once" via the relationship, exactly as
%   python-docx does.
%
%   OWN CONSTRUCTOR + OWN STATIC `load` (the inherited-static trap): MATLAB does
%   not inherit constructors or dispatch inherited static methods to the
%   subclass, so this class declares its own pass-through constructor (chaining
%   to StoryPart) and its own `load` (the PartFactory entry point) constructing a
%   DocumentPart -- the faithful realization of Python's inherited-but-cls-bound
%   XmlPart.load (opc/part.py 229-232, cls=DocumentPart). Without its own `load`
%   the factory would silently build a base XmlPart and `.document` would vanish.
%
%   ARG ORDER (docx): DocumentPart(partname, content_type, element, package) --
%   element third, package last (XmlPart.__init__ order).
%
%   `document` is a PLAIN @property (parts/document.py 58-61) -- a FRESH
%   mat2doc.document.Document on EACH access (NOT cached), so two reads are
%   distinct proxies wrapping the SAME element, exactly as in python-docx.
%
%   Example:
%       tpl = fullfile(fileparts(fileparts(which( ...
%           "mat2doc.package.Package"))), "templates", "default.docx");
%       pkg = mat2doc.package.Package.open(tpl);
%       dp  = pkg.main_document_part();          % a mat2doc.parts.DocumentPart
%       d   = dp.document();                     % a mat2doc.document.Document
%
%   Ported from python-docx v1.2.0: src/docx/parts/document.py::DocumentPart
%   (LIVE: core_properties 52-56, document 58-61, save 111-114, _settings_part
%   143-154, _styles_part 156-169. Feature stubs name their real-phase owner.)

    methods
        function obj = DocumentPart(partname, content_type, element, package)
            % Pass-through to the StoryPart constructor (design.md CT_*/part
            %   constructor contract): forward ALL args, no re-validation. ARG
            %   ORDER element-before-package matches docx XmlPart.__init__.
            obj@mat2doc.parts.StoryPart(partname, content_type, element, package);
        end

        function cp = core_properties(obj)
            % CORE_PROPERTIES (parts/document.py 52-56): a CoreProperties object
            %   with read/write access to the core properties of this document.
            %   Python: return self.package.core_properties. LIVE (P1-7).
            cp = obj.package().core_properties();
        end

        function d = document(obj)
            % DOCUMENT (parts/document.py 58-61, plain @property): a Document
            %   object over this document's content. Python:
            %   return Document(self._element, self). FRESH proxy each access.
            d = mat2doc.document.Document(obj.element(), obj);
        end

        function save(obj, path_or_stream)
            % SAVE (parts/document.py 111-114): save this document to
            %   `path_or_stream`. Python: self.package.save(path_or_stream).
            obj.package().save(path_or_stream);
        end

        % ------------------------------------------------------------------
        % PART-level plumbing -- LIVE at P2-2 (returns real thin parts).
        % ------------------------------------------------------------------

        function sp = styles_part_(obj)
            % _styles_part (document.py 156-169, plain @property): the StylesPart
            %   for this document, creating an empty one if absent. UN-STUBBED at
            %   P2-2 -- returns a real mat2doc.parts.StylesPart. Python catches
            %   ONLY KeyError from part_related_by (a ValueError from a duplicate
            %   rel propagates); the MATLAB try/catch re-raises any non-KeyError.
            %   Underscore rotation: _styles_part -> styles_part_.
            try
                sp = obj.part_related_by(mat2doc.opc.RELATIONSHIP_TYPE.STYLES);
            catch ME
                if ME.identifier ~= "mat2doc:KeyError"
                    rethrow(ME);
                end
                pkg = obj.package();
                sp = mat2doc.parts.StylesPart.default(pkg);
                obj.relate_to(sp, mat2doc.opc.RELATIONSHIP_TYPE.STYLES);
            end
        end

        function sp = settings_part_(obj)
            % _settings_part (document.py 142-154, plain @property): the
            %   SettingsPart for this document, creating a default one if absent.
            %   UN-STUBBED at P2-2 -- returns a real mat2doc.parts.SettingsPart.
            %   KeyError-only catch as _styles_part.
            %   Underscore rotation: _settings_part -> settings_part_.
            try
                sp = obj.part_related_by(mat2doc.opc.RELATIONSHIP_TYPE.SETTINGS);
            catch ME
                if ME.identifier ~= "mat2doc:KeyError"
                    rethrow(ME);
                end
                pkg = obj.package();
                sp = mat2doc.parts.SettingsPart.default(pkg);
                obj.relate_to(sp, mat2doc.opc.RELATIONSHIP_TYPE.SETTINGS);
            end
        end

        % ------------------------------------------------------------------
        % Feature stubs (mat2doc:notYetPorted) -- each names its real-phase
        % owner (FLAG-B corrected: NOT "P2 <x> tier"). NONE is on open->save.
        % ------------------------------------------------------------------

        function [footer_part, rId] = add_footer_part(obj) %#ok<MANU,STOUT>
            % ADD_FOOTER_PART STUB (parts/document.py 35-39). Owner: P5-3b.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.parts.FooterPart (owning WP: P5-3b header/footer tier) " + ...
                "required by mat2doc.parts.DocumentPart.add_footer_part");
        end

        function [header_part, rId] = add_header_part(obj) %#ok<MANU,STOUT>
            % ADD_HEADER_PART STUB (parts/document.py 41-45). Owner: P5-3b.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.parts.HeaderPart (owning WP: P5-3b header/footer tier) " + ...
                "required by mat2doc.parts.DocumentPart.add_header_part");
        end

        function c = comments(obj) %#ok<MANU,STOUT>
            % COMMENTS STUB (parts/document.py 47-50). Owner: P8-2.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.parts.CommentsPart / mat2doc.comments.Comments (owning " + ...
                "WP: P8-2 comments tier) required by mat2doc.parts.DocumentPart.comments");
        end

        function drop_header_part(obj, rId) %#ok<INUSD,MANU>
            % DROP_HEADER_PART STUB (parts/document.py 63-65). Owner: P5-3b.
            %   Faithful body is `self.drop_rel(rId)` (drop_rel is LIVE at P2-2);
            %   the accessor stays stubbed until the P5-3b header/footer tier
            %   wires the caller (Section header/footer handling).
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.parts.DocumentPart.drop_header_part (owning WP: P5-3b " + ...
                "header/footer tier) is not yet ported");
        end

        function fp = footer_part(obj, rId) %#ok<INUSD,MANU,STOUT>
            % FOOTER_PART STUB (parts/document.py 67-69). Owner: P5-3b.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.parts.FooterPart (owning WP: P5-3b header/footer tier) " + ...
                "required by mat2doc.parts.DocumentPart.footer_part");
        end

        function style = get_style(obj, style_id, style_type) %#ok<INUSD,MANU,STOUT>
            % GET_STYLE STUB (parts/document.py 71-77). Owner: P4-7.
            %   OVERRIDES StoryPart.get_style (docx DocumentPart defines its own,
            %   `self.styles.get_by_id(...)`); stays stubbed until P4-7 styles.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.styles.Styles.get_by_id (owning WP: P4-7 styles tier) " + ...
                "required by mat2doc.parts.DocumentPart.get_style");
        end

        function style_id = get_style_id(obj, style_or_name, style_type) %#ok<INUSD,MANU,STOUT>
            % GET_STYLE_ID STUB (parts/document.py 79-87). Owner: P4-7.
            %   OVERRIDES StoryPart.get_style_id (docx DocumentPart defines its
            %   own, `self.styles.get_style_id(...)`).
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.styles.Styles.get_style_id (owning WP: P4-7 styles tier) " + ...
                "required by mat2doc.parts.DocumentPart.get_style_id");
        end

        function hp = header_part(obj, rId) %#ok<INUSD,MANU,STOUT>
            % HEADER_PART STUB (parts/document.py 89-91). Owner: P5-3b.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.parts.HeaderPart (owning WP: P5-3b header/footer tier) " + ...
                "required by mat2doc.parts.DocumentPart.header_part");
        end

        function s = inline_shapes(obj) %#ok<MANU,STOUT>
            % INLINE_SHAPES STUB (parts/document.py 93-96, @lazyproperty). Owner: P7.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.shape.InlineShapes (owning WP: P7 image/shape tier) " + ...
                "required by mat2doc.parts.DocumentPart.inline_shapes");
        end

        function np = numbering_part(obj) %#ok<MANU,STOUT>
            % NUMBERING_PART STUB (parts/document.py 98-109, @lazyproperty). Owner: P8-1.
            %   Kept stubbed at P2-2 (the numbering-definitions feature and
            %   NumberingPart.new -- which raises NotImplementedError in
            %   python-docx itself -- land at P8-1). NumberingPart the PART class
            %   IS ported (thin XmlPart shell) so the WML_NUMBERING flip lands;
            %   this ACCESSOR stays a feature stub.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.parts.NumberingPart numbering definitions (owning WP: " + ...
                "P8-1 numbering tier) required by mat2doc.parts.DocumentPart.numbering_part");
        end

        function s = settings(obj) %#ok<MANU,STOUT>
            % SETTINGS STUB (parts/document.py 116-120). Owner: P5-1.
            %   Faithful body: return self._settings_part.settings. _settings_part
            %   is LIVE (returns a real SettingsPart); SettingsPart.settings (the
            %   Settings proxy) is the P5-1 stub, so the notYetPorted surfaces
            %   there.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.settings.Settings (owning WP: P5-1 settings tier) " + ...
                "required by mat2doc.parts.DocumentPart.settings");
        end

        function s = styles(obj) %#ok<MANU,STOUT>
            % STYLES STUB (parts/document.py 122-126). Owner: P4-7.
            %   Faithful body: return self._styles_part.styles. _styles_part is
            %   LIVE (returns a real StylesPart); StylesPart.styles (the Styles
            %   proxy) is the P4-7 stub.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.styles.Styles (owning WP: P4-7 styles tier) required " + ...
                "by mat2doc.parts.DocumentPart.styles");
        end
    end

    methods (Access = private)
        function cp = comments_part_(obj) %#ok<MANU,STOUT>
            % _comments_part STUB (parts/document.py 128-140, plain @property).
            %   Owner: P8-2. Faithful body fetches the CommentsPart via
            %   part_related_by(RT.COMMENTS) or creates CommentsPart.default. The
            %   CommentsPart class + WML_COMMENTS flip land at P8-2.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.parts.CommentsPart (owning WP: P8-2 comments tier) " + ...
                "required by mat2doc.parts.DocumentPart._comments_part");
        end
    end

    methods (Static)
        function obj = load(partname, content_type, blob, package)
            % LOAD OWN static override (the inherited-static trap) -- the faithful
            %   MATLAB realization of the inherited XmlPart.load with
            %   cls=DocumentPart (opc/part.py 229-232): parse the blob into an
            %   element tree and construct a DocumentPart. WITHOUT this override
            %   DocumentPart.load would resolve to the inherited XmlPart.load and
            %   silently construct a base XmlPart, making the PartFactory flip
            %   WML_DOCUMENT_MAIN -> DocumentPart inert.
            element = mat2doc.oxml.parse_xml(blob);
            obj = mat2doc.parts.DocumentPart( ...
                partname, content_type, element, package);
        end
    end
end

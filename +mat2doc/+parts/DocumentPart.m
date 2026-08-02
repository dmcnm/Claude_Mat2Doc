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
%       get_style / get_style_id -> P4-7 (styles resolution)
%     LIVE at P5-3b (header/footer separate-part wiring):
%       add_header_part / add_footer_part -> Header/FooterPart.new + relate_to
%       drop_header_part -> drop_rel; header_part / footer_part -> related_parts[rId]
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

    properties (Access = private)
        numbering_part_cache_                          % numbering_part lazyproperty cache (P8-1)
        numbering_part_computed_ (1,1) logical = false
    end

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

        function [footer_part, rId] = add_footer_part(obj)
            % ADD_FOOTER_PART (parts/document.py 35-39): the (footer_part, rId)
            %   pair for a newly-created footer part. Python:
            %     footer_part = FooterPart.new(self.package)
            %     rId = self.relate_to(footer_part, RT.FOOTER)
            %     return footer_part, rId
            %   UN-STUBBED at P5-3b. relate_to is LIVE (P2-2); FooterPart.new
            %   creates word/footerN.xml (next_partname) with the default template.
            %
            %   Ported from python-docx v1.2.0: src/docx/parts/document.py::DocumentPart.add_footer_part
            footer_part = mat2doc.parts.FooterPart.new(obj.package());
            rId = obj.relate_to(footer_part, mat2doc.opc.RELATIONSHIP_TYPE.FOOTER);
        end

        function [header_part, rId] = add_header_part(obj)
            % ADD_HEADER_PART (parts/document.py 41-45): the (header_part, rId)
            %   pair for a newly-created header part. Python:
            %     header_part = HeaderPart.new(self.package)
            %     rId = self.relate_to(header_part, RT.HEADER)
            %     return header_part, rId
            %   UN-STUBBED at P5-3b. relate_to is LIVE (P2-2); HeaderPart.new
            %   creates word/headerN.xml (next_partname) with the default template.
            %
            %   Ported from python-docx v1.2.0: src/docx/parts/document.py::DocumentPart.add_header_part
            header_part = mat2doc.parts.HeaderPart.new(obj.package());
            rId = obj.relate_to(header_part, mat2doc.opc.RELATIONSHIP_TYPE.HEADER);
        end

        function c = comments(obj)
            % COMMENTS A Comments object providing access to the comments added to
            %   this document (parts/document.py 47-50, @property). Python:
            %   return self._comments_part.comments. UN-STUBBED at P8-2:
            %   _comments_part (comments_part_) materializes/loads the CommentsPart,
            %   whose `comments` returns the Comments proxy.
            %
            %   Ported from python-docx v1.2.0: src/docx/parts/document.py::DocumentPart.comments
            c = obj.comments_part_().comments;
        end

        function drop_header_part(obj, rId)
            % DROP_HEADER_PART (parts/document.py 63-65): remove the related header
            %   part identified by `rId`. Python: self.drop_rel(rId). UN-STUBBED at
            %   P5-3b. drop_rel is LIVE (P2-2, ref-count < 2 threshold).
            %
            %   ASYMMETRY (preserved verbatim, cross-part audit C): only the HEADER
            %   drop path goes through this DocumentPart method (section.py 460-463,
            %   _Header._drop_definition -> drop_header_part). The FOOTER drop path
            %   (section.py 414-417, _Footer._drop_definition) calls
            %   self._document_part.drop_rel(rId) DIRECTLY -- there is NO
            %   drop_footer_part in python-docx. Do not "tidy" this into symmetry.
            %
            %   Ported from python-docx v1.2.0: src/docx/parts/document.py::DocumentPart.drop_header_part
            obj.drop_rel(rId);
        end

        function fp = footer_part(obj, rId)
            % FOOTER_PART (parts/document.py 67-69): the FooterPart related by
            %   `rId`. Python: return self.related_parts[rId]. UN-STUBBED at P5-3b.
            %   related_parts is the {rId -> target Part} dictionary (values stored
            %   as 1x1 cells, P1-5 currency); [rId] -> unwrap the cell (H5: the
            %   live related part handle).
            %
            %   Ported from python-docx v1.2.0: src/docx/parts/document.py::DocumentPart.footer_part
            rp = obj.related_parts();       % dictionary(string -> cell)
            cellval = rp(rId);              % Python: related_parts[rId]
            fp = cellval{1};
        end

        function style = get_style(obj, style_id, style_type)
            % GET_STYLE The style matching `style_id` (parts/document.py 71-77).
            %   Returns the default style for `style_type` if `style_id` is []
            %   (None) or does not match a defined style of `style_type`. Python:
            %   return self.styles.get_by_id(style_id, style_type). UN-STUBBED at
            %   P4-7a. OVERRIDES StoryPart.get_style (docx DocumentPart defines its
            %   own).
            %
            %   Ported from python-docx v1.2.0: parts/document.py::DocumentPart.get_style
            style = obj.styles().get_by_id(style_id, style_type);
        end

        function style_id = get_style_id(obj, style_or_name, style_type)
            % GET_STYLE_ID The style_id matching `style_or_name`, or [] (parts/document.py 79-87).
            %   Returns [] (None) if the style resolves to the default for
            %   `style_type` or if `style_or_name` is [] (None). Raises if
            %   `style_or_name` is a style of the wrong type or names a style not
            %   present. Python: return self.styles.get_style_id(style_or_name,
            %   style_type). UN-STUBBED at P4-7a. OVERRIDES StoryPart.get_style_id.
            %
            %   Ported from python-docx v1.2.0: parts/document.py::DocumentPart.get_style_id
            style_id = obj.styles().get_style_id(style_or_name, style_type);
        end

        function hp = header_part(obj, rId)
            % HEADER_PART (parts/document.py 89-91): the HeaderPart related by
            %   `rId`. Python: return self.related_parts[rId]. UN-STUBBED at P5-3b.
            %   related_parts is the {rId -> target Part} dictionary (values stored
            %   as 1x1 cells, P1-5 currency); [rId] -> unwrap the cell (H5: the
            %   live related part handle).
            %
            %   Ported from python-docx v1.2.0: src/docx/parts/document.py::DocumentPart.header_part
            rp = obj.related_parts();       % dictionary(string -> cell)
            cellval = rp(rId);              % Python: related_parts[rId]
            hp = cellval{1};
        end

        function s = inline_shapes(obj) %#ok<MANU,STOUT>
            % INLINE_SHAPES STUB (parts/document.py 93-96, @lazyproperty). Owner: P7.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.shape.InlineShapes (owning WP: P7 image/shape tier) " + ...
                "required by mat2doc.parts.DocumentPart.inline_shapes");
        end

        function np = numbering_part(obj)
            % NUMBERING_PART (parts/document.py 98-109, @lazyproperty): the
            %   NumberingPart providing access to the numbering definitions,
            %   creating an empty one if not present. UN-STUBBED at P8-1. Python:
            %     try:
            %         return cast(NumberingPart, self.part_related_by(RT.NUMBERING))
            %     except KeyError:
            %         numbering_part = NumberingPart.new()
            %         self.relate_to(numbering_part, RT.NUMBERING)
            %         return numbering_part
            %   part_related_by raises mat2doc:KeyError when absent (any other
            %   exception propagates). NumberingPart.new() FAITHFULLY raises
            %   mat2doc:NotImplementedError (python-docx v1.2.0 leaves new()
            %   unimplemented), so on a package WITHOUT a numbering part this
            %   accessor raises NotImplementedError -- exactly as python-docx does.
            %   @lazyproperty: cache only a SUCCESSFUL return (an exception is
            %   re-raised on every access, never cached).
            if ~obj.numbering_part_computed_
                RT = mat2doc.opc.RELATIONSHIP_TYPE;
                try
                    np_local = obj.part_related_by(RT.NUMBERING);
                catch ME
                    if ME.identifier ~= "mat2doc:KeyError"
                        rethrow(ME);
                    end
                    np_local = mat2doc.parts.NumberingPart.new();   % raises NotImplementedError (faithful)
                    obj.relate_to(np_local, RT.NUMBERING);
                end
                obj.numbering_part_cache_ = np_local;
                obj.numbering_part_computed_ = true;
            end
            np = obj.numbering_part_cache_;
        end

        function s = settings(obj)
            % SETTINGS The Settings object for this document (parts/document.py
            %   116-120, @property). Python: return self._settings_part.settings.
            %   UN-STUBBED at P5-1: _settings_part (settings_part_) is LIVE and
            %   SettingsPart.settings now returns a real Settings proxy.
            %
            %   Ported from python-docx v1.2.0: src/docx/parts/document.py::DocumentPart.settings
            s = obj.settings_part_().settings;
        end

        function s = styles(obj)
            % STYLES The styles in the styles part of this document (parts/document.py 122-126).
            %   Python: return self._styles_part.styles. UN-STUBBED at P4-7a:
            %   _styles_part is LIVE (a real StylesPart, P2-2) and StylesPart.styles
            %   now returns the real mat2doc.styles.Styles proxy.
            %
            %   Ported from python-docx v1.2.0: parts/document.py::DocumentPart.styles
            s = obj.styles_part_().styles();
        end
    end

    methods (Access = private)
        function cp = comments_part_(obj)
            % _comments_part The CommentsPart for this document, creating a default
            %   one if not present (parts/document.py 128-140, plain @property).
            %   UN-STUBBED at P8-2. Python:
            %     try:
            %         return cast(CommentsPart, self.part_related_by(RT.COMMENTS))
            %     except KeyError:
            %         assert self.package is not None
            %         comments_part = CommentsPart.default(self.package)
            %         self.relate_to(comments_part, RT.COMMENTS)
            %         return comments_part
            %   part_related_by raises mat2doc:KeyError when absent (any other
            %   exception propagates -- the numbering_part precedent). This is a
            %   PLAIN @property (NOT a lazyproperty): re-evaluated each access, but
            %   idempotent -- the first access creates & relates the part, later
            %   accesses find it via part_related_by (H5: relate_to dedupes to the
            %   same rId/part). CommentsPart.default(package) builds word/comments.xml
            %   from the default template on demand.
            %
            %   Ported from python-docx v1.2.0: src/docx/parts/document.py::DocumentPart._comments_part
            RT = mat2doc.opc.RELATIONSHIP_TYPE;
            try
                cp = obj.part_related_by(RT.COMMENTS);   % Python: self.part_related_by(RT.COMMENTS)
                return
            catch ME
                if ME.identifier ~= "mat2doc:KeyError"   % Python: except KeyError
                    rethrow(ME);
                end
            end
            cp = mat2doc.parts.CommentsPart.default(obj.package());  % Python: CommentsPart.default(self.package)
            obj.relate_to(cp, RT.COMMENTS);                          % Python: self.relate_to(comments_part, RT.COMMENTS)
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

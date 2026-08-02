classdef PartFactory
% PARTFACTORY Constructs the registered Part subtype for a content type / reltype.
%
%   part = MAT2DOC.OPC.PARTFACTORY.CREATE(partname, content_type, reltype, blob,
%   package) dispatches to the Part class registered for `reltype` (via the
%   part-class selector) or, failing that, `content_type` (via the content-type
%   map), or the default base Part, and returns Class.load(partname,
%   content_type, blob, package).
%
%   Python's PartFactory is a callable whose __new__ dispatches (part.py 182-196);
%   MATLAB has no __new__, so `create` is the explicit factory (design.md section
%   2: factories are explicit switch/registry tables) and OpcPackage.open passes
%   the handle @mat2doc.opc.PartFactory.create to the Unmarshaller in place of the
%   Python class object. The registration set (part_class_selector + part_type_for
%   + default_part_type) lives in docx/__init__.py (lines 37-51), NOT in
%   opc/part.py; it is baked here as static tables mirroring those lines exactly.
%
%   M1 BYTE-IDENTITY CONTRACT -- the whitespace-collapse decider (frozen M1
%   Finding 1, validation\mat2doc\references\m1_skeleton_target.md):
%   python-docx re-serializes exactly the XML content types it registers to an
%   XmlPart SUBCLASS, and keeps every other part's bytes VERBATIM through the
%   generic base Part. On a plain open->save of default.docx:
%     * document.xml / styles.xml / settings.xml / numbering.xml -> XmlPart
%       subclasses (DocumentPart/StylesPart/SettingsPart/NumberingPart) -> parsed
%       with remove_blank_text and re-serialized -> the template's pretty-print
%       whitespace COLLAPSES -> L1 byte-match with the round-trip oracle.
%     * stylesWithEffects.xml / webSettings.xml / fontTable.xml / theme1.xml /
%       app.xml / customXml / itemProps -> content types NOT in the registry ->
%       fall back to the base Part -> bytes kept VERBATIM (passthrough).
%   Mat2Doc's registry MUST reproduce this split so the same 4 parts collapse and
%   the other XML siblings stay byte-verbatim. This is the concrete M1 requirement.
%
%   XmlPart-vs-Part split (progressively refined P1-8 -> P2-2 -> later): the
%   eight registered part classes are all XmlPart subclasses in python-docx
%   (CorePropertiesPart, CommentsPart, DocumentPart, FooterPart, HeaderPart,
%   NumberingPart, SettingsPart, StylesPart -- each extends XmlPart or
%   StoryPart<XmlPart), and ImagePart (via the part_class_selector for reltype
%   IMAGE) is a plain Part. As each subclass is ported its row is FLIPPED from
%   the base-XmlPart stand-in to the real class; because every one inherits
%   XmlPart.blob unchanged (parse+re-serialize), the emitted bytes are identical
%   -- only the reloaded part's TYPE changes. Ported so far: CorePropertiesPart
%   (P1-7), DocumentPart (P1-8), StylesPart/SettingsPart/NumberingPart (P2-2),
%   HeaderPart/FooterPart (P5-3b). Still base-XmlPart stand-in: WML_COMMENTS
%   (P8-2). The IMAGE selector maps to mat2doc.parts.ImagePart (P7-4 flip; was
%   the base mat2doc.opc.Part stand-in through P7-3). This preserves the M1
%   collapse/passthrough BYTE behavior throughout (ImagePart inherits Part.blob
%   verbatim, so a loaded image part round-trips byte-identically).
%
%   ARG ORDER (docx): create(partname, content_type, reltype, blob, package) and
%   Class.load(partname, content_type, blob, package) -- blob before package.
%
%   UNDERSCORE ROTATION (design.md section 2): private `_part_cls_for` ->
%   part_cls_for_; the registrations `part_type_for` -> part_type_for_ and the
%   selector `part_class_selector` -> part_class_selector_.
%
%   Example:
%       CT = mat2doc.opc.CONTENT_TYPE;
%       f  = @mat2doc.opc.PartFactory.part_cls_for_;
%       disp(f(CT.WML_STYLES))            % "mat2doc.parts.StylesPart" (P2-2 flip)
%       disp(f(CT.WML_WEB_SETTINGS))      % "mat2doc.opc.Part"     (passthrough)
%
%   Ported from python-docx v1.2.0: src/docx/opc/part.py::PartFactory
%   (lines 165-204) + the registration block in src/docx/__init__.py (lines 37-51)

    methods (Static)
        function part = create(partname, content_type, reltype, blob, package)
            % CREATE Construct the registered Part subtype (part.py __new__
            %   182-196). docx ALWAYS registers a part_class_selector
            %   (docx/__init__.py:43), so the `if cls.part_class_selector is not
            %   None` guard is always taken. The selector is looked up through
            %   cls_method_fn -- the getattr(cls, "part_class_selector") analogue
            %   (part.py:192; resolves the P1-5 cls_method_fn VERIFY: cls currency
            %   is the class-name STRING "mat2doc.opc.PartFactory").
            part_class_selector = mat2doc.opc.cls_method_fn( ...
                "mat2doc.opc.PartFactory", "part_class_selector_");
            PartClass = part_class_selector(content_type, reltype);
            if PartClass == ""   % Python: if PartClass is None (H3; "" is None here)
                PartClass = mat2doc.opc.PartFactory.part_cls_for_(content_type);
            end
            % Python: PartClass.load(...). getattr(PartClass, "load") -> cls_method_fn.
            loadfn = mat2doc.opc.cls_method_fn(PartClass, "load");
            part = loadfn(partname, content_type, blob, package);
        end

        function cls = part_cls_for_(content_type)
            % _part_cls_for (part.py 198-204): the class-name registered for
            %   `content_type`, or the default part class if none is registered.
            reg = mat2doc.opc.PartFactory.part_type_for_();
            idx = find(reg(:, 1) == string(content_type), 1);
            if isempty(idx)
                cls = mat2doc.opc.PartFactory.default_part_type_();
            else
                cls = reg(idx, 2);
            end
        end

        function cls = part_class_selector_(content_type, reltype) %#ok<INUSD>
            % PART_CLASS_SELECTOR_ (docx/__init__.py 37-40): reltype IMAGE -> the
            %   image part class; otherwise "" (Python None -> fall through to the
            %   content-type map). P7-4 FLIP: IMAGE -> mat2doc.parts.ImagePart (was
            %   the base mat2doc.opc.Part stand-in through M1..P7-3). Byte-neutral
            %   on LOAD: ImagePart inherits Part.blob unchanged (verbatim image
            %   bytes), so a loaded image part round-trips identically -- only its
            %   TYPE changes, gaining .sha1 for the ImageParts dedupe. Not exercised
            %   on default.docx (no IMAGE relationship; its thumbnail is THUMBNAIL).
            if reltype == mat2doc.opc.RELATIONSHIP_TYPE.IMAGE
                cls = "mat2doc.parts.ImagePart";   % P7-4 flip (was mat2doc.opc.Part)
            else
                cls = "";                   % Python None
            end
        end
    end

    methods (Static, Access = private)
        function reg = part_type_for_()
            % PART_TYPE_FOR_ Nx2 [content_type, class-name] table, a row-for-row
            %   mirror of docx/__init__.py PartFactory.part_type_for assignments
            %   (lines 44-51). ALL eight registered classes are XmlPart subclasses
            %   in python-docx; the base mat2doc.opc.XmlPart is the M1 stand-in for
            %   the six not-yet-ported subclasses (the XmlPart-vs-Part split), P2-2
            %   refines each to its specific subclass (trailing comment) with
            %   byte-identical output.
            %
            %   P1-8 + P2-2 ROW FLIPS (byte-neutral by blob inheritance): each
            %   ported subclass inherits XmlPart.blob unchanged (parse +
            %   serialize_part_xml) and declares its own static `load`
            %   constructing the subclass, so the reloaded part's TYPE changes but
            %   the emitted bytes are IDENTICAL to the base-XmlPart dispatch. The
            %   P1-8 flips: OPC_CORE_PROPERTIES (P1-7 class), WML_DOCUMENT_MAIN
            %   (M1-required). P2-2 flips the three thin XmlPart shells now ported:
            %   WML_STYLES -> StylesPart, WML_SETTINGS -> SettingsPart,
            %   WML_NUMBERING -> NumberingPart (mirrors docx/__init__.py 49-51).
            %   Re-proven byte-neutral by the 17/17 M1 sweep. P5-3b flips
            %   WML_FOOTER -> FooterPart, WML_HEADER -> HeaderPart (StoryPart<XmlPart
            %   shells). P8-2 flips WML_COMMENTS -> CommentsPart (StoryPart<XmlPart
            %   shell; on-demand part -- no M1 fixture loads it, so byte-neutral).
            CT  = mat2doc.opc.CONTENT_TYPE;
            reg = [ ...
                CT.OPC_CORE_PROPERTIES, "mat2doc.opc.parts.CorePropertiesPart";  ...  % P1-8 flip (P1-7 class)
                CT.WML_COMMENTS,        "mat2doc.parts.CommentsPart";   ...  % P8-2 flip (StoryPart<XmlPart shell)
                CT.WML_DOCUMENT_MAIN,   "mat2doc.parts.DocumentPart";   ...  % P1-8 flip (M1-required)
                CT.WML_FOOTER,          "mat2doc.parts.FooterPart";     ...  % P5-3b flip (StoryPart<XmlPart shell)
                CT.WML_HEADER,          "mat2doc.parts.HeaderPart";     ...  % P5-3b flip (StoryPart<XmlPart shell)
                CT.WML_NUMBERING,       "mat2doc.parts.NumberingPart";  ...  % P2-2 flip (thin XmlPart shell)
                CT.WML_SETTINGS,        "mat2doc.parts.SettingsPart";   ...  % P2-2 flip (thin XmlPart shell)
                CT.WML_STYLES,          "mat2doc.parts.StylesPart"];         % P2-2 flip (thin XmlPart shell)
        end

        function cls = default_part_type_()
            % DEFAULT_PART_TYPE (part.py 180): the base Part -- the verbatim
            %   passthrough part used when no registration matches.
            cls = "mat2doc.opc.Part";
        end
    end
end

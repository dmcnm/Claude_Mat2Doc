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
%   M1 STAND-IN (XmlPart-vs-Part split; P2 refines the subclasses): the eight
%   registered part classes are all XmlPart subclasses in python-docx
%   (CorePropertiesPart, CommentsPart, DocumentPart, FooterPart, HeaderPart,
%   NumberingPart, SettingsPart, StylesPart -- verified: each extends XmlPart or
%   StoryPart<XmlPart), and ImagePart (via the part_class_selector for reltype
%   IMAGE) is a plain Part. Those feature subclasses are not ported until P2, so
%   every registered content type here maps to the BASE mat2doc.opc.XmlPart, and
%   the IMAGE selector maps to the BASE mat2doc.opc.Part. This yields the correct
%   M1 collapse/passthrough BYTE behavior (an XmlPart subclass and the base
%   XmlPart share XmlPart.blob = parse+re-serialize; ImagePart and the base Part
%   share the verbatim Part.blob). P2 will REFINE each row to its specific
%   subclass; because those subclasses inherit the same blob unchanged, the
%   emitted bytes are identical -- only the reloaded part's TYPE changes.
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
%       disp(f(CT.WML_STYLES))            % "mat2doc.opc.XmlPart"  (reserialize)
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
            %   content-type map). M1 stand-in: IMAGE -> the base mat2doc.opc.Part
            %   (ImagePart, a plain passthrough Part, is ported in P2). Not
            %   exercised at M1 (default.docx has no IMAGE relationship; its
            %   thumbnail is a THUMBNAIL reltype).
            if reltype == mat2doc.opc.RELATIONSHIP_TYPE.IMAGE
                cls = "mat2doc.opc.Part";   % P2: mat2doc.parts.ImagePart
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
            %   P1-8 ROW FLIPS (byte-neutral by blob inheritance): the two rows
            %   whose subclasses are already ported are flipped to those classes.
            %   Both CorePropertiesPart (P1-7) and DocumentPart (P1-8) inherit
            %   XmlPart.blob unchanged (parse + serialize_part_xml) and their own
            %   static `load` constructs the subclass, so the reloaded part's TYPE
            %   changes but the emitted bytes are IDENTICAL to the base-XmlPart
            %   dispatch. The DocumentPart flip is REQUIRED for M1 (mat2doc.Document
            %   needs main_document_part.document); the CorePropertiesPart flip
            %   closes the P1-6b VERIFY-core-props at zero byte-risk.
            CT  = mat2doc.opc.CONTENT_TYPE;
            XP  = "mat2doc.opc.XmlPart";   % base XmlPart -> parse + re-serialize
            reg = [ ...
                CT.OPC_CORE_PROPERTIES, "mat2doc.opc.parts.CorePropertiesPart";  ...  % P1-8 flip (P1-7 class)
                CT.WML_COMMENTS,        XP;                             ...  % P2-2: CommentsPart (StoryPart<XmlPart)
                CT.WML_DOCUMENT_MAIN,   "mat2doc.parts.DocumentPart";   ...  % P1-8 flip (M1-required)
                CT.WML_FOOTER,          XP;                             ...  % P2-2: FooterPart (StoryPart<XmlPart)
                CT.WML_HEADER,          XP;                             ...  % P2-2: HeaderPart (StoryPart<XmlPart)
                CT.WML_NUMBERING,       XP;                             ...  % P2-2: NumberingPart (XmlPart)
                CT.WML_SETTINGS,        XP;                             ...  % P2-2: SettingsPart (XmlPart)
                CT.WML_STYLES,          XP];                                 % P2-2: StylesPart (XmlPart)
        end

        function cls = default_part_type_()
            % DEFAULT_PART_TYPE (part.py 180): the base Part -- the verbatim
            %   passthrough part used when no registration matches.
            cls = "mat2doc.opc.Part";
        end
    end
end

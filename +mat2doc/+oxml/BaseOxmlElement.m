classdef BaseOxmlElement < mat2doc.oxml.XmlElement
% BASEOXMLELEMENT Effective base class for all custom CT_* element classes.
%
%   Adds standardized behavior to all custom element classes in one place.
%   Every CT_* class ported in later WPs extends this class.
%
%   P1-3a SLICE (part 1 of 2). This is an ACCELERATED RE-PORT of the SOLVED
%   Mat2Ppt engine (Mat2Ppt/+mat2ppt/+oxml/BaseOxmlElement.m), which ports the
%   BYTE-IDENTICAL xmlchemy engine; docx v1.2.0 is the module source of truth
%   and CONFIRMS the same surface, with three descriptor DELTAS noted below.
%   The P1-3a slice ported:
%     * the tree-ops on BaseOxmlElement (xmlchemy.py lines 656-677):
%       first_child_found_in, insert_element_before, remove_all;
%     * xpath() (xmlchemy.py lines 687-692), delegating to the mini-XPath
%       engine +oxml/evaluate_xpath.m (re-ported from Mat2Ppt WP5+WP5-C).
%       HOISTED to mat2doc.oxml.XmlElement at P2-2 (task #60) and inherited
%       here; the override merely injected the fixed nsmap. See XmlElement.xpath;
%     * the ATTRIBUTE descriptor engine (OptionalAttribute / RequiredAttribute,
%       xmlchemy.py lines 154-261): getAttrTyped / setAttrTyped / getAttrRequired
%       / setAttrRequired + the type-dispatch helpers.
%
%   P1-3b SLICE (part 2 of 2, THIS extension). The CHILD-element descriptor
%   engine. xmlchemy's MetaOxmlElement generates, per declared child descriptor
%   (ZeroOrOne / ZeroOrMore / OneAndOnlyOne / OneOrMore / Choice /
%   ZeroOrOneChoice), a family of members (get.x / x_lst / _new_x / _insert_x /
%   _add_x / add_x / get_or_add_x / _remove_x / get_or_change_to_x /
%   _remove_eg_x) whose BODIES all reduce to a handful of generic operations.
%   This slice ports those operations as explicit engine methods every ported
%   CT_* class delegates to (design.md section 2 "BaseOxmlElement engine
%   contract"): getChild, getRequiredChild, getChildList, newChild,
%   insertChildInSequence, addChild, getOrAddChild, removeChild,
%   firstChildFoundIn, removeChildren, getOrChangeToChild. The
%   metaclass/subsref machinery is replaced by checked-in explicit delegating
%   members (no runtime metaprogramming).
%
%   ENGINE-METHOD -> xmlchemy SOURCE MAP (src/docx/oxml/xmlchemy.py, docx v1.2.0):
%    -- P1-3a (tree-ops + xpath + attribute engine) --
%     first_child_found_in  <- BaseOxmlElement.first_child_found_in (656-662)
%     insert_element_before <- BaseOxmlElement.insert_element_before (664-670)
%     remove_all            <- BaseOxmlElement.remove_all           (672-677)
%     xpath                 <- BaseOxmlElement.xpath (687-692) [hoisted to
%                              XmlElement at P2-2, task #60; inherited here]
%     getAttrTyped          <- OptionalAttribute._getter get_attr_value (187-193)
%     setAttrTyped          <- OptionalAttribute._setter set_attr_value (202-212)
%     getAttrRequired       <- RequiredAttribute._getter get_attr_value (240-246)
%     setAttrRequired       <- RequiredAttribute._setter set_attr_value (255-259)
%     resolveTypeCls_       <- BaseAttribute simple_type object       (117-120)
%     attrClarkName_        <- BaseAttribute._clark_name              (139-143)
%    -- P1-3b (child-element descriptor engine) --
%     getChild             <- _BaseChildElement._getter        (380-382)
%     getRequiredChild     <- OneAndOnlyOne._getter            (499-505)
%     getChildList         <- _BaseChildElement._list_getter   (397-398)
%     newChild             <- _BaseChildElement._creator       (366-367)
%     insertChildInSequence<- _BaseChildElement._add_inserter  (319-321)
%     addChild             <- _BaseChildElement._add_adder     (284-291)
%     getOrAddChild        <- ZeroOrOne._add_get_or_adder      (557-562)
%     removeChild          <- ZeroOrOne._add_remover           (572-573)
%     firstChildFoundIn    <- ZeroOrOneChoice._choice_getter   (622-623)
%     removeChildren       <- ZeroOrOneChoice._add_group_remover(610-612)
%     getOrChangeToChild   <- Choice._add_get_or_change_to_method(453-461)
%
%   DOCX-vs-PPTX DELTAS (docx is source of truth; ported to the docx forms,
%   NOT the Mat2Ppt/pptx forms -- see audit_P1-3a_xmlchemy_engine.md and
%   audit_P1-3b_child_descriptors.md):
%    -- attribute engine (P1-3a) --
%     D-delta-1  OptionalAttribute setter guards `value is None OR value ==
%                default` (docx line 203). pptx guards only `value ==
%                default`. When an optional attribute has a NON-None default,
%                assigning None removes it in docx but would fall through to
%                to_xml(None) in pptx.
%     D-delta-2  OptionalAttribute setter, AFTER to_xml, removes the attribute
%                when `str_value is None` (docx lines 208-211). pptx unconditionally
%                sets. Lets a simple type whose to_xml returns None erase the attr.
%     D-delta-3  RequiredAttribute setter, AFTER to_xml, RAISES
%                ValueError("cannot assign {value} to this required attribute")
%                when `str_value is None` (docx lines 257-258). pptx sets
%                unconditionally.
%    -- child-descriptor engine (P1-3b) --
%     D-delta-4  ZeroOrMore.populate_class_members calls _add_public_adder
%                (docx line 536), so a docx ZeroOrMore ALSO generates a public
%                `add_x()`; pptx ZeroOrMore does NOT (only OneOrMore does). This
%                is ENGINE-NEUTRAL: the public add_x() routes through the same
%                addChild primitive as _add_x (see _add_public_adder, docx
%                340-352). The delta only decides WHICH per-class delegating
%                members a future CT_* WP scaffolds, not the engine surface here.
%                Choice / ZeroOrOneChoice behaviour is otherwise IDENTICAL to
%                pptx (verified line-by-line, audit_P1-3b).
%
%   Ported at P1-4 (was deferred by the P1-3a slice):
%     xml property (serialize_for_reading, 679-685)    -> UN-STUBBED at P1-4 (the
%           doc-serialize/OPC WP): the `xml` method now routes through the real
%           pretty-print engine mat2doc.oxml.serialize_for_reading (test-only
%           helper; design.md section 3 forbids pretty-print in the PART
%           serializer, so serialize_for_reading is a distinct helper). No docx
%           PRODUCTION code path reads xmlchemy BaseOxmlElement `.xml` (the
%           production `.xml`-bytes path is the SEPARATE opc/oxml.py
%           CT_Relationships, rotated to xml_file_bytes).
%
%   Deferred, in their owning WPs (NOT ported here):
%     __repr__ (649-654) and _nsptag (694-696)         -> display-only; used ONLY
%           by __repr__ (grep-verified over docx src). Deferred exactly as the
%           Mat2Ppt reference deferred them. NOTE the rotated name `nsptag_`
%           would COLLIDE with XmlElement's private `nsptag_` property, so a
%           direct rotation is not even available; when eventually needed it
%           must take a distinct name.
%
%   Example:
%       p = mat2doc.oxml.BaseOxmlElement("w:p");
%       p.append(mat2doc.oxml.OxmlElement("w:r"));
%       pPr = p.insert_element_before( ...
%           mat2doc.oxml.OxmlElement("w:pPr"), "w:r");   % pPr placed before w:r
%
%   Ported from python-docx v1.2.0: src/docx/oxml/xmlchemy.py::BaseOxmlElement
%   (tree-ops + xpath + attribute engine [P1-3a] + child-element descriptor
%   engine [P1-3b]; metaclass/descriptor machinery replaced by the explicit
%   delegating-member scheme per design.md section 2)

    methods
        function obj = BaseOxmlElement(varargin)
            % BASEOXMLELEMENT Construct a loose custom element (see XmlElement).
            %
            %   TRANSPARENT PASS-THROUGH (design.md section 2 "CT_* constructor
            %   contract (INTEGRATION-CRITICAL)"): forwards ALL positional args
            %   (nsptag, nsmapOrDecls, resolvedUri) verbatim to the base
            %   XmlElement constructor via varargin, with NO re-validation of the
            %   nsmap argument. XmlElement is the SINGLE point that accepts both
            %   the struct nsmap and the Nx2 string decl-pair currencies; the
            %   parser (XmlParser.parseElement_) instantiates registered CT_*
            %   classes via feval(cls, name, ownDecls) where ownDecls is the Nx2
            %   decl-pair, so a struct-typed nsmap arg here would break parsing.
            obj = obj@mat2doc.oxml.XmlElement(varargin{:});
        end

        % =================================================================
        % TREE-OPS (xmlchemy.py BaseOxmlElement, lines 656-677)
        % =================================================================

        function child = first_child_found_in(obj, tagnames)
            % FIRST_CHILD_FOUND_IN First child with tag in tagnames, or [] (None) if not found.
            %
            %   child = e.FIRST_CHILD_FOUND_IN(tagname1, tagname2, ...) where
            %   each tagname is a prefixed tag, e.g. "w:pPr". Search order is
            %   ARGUMENT order (first NAME that matches wins, NOT document
            %   order) -- exactly the Python loop over *tagnames.
            %
            %   Ported from python-docx v1.2.0: src/docx/oxml/xmlchemy.py::
            %   BaseOxmlElement.first_child_found_in (lines 656-662)
            arguments
                obj (1,1) mat2doc.oxml.BaseOxmlElement
            end
            arguments (Repeating)
                tagnames (1,1) string
            end
            for k = 1:numel(tagnames)
                % Python: child = self.find(qn(tagname))  (line 659)
                child = obj.find(mat2doc.oxml.qn(tagnames{k}));
                if ~isequal(child, [])   % Python: if child is not None (H3)
                    return
                end
            end
            child = [];   % Python: return None (H3)
        end

        function elm = insert_element_before(obj, elm, tagnames)
            % INSERT_ELEMENT_BEFORE Insert elm before the first child found among tagnames.
            %
            %   elm = e.INSERT_ELEMENT_BEFORE(elm, tagname1, tagname2, ...)
            %   inserts elm immediately before the first child whose tag is in
            %   tagnames (argument-order search), or APPENDS it when none is
            %   found. Returns elm.
            %
            %   Ported from python-docx v1.2.0: src/docx/oxml/xmlchemy.py::
            %   BaseOxmlElement.insert_element_before (lines 664-670)
            arguments
                obj (1,1) mat2doc.oxml.BaseOxmlElement
                elm (1,1) mat2doc.oxml.XmlElement
            end
            arguments (Repeating)
                tagnames (1,1) string
            end
            successor = obj.first_child_found_in(tagnames{:});
            if ~isequal(successor, [])   % Python: if successor is not None (H3)
                successor.addprevious(elm);
            else
                obj.append(elm);
            end
        end

        function remove_all(obj, tagnames)
            % REMOVE_ALL Remove all child elements whose tag (e.g. "w:p") is in tagnames.
            %
            %   Ported from python-docx v1.2.0: src/docx/oxml/xmlchemy.py::
            %   BaseOxmlElement.remove_all (lines 672-677)
            arguments
                obj (1,1) mat2doc.oxml.BaseOxmlElement
            end
            arguments (Repeating)
                tagnames (1,1) string
            end
            for k = 1:numel(tagnames)
                % Python: matching = self.findall(qn(tagname))  (line 675)
                % H9: findall is already a materialized list in Python, so
                % removing while looping over it is safe in both languages.
                matching = obj.findall(mat2doc.oxml.qn(tagnames{k}));
                for child = matching
                    obj.remove(child);
                end
            end
        end

        % =================================================================
        % XPATH -- HOISTED to XmlElement (task #60, P2-2)
        % =================================================================
        % The fixed-nsmap xpath() (xmlchemy.py 687-692) now lives on
        % mat2doc.oxml.XmlElement (the lxml _Element analogue -- every _Element
        % has .xpath) and is INHERITED here verbatim. python-docx's
        % BaseOxmlElement.xpath override only injects `nsmap`/drops the
        % `namespaces` kwarg; it does not add the capability. Hoisting lets the
        % parser-fallback root class (unregistered parsed roots, e.g. w:document
        % before P2-3) support xpath, as lxml does. See XmlElement.xpath.

        function s = xml(obj)
            % XML Pretty-printed XML for this element, no declaration (test helper).
            %
            %   Python `BaseOxmlElement.xml` (@property, xmlchemy.py 679-685)
            %   returns serialize_for_reading(self) -- etree.tostring with
            %   pretty_print=True and no XML declaration. UN-STUBBED at P1-4 (the
            %   doc-serialize/OPC WP the P1-3a stub deferred to): now routes
            %   through the real engine mat2doc.oxml.serialize_for_reading.
            %
            %   TEST-ONLY surface: no docx PRODUCTION code path reads this
            %   property on a wordprocessing element (grep-verified; the
            %   production `.xml`-bytes path is the SEPARATE opc/oxml.py
            %   CT_Relationships, rotated to xml_file_bytes). Python wraps the
            %   result in XmlString (a str subclass whose __eq__ relaxes to a
            %   canonical comparison for tests); that relaxed equality is a
            %   test-harness concern and is not ported -- this returns a plain
            %   string. Pretty-print WHITESPACE byte-exactness vs lxml is asserted
            %   against live Python, not frozen (see serialize_for_reading.m
            %   VERIFY note).
            arguments
                obj (1,1) mat2doc.oxml.BaseOxmlElement
            end
            s = mat2doc.oxml.serialize_for_reading(obj);
        end
    end

    % =====================================================================
    % ATTRIBUTE DESCRIPTOR ENGINE (OptionalAttribute / RequiredAttribute)
    %
    % xmlchemy's MetaOxmlElement generates, per declared OptionalAttribute /
    % RequiredAttribute, a read/write `.{prop_name}` property whose get/set
    % bodies reduce to these generic operations. Ported as explicit engine
    % methods every CT_* class delegates to (design.md section 2). `name` args
    % are the schema ATTR names ("w:val" or "space"); `type` args are the simple
    % type's (or enum's) class-name token; `default` is the OptionalAttribute
    % default (None -> [], H3).
    % =====================================================================
    methods
        function value = getAttrTyped(obj, name, type, default)
            % GETATTRTYPED Type-converted attribute value, or default when absent (OptionalAttribute getter).
            %   Ported from xmlchemy.py OptionalAttribute._getter's
            %   `get_attr_value` (lines 187-193): attr = obj.get(clark); if None
            %   return default; return simple_type.from_xml(attr). Identical in
            %   docx and pptx (no delta). `name` -> Clark key via attrClarkName_;
            %   `type` -> from_xml static via resolveTypeCls_.
            arguments
                obj (1,1) mat2doc.oxml.BaseOxmlElement
                name (1,1) string
                type (1,1) string
                default = []
            end
            attr = obj.get(obj.attrClarkName_(name));   % [] (None) when absent (H3)
            if isequal(attr, [])                        % Python: if attr_str_value is None
                value = default;
                return
            end
            value = feval(obj.resolveTypeCls_(type) + ".from_xml", attr);
        end

        function setAttrTyped(obj, name, type, value, default)
            % SETATTRTYPED Set/remove a typed attribute (OptionalAttribute setter).
            %   Ported from xmlchemy.py OptionalAttribute._setter's
            %   `set_attr_value` (lines 202-212, DOCX form):
            %       if value is None or value == self._default:
            %           if clark in obj.attrib: del obj.attrib[clark]
            %           return
            %       str_value = simple_type.to_xml(value)
            %       if str_value is None:
            %           if clark in obj.attrib: del obj.attrib[clark]
            %           return
            %       obj.set(clark, str_value)
            %   DOCX DELTAS vs pptx (see class header D-delta-1/-2): (1) the
            %   explicit `value is None` short-circuit -- removes even when the
            %   default is NON-None; (2) the post-to_xml `str_value is None`
            %   removal. isequal(.,[]) is the `is None` analogue (H3: "" is a real
            %   string, isequal("",[]) is false, so "" is stored, not removed).
            arguments
                obj (1,1) mat2doc.oxml.BaseOxmlElement
                name (1,1) string
                type (1,1) string
                value
                default = []
            end
            clark = obj.attrClarkName_(name);
            % Python: if value is None or value == self._default   (docx line 203)
            if isequal(value, []) || isequal(value, default)
                if obj.has_attrib(clark)    % Python: if clark in obj.attrib
                    obj.remove_attrib(clark);
                end
                return
            end
            str_value = feval(obj.resolveTypeCls_(type) + ".to_xml", value);
            % Python: if str_value is None: (del if present) return   (docx lines 208-211)
            if isequal(str_value, [])
                if obj.has_attrib(clark)
                    obj.remove_attrib(clark);
                end
                return
            end
            obj.set(clark, str_value);
        end

        function value = getAttrRequired(obj, name, type)
            % GETATTRREQUIRED Type-converted required attribute value; raises when absent (RequiredAttribute getter).
            %   Ported from xmlchemy.py RequiredAttribute._getter's
            %   `get_attr_value` (lines 240-246): attr = obj.get(clark); if None
            %   raise InvalidXmlError("required '%s' attribute not present on
            %   element %s" % (attr_name, obj.tag)); else simple_type.from_xml.
            %   Identical in docx and pptx (no delta). No default -- a missing
            %   required attribute is a malformed-XML error; obj.tag is the Clark
            %   name (XmlElement get.tag), matching lxml's obj.tag in the message.
            %   Enum-aware via resolveTypeCls_.
            arguments
                obj (1,1) mat2doc.oxml.BaseOxmlElement
                name (1,1) string
                type (1,1) string
            end
            attr = obj.get(obj.attrClarkName_(name));   % [] (None) when absent (H3)
            if isequal(attr, [])                        % Python: if attr_str_value is None
                error("mat2doc:InvalidXmlError", ...
                    "required '%s' attribute not present on element %s", name, obj.tag);
            end
            value = feval(obj.resolveTypeCls_(type) + ".from_xml", attr);
        end

        function setAttrRequired(obj, name, type, value)
            % SETATTRREQUIRED Set a required typed attribute; never removes (RequiredAttribute setter).
            %   Ported from xmlchemy.py RequiredAttribute._setter's
            %   `set_attr_value` (lines 255-259, DOCX form):
            %       str_value = simple_type.to_xml(value)
            %       if str_value is None:
            %           raise ValueError(f"cannot assign {value} to this required attribute")
            %       obj.set(clark, str_value)
            %   DOCX DELTA vs pptx (see class header D-delta-3): the
            %   `str_value is None` ValueError guard (pptx sets unconditionally).
            %   A required attribute has NO default, so this ALWAYS to_xml's and
            %   sets -- it never removes. VERIFY (audit): the "{value}" repr in
            %   the ValueError message follows MATLAB string(value), which may
            %   differ from CPython str(value) for non-string values (message-text
            %   only; error path; carried under the D-005 message-token class).
            arguments
                obj (1,1) mat2doc.oxml.BaseOxmlElement
                name (1,1) string
                type (1,1) string
                value
            end
            str_value = feval(obj.resolveTypeCls_(type) + ".to_xml", value);
            if isequal(str_value, [])   % Python: if str_value is None (docx line 257)
                error("mat2doc:ValueError", ...
                    "cannot assign %s to this required attribute", ...
                    mat2doc.oxml.BaseOxmlElement.valueRepr_(value));
            end
            obj.set(obj.attrClarkName_(name), str_value);
        end
    end

    % =====================================================================
    % P1-3b CHILD-ELEMENT descriptor ENGINE
    %
    % Generic operations the generated CT_* child-descriptor delegating members
    % call. Each is the faithful port of one xmlchemy descriptor method body
    % (see the engine-method -> source map in the class header). `tag` args are
    % namespace-prefixed tagnames (e.g. "w:pPr"); `successors` / `tags` /
    % `groupTags` args are string arrays of prefixed tagnames drawn from a
    % class's Constant SUCCESSORS / group-members table.
    %
    % OVERRIDE NOTE (design.md section 2 porter alert; 8 docx override/inherit
    % flag instances): where a CT_* class defines an explicit `_new_x`/`_insert_x`
    % (Python `hasattr` guard in _add_to_class makes it WIN over the generated
    % one), that class's add_x_/get_or_add_x/get_or_change_to_x members must
    % route through the class's OWN new_x_/insert_x_ rather than the generic
    % addChild below -- the generic path here always uses the DEFAULT creator
    % (newChild). Those override spots are hand-ported in the affected classes.
    %
    % H11 (child ordering): insertChildInSequence -> insert_element_before ->
    % first_child_found_in scans the successors in ARGUMENT order and inserts
    % addprevious the FIRST present successor tag (else append). The successor
    % slice logic lives in the P1-3a tree-ops (insert_element_before,
    % first_child_found_in); this engine merely expands the class SUCCESSORS
    % Constant into their repeating tagname args.
    % =====================================================================
    methods
        function child = getChild(obj, tag)
            % GETCHILD Child with prefixed tag, or [] (None) if absent (ZeroOrOne / Choice getter).
            %   Ported from xmlchemy.py _BaseChildElement._getter's
            %   get_child_element (docx lines 380-382): `return obj.find(qn(nsptag))`.
            %   Also serves the Choice getter (Choice adds _add_getter, same body).
            arguments
                obj (1,1) mat2doc.oxml.BaseOxmlElement
                tag (1,1) string
            end
            child = obj.find(mat2doc.oxml.qn(tag));   % [] when absent (H3)
        end

        function child = getRequiredChild(obj, tag)
            % GETREQUIREDCHILD Child with prefixed tag; error if absent (OneAndOnlyOne getter).
            %   Ported from xmlchemy.py OneAndOnlyOne._getter's get_child_element
            %   (docx lines 499-505): find; if None raise InvalidXmlError. The
            %   literal RST double-backticks in the message are reproduced
            %   verbatim; id mat2doc:InvalidXmlError.
            arguments
                obj (1,1) mat2doc.oxml.BaseOxmlElement
                tag (1,1) string
            end
            child = obj.find(mat2doc.oxml.qn(tag));
            if isequal(child, [])   % Python: if child is None (H3)
                error("mat2doc:InvalidXmlError", ...
                    "required ``<%s>`` child element not present", tag);
            end
        end

        function list = getChildList(obj, tag)
            % GETCHILDLIST All children with prefixed tag, in document order (ZeroOrMore/OneOrMore x_lst).
            %   Ported from xmlchemy.py _BaseChildElement._list_getter's
            %   get_child_element_list (docx lines 397-398):
            %   `return obj.findall(qn(nsptag))`. Returns a (1,N) (possibly 1x0)
            %   XmlElement array -- Python returns a LIST, so "none present" is an
            %   EMPTY ARRAY (materialized, H9), NOT [] (None, H3). (H1: findall is
            %   already 1-based document order; no index shift.)
            arguments
                obj (1,1) mat2doc.oxml.BaseOxmlElement
                tag (1,1) string
            end
            list = obj.findall(mat2doc.oxml.qn(tag));
        end

        function child = newChild(obj, tag)
            % NEWCHILD Loose child element of the correct type, no attrs (default creator _new_x).
            %   Ported from xmlchemy.py _BaseChildElement._creator's
            %   new_child_element (docx lines 366-367): `return OxmlElement(nsptag)`.
            %   Applies the registry (registered CT_* class or plain XmlElement
            %   fallback), exactly as the Python creator does. obj is unused
            %   (matches Python: the default creator ignores obj) but kept for
            %   method-call form.
            arguments
                obj (1,1) mat2doc.oxml.BaseOxmlElement %#ok<INUSA>
                tag (1,1) string
            end
            child = mat2doc.oxml.OxmlElement(tag);
        end

        function child = insertChildInSequence(obj, child, successors)
            % INSERTCHILDINSEQUENCE Insert child before the first present successor tag; return child.
            %   Ported from xmlchemy.py _BaseChildElement._add_inserter's
            %   _insert_child (docx lines 319-321):
            %     obj.insert_element_before(child, *self._successors); return child
            %   successors is a class SUCCESSORS Constant (string array); it is
            %   expanded into the repeating tagname args of insert_element_before
            %   (the P1-3a tree-op that carries the H11 successor-slice logic).
            arguments
                obj (1,1) mat2doc.oxml.BaseOxmlElement
                child (1,1) mat2doc.oxml.XmlElement
                successors (1,:) string
            end
            sc = num2cell(successors);
            child = obj.insert_element_before(child, sc{:});
        end

        function child = addChild(obj, tag, successors, varargin)
            % ADDCHILD Create a child (default creator), set attrs, insert in sequence (_add_x).
            %   Ported from xmlchemy.py _BaseChildElement._add_adder's
            %   _add_child(obj, **attrs) (docx lines 284-291): new_method(); for
            %   key,value in attrs: setattr(child, key, value); insert; return.
            %   attrs arrive as trailing name-value pairs (Python **attrs);
            %   `child.(name) = value` is the setattr analogue (typed-property
            %   set). This generic path uses the DEFAULT creator -- classes with a
            %   `_new_x` override do NOT delegate here (see OVERRIDE NOTE). Also
            %   the primitive the public add_x() adder routes to (both OneOrMore
            %   and, in docx, ZeroOrMore -- D-delta-4).
            arguments
                obj (1,1) mat2doc.oxml.BaseOxmlElement
                tag (1,1) string
                successors (1,:) string
            end
            arguments (Repeating)
                varargin
            end
            child = obj.newChild(tag);                 % default creator
            for k = 1:2:numel(varargin)                % Python: setattr per attr
                child.(varargin{k}) = varargin{k + 1};
            end
            child = obj.insertChildInSequence(child, successors);
        end

        function child = getOrAddChild(obj, tag, successors)
            % GETORADDCHILD Return the child, adding it (in sequence) if absent (get_or_add_x).
            %   Ported from xmlchemy.py ZeroOrOne._add_get_or_adder's
            %   get_or_add_child (docx lines 557-562): child = getter; if None:
            %   child = _add_x(). Generic (default-creator) path. H5: when the
            %   child is already present, the SAME handle is returned on every
            %   call (getChild -> find returns the live node, not a copy).
            arguments
                obj (1,1) mat2doc.oxml.BaseOxmlElement
                tag (1,1) string
                successors (1,:) string
            end
            child = obj.getChild(tag);
            if isequal(child, [])   % Python: if child is None (H3)
                child = obj.addChild(tag, successors);
            end
        end

        function removeChild(obj, tag)
            % REMOVECHILD Remove ALL children with prefixed tag (_remove_x).
            %   Ported from xmlchemy.py ZeroOrOne._add_remover's _remove_child
            %   (docx lines 572-573): `obj.remove_all(nsptagname)` -- removes
            %   EVERY matching child, not just the first (faithful to remove_all).
            arguments
                obj (1,1) mat2doc.oxml.BaseOxmlElement
                tag (1,1) string
            end
            obj.remove_all(tag);
        end

        function child = firstChildFoundIn(obj, tags)
            % FIRSTCHILDFOUNDIN First child whose tag is in tags, else [] (choice-group getter).
            %   Ported from xmlchemy.py ZeroOrOneChoice._choice_getter's
            %   get_group_member_element (docx lines 622-623):
            %   `return obj.first_child_found_in(*member_nsptagnames)`.
            %   tags is the class's group-members Constant (string array). Search
            %   is by ARGUMENT order (first NAME that matches), the P1-3a
            %   first_child_found_in semantics.
            arguments
                obj (1,1) mat2doc.oxml.BaseOxmlElement
                tags (1,:) string
            end
            tc = num2cell(tags);
            child = obj.first_child_found_in(tc{:});
        end

        function removeChildren(obj, tags)
            % REMOVECHILDREN Remove whichever of the group-member tags are present (_remove_eg_x).
            %   Ported from xmlchemy.py ZeroOrOneChoice._add_group_remover's
            %   _remove_choice_group (docx lines 610-612): for tagname in members:
            %   obj.remove_all(tagname). remove_all already accepts many tags, so
            %   one call over the whole group is the faithful equivalent.
            arguments
                obj (1,1) mat2doc.oxml.BaseOxmlElement
                tags (1,:) string
            end
            tc = num2cell(tags);
            obj.remove_all(tc{:});
        end

        function child = getOrChangeToChild(obj, tag, groupTags, successors)
            % GETORCHANGETOCHILD Return this group member, replacing any other present member (get_or_change_to_x).
            %   Ported from xmlchemy.py Choice._add_get_or_change_to_method's
            %   get_or_change_to_child (docx lines 453-461): child = getter; if
            %   not None return; remove_group(); child = _add_x(). Generic
            %   (default-creator) path; override members hand-route (see NOTE).
            %   groupTags is the choice group's member Constant (the _remove_eg_x
            %   removal set); tag is THIS choice's tag; successors is the group's
            %   SUCCESSORS Constant.
            arguments
                obj (1,1) mat2doc.oxml.BaseOxmlElement
                tag (1,1) string
                groupTags (1,:) string
                successors (1,:) string
            end
            child = obj.getChild(tag);
            if ~isequal(child, [])   % Python: if child is not None (H3)
                return
            end
            obj.removeChildren(groupTags);
            child = obj.addChild(tag, successors);
        end
    end

    methods (Access = private)
        function cls = resolveTypeCls_(~, type)
            % RESOLVETYPECLS_ Fully-qualified class name for a schema TYPE token.
            %   xmlchemy attribute descriptors hold the actual simple-type OR
            %   enum type OBJECT; the port stores the type's NAME string in the
            %   class's Constant schema table and resolves it to a class here.
            %   Two currencies (backward compatible):
            %     * a BARE short name ("ST_String") is a simple type in
            %       +oxml/+simpletypes -- the default; prefix applied.
            %     * a name that ALREADY carries a package ("." present), e.g.
            %       "mat2doc.enum.WD_UNDERLINE", is used VERBATIM -- how enum
            %       simple-types dispatch to +enum. Both surfaces expose
            %       from_xml/to_xml statics with (xml_value)/(value) signatures,
            %       so the feval dispatch in getAttr*/setAttr* is uniform.
            %   NOTE (P1-3a): +oxml/+simpletypes is P3 and does not exist yet;
            %   this is descriptor PLUMBING that CALLS a converter (design.md
            %   section scope). No CT_* class delegates here in this slice, so the
            %   dispatch is dormant until the ST_* types land.
            if contains(type, ".")
                cls = type;                                   % already fully qualified (enum)
            else
                cls = "mat2doc.oxml.simpletypes." + type;     % simple-type short name (default)
            end
        end

        function clark = attrClarkName_(~, name)
            % ATTRCLARKNAME_ Clark key for an attribute name (BaseAttribute._clark_name, xmlchemy.py 139-143).
            %   Prefixed name ("w:val") -> qn(name); plain name ("space") -> name.
            if contains(name, ":")
                clark = mat2doc.oxml.qn(name);
            else
                clark = name;
            end
        end
    end

    methods (Static, Access = private)
        function s = valueRepr_(value)
            % VALUEREPR_ Best-effort str(value) for the RequiredAttribute
            %   ValueError message (docx-delta-3). Faithful for string/char/
            %   numeric scalars; falls back to the class token otherwise. See the
            %   VERIFY note on setAttrRequired -- error-path message text only.
            if isequal(value, [])
                s = "None";
            elseif isstring(value) && isscalar(value) && ~ismissing(value)
                s = value;
            elseif ischar(value) && (isrow(value) || isempty(value))
                s = string(value);
            elseif (isnumeric(value) || islogical(value)) && isscalar(value)
                s = string(value);
            else
                s = "<" + class(value) + ">";
            end
        end
    end
end

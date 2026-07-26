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
%   This slice ports ONLY:
%     * the tree-ops on BaseOxmlElement (xmlchemy.py lines 677-698):
%       first_child_found_in, insert_element_before, remove_all;
%     * xpath() (xmlchemy.py lines 708-713), delegating to the mini-XPath
%       engine +oxml/evaluate_xpath.m (re-ported from Mat2Ppt WP5+WP5-C);
%     * the ATTRIBUTE descriptor engine (OptionalAttribute / RequiredAttribute,
%       xmlchemy.py lines 177-274): getAttrTyped / setAttrTyped / getAttrRequired
%       / setAttrRequired + the type-dispatch helpers.
%
%   NOT in this slice -- the CHILD-element descriptor engine (ZeroOrOne,
%   ZeroOrMore, OneAndOnlyOne, OneOrMore, Choice, ZeroOrOneChoice: getChild /
%   getOrAddChild / newChild / insertChildInSequence / addChild / removeChild /
%   firstChildFoundIn / removeChildren / getOrChangeToChild) is P1-3b. It is
%   NOT ported here even though the Mat2Ppt reference carries it in one class.
%
%   ENGINE-METHOD -> xmlchemy SOURCE MAP (src/docx/oxml/xmlchemy.py):
%     first_child_found_in  <- BaseOxmlElement.first_child_found_in (677-683)
%     insert_element_before <- BaseOxmlElement.insert_element_before (685-691)
%     remove_all            <- BaseOxmlElement.remove_all           (693-698)
%     xpath                 <- BaseOxmlElement.xpath                (708-713)
%     getAttrTyped          <- OptionalAttribute._getter get_attr_value (205-209)
%     setAttrTyped          <- OptionalAttribute._setter set_attr_value (218-226)
%     getAttrRequired       <- RequiredAttribute._getter get_attr_value (244-250)
%     setAttrRequired       <- RequiredAttribute._setter set_attr_value (270-274)
%     resolveTypeCls_       <- BaseAttribute simple_type object       (137-139)
%     attrClarkName_        <- BaseAttribute._clark_name              (160-164)
%
%   DOCX-vs-PPTX DELTAS (docx is source of truth; ported to the docx forms,
%   NOT the Mat2Ppt/pptx forms -- see audit_P1-3a_xmlchemy_engine.md):
%     D-delta-1  OptionalAttribute setter guards `value is None OR value ==
%                default` (docx line 218-224). pptx guards only `value ==
%                default`. When an optional attribute has a NON-None default,
%                assigning None removes it in docx but would fall through to
%                to_xml(None) in pptx.
%     D-delta-2  OptionalAttribute setter, AFTER to_xml, removes the attribute
%                when `str_value is None` (docx line 221-224). pptx unconditionally
%                sets. Lets a simple type whose to_xml returns None erase the attr.
%     D-delta-3  RequiredAttribute setter, AFTER to_xml, RAISES
%                ValueError("cannot assign {value} to this required attribute")
%                when `str_value is None` (docx line 271-272). pptx sets
%                unconditionally.
%
%   Deferred, in their owning WPs (NOT ported here):
%     child-element descriptor engine (getChild ...)  -> P1-3b
%     xml property (serialize_for_reading, 700-706)    -> the doc-serialize/OPC
%           WP (pretty-print test helper; design.md section 3 forbids pretty-print
%           in the PART serializer, so serialize_for_reading is a distinct
%           helper). No docx PRODUCTION code path reads xmlchemy BaseOxmlElement
%           `.xml` (the production `.xml` at opc/rel.py is the SEPARATE opc/oxml.py
%           base class); it is a test-only helper. Provided as a clean
%           notYetPorted stub so an accidental caller gets a named error, NOT a
%           missing-method error. Mirrors the Mat2Ppt deferral of xml to WP4.
%     __repr__ (670-675) and _nsptag (715-717)         -> display-only; used ONLY
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
%   (tree-ops + xpath slice; metaclass/descriptor machinery replaced by the
%   explicit delegating-member scheme per design.md section 2; attribute
%   descriptor engine from OptionalAttribute/RequiredAttribute)

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
        % TREE-OPS (xmlchemy.py BaseOxmlElement, lines 677-698)
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
            %   BaseOxmlElement.first_child_found_in (lines 677-683)
            arguments
                obj (1,1) mat2doc.oxml.BaseOxmlElement
            end
            arguments (Repeating)
                tagnames (1,1) string
            end
            for k = 1:numel(tagnames)
                % Python: child = self.find(qn(tagname))  (line 680)
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
            %   BaseOxmlElement.insert_element_before (lines 685-691)
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
            %   BaseOxmlElement.remove_all (lines 693-698)
            arguments
                obj (1,1) mat2doc.oxml.BaseOxmlElement
            end
            arguments (Repeating)
                tagnames (1,1) string
            end
            for k = 1:numel(tagnames)
                % Python: matching = self.findall(qn(tagname))  (line 696)
                % H9: findall is already a materialized list in Python, so
                % removing while looping over it is safe in both languages.
                matching = obj.findall(mat2doc.oxml.qn(tagnames{k}));
                for child = matching
                    obj.remove(child);
                end
            end
        end

        % =================================================================
        % XPATH (xmlchemy.py BaseOxmlElement.xpath, lines 708-713)
        % =================================================================

        function result = xpath(obj, xpath_str, namespaces)
            % XPATH Evaluate an XPath expression over this element's tree.
            %
            %   nodes = e.XPATH(expr) evaluates expr (a string) against e using
            %   the fixed WordprocessingML namespace map (mat2doc.oxml.nsmap =
            %   Python `nsmap`), mirroring
            %   docx.oxml.xmlchemy.BaseOxmlElement.xpath (xmlchemy.py:708-713),
            %   which centralizes the namespace mapping. NOTE (docx-vs-pptx): docx
            %   injects the PUBLIC `nsmap`; pptx injects the private `_nsmap`.
            %
            %   nodes = e.XPATH(expr, ns) resolves prefixes against the scalar
            %   struct ns (prefix -> URI) instead. python-docx's own override
            %   drops lxml's `namespaces` kwarg and always injects `nsmap`, so no
            %   upstream call site passes a custom map; this parameter defaults to
            %   that same fixed map, making `e.xpath(expr)` byte-for-byte the
            %   upstream behavior.
            %
            %   Return type mirrors lxml exactly (H3 -- empty match is an EMPTY
            %   ARRAY of the correct type, decided from the AST terminal step,
            %   NEVER [] (None); callers use numel/~isempty/arr(1)):
            %     - expr ending in /@attr  -> (1,N) string  (attribute values),
            %                                 string.empty(1,0) on no match
            %     - expr ending in /text() -> (1,N) string  (text nodes),
            %                                 string.empty(1,0) on no match
            %     - otherwise              -> (1,N) mat2doc.oxml.XmlElement,
            %                                 XmlElement.empty(1,0) on no match
            %
            %   Supported subset and error contract: see +oxml/evaluate_xpath.m.
            %   Anything outside the subset raises mat2doc:XPathError rather than
            %   being silently mis-evaluated (design.md section XPath / 7).
            %
            %   Ported from python-docx v1.2.0: src/docx/oxml/xmlchemy.py::
            %   BaseOxmlElement.xpath (lines 708-713)
            arguments
                obj (1,1) mat2doc.oxml.BaseOxmlElement
                xpath_str (1,1) string
                namespaces (1,1) struct = mat2doc.oxml.nsmap()
            end
            result = mat2doc.oxml.evaluate_xpath(obj, xpath_str, namespaces);
        end

        function s = xml(obj) %#ok<STOUT,INUSD>
            % XML DEFERRED test-only serialize_for_reading (pretty-printed, no decl).
            %
            %   Python `BaseOxmlElement.xml` (@property, xmlchemy.py 700-706)
            %   returns serialize_for_reading(self) -- etree.tostring with
            %   pretty_print=True and no XML declaration, wrapped in XmlString
            %   (a test-comparison helper). No docx PRODUCTION code path reads
            %   this property on a wordprocessing element (grep-verified; the
            %   production `.xml` at opc/rel.py is the SEPARATE opc/oxml.py base
            %   class). It is deferred to the doc-serialize/OPC WP because the
            %   pretty-printer is a distinct helper from the byte-exact PART
            %   serializer (design.md section 3 forbids pretty-print there).
            %
            %   Clean stub per design.md ("dependencies not yet ported: clean
            %   stubs raising mat2doc:notYetPorted naming the target symbol and
            %   its owning WP").
            error("mat2doc:notYetPorted", ...
                "%s", "BaseOxmlElement.xml (serialize_for_reading, " + ...
                "src/docx/oxml/xmlchemy.py::serialize_for_reading) is deferred " + ...
                "to the doc-serialize/OPC WP");
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
            %   `get_attr_value` (lines 205-209): attr = obj.get(clark); if None
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
            %   `set_attr_value` (lines 218-226, DOCX form):
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
            % Python: if value is None or value == self._default   (docx line 218)
            if isequal(value, []) || isequal(value, default)
                if obj.has_attrib(clark)    % Python: if clark in obj.attrib
                    obj.remove_attrib(clark);
                end
                return
            end
            str_value = feval(obj.resolveTypeCls_(type) + ".to_xml", value);
            % Python: if str_value is None: (del if present) return   (docx line 221)
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
            %   `get_attr_value` (lines 244-250): attr = obj.get(clark); if None
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
            %   `set_attr_value` (lines 270-274, DOCX form):
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
            if isequal(str_value, [])   % Python: if str_value is None (docx line 271)
                error("mat2doc:ValueError", ...
                    "cannot assign %s to this required attribute", ...
                    mat2doc.oxml.BaseOxmlElement.valueRepr_(value));
            end
            obj.set(obj.attrClarkName_(name), str_value);
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
            % ATTRCLARKNAME_ Clark key for an attribute name (BaseAttribute._clark_name, xmlchemy.py 160-164).
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

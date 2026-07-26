classdef XmlElement < handle & matlab.mixin.Heterogeneous
% XMLELEMENT XML element node -- the toolbox's lxml `_Element` replacement.
%
%   e = MAT2DOC.OXML.XMLELEMENT(nsptag) creates a loose element with the
%   prefixed tag nsptag (e.g. "w:p") and no namespace declarations.
%
%   e = MAT2DOC.OXML.XMLELEMENT(nsptag, nsmap) additionally records the
%   namespace declarations in scalar struct nsmap (<prefix> -> URI, field order
%   preserved, H11) on the element, mirroring lxml
%   `parser.makeelement(clark_name, nsmap=...)`.
%
%   DESIGN-REALIZATION (design.md section 3, D-001): this class is not the port
%   of a single Python symbol; it realizes the lxml `etree._Element` surface
%   that python-docx v1.2.0 actually uses, with every semantic matched to lxml
%   5.3.0 probe evidence. The design is the SOLVED Mat2Ppt +oxml reference,
%   re-ported here (no shared code); docx v1.2.0 is the source of truth for the
%   module contents.
%
%   Tree shape (design.md section 3):
%     - tag identity: prefixed tag as written (nsptag_str, e.g. "w:p") plus
%       resolved namespace URI; the `tag` property presents the lxml Clark-form
%       string "{uri}local" so ported call sites comparing against qn(...) work
%       verbatim.
%     - ordered attribute list: insertion-ordered name/value string pairs; names
%       keyed exactly as lxml stores them -- Clark form for namespaced
%       attributes, plain local name otherwise.
%     - ordered heterogeneous children (this class derives
%       matlab.mixin.Heterogeneous so CT_* subclasses mix in one vector).
%     - text/tail: string scalar or [] (None, H3 tri-state; "" is NOT None).
%     - parent backref; ordered namespace-declaration list (prefix/URI pairs;
%       prefix "" reserved for a default xmlns declaration).
%
%   ELEMENT IDENTITY = HANDLE IDENTITY (H5): the wrapper IS the node; `==`
%   replicates lxml same-element identity. Elements are never copied implicitly;
%   append/insert/addnext/addprevious MOVE an element out of its old parent, and
%   its text/tail travel with it.
%
%   Iteration/len dunder mapping (design.md section 2):
%     len(e)        -> numel(e.to_array())
%     for c in e    -> for c = e.to_array()
%     e[i]          -> arr = e.to_array(); arr(i+1)   % IDX
%     k in e.attrib -> e.has_attrib(k)
%     del e.attrib[k] -> e.remove_attrib(k)
%
%   INDEXING (H1): insert() and index() are 1-BASED on this MATLAB surface.
%   Ported Python call sites must shift: `elm.insert(i, c)` -> `elm.insert(i+1, c)`
%   and data uses of `parent.index(child)` need `-1`. See method help.
%
%   TYPE ERRORS (D-005): a non-text value passed to set(name, value) or assigned
%   to text/tail raises mat2doc:TypeError whose message reports the MATLAB type
%   token (e.g. 'double') where lxml reports 'int'; the exception class and the
%   message template ("Argument must be bytes or unicode, got '<type>'") are
%   faithful, only the token differs -- a dead-path, API-invisible divergence
%   (deviation ledger D-005, adopted for Mat2Doc).
%
%   XPATH (task #60 hoist, P2-2): every lxml `_Element` has `.xpath`, so the
%   method lives HERE on XmlElement (the _Element analogue), NOT only on the
%   BaseOxmlElement subclass. python-docx overrides `BaseOxmlElement.xpath`
%   (xmlchemy.py 687-692) merely to inject the fixed `nsmap` and drop lxml's
%   `namespaces` kwarg -- it does not ADD the capability; the underlying
%   lxml._Element.xpath is universal. At P1-3 the method was placed on
%   BaseOxmlElement only; #60 HOISTS it to XmlElement so the parser-fallback root
%   class (an UNREGISTERED root, e.g. the not-yet-registered `w:document` before
%   P2-3, or any generic parsed element) also supports xpath -- required by
%   XmlPart.rel_ref_count_ (`element.xpath("//@r:id")`, opc/part.py 246) and
%   StoryPart.next_id (`element.xpath("//@id")`, parts/story.py 84), whose
%   `element` is a plain XmlElement until the corresponding CT_* is registered.
%   Byte-neutral (adds a method; changes no serialization). Both spellings return
%   IDENTICALLY -- BaseOxmlElement inherits this method verbatim.
%
%   Example:
%       p = mat2doc.oxml.XmlElement("w:p");
%       r = mat2doc.oxml.XmlElement("w:r");
%       p.append(r);                                   % append MOVES r under p
%       tf = (p.find(mat2doc.oxml.qn("w:r")) == r);    % identity is true (H5)
%
%   Design-realization of the lxml _Element surface used by python-docx v1.2.0
%   (see design.md section 3).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/parser.py::oxml_parser
%   (lxml _Element surface; design-realization, D-001)

    properties (Dependent, SetAccess = private)
        tag         % Clark-form tag string, "{uri}local" (or bare local part when the element has no namespace) -- exactly what lxml `.tag` returns, so comparisons against qn(...) port verbatim
        nsptag_str  % the prefixed tag as written, e.g. "w:p" (serializer-side identity; "" prefix possible for default-ns elements)
        local_part  % local part of the tag, e.g. "p"
        nspfx       % namespace prefix, e.g. "w"; "" when the tag has none
        nsuri       % namespace URI resolved for this element's tag; "" when none
        nsdecls     % Nx2 string array [prefix, URI] of namespace declarations stored on THIS element, in insertion order (prefix "" = default xmlns)
    end

    properties (Dependent)
        text        % element text: string scalar or [] (None). Tri-state per H3: "" is a real empty string, distinct from [] (lxml serializes text="" as <t></t>, text=None as <t/>)
        tail        % text following this element's end tag, before the next sibling: string scalar or [] (None). Travels with the element on moves
    end

    properties (Access = private)
        nsptag_ (1,1) string = ""        % prefixed tag as written
        pfx_ (1,1) string = ""           % tag prefix ("" if none)
        local_ (1,1) string = ""         % tag local part
        nsuri_ (1,1) string = ""         % resolved tag namespace URI ("" if none)
        attr_names_ (1,:) string = strings(1, 0)   % attribute keys, lxml keying (Clark or plain), insertion order
        attr_values_ (1,:) string = strings(1, 0)  % parallel attribute values
        children_ (1,:) mat2doc.oxml.XmlElement    % ordered children (heterogeneous handle vector)
        parent_ = []                     % parent XmlElement handle, or [] (None) when detached/root
        nsdecls_ (:,2) string = strings(0, 2)      % ordered namespace declarations [prefix, URI]
        text_ = []                       % string scalar or [] (None)
        tail_ = []                       % string scalar or [] (None)
        nsVerbatim_ (1,1) logical = false % true iff this element's stored ns-decls must be serialized VERBATIM (never suppressed). Set by the parser (markNsVerbatim_); CLEARED when a public move (append/insert/addnext/addprevious) reconciles the element into a new scope. Default false = subject to serialize-time suppression. See serialize_part_xml step 1
    end

    methods
        function obj = XmlElement(nsptag, nsmap, resolvedUri)
            % XMLELEMENT Construct a loose element (lxml makeelement analogue).
            %
            %   Inputs:  nsptag - (1,1) string, prefixed tag, e.g. "w:p"
            %            nsmap  - optional namespace declarations to store on the
            %                     element, in EITHER currency:
            %                       (1,1) struct: <prefix> -> URI, field order
            %                             preserved (H11); or
            %                       (N,2) string: [prefix, URI] pairs in insertion
            %                             order. This is the decl currency used by
            %                             the parser; it is REQUIRED to express a
            %                             DEFAULT namespace declaration (prefix
            %                             ""), which a struct field cannot encode,
            %                             e.g. <Types xmlns="...">.
            %                     Default: none.
            %            resolvedUri - optional (1,1) string: the namespace URI the
            %                     parser resolved for this element's tag from the
            %                     in-scope declaration stack. When supplied (not
            %                     <missing>), it is used verbatim as the tag URI,
            %                     BYPASSING the own-decls-then-fixed-map resolution
            %                     below -- REQUIRED for elements whose prefix (or
            %                     default xmlns) is declared on an ANCESTOR, or
            %                     whose prefix is not in the fixed nsmap (e.g.
            %                     document.xml's mc/wp14/wps). Default: <missing>.
            %   Outputs: obj - scalar XmlElement handle
            arguments
                nsptag (1,1) string
                nsmap = struct()
                resolvedUri (1,1) string = string(missing)
            end
            parts = split(nsptag, ":");
            if numel(parts) == 1
                obj.pfx_ = "";
                obj.local_ = parts(1);
            elseif numel(parts) == 2
                obj.pfx_ = parts(1);
                obj.local_ = parts(2);
            else
                error("mat2doc:ValueError", "Invalid tag name '%s'", nsptag);
            end
            obj.nsptag_ = nsptag;
            % -- namespace declarations: struct fields in order (H11), or the
            %    (N,2) string decl pairs verbatim --
            if isstruct(nsmap) && isscalar(nsmap)
                f = string(fieldnames(nsmap));
                decls = strings(numel(f), 2);
                for k = 1:numel(f)
                    decls(k, :) = [f(k), string(nsmap.(f(k)))];
                end
            elseif isstring(nsmap) && ismatrix(nsmap) && ...
                    (size(nsmap, 2) == 2 || isempty(nsmap)) && ~any(ismissing(nsmap(:)))
                decls = reshape(nsmap, [], 2);
            else
                error("mat2doc:TypeError", ...
                    "nsmap must be a scalar struct or an Nx2 string array");
            end
            obj.nsdecls_ = decls;
            % -- parser path: authoritative in-scope-resolved URI supplied --
            if ~ismissing(resolvedUri)
                obj.nsuri_ = resolvedUri;
                return
            end
            % -- resolve tag namespace URI: own decls first, then fixed map --
            % (For the OxmlElement factory path nsmap is always a subset of the
            % fixed map, so precedence is unobservable there; own-decls-first is
            % required for the parser path, where document declarations may bind
            % arbitrary prefixes.)
            if obj.pfx_ ~= ""
                i = find(decls(:, 1) == obj.pfx_, 1);
                if ~isempty(i)
                    obj.nsuri_ = decls(i, 2);
                else
                    fixed = mat2doc.oxml.nsmap();
                    if ~isfield(fixed, obj.pfx_)
                        error("mat2doc:KeyError", "'%s'", obj.pfx_);
                    end
                    obj.nsuri_ = fixed.(obj.pfx_);
                end
            else
                % Unprefixed tag: in the namespace of an OWN default ("")
                % declaration when one is given (XML default-ns semantics, e.g.
                % <Types xmlns="...">); otherwise no namespace. In-scope
                % (ancestor) default-ns resolution at parse time is the parser's
                % job.
                i = find(decls(:, 1) == "", 1);
                if ~isempty(i)
                    obj.nsuri_ = decls(i, 2);
                else
                    obj.nsuri_ = "";
                end
            end
        end

        % ------------------------------------------------------------------
        % attributes (lxml .get / .set / .attrib)
        % ------------------------------------------------------------------

        function value = get(obj, name, default)
            % GET Attribute value by name, or default when absent.
            %
            %   value = e.GET(name) returns the attribute value (string) or []
            %   (None) when no such attribute exists.
            %   value = e.GET(name, default) returns default when absent.
            %
            %   name uses lxml keying: Clark form (qn("w:val")) for namespaced
            %   attributes, plain local name ("val") otherwise. A prefixed name
            %   like "w:val" can never match (lxml returns None for it); no prefix
            %   translation is performed.
            arguments
                obj (1,1) mat2doc.oxml.XmlElement
                name (1,1) string
                default = []
            end
            i = find(obj.attr_names_ == name, 1);
            if isempty(i)
                value = default;
            else
                value = obj.attr_values_(i);
            end
        end

        function set(obj, name, value)
            % SET Set attribute name to value (string), lxml keying.
            %
            %   e.SET(name, value): name is Clark form for namespaced attributes
            %   or a plain local name. A prefixed name such as "w:val" raises
            %   mat2doc:ValueError exactly as lxml does (ValueError "Invalid
            %   attribute name 'w:val'"). A non-text value raises mat2doc:TypeError.
            %   Setting an existing attribute overwrites its value in place --
            %   insertion position is retained.
            arguments
                obj (1,1) mat2doc.oxml.XmlElement
                name (1,1) string
                value
            end
            if ismissing(name)
                error("mat2doc:TypeError", "Argument must be bytes or unicode, got 'NoneType'");
            end
            if ~((isstring(value) && isscalar(value) && ~ismissing(value)) || ...
                    (ischar(value) && (isrow(value) || isempty(value))))
                error("mat2doc:TypeError", ...
                    "Argument must be bytes or unicode, got '%s'", class(value));
            end
            value = string(value);
            % lxml name-form rule: Clark names accepted verbatim (URI need not be
            % in the fixed map); prefixed names rejected; plain names accepted.
            % (Full XML NCName character validation is NOT ported -- docx never
            % passes an invalid name.)
            if ~startsWith(name, "{") && contains(name, ":")
                error("mat2doc:ValueError", "Invalid attribute name '%s'", name);
            end
            i = find(obj.attr_names_ == name, 1);
            if isempty(i)
                obj.attr_names_(end + 1) = name;
                obj.attr_values_(end + 1) = value;
            else
                obj.attr_values_(i) = value;   % position retained
            end
        end

        function tf = has_attrib(obj, name)
            % HAS_ATTRIB True when attribute name is present.
            %
            %   Dunder mapping: Python `key in element.attrib` ports as
            %   `element.has_attrib(key)`.
            arguments
                obj (1,1) mat2doc.oxml.XmlElement
                name (1,1) string
            end
            tf = any(obj.attr_names_ == name);
        end

        function remove_attrib(obj, name)
            % REMOVE_ATTRIB Delete attribute name.
            %
            %   Dunder mapping: Python `del element.attrib[key]` ports as
            %   `element.remove_attrib(key)`. Missing key raises mat2doc:KeyError
            %   exactly as Python.
            arguments
                obj (1,1) mat2doc.oxml.XmlElement
                name (1,1) string
            end
            i = find(obj.attr_names_ == name, 1);
            if isempty(i)
                error("mat2doc:KeyError", "'%s'", name);
            end
            obj.attr_names_(i) = [];
            obj.attr_values_(i) = [];
        end

        function set_nsdecl_(obj, prefix, uri)
            % SET_NSDECL_ Ensure a namespace declaration [prefix, uri] is stored
            %   on THIS element (appended in insertion order if not already
            %   present with that exact prefix+URI).
            %
            %   This reproduces the effect of lxml's namespace RECONCILIATION,
            %   which python-docx relies on when it sets a namespaced attribute in
            %   order to force a namespace declaration onto an element (e.g.
            %   coreprops stamps xsi:type and needs xmlns:xsi placed ONCE on the
            %   root). XmlElement.set does not reconcile declarations, so this
            %   explicit accessor is the faithful analogue. The serializer then
            %   renders the standard prefix (e.g. "xsi") instead of an invented
            %   nsN (see serialize_part_xml H8 notes).
            %
            %   Ported from python-docx v1.2.0: the lxml namespace-reconciliation
            %   behaviour exercised by src/docx/oxml/coreprops.py (xsi:type stamp).
            arguments
                obj (1,1) mat2doc.oxml.XmlElement
                prefix (1,1) string
                uri (1,1) string
            end
            if isempty(obj.nsdecls_) || ...
                    ~any(obj.nsdecls_(:, 1) == prefix & obj.nsdecls_(:, 2) == uri)
                obj.nsdecls_(end + 1, :) = [prefix, uri];
            end
        end

        function names = attrib_names(obj)
            % ATTRIB_NAMES Attribute names, insertion-ordered (lxml attrib.keys()).
            %
            %   1xN string of attribute keys in insertion/document order -- the
            %   order lxml serializes them. Iterate names + get(name) to enumerate
            %   attributes (serializer, xmlchemy engine).
            arguments
                obj (1,1) mat2doc.oxml.XmlElement
            end
            names = obj.attr_names_;
        end

        % ------------------------------------------------------------------
        % child search (lxml .find / .findall)
        % ------------------------------------------------------------------

        function child = find(obj, tag_name)
            % FIND First DIRECT child whose tag equals tag_name, or [] (None).
            %
            %   tag_name is a Clark name (qn("w:pPr")) or a plain name for
            %   no-namespace children. Direct children only -- grandchildren are
            %   NOT searched. No path syntax: docx passes only qn() results here;
            %   XPath is a separate WP.
            arguments
                obj (1,1) mat2doc.oxml.XmlElement
                tag_name (1,1) string
            end
            for i = 1:numel(obj.children_)
                c = obj.children_(i);
                if c.tag == tag_name
                    child = c;
                    return
                end
            end
            child = [];   % Python None (H3)
        end

        function matching = findall(obj, tag_name)
            % FINDALL All DIRECT children whose tag equals tag_name.
            %
            %   Returns a 1xN (possibly 1x0) XmlElement array -- Python returns a
            %   list, so "no matches" is an EMPTY ARRAY here, not [] (contrast
            %   find, which maps None -> []). Iterate with `for c = e.findall(...)`.
            arguments
                obj (1,1) mat2doc.oxml.XmlElement
                tag_name (1,1) string
            end
            n = numel(obj.children_);
            keep = false(1, n);
            for i = 1:n
                keep(i) = (obj.children_(i).tag == tag_name);
            end
            matching = obj.children_(keep);
        end

        % ------------------------------------------------------------------
        % xpath (lxml _Element.xpath; docx BaseOxmlElement.xpath override)
        % ------------------------------------------------------------------

        function result = xpath(obj, xpath_str, namespaces)
            % XPATH Evaluate an XPath expression over this element's tree.
            %
            %   nodes = e.XPATH(expr) evaluates expr (a string) against e using
            %   the fixed WordprocessingML namespace map (mat2doc.oxml.nsmap =
            %   Python `nsmap`), mirroring
            %   docx.oxml.xmlchemy.BaseOxmlElement.xpath (xmlchemy.py:687-692),
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
            %   TASK #60 HOIST (P2-2): defined on XmlElement (not just
            %   BaseOxmlElement) because every lxml _Element has xpath. The
            %   parser-fallback root class -- an unregistered parsed root such as
            %   `w:document` before P2-3, or any generic parsed element -- reaches
            %   this method directly. BaseOxmlElement INHERITS it unchanged.
            %   Consumers requiring it on the plain class: XmlPart.rel_ref_count_
            %   (`element.xpath("//@r:id")`, opc/part.py 246) and StoryPart.next_id
            %   (`element.xpath("//@id")`, parts/story.py 84).
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
            %   BaseOxmlElement.xpath (lines 687-692), hoisted to the _Element
            %   analogue per task #60.
            arguments
                obj (1,1) mat2doc.oxml.XmlElement
                xpath_str (1,1) string
                namespaces (1,1) struct = mat2doc.oxml.nsmap()
            end
            result = mat2doc.oxml.evaluate_xpath(obj, xpath_str, namespaces);
        end

        % ------------------------------------------------------------------
        % structure mutation (lxml .append / .insert / .remove / .addnext /
        % .addprevious)
        % ------------------------------------------------------------------

        function append(obj, child)
            % APPEND Add child as the last child of this element, MOVING it.
            %
            %   lxml MOVES on append: an element appended here is detached from
            %   its previous parent -- including from THIS element (re-append
            %   moves to the end). The child's text/tail and namespace
            %   declarations travel with it.
            arguments
                obj (1,1) mat2doc.oxml.XmlElement
                child (1,1) mat2doc.oxml.XmlElement
            end
            obj.assertNotAncestorOrSelf_(child);
            child.detach_();
            obj.children_ = [obj.children_, child];
            child.parent_ = obj;
            child.clearNsVerbatim_();   % reconcile moved subtree (H8; see markNsVerbatim_)
        end

        function insert(obj, idx, child)
            % INSERT Insert child before the element currently at position idx.
            %
            %   e.INSERT(idx, child) -- idx is 1-BASED (H1 MATLAB convention): the
            %   child is moved to come immediately before the current idx-th
            %   child; idx == numel(children)+1 or greater appends (lxml clamps).
            %
            %   % IDX PORTING RULE: Python/lxml `elm.insert(i, c)` is 0-based;
            %   port as `elm.insert(i+1, c)`. docx call sites e.g.
            %   oxml/table.py:133 `self.insert(0, tblPrEx)` -> `obj.insert(1, ...)`;
            %   oxml/text/paragraph.py:105, oxml/text/run.py:146.
            %
            %   Python NEGATIVE indices are NOT ported (unused by docx); idx < 1
            %   raises mat2doc:ValueError.
            %
            %   Like append, insert MOVES the child. When the child is already a
            %   child of this element, the reference position is the one BEFORE
            %   the child is unlinked, matching lxml.
            arguments
                obj (1,1) mat2doc.oxml.XmlElement
                idx (1,1) double {mustBeInteger}
                child (1,1) mat2doc.oxml.XmlElement
            end
            if idx < 1
                error("mat2doc:ValueError", ...
                    "insert index must be >= 1 (1-based; Python negative indices not ported)");
            end
            obj.assertNotAncestorOrSelf_(child);
            n = numel(obj.children_);
            if idx > n
                obj.append(child);   % lxml clamps over-large index to append
                return
            end
            ref = obj.children_(idx);   % reference resolved BEFORE unlink
            if ref == child
                return   % inserting an element before itself: order unchanged
            end
            child.detach_();
            i = find(obj.children_ == ref, 1);
            obj.children_ = [obj.children_(1:i - 1), child, obj.children_(i:end)];
            child.parent_ = obj;
            child.clearNsVerbatim_();   % reconcile moved subtree (H8; see markNsVerbatim_)
        end

        function remove(obj, child)
            % REMOVE Remove child from this element (identity-based, H5).
            %
            %   Matching is by HANDLE IDENTITY (==), never by content. A non-child
            %   raises mat2doc:ValueError with lxml's message. The removed element
            %   keeps its text, tail, attributes and children, and its getparent()
            %   becomes [].
            arguments
                obj (1,1) mat2doc.oxml.XmlElement
                child (1,1) mat2doc.oxml.XmlElement
            end
            i = find(obj.children_ == child, 1);
            if isempty(i)
                error("mat2doc:ValueError", "Element is not a child of this node.");
            end
            obj.children_ = [obj.children_(1:i - 1), obj.children_(i + 1:end)];
            child.parent_ = [];
        end

        function addnext(obj, sibling)
            % ADDNEXT Move sibling to be the following sibling of this element.
            %
            %   Mirrors lxml _Element.addnext: the sibling element is MOVED
            %   (detached from any previous parent) and placed as the sibling
            %   element immediately after this one; its tail travels with it.
            %   Calling on a parentless element raises mat2doc:TypeError with
            %   lxml's message.
            %
            %   docx call sites: oxml/table.py:138, oxml/text/run.py:97.
            arguments
                obj (1,1) mat2doc.oxml.XmlElement
                sibling (1,1) mat2doc.oxml.XmlElement
            end
            parent = obj.parent_;
            if isequal(parent, [])
                error("mat2doc:TypeError", ...
                    "Only processing instructions and comments can be siblings of the root element");
            end
            if sibling == obj
                return
            end
            parent.assertNotAncestorOrSelf_(sibling);
            sibling.detach_();
            i = find(parent.children_ == obj, 1);
            parent.children_ = [parent.children_(1:i), sibling, parent.children_(i + 1:end)];
            sibling.parent_ = parent;
            sibling.clearNsVerbatim_();   % reconcile moved subtree (H8; see markNsVerbatim_)
        end

        function addprevious(obj, sibling)
            % ADDPREVIOUS Move sibling to be the preceding sibling of this element.
            %
            %   Mirrors lxml _Element.addprevious; move semantics and tail travel
            %   as addnext. Parentless raises mat2doc:TypeError.
            %
            %   docx call sites: oxml/text/paragraph.py:36, oxml/text/parfmt.py:389,
            %   oxml/text/run.py:102, oxml/xmlchemy.py:667.
            arguments
                obj (1,1) mat2doc.oxml.XmlElement
                sibling (1,1) mat2doc.oxml.XmlElement
            end
            parent = obj.parent_;
            if isequal(parent, [])
                error("mat2doc:TypeError", ...
                    "Only processing instructions and comments can be siblings of the root element");
            end
            if sibling == obj
                return
            end
            parent.assertNotAncestorOrSelf_(sibling);
            sibling.detach_();
            i = find(parent.children_ == obj, 1);
            parent.children_ = [parent.children_(1:i - 1), sibling, parent.children_(i:end)];
            sibling.parent_ = parent;
            sibling.clearNsVerbatim_();   % reconcile moved subtree (H8; see markNsVerbatim_)
        end

        % ------------------------------------------------------------------
        % navigation / iteration (lxml .getparent / list(e) / len(e) / .index)
        % ------------------------------------------------------------------

        function parent = getparent(obj)
            % GETPARENT THE parent element handle, or [] (None) for a root/detached element (H5).
            arguments
                obj (1,1) mat2doc.oxml.XmlElement
            end
            parent = obj.parent_;
        end

        function children = to_array(obj)
            % TO_ARRAY Ordered children as a 1xN (possibly 1x0) heterogeneous array.
            %
            %   The iteration idiom (design.md section 2):
            %     Python `for child in element` -> `for child = element.to_array()`
            %     Python `len(element)`         -> `numel(element.to_array())`
            %     Python `element[i]`           -> `arr = element.to_array(); arr(i+1)` % IDX
            arguments
                obj (1,1) mat2doc.oxml.XmlElement
            end
            children = obj.children_;
        end

        function children = getchildren(obj)
            % GETCHILDREN Alias of to_array (lxml _Element.getchildren / list(element)).
            arguments
                obj (1,1) mat2doc.oxml.XmlElement
            end
            children = obj.children_;
        end

        function children = iterchildren(obj)
            % ITERCHILDREN Alias of to_array (lxml _Element.iterchildren()).
            %
            %   H9: lxml returns a lazy iterator; docx call sites only read tags
            %   without mutating during iteration, so a materialized array is
            %   unobservable.
            arguments
                obj (1,1) mat2doc.oxml.XmlElement
            end
            children = obj.children_;
        end

        function i = index(obj, child)
            % INDEX 1-BASED position of child among this element's children.
            %
            %   % IDX PORTING RULE (H1): lxml `parent.index(child)` is 0-based;
            %   where the Python VALUE is used as data, port with an explicit -1:
            %   e.g. `self.getparent().index(self)` -> `obj.getparent().index(obj) - 1`.
            %   A non-child raises mat2doc:ValueError with lxml's message.
            arguments
                obj (1,1) mat2doc.oxml.XmlElement
                child (1,1) mat2doc.oxml.XmlElement
            end
            i = find(obj.children_ == child, 1);
            if isempty(i)
                error("mat2doc:ValueError", "Element is not a child of this node.");
            end
        end

        % ------------------------------------------------------------------
        % property access methods
        % ------------------------------------------------------------------

        function value = get.tag(obj)
            if obj.nsuri_ == ""
                value = obj.local_;
            else
                value = "{" + obj.nsuri_ + "}" + obj.local_;
            end
        end

        function value = get.nsptag_str(obj)
            value = obj.nsptag_;
        end

        function value = get.local_part(obj)
            value = obj.local_;
        end

        function value = get.nspfx(obj)
            value = obj.pfx_;
        end

        function value = get.nsuri(obj)
            value = obj.nsuri_;
        end

        function value = get.nsdecls(obj)
            value = obj.nsdecls_;
        end

        function value = get.text(obj)
            value = obj.getText_();
        end

        function set.text(obj, value)
            obj.setText_(value);
        end

        function value = get.tail(obj)
            value = obj.tail_;
        end

        function set.tail(obj, value)
            obj.tail_ = mat2doc.oxml.XmlElement.normalizeTextValue_(value);
        end
    end

    % ------------------------------------------------------------------
    % Heterogeneous-array identity operators (design.md section 2
    % "Heterogeneous-array Sealed-method contract (INTEGRATION-CRITICAL)").
    %
    % XmlElement is the matlab.mixin.Heterogeneous root; once the registry is
    % populated, a parsed tree's child vector is a HETEROGENEOUS array (mixed
    % XmlElement + CT_* subclasses). MATLAB refuses to dispatch a method on a
    % heterogeneous array unless that method is Sealed in the root class. The
    % tree engine locates a node inside its parent's child vector by HANDLE
    % IDENTITY -- find(arr == node, 1) -- in every one of remove / index / insert
    % / addnext / addprevious / detach_. Those `==`/`~=` are inherited from
    % `handle` (identity comparison) but are NOT Sealed there, so `arr == node`
    % throws MATLAB:class:UnsealedMethod the instant arr holds more than one
    % concrete class. Sealing eq/ne here -- forwarding verbatim to the handle
    % builtin -- makes `==`/`~=` dispatch on the mixed array while preserving
    % EXACT handle-identity semantics (H5). This is IDENTITY only; no content
    % comparison is introduced.
    methods (Sealed)
        function tf = eq(a, b)
            % EQ Element identity (==): true iff a and b are the SAME node (H5).
            %   Sealed so `==` dispatches on a heterogeneous XmlElement+CT_* array
            %   (the parsed-tree child vector). Forwards to handle's identity `eq`.
            tf = eq@handle(a, b);
        end

        function tf = ne(a, b)
            % NE Element non-identity (~=): true iff a and b are DIFFERENT nodes.
            %   Sealed counterpart of eq. Forwards to handle's identity `ne`.
            tf = ne@handle(a, b);
        end

        function c = deepcopy(obj)
            % DEEPCOPY Independent deep copy of this element's whole subtree.
            %
            %   c = e.DEEPCOPY() returns a NEW, DETACHED element that duplicates
            %   e's tag, namespace declarations, attributes (names, values and
            %   ORDER), text, tail and every descendant -- sharing NO handle with
            %   e. c.getparent() is [] (None). Realizes Python
            %   `copy.deepcopy(element)` / lxml `_Element.__deepcopy__`.
            %
            %   docx call sites: oxml/section.py:192 (`cloned_sectPr =
            %   deepcopy(self)`) and oxml/text/pagebreak.py:154/182/218/247
            %   (`copy.deepcopy(self._enclosing_p)`), each cloning a subtree and
            %   re-homing it WITHIN the same document.xml part.
            %
            %   It is a Sealed method on the root class (not a package function)
            %   because (1) it must reproduce PRIVATE tree state (attr_names_,
            %   attr_values_, nsdecls_, text_, tail_, nsVerbatim_, the child
            %   vector) that only code inside XmlElement may touch, and (2) the
            %   recursion runs over the HETEROGENEOUS (XmlElement + CT_*) child
            %   vector, which requires Sealed (design.md section 2). SEALED IS
            %   MANDATORY: an unsealed version appears to work on a HOMOGENEOUS
            %   hand-built tree and throws MATLAB:class:UnsealedMethod the instant
            %   it meets a PARSED mixed tree.
            %
            %   COMPLETENESS: XmlElement has exactly twelve stateful fields; every
            %   one is accounted for -- nsptag_ / pfx_ / local_ (rebuilt by the
            %   ctor from the prefixed tag), nsuri_ (passed verbatim as
            %   resolvedUri), nsdecls_ (ctor arg), attr_names_ / attr_values_ /
            %   text_ / tail_ / nsVerbatim_ (copied below), children_ (recursed),
            %   parent_ (deliberately left [] -- a clone is detached). No CT_*
            %   subclass declares stateful properties, so this is a COMPLETE copy
            %   for the whole hierarchy. IF A FUTURE CT_* ADDS INSTANCE STATE THIS
            %   METHOD MUST GROW AN OVERRIDABLE HOOK OR IT WILL SILENTLY DROP IT.
            %
            %   CLASS IS PRESERVED (H10): built via feval(class(obj), ...) using
            %   the CT_* transparent-pass-through constructor contract, so a `w:p`
            %   clones as its CT_P class. lxml likewise preserves the custom
            %   element class across a deepcopy.
            %
            %   H3 tri-state: text_/tail_ are copied to the RAW fields (never
            %   through the .text property), so a CT_* getText_/setText_ shadow
            %   cannot alter what is copied, and "" vs [] (None) is preserved
            %   exactly. TAIL IS COPIED (lxml's deepcopy carries the tail; in the
            %   remove_blank_text path the tail is None anyway).
            %
            %   NAMESPACE BEHAVIOR (H8): this port copies only the declarations
            %   stored on the element ITSELF and passes the resolved tag URI
            %   verbatim (third ctor arg), copying children WITHOUT reconciliation
            %   (appendParsed_); lxml instead reconciles ancestor decls onto the
            %   copy root. The difference is unobservable while a clone is re-homed
            %   into the SAME part scope (as all docx callers do -- document.xml's
            %   root declares every prefix used). nsVerbatim_ is copied faithfully,
            %   so a clone of a parsed node starts verbatim exactly as its source;
            %   the public move that re-inserts it then clears the flag on the
            %   clone only. ^ VERIFY FOR ANY FUTURE CALLER that moves a clone into
            %   a DIFFERENT part or namespace scope -- add a dedicated namespace
            %   probe before merging (re-verify the docx section/pagebreak callers
            %   at their WPs).
            %
            %   Sealed for the heterogeneous-dispatch reason above.
            %
            %   Inputs:  none (operates on the element it is called on)
            %   Outputs: c - a new DETACHED element of the same class as OBJ,
            %                sharing no handle with OBJ or any descendant
            %
            %   Design-realization of lxml `_Element.__deepcopy__` (the target of
            %   `copy.deepcopy`), part of the lxml _Element surface used by
            %   python-docx v1.2.0. Call sites: src/docx/oxml/section.py:192,
            %   src/docx/oxml/text/pagebreak.py:154+.
            arguments
                obj (1,1) mat2doc.oxml.XmlElement
            end
            % Same class, same written prefixed tag, same OWN ns-decls, and the
            % already-resolved tag URI passed verbatim as `resolvedUri` so the
            % ctor does NOT re-resolve against the fixed nsmap (which would fail
            % for a prefix declared on an ancestor, e.g. a parsed mc/wps element).
            c = feval(class(obj), obj.nsptag_, obj.nsdecls_, obj.nsuri_);
            % Attributes: names, values AND ORDER (H8/H11 -- order is byte-visible).
            c.attr_names_ = obj.attr_names_;
            c.attr_values_ = obj.attr_values_;
            % text/tail preserve the [] (None) vs "" tri-state exactly (H3).
            c.text_ = obj.text_;
            c.tail_ = obj.tail_;
            c.nsVerbatim_ = obj.nsVerbatim_;
            % Children: recursive clone, attached WITHOUT namespace reconciliation
            % (appendParsed_), because lxml's copy builds the subtree in place and
            % reconciles only at the destination on move.
            for i = 1:numel(obj.children_)
                c.appendParsed_(obj.children_(i).deepcopy());
            end
            % c.parent_ is left [] -- lxml's deepcopy yields a detached root.
        end
    end

    methods (Sealed, Hidden)
        function t = text_raw_(obj)
            % TEXT_RAW_ Raw stored element text: string scalar or [] (None).
            %
            %   Serialization accessor. lxml's serializer reads the C-level
            %   _Element text, NEVER a Python-level `.text` property shadow (docx
            %   oxml/table.py CT_Tc.text shadows lxml _Element.text yet serializes
            %   its stored text). The MATLAB serializer must therefore bypass
            %   getText_/setText_ overrides (D10). Sealed so CT_* subclasses
            %   cannot alter serialization behavior.
            t = obj.text_;
        end

        % ---- parser entry points -----------------------------------------
        % The parser builds the tree bottom-up from raw document bytes. It must
        % set the C-level text/tail and the in-scope-resolved namespace URI
        % directly, exactly as lxml's C parser does -- BYPASSING any CT_* `.text`
        % property shadow (D10) and the constructor's fixed-map URI resolution.
        % Sealed so CT_* subclasses cannot alter parse behavior; Hidden so they
        % do not widen the API.

        function setTextRaw_(obj, value)
            % SETTEXTRAW_ Set the raw element text (parser), bypassing setText_
            %   shadows. value: string scalar or [] (None, H3).
            obj.text_ = mat2doc.oxml.XmlElement.normalizeTextValue_(value);
        end

        function setTailRaw_(obj, value)
            % SETTAILRAW_ Set the raw element tail (parser). tail is never
            %   shadowed (D10), but provided for symmetry. value: string or [].
            obj.tail_ = mat2doc.oxml.XmlElement.normalizeTextValue_(value);
        end

        function markNsVerbatim_(obj)
            % MARKNSVERBATIM_ Mark this element's stored ns-decls to be serialized
            %   VERBATIM (never suppressed). XmlParser.parseElement_ calls this on
            %   every node it builds.
            %
            %   WHY (fixes D-serializer-nsdecl): lxml serializes a PARSED tree's
            %   namespace declarations VERBATIM. lxml's "suppression" of a
            %   redundant nested decl is a MOVE-time reconciliation side effect
            %   (moveNodeToDocument, fired by the public append/insert/addnext/
            %   addprevious) -- it NEVER applies to declarations that arrived
            %   through parsing, because the C parser builds the tree in place and
            %   reconciles nothing. So the discriminator is NOT parsed vs generated
            %   (generation itself builds fragments via parse_xml on XML literals)
            %   but VERBATIM-UNTIL-MOVED:
            %     - the parser marks each node verbatim (this method), attaching
            %       children via appendParsed_ (no reconciliation);
            %     - a node stays verbatim until a PUBLIC move reconciles it into a
            %       new scope, at which point clearNsVerbatim_ clears it (and its
            %       subtree) and the serializer's ancestor-URI suppression applies
            %       again.
            %   Net: a file opened and re-saved unchanged round-trips verbatim
            %   (redundant nested decls preserved exactly as python-docx does);
            %   everything Mat2Doc builds is moved into place and therefore dedup'd
            %   byte-for-byte as before.
            %
            %   Sealed so CT_* subclasses cannot alter this; Hidden so it does not
            %   widen the public API. The move-time reconciliation model is the
            %   design.md section 3 "Serialize" contract.
            arguments
                obj (1,1) mat2doc.oxml.XmlElement
            end
            obj.nsVerbatim_ = true;
        end

        function tf = isNsVerbatim_(obj)
            % ISNSVERBATIM_ True iff this element's stored ns-decls must be
            %   serialized verbatim (see markNsVerbatim_). Read by
            %   serialize_part_xml step 1 to choose verbatim emission vs
            %   ancestor-URI suppression.
            arguments
                obj (1,1) mat2doc.oxml.XmlElement
            end
            tf = obj.nsVerbatim_;
        end

        function appendParsed_(obj, child)
            % APPENDPARSED_ Parser-internal attach: place child as the last child
            %   WITHOUT the namespace reconciliation the public append performs.
            %   lxml's C parser builds the tree in place and never reconciles;
            %   only the public _Element.append/insert/addnext/addprevious
            %   (moveNodeToDocument) do. The parser therefore attaches through this
            %   method so a parsed child keeps its verbatim ns-decls
            %   (markNsVerbatim_). child is a freshly parsed element, never
            %   previously attached.
            arguments
                obj (1,1) mat2doc.oxml.XmlElement
                child (1,1) mat2doc.oxml.XmlElement
            end
            obj.children_ = [obj.children_, child];
            child.parent_ = obj;
        end

        function clearNsVerbatim_(obj)
            % CLEARNSVERBATIM_ Clear the verbatim-ns flag on this element and its
            %   whole current subtree (recursively). Called by the public move
            %   methods (append/insert/addnext/addprevious): lxml reconciles the
            %   ENTIRE moved subtree against the destination scope
            %   (moveNodeToDocument), so a moved element -- even one that arrived
            %   by parsing -- is no longer serialized verbatim; the serializer's
            %   ancestor-URI suppression applies to it instead. Sealed so it
            %   dispatches over the heterogeneous (XmlElement + CT_*) child vector.
            obj.nsVerbatim_ = false;
            for i = 1:numel(obj.children_)
                obj.children_(i).clearNsVerbatim_();
            end
        end

        function setResolvedNsuri_(obj, uri)
            % SETRESOLVEDNSURI_ Override the element's namespace URI with the value
            %   the parser resolved from the in-scope declaration stack.
            %
            %   REQUIRED because the public constructor resolves a tag's URI from
            %   the element's OWN declarations first, then the fixed nsmap --
            %   neither of which covers an element whose prefix (or default xmlns)
            %   is declared on an ANCESTOR, e.g. a <Default> child inheriting the
            %   default namespace of <Types xmlns="...">. The parser stores only
            %   the declarations the document wrote on THIS element (nsdecls) and
            %   supplies the authoritative in-scope-resolved URI here.
            %   uri: string scalar ("" = none).
            arguments
                obj (1,1) mat2doc.oxml.XmlElement
                uri (1,1) string
            end
            obj.nsuri_ = uri;
        end
    end

    methods (Access = protected)
        % text is routed through overridable protected methods because Python
        % CT_* subclasses SHADOW the lxml `.text` attribute with a property (docx
        % oxml/table.py CT_Tc.text, oxml/text/*.py -- "note this shadows lxml
        % _Element.text"). MATLAB forbids redefining a superclass property, so
        % those ports override getText_/setText_ instead.

        function value = getText_(obj)
            value = obj.text_;
        end

        function setText_(obj, value)
            obj.text_ = mat2doc.oxml.XmlElement.normalizeTextValue_(value);
        end
    end

    methods (Access = private)
        function detach_(obj)
            % DETACH_ Unlink this element from its parent (no error if loose).
            % Text/tail/attributes/children/nsdecls are NOT touched -- they travel
            % with the element on moves.
            p = obj.parent_;
            if isequal(p, [])
                return
            end
            i = find(p.children_ == obj, 1);
            p.children_ = [p.children_(1:i - 1), p.children_(i + 1:end)];
            obj.parent_ = [];
        end

        function assertNotAncestorOrSelf_(obj, child)
            % ASSERTNOTANCESTORORSELF_ Guard against creating a tree cycle. lxml
            % raises ValueError("cannot append parent to itself") for both
            % ancestor-append and self-append. docx never triggers this guard, but
            % its message matches lxml verbatim.
            anc = obj;
            while ~isempty(anc)
                if anc == child
                    error("mat2doc:ValueError", "cannot append parent to itself");
                end
                anc = anc.parent_;
            end
        end
    end

    methods (Static, Access = private)
        function v = normalizeTextValue_(value)
            % NORMALIZETEXTVALUE_ Validate/normalize a text/tail assignment.
            % Accepts string scalar, char row/'' (-> string), or [] (None, H3).
            % Anything else raises mat2doc:TypeError exactly as lxml (TypeError
            % "Argument must be bytes or unicode, got 'int'").
            if isequal(value, [])
                v = [];
            elseif ischar(value) && (isrow(value) || isempty(value))
                v = string(value);
            elseif isstring(value) && isscalar(value) && ~ismissing(value)
                v = value;
            else
                error("mat2doc:TypeError", ...
                    "Argument must be bytes or unicode, got '%s'", class(value));
            end
        end
    end
end

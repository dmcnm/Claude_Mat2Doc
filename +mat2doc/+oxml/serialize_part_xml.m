function xml_bytes = serialize_part_xml(part_elm)
% SERIALIZE_PART_XML Produce XML-file bytes for part_elm, suitable for writing directly to a .xml file.
%
%   xml_bytes = MAT2DOC.OXML.SERIALIZE_PART_XML(part_elm) serializes the
%   XmlElement tree rooted at part_elm to UTF-8 bytes, including the XML
%   declaration header -- the MATLAB replacement for
%   `etree.tostring(part_elm, encoding="UTF-8", standalone=True)`.
%
%   Inputs:  part_elm - (1,1) mat2doc.oxml.XmlElement, the part root element
%   Outputs: xml_bytes - 1xN uint8, exact file bytes (Python `bytes`)
%
%   BYTE-MATCHED to lxml 5.3.0 (libxml2 2.13.9), the serializer python-docx
%   v1.2.0 uses for every part it writes. Frozen conventions (golden fixtures
%   harness\common\golden\*_docx.bin):
%
%   - Declaration: <?xml version='1.0' encoding='UTF-8' standalone='yes'?> with
%     SINGLE quotes, followed by exactly one LF (golden declaration_docx.bin;
%     proven byte-identical to the pptx declaration.bin).
%   - Attributes: insertion order (golden attr_order_docx.bin), double-quoted.
%   - Escaping (H7, goldens escaping_text_docx.bin / escaping_attr_docx.bin):
%     text escapes & < > and CR (as &#13;); attribute values additionally escape
%     " (as &quot;), LF (&#10;) and TAB (&#9;). Decimal character references
%     only; apostrophe NEVER escaped; all non-ASCII emitted as raw UTF-8.
%   - Empty element: self-closing <tag/> (no space before "/>") ONLY when the
%     element has no children AND text is [] (None); text == "" (empty string)
%     serializes as <tag></tag> (H3 tri-state; golden empty_element_docx.bin).
%   - Namespace declarations (H8): emitted on the owning element, in stored
%     order, BEFORE ordinary attributes. Verbatim-until-moved (D-serializer-nsdecl):
%       * MOVED element (default, everything Mat2Doc builds) -- a stored
%         declaration is SUPPRESSED when its URI is already reachable through a
%         valid (non-shadowed) in-scope binding of an ANCESTOR: lxml's move-time
%         reconciliation merge, reproduced at serialize time (golden
%         nsdecl_placement_docx.bin). Declarations on one element never suppress
%         each other.
%       * VERBATIM element (XmlElement.isNsVerbatim_ -- parsed and not yet
%         re-moved) -- stored declarations are emitted VERBATIM, never
%         suppressed. lxml serializes a parsed tree's declarations as written
%         (its suppression is a move-time side effect that never fires on a tree
%         the parser built in place); reproducing that restores L1 on foreign
%         parts carrying redundant nested decls. See serializer step 1 and
%         XmlElement.markNsVerbatim_.
%     The move-time-vs-serialize-time distinction (lxml moveNodeToDocument is the
%     ONLY reconciliation trigger; the C parser reconciles nothing) is the
%     design.md section 3 "Serialize" contract. Residual foreign-file
%     prefix-rendering divergence is ledgered as D-nsprefix-rewrite (L2, dead on
%     generation), NOT re-introduced here.
%   - Prefix rendering: an element/attribute in namespace URI is rendered with
%     the prefix of the first valid in-scope binding for that URI -- innermost
%     element first, stored order within an element, bindings shadowed by an
%     inner rebinding of the same prefix skipped. The stored prefixed tag is
%     used only as a fallback: a tag URI reachable through NO binding
%     materializes a declaration of the element's stored prefix on the element
%     itself (lxml analogue: detached elements re-materialize their binding).
%   - Attributes NEVER use the default ("") namespace binding; a namespaced
%     attribute whose URI has no prefixed in-scope binding gets an INVENTED nsN
%     prefix (smallest N >= 0 whose prefix "nsN" is unbound in scope), declared
%     on the same element after the stored declarations, in attribute-encounter
%     order -- exactly lxml.
%   - The root element's tail, when set, is emitted after the closing tag. No
%     pretty-printing, no xmlns="" un-declaration for no-ns elements under a
%     default-ns scope.
%
%   NOTE: like lxml's C-level writer, serialization reads the RAW stored element
%   text (text_raw_), never a CT_* subclass `.text` property shadow (D10).
%
%   SCOPE: part_elm is serialized as a document root. python-docx only ever
%   passes part roots and the .rels root to this tostring call. Serializing an
%   element still attached under ancestors whose declarations it inherits is NOT
%   part of the used surface.
%
%   Example:
%       t = mat2doc.oxml.XmlElement("w:t", struct("w", mat2doc.oxml.nsmap().w));
%       t.text = "hello";
%       bytes = mat2doc.oxml.serialize_part_xml(t);
%       % <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
%       % <w:t xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">hello</w:t>
%
%   Ported from python-docx v1.2.0: src/docx/opc/oxml.py::serialize_part_xml

arguments
    part_elm (1,1) mat2doc.oxml.XmlElement
end

% Golden declaration_docx.bin: single quotes, exact case, one trailing LF (0x0A).
declaration = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>" + newline;

s = writeElement_(part_elm, {});

% Root tail is included by lxml tostring; never set on docx part roots, carried
% for tree fidelity.
root_tail = part_elm.tail;
if ~isequal(root_tail, [])
    s = s + escapeText_(root_tail);
end

% H2: text -> bytes ONLY via unicode2native; UTF-16 surrogate pairs (emoji)
% become correct 4-byte UTF-8 sequences.
xml_bytes = unicode2native(declaration + s, "UTF-8");
end

% ==========================================================================
% local functions
% ==========================================================================

function s = writeElement_(elm, scope)
% WRITEELEMENT_ Serialize one element and its subtree.
%   scope is a cell row vector of Nx2 string arrays [prefix, URI]: the namespace
%   declarations EMITTED on each ancestor, outermost first. Empty for the
%   document root.

% -- 1. namespace declarations: verbatim-until-moved (H8) --------------------
% lxml's suppression of a redundant nested decl is a MOVE-time reconciliation
% side effect (moveNodeToDocument, fired by append/insert/addnext/addprevious);
% it NEVER applies to declarations that arrived through parsing (the C parser
% builds the tree in place and reconciles nothing). The port mirrors this with
% XmlElement.nsVerbatim_ (set by the parser via markNsVerbatim_, CLEARED on the
% moved subtree by any public move) -- NOT a parsed-vs-generated distinction,
% because generation itself builds fragments via parse_xml on XML literals and
% then MOVES them into place (which clears the flag). So:
%
%   - isNsVerbatim_ (parsed, not yet moved): emit every stored decl verbatim, in
%     stored order. The prefix-rendering steps below then bind each tag/attr to
%     the innermost stored binding, so a redundant z-decl re-renders z:tag /
%     z:attr exactly as the source wrote it -- never an invented nsN. Restores
%     L1 byte-match with python-docx's own open->save on foreign parts carrying
%     redundant nested decls (fixes D-serializer-nsdecl). NO-OP on the common
%     parsed case: a normal parsed child stores NO decls (inherited from an
%     ancestor), so emitted is empty exactly as before.
%   - else (moved into place -- everything Mat2Doc builds, or a re-moved parsed
%     element): reproduce lxml's move-time reconciliation net effect at
%     serialize time -- suppress a stored decl whose URI is already reachable via
%     a valid (non-shadowed) ancestor binding (shadow-aware). Same-element decls
%     never suppress each other, so the test runs against the ANCESTOR scope
%     only. This branch is byte-for-byte the pre-fix code, so every generated
%     part is byte-neutral.
stored = elm.nsdecls;
if elm.isNsVerbatim_()
    emitted = stored;
else
    emitted = strings(0, 2);
    for k = 1:size(stored, 1)
        [~, found] = lookupPrefix_(scope, stored(k, 2), true);
        if ~found
            emitted(end + 1, :) = stored(k, :); %#ok<AGROW>
        end
    end
end
scope2 = [scope, {emitted}];

% -- 2. tag rendering: prefix from nearest valid in-scope binding -----------
if elm.nsuri == ""
    % No-namespace element: bare local name, even under a default-ns scope
    % (lxml emits no xmlns="" un-declaration).
    name = elm.local_part;
elseif elm.nsuri == "http://www.w3.org/XML/1998/namespace"
    % Reserved 'xml' prefix (W3C Namespaces section 3): always in scope, never
    % declared -- lxml renders xml:local with NO xmlns:xml.
    name = "xml:" + elm.local_part;
else
    [pfx, found] = lookupPrefix_(scope2, elm.nsuri, true);
    if ~found
        % No binding reaches the tag URI: materialize the element's stored prefix
        % on this element (lxml analogue: a detached element re-materializes its
        % own binding). Reachable only for elements built without an nsmap;
        % OxmlElement always stores the tag's own declaration.
        pfx = elm.nspfx;
        emitted(end + 1, :) = [pfx, elm.nsuri];
        scope2{end} = emitted;
    end
    if pfx == ""
        name = elm.local_part;   % default-ns binding: unprefixed
    else
        name = pfx + ":" + elm.local_part;
    end
end

% -- 3. attributes: insertion order; Clark names -> in-scope prefix ---------
attr_names = elm.attrib_names();
attrs_str = "";
for k = 1:numel(attr_names)
    n = attr_names(k);
    if startsWith(n, "{")
        uri = extractBetween(n, "{", "}");
        local = extractAfter(n, "}");
        if uri == "http://www.w3.org/XML/1998/namespace"
            % Reserved 'xml' prefix: render xml:local, NEVER declare xmlns:xml
            % (Word writes <w:t ... xml:space="preserve"> with no xmlns:xml).
            % Must precede the invent-prefix path, which would otherwise
            % materialise a bogus nsN declaration for this URI.
            rendered = "xml:" + local;
        else
            % Attributes never use the default ("") binding.
            [apfx, afound] = lookupPrefix_(scope2, uri, false);
            if ~afound
                % Invent nsN exactly as lxml: smallest N with "nsN" unbound in
                % scope; declared on this element after the stored declarations,
                % in attribute-encounter order.
                apfx = inventPrefix_(scope2);
                emitted(end + 1, :) = [apfx, uri]; %#ok<AGROW>
                scope2{end} = emitted;
            end
            rendered = apfx + ":" + local;
        end
    else
        rendered = n;
    end
    attrs_str = attrs_str + " " + rendered + "=""" + ...
        escapeAttr_(elm.get(n)) + """";
end

% -- 4. declaration strings (stored-emitted, then materialized/invented) ----
decls_str = "";
for k = 1:size(emitted, 1)
    if emitted(k, 1) == ""
        decls_str = decls_str + " xmlns=""" + escapeAttr_(emitted(k, 2)) + """";
    else
        decls_str = decls_str + " xmlns:" + emitted(k, 1) + "=""" + ...
            escapeAttr_(emitted(k, 2)) + """";
    end
end

% -- 5. content: self-close ONLY for no-children AND text None (H3) ---------
open_tag = "<" + name + decls_str + attrs_str;
kids = elm.to_array();
txt = elm.text_raw_();   % raw stored text; bypasses CT_* .text shadowing (D10)
if isempty(kids) && isequal(txt, [])
    s = open_tag + "/>";
else
    inner = "";
    if ~isequal(txt, [])
        inner = escapeText_(txt);
    end
    for k = 1:numel(kids)
        c = kids(k);
        inner = inner + writeElement_(c, scope2);
        ct = c.tail;
        if ~isequal(ct, [])
            inner = inner + escapeText_(ct);
        end
    end
    s = open_tag + ">" + inner + "</" + name + ">";
end
end

function [pfx, found] = lookupPrefix_(scope, uri, allow_default)
% LOOKUPPREFIX_ First valid in-scope binding for uri, or found = false.
%   Search order: innermost element level first, stored order within a level. A
%   binding is skipped when its prefix is rebound at any inner level (shadowed),
%   or when it is the default ("") binding and allow_default is false (attribute
%   lookup).
pfx = "";
found = false;
for lvl = numel(scope):-1:1
    level = scope{lvl};
    for k = 1:size(level, 1)
        if level(k, 2) ~= uri
            continue
        end
        p = level(k, 1);
        if ~allow_default && p == ""
            continue
        end
        if isShadowed_(scope, lvl, p)
            continue
        end
        pfx = p;
        found = true;
        return
    end
end
end

function tf = isShadowed_(scope, lvl, p)
% ISSHADOWED_ True when prefix p is redeclared at a level inner than lvl.
tf = false;
for L = lvl + 1:numel(scope)
    if any(scope{L}(:, 1) == p)
        tf = true;
        return
    end
end
end

function pfx = inventPrefix_(scope)
% INVENTPREFIX_ Smallest nsN prefix not bound anywhere in scope (an in-scope ns0
%   forces ns1; numbering restarts per element because sibling declarations are
%   out of scope).
n = 0;
while true
    pfx = "ns" + n;
    bound = false;
    for L = 1:numel(scope)
        if any(scope{L}(:, 1) == pfx)
            bound = true;
            break
        end
    end
    if ~bound
        return
    end
    n = n + 1;
end
end

function out = escapeText_(s)
% ESCAPETEXT_ lxml text-context escaping (H7, golden escaping_text_docx.bin):
%   & < > escaped as named entities; CR as decimal &#13;. LF, TAB, quotes and
%   apostrophes pass through raw; non-ASCII stays raw (UTF-8 at byte stage).
%   Order matters: "&" first so entity ampersands are not re-escaped.
out = replace(s, ["&", "<", ">", string(char(13))], ...
    ["&amp;", "&lt;", "&gt;", "&#13;"]);
end

function out = escapeAttr_(s)
% ESCAPEATTR_ lxml attribute-context escaping (H7, golden escaping_attr_docx.bin):
%   text table plus " (&quot;), LF (&#10;) and TAB (&#9;); decimal references
%   only; apostrophe never escaped (lxml always double-quotes values).
out = replace(s, ...
    ["&", "<", ">", """", string(newline), string(char(9)), string(char(13))], ...
    ["&amp;", "&lt;", "&gt;", "&quot;", "&#10;", "&#9;", "&#13;"]);
end

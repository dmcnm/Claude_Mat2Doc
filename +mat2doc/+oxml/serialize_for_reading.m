function xml = serialize_for_reading(element)
% SERIALIZE_FOR_READING Human-readable (pretty-printed) XML for `element`, no declaration.
%
%   xml = MAT2DOC.OXML.SERIALIZE_FOR_READING(element) returns a string scalar
%   holding the XML serialization of the tree rooted at `element`, pretty-printed
%   for readability and WITHOUT an XML declaration -- the MATLAB replacement for
%   `etree.tostring(element, encoding="unicode", pretty_print=True)`
%   (docx/oxml/xmlchemy.py:22-27). Backs the `.xml` test-helper property of both
%   BaseOxmlElement base classes (mat2doc.oxml.BaseOxmlElement and
%   mat2doc.opc.oxml.BaseOxmlElement).
%
%   TEST-ONLY SURFACE (design.md section 3; plan-audit planaudit_2026-07-25 H-risk
%   #1). No docx PRODUCTION code path reads this: the byte-exact part serializer
%   is mat2doc.oxml.serialize_part_xml (declaration + no insignificant
%   whitespace), and the only production `.xml`-bytes path (CT_Relationships) is
%   the rotated xml_file_bytes -> serialize_part_xml. This pretty printer exists
%   for Gate-4 test comparisons that mirror python-docx's own `element.xml`.
%
%   Whitespace semantics (verified against LIVE lxml 5.3.0 / libxml2 at the P1-4
%   Gate-2 audit, per plan-audit H-risk #1 -- asserted live, not frozen as a
%   golden). The libxml2 format algorithm reproduced here:
%     - 2-space indent per level; one trailing LF after the root (after the
%       root's tail, when set -- tostring includes the element's own tail).
%     - An element is indented (format ON) only when ALL its children are
%       element nodes; a text/CDATA/entity-ref child among the children turns
%       formatting OFF, so leaf text like <w:t>hi</w:t> stays inline.
%     - Format-off PROPAGATES to the ENTIRE subtree of a mixed-content element:
%       libxml2 zeroes the format flag on descent and never re-enables it, so
%       '<a>t<b><c><d/></c></b></a>' serializes fully inline even though <c>
%       has all-element children (Gate-2 probe vector c2/c3, 2026-07-25).
%   This function is deliberately kept OUT of the L1 byte path and
%   does NOT reuse/alter the byte-critical serialize_part_xml (it duplicates that
%   file's ns/attr/escape rendering, cross-referenced below, so the audited P1-2
%   serializer is untouched).
%
%   NOTE (XmlString): Python wraps the result in XmlString (a str subclass whose
%   __eq__ relaxes to a canonical comparison for tests). That relaxed-equality
%   comparison is a test-harness concern and is NOT ported here; this returns a
%   plain string scalar.
%
%   Inputs:  element - (1,1) mat2doc.oxml.XmlElement
%   Outputs: xml     - (1,1) string, pretty-printed, no declaration, trailing LF
%
%   Ported from python-docx v1.2.0: src/docx/oxml/xmlchemy.py::serialize_for_reading
%   (lines 22-27)

arguments
    element (1,1) mat2doc.oxml.XmlElement
end
% encoding="unicode" -> no XML declaration. pretty_print=True -> libxml2 format
% mode; lxml appends exactly one trailing LF after the root (and its tail).
xml = writePretty_(element, {}, 0, true);
% tostring includes the element's own tail (with_tail=True default): lxml emits
% '<a/>TAIL\n' for a tailed element (Gate-2 probe, 2026-07-25); mirror
% serialize_part_xml's root-tail step. Never set on docx part roots.
root_tail = element.tail;
if ~isequal(root_tail, [])
    xml = xml + escapeText_(root_tail);
end
xml = xml + newline;
end

% ==========================================================================
% local functions
%
% The namespace/attribute/escape rendering below MIRRORS
% mat2doc.oxml.serialize_part_xml's writeElement_ helpers EXACTLY (verbatim
% copies of lookupPrefix_ / isShadowed_ / inventPrefix_ / escapeText_ /
% escapeAttr_ and the same verbatim-until-moved nsdecl logic, H7/H8); only the
% element-content emission adds libxml2 pretty-print indentation. Kept separate
% so the byte-critical serialize_part_xml is not modified.
% ==========================================================================

function s = writePretty_(elm, scope, level, fmt)
% WRITEPRETTY_ Serialize one element and its subtree.
%   fmt is the libxml2 format flag as INHERITED from the parent: false when any
%   ancestor was mixed-content (libxml2 zeroes format on descent into a mixed
%   element and never re-enables it -- Gate-2 probe vectors c2/c3). When fmt is
%   true this element still turns formatting off for its own subtree if it has
%   a text-node child.
% -- 1. namespace declarations: verbatim-until-moved (H8; see serialize_part_xml)
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
    name = elm.local_part;
elseif elm.nsuri == "http://www.w3.org/XML/1998/namespace"
    name = "xml:" + elm.local_part;
else
    [pfx, found] = lookupPrefix_(scope2, elm.nsuri, true);
    if ~found
        pfx = elm.nspfx;
        emitted(end + 1, :) = [pfx, elm.nsuri];
        scope2{end} = emitted;
    end
    if pfx == ""
        name = elm.local_part;
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
            rendered = "xml:" + local;
        else
            [apfx, afound] = lookupPrefix_(scope2, uri, false);
            if ~afound
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

open_tag = "<" + name + decls_str + attrs_str;

% -- 5. content + libxml2 pretty-print indentation --------------------------
kids = elm.to_array();
txt = elm.text_raw_();   % raw stored text (bypass CT_* .text shadow, D10)

% libxml2 format decision: an element is indented (format=1) ONLY when the
% inherited fmt flag is still on AND it has NO text-node children. In this tree
% model a "text node" is a non-None leading text OR a non-None tail on any
% child. A text node among the children turns formatting OFF for THIS element
% AND its whole subtree (fmt=false propagates down and is never re-enabled --
% libxml2 zeroes the flag on descent; Gate-2 probe vectors c2/c3).
has_text_node = ~isequal(txt, []);
if ~has_text_node
    for k = 1:numel(kids)
        if ~isequal(kids(k).tail, [])
            has_text_node = true;
            break
        end
    end
end
child_fmt = fmt && ~has_text_node;

if isempty(kids) && isequal(txt, [])
    % empty element (no children, text None): self-closing (H3 tri-state)
    s = open_tag + "/>";
    return
end

if isempty(kids)
    % leaf with text (or text == "" empty string): inline, no indentation
    inner = "";
    if ~isequal(txt, [])
        inner = escapeText_(txt);
    end
    s = open_tag + ">" + inner + "</" + name + ">";
    return
end

if ~child_fmt
    % format OFF (mixed/text content here, or inherited from a mixed ancestor):
    % emit inline exactly like the compact serializer; fmt=false propagates to
    % the entire subtree (a descendant with all-element children stays inline).
    inner = "";
    if ~isequal(txt, [])
        inner = escapeText_(txt);
    end
    for k = 1:numel(kids)
        inner = inner + writePretty_(kids(k), scope2, level + 1, false);
        ct = kids(k).tail;
        if ~isequal(ct, [])
            inner = inner + escapeText_(ct);
        end
    end
    s = open_tag + ">" + inner + "</" + name + ">";
    return
end

% all-element children, format ON: pretty-print -- each child on its own
% indented line, closing tag indented to this element's level.
child_indent = indent_(level + 1);
inner = "";
for k = 1:numel(kids)
    inner = inner + newline + child_indent + writePretty_(kids(k), scope2, level + 1, true);
end
s = open_tag + ">" + inner + newline + indent_(level) + "</" + name + ">";
end

function ind = indent_(level)
% INDENT_ libxml2 default indent: two spaces per level.
if level <= 0
    ind = "";
else
    ind = string(repmat(' ', 1, 2 * level));
end
end

function [pfx, found] = lookupPrefix_(scope, uri, allow_default)
% Verbatim copy of serialize_part_xml.lookupPrefix_.
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
% Verbatim copy of serialize_part_xml.isShadowed_.
tf = false;
for L = lvl + 1:numel(scope)
    if any(scope{L}(:, 1) == p)
        tf = true;
        return
    end
end
end

function pfx = inventPrefix_(scope)
% Verbatim copy of serialize_part_xml.inventPrefix_.
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
% Verbatim copy of serialize_part_xml.escapeText_ (H7).
out = replace(s, ["&", "<", ">", string(char(13))], ...
    ["&amp;", "&lt;", "&gt;", "&#13;"]);
end

function out = escapeAttr_(s)
% Verbatim copy of serialize_part_xml.escapeAttr_ (H7).
out = replace(s, ...
    ["&", "<", ">", """", string(newline), string(char(9)), string(char(13))], ...
    ["&amp;", "&lt;", "&gt;", "&quot;", "&#10;", "&#9;", "&#13;"]);
end

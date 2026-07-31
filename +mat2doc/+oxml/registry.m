function cls_name = registry(clark_name)
% REGISTRY Custom element-class lookup: Clark tag name -> MATLAB class name.
%
%   cls_name = MAT2DOC.OXML.REGISTRY(clark_name) returns the fully qualified
%   MATLAB class name (string) registered for the element tag clark_name
%   ("{uri}local"), or "" when no custom class is registered -- the caller
%   (createElement / the parser) then falls back to plain XmlElement, mirroring
%   lxml's fallback to _Element for unregistered tags.
%
%   This is the MATLAB analogue of python-docx's parse-time class lookup:
%   `element_class_lookup = etree.ElementNamespaceClassLookup()` configured on
%   `oxml_parser` (docx/oxml/parser.py lines 18-20), populated by the
%   `register_element_cls(tag, cls)` calls. lxml keys the lookup by
%   (namespace URI, local part) -- register_element_cls lines 39-41 -- which is
%   exactly the Clark name, so the table here is keyed by Clark name.
%
%   REGISTRATION TABLE POLICY (design.md section 2 "Factories / registries"):
%   an explicit static table, one line per Python `register_element_cls` call,
%   kept in Python source order, audited line-by-line. Each WP that ports CT_*
%   classes appends its registration lines. Count guard (H10): 120
%   register_element_cls calls TOTAL for docx, all in docx/oxml/__init__.py
%   (catalogs\docx_catalog.json). The SEPARATE 5-class OPC lookup in
%   docx/opc/oxml.py (<Default>/<Override>/<Types>/<Relationship>/
%   <Relationships>, bound via ct_namespace[...] / pr_namespace[...] not
%   register_element_cls) is its OWN element_class_lookup in docx and lands
%   with the OPC-layer WP.
%
%   TABLE CONTENT: the 120 main-map rows (docx/oxml/__init__.py
%   register_element_cls calls) are appended by their CT_* WPs in
%   docx/oxml/__init__.py source order and remain EMPTY here until those WPs
%   land. SEPARATELY, P1-4 merges the 5 OPC element classes (CT_Default /
%   CT_Override / CT_Types / CT_Relationship / CT_Relationships) into this same
%   lookup, keyed by their RAW CLARK NAMES (plan-audit planaudit_2026-07-25
%   condition B2, option A). In docx these 5 live in a SEPARATE
%   element_class_lookup on the OPC parser (docx/opc/oxml.py:240-247), bound via
%   ct_namespace[...]/pr_namespace[...] rather than register_element_cls. Merging
%   them here is behavior-preserving because the ct/pr namespaces are DISJOINT
%   from the main w:* namespaces, so no tag resolves differently under one merged
%   table than under docx's two separate lookups (see +opc/+oxml/parse_xml.m).
%   The 5 OPC rows are added by raw Clark name via registerClark_ -- NOT via
%   registerElementCls_, which would resolve the ct/pr prefixes through the main
%   nsmap (mat2doc.oxml.nsmap has NO ct/pr) and hard-error.
%
%   COUNT GUARD (H10): 120 main-map rows (target, tracked as CT_* WPs land) + 5
%   OPC rows (present now) = 125 total. The two groups are tracked separately.
%
%   Inputs:  clark_name - (1,1) string, e.g. "{http://.../main}p"
%   Outputs: cls_name   - (1,1) string, e.g. "mat2doc.oxml.CT_P", or ""
%
%   Example:
%       cls = mat2doc.oxml.registry(mat2doc.oxml.qn("w:p"));   % "" until P1-x
%
%   Ported from python-docx v1.2.0: src/docx/oxml/parser.py::register_element_cls
%   + element_class_lookup (lines 18-41; registration blocks in
%   docx/oxml/__init__.py pending their CT_* WPs)

arguments
    clark_name (1,1) string
end
persistent map
if isempty(map)
    map = buildRegistry_();
end
if isKey(map, clark_name)
    cls_name = map(clark_name);
else
    cls_name = "";
end
end

function map = buildRegistry_()
% BUILDREGISTRY_ The explicit registration table (built once).
map = dictionary(string.empty(0, 1), string.empty(0, 1));
% -------------------------------------------------------------------------
% MAIN-MAP rows (docx/oxml/__init__.py): one registerElementCls_ line per Python
% register_element_cls call, in docx/oxml/__init__.py source order. Lines are
% appended by the WP that ports the corresponding CT_* class, e.g.:
%
%   map = registerElementCls_(map, "w:document", "mat2doc.oxml.CT_Document");
%
% The H10 dispatch-matrix probe (row count vs Python, target 120) applies as
% CT_* rows land. FIRST row added by P1-7 (coreprops): cp:coreProperties. The
% remaining 119 are appended by their CT_* WPs in docx/oxml/__init__.py order.
map = registerElementCls_(map, "cp:coreProperties", ...
    "mat2doc.oxml.coreprops.CT_CoreProperties");   % __init__.py:96 (P1-7)
map = registerElementCls_(map, "w:body", ...
    "mat2doc.oxml.document.CT_Body");              % __init__.py:100 (P2-3)
map = registerElementCls_(map, "w:document", ...
    "mat2doc.oxml.document.CT_Document");          % __init__.py:101 (P2-3)
% -- font block (docx/oxml/__init__.py:198-225), P4-1a. 28 rows in __init__.py
% -- source order (alphabetical by tag). 21 target the shared CT_OnOff/CT_String
% -- (C3 fold-in), the other 7 target the +text font children / CT_RPr.
map = registerElementCls_(map, "w:b",          "mat2doc.oxml.shared.CT_OnOff");            % __init__.py:198
map = registerElementCls_(map, "w:bCs",        "mat2doc.oxml.shared.CT_OnOff");            % __init__.py:199
map = registerElementCls_(map, "w:caps",       "mat2doc.oxml.shared.CT_OnOff");            % __init__.py:200
map = registerElementCls_(map, "w:color",      "mat2doc.oxml.text.CT_Color");              % __init__.py:201
map = registerElementCls_(map, "w:cs",         "mat2doc.oxml.shared.CT_OnOff");            % __init__.py:202
map = registerElementCls_(map, "w:dstrike",    "mat2doc.oxml.shared.CT_OnOff");            % __init__.py:203
map = registerElementCls_(map, "w:emboss",     "mat2doc.oxml.shared.CT_OnOff");            % __init__.py:204
map = registerElementCls_(map, "w:highlight",  "mat2doc.oxml.text.CT_Highlight");          % __init__.py:205
map = registerElementCls_(map, "w:i",          "mat2doc.oxml.shared.CT_OnOff");            % __init__.py:206
map = registerElementCls_(map, "w:iCs",        "mat2doc.oxml.shared.CT_OnOff");            % __init__.py:207
map = registerElementCls_(map, "w:imprint",    "mat2doc.oxml.shared.CT_OnOff");            % __init__.py:208
map = registerElementCls_(map, "w:noProof",    "mat2doc.oxml.shared.CT_OnOff");            % __init__.py:209
map = registerElementCls_(map, "w:oMath",      "mat2doc.oxml.shared.CT_OnOff");            % __init__.py:210
map = registerElementCls_(map, "w:outline",    "mat2doc.oxml.shared.CT_OnOff");            % __init__.py:211
map = registerElementCls_(map, "w:rFonts",     "mat2doc.oxml.text.CT_Fonts");              % __init__.py:212
map = registerElementCls_(map, "w:rPr",        "mat2doc.oxml.text.CT_RPr");                % __init__.py:213
map = registerElementCls_(map, "w:rStyle",     "mat2doc.oxml.shared.CT_String");           % __init__.py:214
map = registerElementCls_(map, "w:rtl",        "mat2doc.oxml.shared.CT_OnOff");            % __init__.py:215
map = registerElementCls_(map, "w:shadow",     "mat2doc.oxml.shared.CT_OnOff");            % __init__.py:216
map = registerElementCls_(map, "w:smallCaps",  "mat2doc.oxml.shared.CT_OnOff");            % __init__.py:217
map = registerElementCls_(map, "w:snapToGrid", "mat2doc.oxml.shared.CT_OnOff");            % __init__.py:218
map = registerElementCls_(map, "w:specVanish", "mat2doc.oxml.shared.CT_OnOff");            % __init__.py:219
map = registerElementCls_(map, "w:strike",     "mat2doc.oxml.shared.CT_OnOff");            % __init__.py:220
map = registerElementCls_(map, "w:sz",         "mat2doc.oxml.text.CT_HpsMeasure");         % __init__.py:221
map = registerElementCls_(map, "w:u",          "mat2doc.oxml.text.CT_Underline");          % __init__.py:222
map = registerElementCls_(map, "w:vanish",     "mat2doc.oxml.shared.CT_OnOff");            % __init__.py:223
map = registerElementCls_(map, "w:vertAlign",  "mat2doc.oxml.text.CT_VerticalAlignRun");   % __init__.py:224
map = registerElementCls_(map, "w:webHidden",  "mat2doc.oxml.shared.CT_OnOff");            % __init__.py:225
% -- run block (docx/oxml/__init__.py:72-78), P4-1b. The 6 run + run-content
% -- element classes from oxml/text/run.py. (w:drawing @__init__.py:58 ->
% -- CT_Drawing is OUT of scope, owned by the drawing WP; w:lastRenderedPageBreak
% -- @:74 -> CT_LastRenderedPageBreak is now registered in the P4-3 block below;
% -- w:tab @:249 -> CT_TabStop is owned by the parfmt WP.) Listed in __init__.py
% -- source order.
map = registerElementCls_(map, "w:br",             "mat2doc.oxml.text.CT_Br");             % __init__.py:72
map = registerElementCls_(map, "w:cr",             "mat2doc.oxml.text.CT_Cr");             % __init__.py:73
map = registerElementCls_(map, "w:noBreakHyphen",  "mat2doc.oxml.text.CT_NoBreakHyphen");  % __init__.py:75
map = registerElementCls_(map, "w:ptab",           "mat2doc.oxml.text.CT_PTab");           % __init__.py:76
map = registerElementCls_(map, "w:r",              "mat2doc.oxml.text.CT_R");              % __init__.py:77
map = registerElementCls_(map, "w:t",              "mat2doc.oxml.text.CT_Text");           % __init__.py:78
% -- paragraph block (docx/oxml/__init__.py:227-229), P4-2. The one CT_* from
% -- oxml/text/paragraph.py.
map = registerElementCls_(map, "w:p",   "mat2doc.oxml.text.CT_P");   % __init__.py:229
% -- parfmt block (docx/oxml/__init__.py:231-251), P4-2. The 6 parfmt CT_* classes
% -- + the shared CT_OnOff/CT_String rows this block registers, in __init__.py
% -- source order. NOTE w:outlineLvl @:244 -> CT_DecimalNumber was DEFERRED by
% -- P4-2 (the shared CT_DecimalNumber class was not yet ported); P4-6 PORTS
% -- CT_DecimalNumber (oxml/shared.py) and CLOSES that deferral -- the row below
% -- is now LIVE. w:outlineLvl still round-trips byte-identically (CT_PPr never
% -- reads .val on it); the only change is that it now transits CT_DecimalNumber.
map = registerElementCls_(map, "w:ind",             "mat2doc.oxml.text.CT_Ind");        % __init__.py:240
map = registerElementCls_(map, "w:jc",              "mat2doc.oxml.text.CT_Jc");         % __init__.py:241
map = registerElementCls_(map, "w:keepLines",       "mat2doc.oxml.shared.CT_OnOff");    % __init__.py:242
map = registerElementCls_(map, "w:keepNext",        "mat2doc.oxml.shared.CT_OnOff");    % __init__.py:243
map = registerElementCls_(map, "w:outlineLvl",      "mat2doc.oxml.shared.CT_DecimalNumber"); % __init__.py:244 (deferral CLOSED by P4-6)
map = registerElementCls_(map, "w:pageBreakBefore", "mat2doc.oxml.shared.CT_OnOff");    % __init__.py:245
map = registerElementCls_(map, "w:pPr",             "mat2doc.oxml.text.CT_PPr");        % __init__.py:246
map = registerElementCls_(map, "w:pStyle",          "mat2doc.oxml.shared.CT_String");   % __init__.py:247
map = registerElementCls_(map, "w:spacing",         "mat2doc.oxml.text.CT_Spacing");    % __init__.py:248
map = registerElementCls_(map, "w:tab",             "mat2doc.oxml.text.CT_TabStop");    % __init__.py:249
map = registerElementCls_(map, "w:tabs",            "mat2doc.oxml.text.CT_TabStops");   % __init__.py:250
map = registerElementCls_(map, "w:widowControl",    "mat2doc.oxml.shared.CT_OnOff");    % __init__.py:251
% -- hyperlink + pagebreak block (docx/oxml/__init__.py:67, :74), P4-3 -- the LAST
% -- oxml WP of P4. Closes the two P4 carry-forwards: w:hyperlink (CT_P.text over a
% -- hyperlink now returns the hyperlink's concatenated run text) and
% -- w:lastRenderedPageBreak (the rendered-page-break element + its split
% -- machinery; also un-stubs the CLASS resolution in CT_R/CT_P.lastRenderedPageBreaks).
% -- Listed in __init__.py source order (hyperlink block @:67 precedes the
% -- text-related block @:74).
map = registerElementCls_(map, "w:hyperlink", ...
    "mat2doc.oxml.text.CT_Hyperlink");             % __init__.py:67 (P4-3)
map = registerElementCls_(map, "w:lastRenderedPageBreak", ...
    "mat2doc.oxml.text.CT_LastRenderedPageBreak"); % __init__.py:74 (P4-3)
% -- styles block (docx/oxml/__init__.py:136-149), P4-6 -- FIRST WP of the styles
% -- chain. The 4 styles CT_* classes (CT_LatentStyles/CT_LsdException/CT_Style/
% -- CT_Styles) + the shared CT_String/CT_OnOff/CT_DecimalNumber rows this block
% -- registers, in __init__.py source order. ALL 12 tags appear in word/styles.xml
% -- (see boundary audit "12 styles tags"); registering them makes styles.xml
% -- transit these CT_* classes on load (the M1 parse path) and lights up
% -- CT_Style's .val accessors (name_val/basedOn_val/next_style read .val on the
% -- CT_String/CT_DecimalNumber/CT_OnOff children). w:uiPriority is
% -- CT_DecimalNumber's FIRST registration. The numbering/table CT_DecimalNumber
% -- rows (w:gridSpan/w:numId/...) stay DEFERRED to P6/P8 (they do not appear in
% -- styles.xml). ZERO exact-class XmlElement pins target these 12 tags (boundary
% -- audit), so no stale-pin flip.
map = registerElementCls_(map, "w:basedOn",         "mat2doc.oxml.shared.CT_String");        % __init__.py:138
map = registerElementCls_(map, "w:latentStyles",    "mat2doc.oxml.styles.CT_LatentStyles");  % __init__.py:139
map = registerElementCls_(map, "w:locked",          "mat2doc.oxml.shared.CT_OnOff");         % __init__.py:140
map = registerElementCls_(map, "w:lsdException",    "mat2doc.oxml.styles.CT_LsdException");   % __init__.py:141
map = registerElementCls_(map, "w:name",            "mat2doc.oxml.shared.CT_String");        % __init__.py:142
map = registerElementCls_(map, "w:next",            "mat2doc.oxml.shared.CT_String");        % __init__.py:143
map = registerElementCls_(map, "w:qFormat",         "mat2doc.oxml.shared.CT_OnOff");         % __init__.py:144
map = registerElementCls_(map, "w:semiHidden",      "mat2doc.oxml.shared.CT_OnOff");         % __init__.py:145
map = registerElementCls_(map, "w:style",           "mat2doc.oxml.styles.CT_Style");          % __init__.py:146
map = registerElementCls_(map, "w:styles",          "mat2doc.oxml.styles.CT_Styles");         % __init__.py:147
map = registerElementCls_(map, "w:uiPriority",      "mat2doc.oxml.shared.CT_DecimalNumber");  % __init__.py:148
map = registerElementCls_(map, "w:unhideWhenUsed",  "mat2doc.oxml.shared.CT_OnOff");         % __init__.py:149
% -------------------------------------------------------------------------
% OPC rows (P1-4; docx/opc/oxml.py:240-247): the 5 OPC element classes, keyed by
% RAW CLARK NAME (condition B2). The Clark URIs come from mat2doc.opc.NAMESPACE
% so they are IDENTICAL to the xmlns the CT_*.new factories emit and the parser
% resolves -- guaranteeing the parsed element's Clark key matches this row. NOT
% routed through registerElementCls_ (ct/pr are absent from the main nsmap).
NS = mat2doc.opc.NAMESPACE;
map = registerClark_(map, "{" + NS.OPC_CONTENT_TYPES + "}Default", ...
    "mat2doc.opc.oxml.CT_Default");        % ct_namespace["Default"]  opc/oxml.py:241
map = registerClark_(map, "{" + NS.OPC_CONTENT_TYPES + "}Override", ...
    "mat2doc.opc.oxml.CT_Override");       % ct_namespace["Override"] opc/oxml.py:242
map = registerClark_(map, "{" + NS.OPC_CONTENT_TYPES + "}Types", ...
    "mat2doc.opc.oxml.CT_Types");          % ct_namespace["Types"]    opc/oxml.py:243
map = registerClark_(map, "{" + NS.OPC_RELATIONSHIPS + "}Relationship", ...
    "mat2doc.opc.oxml.CT_Relationship");   % pr_namespace["Relationship"]  opc/oxml.py:246
map = registerClark_(map, "{" + NS.OPC_RELATIONSHIPS + "}Relationships", ...
    "mat2doc.opc.oxml.CT_Relationships");  % pr_namespace["Relationships"] opc/oxml.py:247
end

function map = registerClark_(map, clark, cls_name)
% REGISTERCLARK_ Register cls_name for a RAW Clark tag key (no prefix resolution).
%   Used for OPC classes whose ct/pr prefixes are not in the main nsmap, so the
%   NamespacePrefixedTag path in registerElementCls_ cannot be used.
arguments
    map dictionary
    clark (1,1) string
    cls_name (1,1) string
end
if isKey(map, clark)
    error("mat2doc:InternalError", ...
        "duplicate element-class registration for '%s'", clark);
end
map(clark) = cls_name;
end

function map = registerElementCls_(map, tag, cls_name) %#ok<DEFNU> -- invoked by table lines as CT_* WPs add them
% REGISTERELEMENTCLS_ Register cls_name to be constructed for elements having tag.
%
%   tag is a string of the form "nspfx:tagroot", e.g. "w:document".
%
%   Ported from python-docx v1.2.0: src/docx/oxml/parser.py::register_element_cls
%   (lines 32-41)
arguments
    map dictionary
    tag (1,1) string
    cls_name (1,1) string
end
% Python (lines 39-41): keys the class into the lookup by
% (nsmap[nspfx], tagroot) -- equivalent to the Clark name.
nsptag = mat2doc.oxml.NamespacePrefixedTag(tag);
key = nsptag.clark_name;
% Python dict assignment would silently overwrite; upstream never registers the
% same tag twice, so a duplicate here can only be a table-generation defect --
% fail loudly (build-time guard, not output-visible behavior).
if isKey(map, key)
    error("mat2doc:InternalError", ...
        "duplicate element-class registration for '%s'", tag);
end
map(key) = cls_name;
end

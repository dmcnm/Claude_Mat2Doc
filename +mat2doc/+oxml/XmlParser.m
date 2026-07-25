classdef XmlParser < handle
% XMLPARSER Own OOXML-subset XML parser (D-001) -- the read side of the XML layer.
%
%   p = MAT2DOC.OXML.XMLPARSER(text) wraps a decoded XML character string;
%   p.parse() returns the root mat2doc.oxml.XmlElement of an order-preserving
%   tree. Use the package function mat2doc.oxml.parse_xml (the analogue of
%   python-docx's parse_xml) as the public entry point -- this class is the
%   internal engine it drives.
%
%   WHY OWN PARSER (D-001, adopted for Mat2Doc per
%   validation\summary\decision_2026-07-25_mat2doc_deviation_preadoption.md):
%   Spike S1 proved matlab.io.xml.dom returns attributes ALPHABETICALLY SORTED
%   and its writer re-scrambles them, so byte-conservative round-tripping of
%   OOXML parts (attribute order identical to what Word wrote) is impossible
%   through the DOM. This purpose-built parser preserves document order
%   end-to-end, closing the parse->tree->serialize->bytes round-trip (design.md
%   section 3; deviation_ledger.md D-001).
%
%   SCOPE (design.md section 3 -- the OOXML subset): XML declaration (parsed and
%   skipped; the tree does not store it -- the serializer re-emits the canonical
%   declaration), UTF-8 decode (H2), elements and attributes IN DOCUMENT ORDER
%   (H11 -- the whole point), character data with the text/tail tri-state (H3),
%   the five predefined entities plus decimal AND hex numeric character
%   references ONLY (resolve_entities=False -> NO DTD/custom entities), CDATA
%   sections (unwrapped to text), namespace declarations recorded as ordered
%   decls stored WHERE DECLARED, and replication of remove_blank_text=True.
%   Matched to lxml 5.3.0 (libxml2 2.13.9) under the docx parser config
%   oxml_parser = etree.XMLParser(remove_blank_text=True,
%   resolve_entities=False) (docx/oxml/parser.py:19).
%
%   RESERVED xml PREFIX. The prefix xml is predefined by W3C Namespaces section
%   3 (bound to http://www.w3.org/XML/1998/namespace) and resolves with NO
%   declaration -- always in scope, never recorded as an xmlns:xml decl, never
%   emitted by the serializer (Word writes <w:t xml:space="preserve"> and
%   xml:lang). See resolvePrefix_.
%
%   CHAR-RANGE GUARD. Every numeric character reference is range-checked against
%   the XML 1.0 Char production (validateCharRef_), so the parser REJECTS
%   exactly what lxml rejects (&#0; &#xB; &#xD800; &#xFFFF; &#x110000;, and a
%   literal < in an attribute value) rather than emitting bytes that are
%   themselves invalid XML.
%
%   DELIBERATE DIVERGENCES from lxml (design.md section 3), both DEAD PATHS in
%   OOXML (no OOXML part carries either construct), documented not silent:
%     - Comments and processing instructions are SKIPPED (lxml keeps them as
%       tree nodes and re-serializes them). The XmlElement tree has no
%       comment/PI node type. Accepted-unreachable, NO D-number -> V-CMT.
%     - A DOCTYPE declaration is REJECTED with mat2doc:XMLSyntaxError (lxml
%       accepts and silently ignores it). No DTD support; resolve_entities=False
%       could not honor DTD entities anyway; avoids XXE. Asymmetric
%       reject-vs-accept -> ledgered D-006 (adopted for Mat2Doc, SIGNED-PROVISIONAL).
%
%   The parse walk applies the class registry (registry.m / createElement's
%   dispatch) at each element so registered CT_* classes instantiate; the
%   registry is empty before the CT_* WPs, so every node is an XmlElement now,
%   but the hook is wired.
%
%   Inputs:  text - (1,1) string / char row: already UTF-8-decoded XML content
%                   (mat2doc.oxml.parse_xml does the bytes -> string decode).
%   Outputs: p    - scalar mat2doc.oxml.XmlParser; call p.parse() for the root.
%
%   Example:
%       % Drive the engine directly (parse_xml is the normal entry point):
%       p    = mat2doc.oxml.XmlParser("<w:r xmlns:w=""urn:w""><w:t>hi</w:t></w:r>");
%       root = p.parse();
%       root.tag                 % "{urn:w}r"
%       kid  = root.to_array();  % child inherits prefix w from the scope stack
%       kid.tag                  % "{urn:w}t"
%       kid.text                 % "hi"
%
%   Design-realization of the OOXML-subset read side (D-001; see design.md
%   section 3) -- no single Python def is its origin; the behavior realized is
%   lxml's etree.fromstring(xml, oxml_parser) read path.
%
%   Ported from python-docx v1.2.0: src/docx/oxml/parser.py::parse_xml /
%   oxml_parser (lines 18-29; design-realization, D-001)

    properties (Access = private)
        buf (1,:) char = ''      % decoded, line-end-normalized document text
        n (1,1) double = 0       % numel(buf)
        pos (1,1) double = 1     % 1-based cursor
        scope (1,:) cell = {}     % stack of Nx2 string [prefix, URI] own-decls, outermost first
    end

    properties (Constant, Access = private)
        WS = sprintf(' \t\n\r')   % XML whitespace: space, tab, LF, CR
    end

    methods
        function obj = XmlParser(text)
            % XMLPARSER Wrap a decoded XML string for parsing.
            %   text: (1,1) string or char row -- already UTF-8-decoded content.
            arguments
                text (1,1) string
            end
            c = char(text);
            % Strip a leading BOM (U+FEFF) if present.
            if ~isempty(c) && double(c(1)) == 65279
                c(1) = [];
            end
            % XML 1.0 line-end normalization: CRLF -> LF, then lone CR -> LF.
            c = strrep(c, sprintf('\r\n'), sprintf('\n'));
            c = strrep(c, sprintf('\r'), sprintf('\n'));
            obj.buf = c;
            obj.n = numel(c);
            obj.pos = 1;
        end

        function root = parse(obj)
            % PARSE Parse the whole document; return the root XmlElement.
            obj.skipProlog_();
            if obj.pos > obj.n || obj.buf(obj.pos) ~= '<'
                if obj.pos > obj.n
                    obj.err_("Document is empty");
                else
                    obj.err_("Start tag expected, '<' not found");
                end
            end
            root = obj.parseElement_();
            obj.skipMisc_();
            if obj.pos <= obj.n
                obj.err_("Extra content at the end of the document");
            end
        end
    end

    methods (Access = private)

        % ------------------------------------------------------------------
        % prolog / misc
        % ------------------------------------------------------------------

        function skipProlog_(obj)
            % Optional XML declaration must be the very first thing.
            if obj.startsWith_("<?xml") && obj.pos + 5 <= obj.n + 1
                nxt = obj.buf(obj.pos + 5);
                if any(nxt == obj.WS) || nxt == '?'
                    e = strfind(obj.buf(obj.pos:end), '?>');
                    if isempty(e)
                        obj.err_("Malformed XML declaration");
                    end
                    obj.pos = obj.pos + e(1) + 1;   % past "?>"
                end
            end
            obj.skipMisc_();
        end

        function skipMisc_(obj)
            % Skip whitespace, comments, PIs (and reject DOCTYPE) between the
            % declaration and the root, and after the root.
            while obj.pos <= obj.n
                if any(obj.buf(obj.pos) == obj.WS)
                    obj.pos = obj.pos + 1;
                elseif obj.startsWith_("<!--")
                    obj.skipComment_();
                elseif obj.startsWith_("<!DOCTYPE") || obj.startsWith_("<!doctype")
                    obj.err_("DOCTYPE declarations are not supported " + ...
                        "(no DTD support; OOXML parts carry no DTD)");
                elseif obj.startsWith_("<?")
                    obj.skipPI_();
                else
                    break
                end
            end
        end

        function skipComment_(obj)
            e = strfind(obj.buf(obj.pos:end), '-->');
            if isempty(e)
                obj.err_("Comment not terminated");
            end
            obj.pos = obj.pos + e(1) + 2;   % past "-->" (3-char terminator)
        end

        function skipPI_(obj)
            e = strfind(obj.buf(obj.pos:end), '?>');
            if isempty(e)
                obj.err_("Processing instruction not terminated");
            end
            obj.pos = obj.pos + e(1) + 1;   % past "?>"
        end

        % ------------------------------------------------------------------
        % element
        % ------------------------------------------------------------------

        function elm = parseElement_(obj)
            % pos is at '<' of a start tag.
            obj.pos = obj.pos + 1;                 % consume '<'
            name = obj.parseName_();
            if name == ""
                obj.err_("StartTag: invalid element name");
            end
            [ownDecls, attrRawNames, attrRawVals, selfClose] = obj.parseAttrs_();

            % Push this element's own declarations before resolving its tag and
            % attribute prefixes (an xmlns on the element applies to itself).
            obj.scope{end + 1} = ownDecls;

            % Resolve tag namespace from the in-scope stack (handles ancestor-
            % declared and non-fixed prefixes such as mc/wp14/wps).
            [pfx, local] = obj.splitName_(name);
            uri = obj.resolvePrefix_(pfx, name);

            % Apply the class registry (createElement/registry dispatch) keyed by
            % the RESOLVED Clark name -- correct for default-ns children.
            if uri == ""
                clark = local;
            else
                clark = "{" + uri + "}" + local;
            end
            cls = mat2doc.oxml.registry(clark);
            if cls == ""
                % lxml fallback to the plain element class. The 3-arg ctor takes
                % the parser-resolved URI verbatim.
                elm = mat2doc.oxml.XmlElement(name, ownDecls, uri);
            else
                elm = feval(cls, name, ownDecls);
                elm.setResolvedNsuri_(uri);   % make parser resolution authoritative
            end
            % Verbatim ns-decls for the serializer (fixes D-serializer-nsdecl):
            % a PARSED element's stored ns-decls are serialized VERBATIM -- lxml's
            % redundant-decl suppression is a MOVE-time reconciliation effect that
            % never applies to a tree built in place by the parser. The parser
            % attaches children via appendParsed_ (below), which does NOT
            % reconcile, so this flag survives until a public move clears it. This
            % move-time-only reconciliation is the design.md section 3 "Serialize"
            % contract. See XmlElement.markNsVerbatim_ and serialize_part_xml
            % step 1.
            elm.markNsVerbatim_();

            % Attributes: insertion order; resolve prefixed names to Clark keys.
            seen = strings(1, 0);
            for k = 1:numel(attrRawNames)
                an = attrRawNames(k);
                if an == "xmlns" || startsWith(an, "xmlns:")
                    continue   % declarations already captured in ownDecls
                end
                [apfx, alocal] = obj.splitName_(an);
                if apfx == ""
                    key = an;   % unprefixed attrs are NOT in the default ns
                else
                    auri = obj.resolvePrefix_(apfx, an);
                    key = "{" + auri + "}" + alocal;
                end
                if any(seen == key)
                    obj.err_("Attribute " + an + " redefined");
                end
                seen(end + 1) = key; %#ok<AGROW>
                elm.set(key, obj.normalizeAttrValue_(attrRawVals(k)));
            end

            if ~selfClose
                obj.parseContent_(elm, name);
                obj.finalizeBlankText_(elm);   % remove_blank_text replication
            end

            obj.scope(end) = [];   % pop
        end

        function parseContent_(obj, elm, name)
            % Parse children/char-data until the matching end tag. Character data
            % (text runs + CDATA) accumulate into pendingText and are flushed to
            % elm.text (before the first child) or the previous child's tail on
            % the next child / end tag.
            pending = '';
            hasPending = false;
            lastChild = mat2doc.oxml.XmlElement.empty(1, 0);
            while true
                if obj.pos > obj.n
                    obj.err_("Premature end of data in tag " + name);
                end
                if obj.startsWith_("</")
                    obj.flushPending_(elm, lastChild, pending, hasPending);
                    obj.parseEndTag_(name);
                    return
                elseif obj.startsWith_("<![CDATA[")
                    [pending, hasPending] = obj.readCData_(pending);
                elseif obj.startsWith_("<!--")
                    obj.skipComment_();     % SKIP (design section 3); dead path in OOXML
                elseif obj.startsWith_("<?")
                    obj.skipPI_();          % SKIP; dead path in OOXML
                elseif obj.buf(obj.pos) == '<'
                    % child element: flush pending text to its slot, then recurse
                    obj.flushPending_(elm, lastChild, pending, hasPending);
                    pending = ''; hasPending = false;
                    child = obj.parseElement_();
                    elm.appendParsed_(child);   % non-reconciling attach: keep parsed verbatim ns-decls (H8)
                    lastChild = child;
                else
                    seg = obj.readText_();
                    pending = [pending, seg]; %#ok<AGROW>
                    hasPending = true;
                end
            end
        end

        function flushPending_(~, elm, lastChild, pending, hasPending)
            if ~hasPending
                return   % no char content -> leave slot as None (H3)
            end
            if isempty(lastChild)
                elm.setTextRaw_(string(pending));   % raw (bypass CT_* .text shadow, D10)
            else
                lastChild.setTailRaw_(string(pending));
            end
        end

        function parseEndTag_(obj, name)
            obj.pos = obj.pos + 2;                 % consume "</"
            endName = obj.parseName_();
            obj.skipWs_();
            if obj.pos > obj.n || obj.buf(obj.pos) ~= '>'
                obj.err_("Expected '>' to close end tag " + name);
            end
            obj.pos = obj.pos + 1;                 % consume '>'
            if endName ~= name
                obj.err_("Opening and ending tag mismatch: " + name + ...
                    " and " + endName);
            end
        end

        function finalizeBlankText_(obj, elm)
            % REMOVE_BLANK_TEXT replication (design section 3): drop
            % whitespace-only text where the element has ELEMENT CHILDREN.
            %
            % Rule (probe-derived, positional): for an element with >=1 child
            % element, a leading text (elm.text) that is whitespace-only is always
            % dropped; a child's tail that is whitespace-only is dropped UNLESS a
            % non-whitespace text/tail has already been seen earlier in the element
            % (mixed content, in which case interior blanks are kept). An element
            % with NO element children keeps its text verbatim (whitespace-only
            % leaf preserved, e.g. <w:t>   </w:t>). This is a NO-OP on every stored
            % OOXML part that carries no inter-element whitespace; it matches lxml
            % exactly for pure-element content, pure-text leaves and leading/simple
            % mixed content.
            kids = elm.to_array();
            if isempty(kids)
                return
            end
            mixed = false;
            t = elm.text_raw_();
            if ~isequal(t, [])
                if obj.isWsOnly_(t)
                    elm.setTextRaw_([]);
                else
                    mixed = true;
                end
            end
            for k = 1:numel(kids)
                tl = kids(k).tail;
                if isequal(tl, [])
                    continue
                end
                if obj.isWsOnly_(tl)
                    if ~mixed
                        kids(k).setTailRaw_([]);
                    end
                else
                    mixed = true;
                end
            end
        end

        % ------------------------------------------------------------------
        % attributes / names
        % ------------------------------------------------------------------

        function [ownDecls, names, vals, selfClose] = parseAttrs_(obj)
            names = strings(1, 0);
            vals = strings(1, 0);
            declList = strings(0, 2);
            while true
                obj.skipWs_();
                if obj.pos > obj.n
                    obj.err_("Unexpected end of data in start tag");
                end
                ch = obj.buf(obj.pos);
                if ch == '/'
                    if obj.startsWith_("/>")
                        obj.pos = obj.pos + 2;
                        selfClose = true;
                        ownDecls = declList;
                        return
                    end
                    obj.err_("Malformed start tag (expected '/>')");
                elseif ch == '>'
                    obj.pos = obj.pos + 1;
                    selfClose = false;
                    ownDecls = declList;
                    return
                end
                aname = obj.parseName_();
                if aname == ""
                    obj.err_("error parsing attribute name");
                end
                obj.skipWs_();
                if obj.pos > obj.n || obj.buf(obj.pos) ~= '='
                    obj.err_("Specification mandates value for attribute " + aname);
                end
                obj.pos = obj.pos + 1;             % consume '='
                obj.skipWs_();
                rawVal = obj.readAttrValue_();
                if aname == "xmlns"
                    declList(end + 1, :) = ["", obj.normalizeAttrValue_(rawVal)]; %#ok<AGROW>
                elseif startsWith(aname, "xmlns:")
                    declList(end + 1, :) = [extractAfter(aname, "xmlns:"), ...
                        obj.normalizeAttrValue_(rawVal)]; %#ok<AGROW>
                else
                    names(end + 1) = aname;         %#ok<AGROW>
                    vals(end + 1) = rawVal;         %#ok<AGROW> normalized later
                end
            end
        end

        function name = parseName_(obj)
            start = obj.pos;
            while obj.pos <= obj.n
                ch = obj.buf(obj.pos);
                if any(ch == obj.WS) || ch == '/' || ch == '>' || ch == '=' || ch == '<'
                    break
                end
                obj.pos = obj.pos + 1;
            end
            name = string(obj.buf(start:obj.pos - 1));
        end

        function rawVal = readAttrValue_(obj)
            if obj.pos > obj.n
                obj.err_("AttValue: expected quote");
            end
            q = obj.buf(obj.pos);
            if q ~= '"' && q ~= ''''
                obj.err_("AttValue: '""' or ''' expected");
            end
            obj.pos = obj.pos + 1;
            e = find(obj.buf(obj.pos:end) == q, 1);
            if isempty(e)
                obj.err_("AttValue: quote not closed");
            end
            val = obj.buf(obj.pos:obj.pos + e - 2);
            % A literal '<' is never allowed in an attribute value (XML 1.0
            % AttValue production); it must be written &lt;. lxml rejects it --
            % match. &lt; (a reference) is fine: it is not a literal '<' and is
            % expanded later.
            if any(val == '<')
                obj.err_("Unescaped '<' not allowed in attributes values");
            end
            rawVal = string(val);
            obj.pos = obj.pos + e;                 % past closing quote
        end

        % ------------------------------------------------------------------
        % text / cdata / references
        % ------------------------------------------------------------------

        function seg = readText_(obj)
            % Character data up to the next '<'; entity/char refs expanded.
            e = find(obj.buf(obj.pos:end) == '<', 1);
            if isempty(e)
                obj.err_("Premature end of data (unterminated content)");
            end
            raw = obj.buf(obj.pos:obj.pos + e - 2);
            obj.pos = obj.pos + e - 1;             % at the '<'
            if any(raw == '&')
                seg = char(obj.expandRefs_(string(raw)));
            else
                seg = raw;
            end
        end

        function [pending, hasPending] = readCData_(obj, pending)
            obj.pos = obj.pos + 9;                 % consume "<![CDATA["
            e = strfind(obj.buf(obj.pos:end), ']]>');
            if isempty(e)
                obj.err_("CDATA section not terminated");
            end
            content = obj.buf(obj.pos:obj.pos + e(1) - 2);   % literal, no ref expansion
            obj.pos = obj.pos + e(1) + 2;          % past "]]>" (3-char terminator)
            pending = [pending, content];
            hasPending = true;
        end

        function v = normalizeAttrValue_(obj, rawVal)
            % XML attribute-value normalization (CDATA type, no DTD): each literal
            % TAB and LF (CR already LF-normalized globally) becomes a single
            % space; THEN entity/char refs are expanded, so a char ref such as
            % &#9; survives as a real TAB.
            c = char(rawVal);
            c = strrep(c, sprintf('\t'), ' ');
            c = strrep(c, sprintf('\n'), ' ');
            if any(c == '&')
                v = obj.expandRefs_(string(c));
            else
                v = string(c);
            end
        end

        function out = expandRefs_(obj, s)
            % Expand the 5 predefined entities and decimal/hex numeric character
            % references (resolve_entities=False -> nothing else; a named
            % non-predefined entity or a bare '&' is a well-formedness error,
            % matching lxml).
            c = char(s);
            m = numel(c);
            outc = blanks(0);
            i = 1;
            while i <= m
                ch = c(i);
                if ch ~= '&'
                    outc = [outc, ch]; %#ok<AGROW>
                    i = i + 1;
                    continue
                end
                semi = i + find(c(i + 1:end) == ';', 1);
                if isempty(find(c(i + 1:end) == ';', 1))
                    obj.err_("xmlParseEntityRef: no name");
                end
                token = c(i + 1:semi - 1);
                if isempty(token)
                    obj.err_("xmlParseEntityRef: no name");
                end
                if token(1) == '#'
                    if numel(token) >= 2 && (token(2) == 'x' || token(2) == 'X')
                        cp = obj.hex2cp_(token(3:end));
                    else
                        cp = obj.dec2cp_(token(2:end));
                    end
                    obj.validateCharRef_(cp);   % XML 1.0 Char-range guard
                    outc = [outc, obj.cp2chars_(cp)]; %#ok<AGROW>
                else
                    switch token
                        case 'amp';  outc = [outc, '&']; %#ok<AGROW>
                        case 'lt';   outc = [outc, '<']; %#ok<AGROW>
                        case 'gt';   outc = [outc, '>']; %#ok<AGROW>
                        case 'quot'; outc = [outc, '"']; %#ok<AGROW>
                        case 'apos'; outc = [outc, '''']; %#ok<AGROW>
                        otherwise
                            obj.err_("Entity '" + string(token) + "' not defined");
                    end
                end
                i = semi + 1;
            end
            out = string(outc);
        end

        function validateCharRef_(obj, cp)
            % XML 1.0 Char production guard for numeric character references
            % (matches lxml/libxml2 rejection exactly). A numeric char ref must
            % denote a valid XML character:
            %   #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] |
            %   [#x10000-#x10FFFF].
            % NUL, C0 controls other than TAB/LF/CR (incl. #xB, #xC, #x1F),
            % surrogates (#xD800-#xDFFF), #xFFFE/#xFFFF and anything > #x10FFFF
            % are rejected. lxml uses two message cores; both reproduced here.
            if cp > 1114111                        % > #x10FFFF
                obj.err_("xmlParseCharRef: character reference out of bounds");
            end
            valid = cp == 9 || cp == 10 || cp == 13 || ...
                (cp >= 32 && cp <= 55295) || ...       % #x20 - #xD7FF
                (cp >= 57344 && cp <= 65533) || ...    % #xE000 - #xFFFD
                (cp >= 65536 && cp <= 1114111);        % #x10000 - #x10FFFF
            if ~valid
                obj.err_("xmlParseCharRef: invalid xmlChar value " + cp);
            end
        end

        function cp = dec2cp_(obj, digits)
            if isempty(digits) || any(digits < '0' | digits > '9')
                obj.err_("xmlParseCharRef: invalid decimal value");
            end
            cp = str2double(digits);
        end

        function cp = hex2cp_(obj, digits)
            ok = ~isempty(digits) && all((digits >= '0' & digits <= '9') | ...
                (lower(digits) >= 'a' & lower(digits) <= 'f'));
            if ~ok
                obj.err_("xmlParseCharRef: invalid hexadecimal value");
            end
            cp = hex2dec(digits);
        end

        function chs = cp2chars_(~, cp)
            % Unicode code point -> MATLAB (UTF-16) chars; astral via surrogate
            % pair so serialization to UTF-8 (unicode2native) is correct (H2).
            if cp <= 65535
                chs = char(cp);
            else
                v = cp - 65536;
                hi = 55296 + floor(v / 1024);
                lo = 56320 + mod(v, 1024);
                chs = char([hi, lo]);
            end
        end

        % ------------------------------------------------------------------
        % namespaces
        % ------------------------------------------------------------------

        function uri = resolvePrefix_(obj, pfx, name)
            % Reserved prefix 'xml' is predefined by W3C Namespaces section 3,
            % bound to http://www.w3.org/XML/1998/namespace; it needs NO
            % declaration and is ALWAYS in scope (lxml resolves it; Word writes
            % <w:t xml:space="preserve"> and xml:lang). It is never stored as a
            % namespace declaration on the element (nsdecls), and the serializer
            % never emits xmlns:xml for it. The other reserved prefix 'xmlns' is
            % deliberately NOT pre-bound: lxml rejects it as an element/attribute
            % prefix ("Namespace prefix xmlns on ... is not defined"), which the
            % fall-through reject below already reproduces; xmlns / xmlns:*
            % attribute NAMES are consumed as declarations in parseAttrs_ before
            % resolution.
            if pfx == "xml"
                uri = "http://www.w3.org/XML/1998/namespace";
                return
            end
            for lvl = numel(obj.scope):-1:1
                L = obj.scope{lvl};
                for k = 1:size(L, 1)
                    if L(k, 1) == pfx
                        uri = L(k, 2);
                        return
                    end
                end
            end
            if pfx == ""
                uri = "";   % no default namespace in scope -> no namespace
            else
                obj.err_("Namespace prefix " + pfx + " on " + name + ...
                    " is not defined");
            end
        end

        % ------------------------------------------------------------------
        % small helpers
        % ------------------------------------------------------------------

        function [pfx, local] = splitName_(~, name)
            parts = split(name, ":");
            if numel(parts) == 1
                pfx = "";
                local = parts(1);
            else
                pfx = parts(1);
                local = parts(2);
            end
        end

        function tf = startsWith_(obj, lit)
            L = char(lit);
            k = numel(L);
            tf = obj.pos + k - 1 <= obj.n && ...
                strcmp(obj.buf(obj.pos:obj.pos + k - 1), L);
        end

        function skipWs_(obj)
            while obj.pos <= obj.n && any(obj.buf(obj.pos) == obj.WS)
                obj.pos = obj.pos + 1;
            end
        end

        function tf = isWsOnly_(obj, s)
            c = char(s);
            tf = ~isempty(c) && all(c == ' ' | c == sprintf('\t') | ...
                c == sprintf('\n') | c == sprintf('\r'));
        end

        function err_(obj, msg)
            error("mat2doc:XMLSyntaxError", "%s (byte %d)", string(msg), obj.pos);
        end
    end
end

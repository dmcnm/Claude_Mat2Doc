function result = evaluate_xpath(context, xpath_str, ns)
% EVALUATE_XPATH Mini XPath-1.0 engine over the XmlElement tree.
%
%   result = MAT2DOC.OXML.EVALUATE_XPATH(context, xpath_str, ns) evaluates
%   xpath_str against XmlElement context with prefix->URI mapping ns (scalar
%   struct, the P1-2 nsmap() currency), returning EXACTLY what lxml's
%   `_Element.xpath(expr, namespaces=ns)` returns for the closed subset of
%   expressions the docx/lxml surface uses (design.md section XPath):
%
%     - prefixed child paths            w:body/w:p/w:r
%     - self / relative-descendant      .   ./w:x   .//w:x
%     - absolute / absolute-descendant  /w:document   //w:p   //@r:id
%     - wildcard child                  ./*[1]
%     - positional predicate            w:p[1]          (1-based; H1: never shift)
%     - attribute-equality predicate    w:t[@xml:space="preserve"]   w:x[@w:val=%d]
%     - element-existence predicate     w:p[w:pPr]    w:x[w:y[@w:val="%d"]]
%     - attribute result                ./w:pPr/w:rPr/@w:val   //@r:id
%     - text() result                   ./w:p//w:t/text()
%     - ancestor axis                   ancestor::w:tbl
%     - parent step + union group       (../w:x | ../w:y)/w:z[@w:val="%d"]
%
%   P1-3x EXTENSION (docx call-site patterns the pptx-derived engine did not
%   cover; each byte/value-verified against lxml 5.3.0 -- see the audit record
%   validation\mat2doc\audit_P1-3x_xpath_extension.md):
%     - bare top-level union            ./w:p | ./w:tbl   (docx's commonest form;
%                                       operands may be relative, absolute, or
%                                       grouped; result is doc-ordered + deduped)
%     - not() predicate                 ./*[not(self::w:sectPr)]
%     - following-sibling / preceding-sibling axes   ./following-sibling::w:tc
%     - preceding axis                  ./preceding::w:sectPr[1]
%     - position()=n / [last()]         w:p[position()=1]   (./w:r)[last()]
%     - group predicate                 (./w:comment[@w:id='5'])[1]
%     - predicate attribute sub-path    w:style[w:name/@w:val="Heading 1"]
%     - positional predicate on a       //@r:id[2]   ([] -- one attr per element)
%       terminal @attr / text() step
%   H1: positional predicates are 1-based and, on the REVERSE axes
%   (preceding-sibling, preceding, ancestor), position() counts in axis order
%   (nearest first) -- never shifted. H11: unions are document-ordered and
%   identity-deduped (the union trap -- lxml node-sets are sets).
%
%   RETURN TYPE (matches lxml exactly):
%     - expression terminating in /@attr   -> (1,N) string array (attr values)
%     - expression terminating in /text()  -> (1,N) string array (text nodes)
%     - otherwise                          -> (1,N) mat2doc.oxml.XmlElement
%     - NO MATCH -> EMPTY ARRAY of that type (H3: lxml xpath() returns [],
%       NEVER None; an empty node-set is string.empty(1,0) / XmlElement.empty(1,0))
%
%   This is NOT a general XPath engine: any construct outside the verified
%   subset raises mat2doc:XPathError rather than being silently mis-evaluated
%   (design.md section 7).
%
%   RE-PORT PROVENANCE (P1-3a, accelerated re-port): this is the SOLVED Mat2Ppt
%   +oxml xpath evaluator (WP5 + the WP5-C corrective fixes) re-ported verbatim
%   with mat2doc namespacing; docx v1.2.0 is the module source of truth. The
%   evaluator is tree/namespace-agnostic within its subset, so no logic changes
%   were needed -- only mat2ppt->mat2doc identifiers and error ids. The four
%   WP5-C corrective semantics (adopted here mat2doc:-namespaced -- shared
%   deviation, no new D-number) that bring the engine to lxml fidelity so
%   design.md's "never silently mis-evaluated" rule holds with ZERO silent
%   mis-evaluations even WITHIN the subset:
%     - F1  text() reads the RAW text node plus each child element's .tail
%           (the lxml C-level text nodes), never the overridable getText_ shadow
%           (D10): a shadowing w:r / w:br / w:fld returns [] and mixed content
%           yields every self-text + child-tail fragment.
%     - F2  parseable-but-out-of-subset forms still RAISE mat2doc:XPathError
%           instead of mis-evaluating -- an absolute sub-path in a predicate
%           (w:p[//w:a]) or in a group ((/w:document)/w:body), a nested group
%           ((.//w:p)), and a wildcard node test on the ancestor axis
%           (ancestor::*). NOTE (P1-3x): a POSITIONAL predicate on a terminal
%           string step (//@id[2]) is no longer in this raise set -- it is now
%           evaluated positionally per context element (lxml-faithful; -> []).
%     - F3  string (attr / text) results are identity-deduped by node, matching
%           lxml node-set semantics (//w:x//@id -> ["X1","X2"], not "X2" twice).
%     - WPC-F1  a tail text node sorts AFTER its element's whole subtree (sort
%           key [docKey(kid), Inf]), giving lxml document order for nested text().
%
%   Example:
%       c   = mat2doc.oxml.nsmap();     % the fixed WordprocessingML prefix->URI map
%       p   = mat2doc.oxml.parse_xml( ...
%           "<w:p xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>" + ...
%           "<w:r><w:t>hi</w:t></w:r></w:p>");
%       t   = mat2doc.oxml.evaluate_xpath(p, 'w:r/w:t', c);
%       disp(t.nsptag_str)               % "w:t"  (1x1 XmlElement)
%       none = mat2doc.oxml.evaluate_xpath(p, 'w:r/w:br', c);
%       disp(isempty(none))              % 1  -- typed EMPTY XmlElement, not [] (H3)
%
%   Design-realization of the mini-XPath subset delegated to by
%   src/docx/oxml/xmlchemy.py::BaseOxmlElement.xpath (fixed nsmap). Re-ported
%   from the corrected Mat2Ppt +oxml evaluate_xpath (no shared code).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/xmlchemy.py::BaseOxmlElement.xpath
%   (the lxml _Element.xpath subset it narrows; design-realization, D-001)

arguments
    context (1,1) mat2doc.oxml.XmlElement
    xpath_str (1,1) string
    ns (1,1) struct
end

ast = parseExpr(char(xpath_str));

if strcmp(ast.absolute, "union")
    % Bare top-level union `A | B | ...` (docx's single most common pattern:
    % `./w:p | ./w:tbl`). Each operand is a full location path -- relative
    % (`./w:p`, `w:r`), absolute (`/w:document/.../w:sectPr`), or grouped. lxml
    % merges the operand node-sets and returns them in DOCUMENT ORDER with
    % identity duplicates removed (H11 -- the union trap; node-sets are sets).
    acc = {};
    for oi = 1:numel(ast.operands)
        acc = [acc, pathNodeset(ast.operands{oi}, context, ns)]; %#ok<AGROW>
    end
    result = cellToElemArray(docSortDedupe(acc));
    return
end

% -- single location path (relative / absolute / grouped) --
% decide the terminal result kind from the last step of the AST
steps = ast.steps;
if isempty(steps)
    lastIsString = false;
else
    last = steps{end};
    lastIsString = strcmp(last.axis, "attribute") || strcmp(last.ntKind, "text");
end

nodeset = initialNodeset(ast, context, ns);

if lastIsString
    prefix = evalElemSteps(nodeset, steps(1:end-1), ns);
    result = evalStringStep(prefix, steps{end}, ns);
else
    finalCell = evalElemSteps(nodeset, steps, ns);
    result = cellToElemArray(finalCell);
end
end

% ======================================================================
% LOCATION-PATH DISPATCH (single path / union operand)
% ======================================================================

function nodeset = initialNodeset(node, context, ns)
% The initial node-set for a single (non-union) location path: the context node
% for a relative path, the XPath document node for an absolute path, or the
% (predicated) grouped node-set for a `(...)` path.
switch node.absolute
    case "rel"
        nodeset = {context};
    case "root"
        nodeset = {docNode(rootOf(context))};
    case "group"
        nodeset = {};
        for gi = 1:numel(node.group)
            g = node.group{gi};
            if ~strcmp(g.absolute, "rel")
                % F2 (design.md section XPath, "never silently mis-evaluated"):
                % a group member that is itself absolute (/, //) or a nested
                % group is parseable but the evaluator only implements RELATIVE
                % group members. Raise rather than evaluate an absolute sub-path
                % as relative ((/w:document)/w:body) or collapse a nested group.
                error("mat2doc:XPathError", ...
                    "Unsupported XPath: non-relative sub-path inside a group '(...)'");
            end
            nodeset = [nodeset, evalElemSteps({context}, g.steps, ns)]; %#ok<AGROW>
        end
        nodeset = docSortDedupe(nodeset);
        % Predicates on the group itself, e.g. `(./w:r)[last()]` /
        % `(./w:comment[@w:id='x'])[1]`: applied to the WHOLE grouped node-set in
        % document order -- a grouped primary expression re-numbers positions by
        % document order (H1: 1-based, never shifted).
        nodeset = applyElemPredicates(nodeset, node.predicates, ns);
    otherwise
        error("mat2doc:XPathError", "Unsupported location-path kind '%s'", node.absolute);
end
end

function outCell = pathNodeset(node, context, ns)
% Evaluate a single location path (relative / absolute / grouped) to an ELEMENT
% node-set (cell of handles). Used for union operands, which in the docx surface
% are always element-producing; a string-terminal (@attr / text()) operand is
% out of subset and raises rather than being silently mis-evaluated.
base = initialNodeset(node, context, ns);
steps = node.steps;
if ~isempty(steps)
    last = steps{end};
    if strcmp(last.axis, "attribute") || strcmp(last.ntKind, "text")
        error("mat2doc:XPathError", ...
            "Unsupported XPath: string-terminal (@attr/text()) union operand");
    end
end
outCell = evalElemSteps(base, steps, ns);
end

% ======================================================================
% PARSER
% ======================================================================

function ast = parseExpr(s)
% Recursive-descent parse of the closed subset into an AST location path.
toks = tokenize(s);
pos = 1;
ast = parseUnion();
if pos <= numel(toks)
    error("mat2doc:XPathError", "Unsupported XPath (trailing tokens): '%s'", s);
end

    function t = peek()
        if pos <= numel(toks)
            t = toks(pos).t;
        else
            t = "";
        end
    end

    function v = peekVal()
        if pos <= numel(toks)
            v = toks(pos).v;
        else
            v = "";
        end
    end

    function t = peek2()
        if pos + 1 <= numel(toks)
            t = toks(pos + 1).t;
        else
            t = "";
        end
    end

    function tok = advance()
        tok = toks(pos);
        pos = pos + 1;
    end

    function expect(tt)
        if pos > numel(toks) || toks(pos).t ~= tt
            error("mat2doc:XPathError", "Expected token '%s' in '%s'", tt, s);
        end
        pos = pos + 1;
    end

    function node = parseUnion()
        % UnionExpr := PathExpr ('|' PathExpr)*  -- a bare top-level union is
        % returned as an "union" AST node whose operands are full location paths;
        % a single path is returned unwrapped.
        subs = {parsePath()};
        while peek() == "union"
            advance();
            subs{end+1} = parsePath(); %#ok<AGROW>
        end
        if numel(subs) == 1
            node = subs{1};
        else
            node = struct("absolute", "union", "operands", {subs});
        end
    end

    function node = parsePath()
        node = struct("absolute", "rel", "group", {{}}, "steps", {{}}, "predicates", {{}});
        if peek() == "lp"
            advance();                          % '('
            inner = parseUnion();               % union (or single path) inside ()
            expect("rp");
            node.absolute = "group";
            if isfield(inner, "operands")
                node.group = inner.operands;
            else
                node.group = {inner};
            end
            node.predicates = parsePredicateList();   % (...)[1] / (...)[last()]
            node.steps = parseTrailingSteps();
            return
        end
        if peek() == "dslash"
            advance();
            node.absolute = "root";
            node.steps{end+1} = parseStep(true);
        elseif peek() == "slash"
            advance();
            node.absolute = "root";
            node.steps{end+1} = parseStep(false);
        else
            node.absolute = "rel";
            node.steps{end+1} = parseStep(false);
        end
        rest = parseTrailingSteps();
        node.steps = [node.steps, rest];
    end

    function stepsC = parseTrailingSteps()
        stepsC = {};
        while peek() == "slash" || peek() == "dslash"
            via = (advance().t == "dslash");
            stepsC{end+1} = parseStep(via); %#ok<AGROW>
        end
    end

    function st = parseStep(via)
        st = struct("viaDescendant", via, "axis", "child", ...
            "ntKind", "", "ntName", "", "predicates", {{}});
        t = peek();
        switch t
            case "dotdot"
                advance(); st.axis = "parent"; st.ntKind = "node";
            case "dot"
                advance(); st.axis = "self"; st.ntKind = "node";
            case "at"
                advance();
                nm = expectName();
                st.axis = "attribute"; st.ntKind = "attr"; st.ntName = nm;
            case "star"
                advance(); st.axis = "child"; st.ntKind = "wildcard";
            case "name"
                nm = advance().v;
                if peek() == "axis"                 % NAME '::' nodetest  (e.g. ancestor::w:tbl)
                    advance();
                    st.axis = nm;
                    if peek() == "star"
                        advance(); st.ntKind = "wildcard";
                    else
                        st.ntName = expectName(); st.ntKind = "name";
                    end
                elseif peek() == "lp"               % function call, text() only
                    advance(); expect("rp");
                    if nm ~= "text"
                        error("mat2doc:XPathError", "Unsupported function '%s()'", nm);
                    end
                    st.axis = "child"; st.ntKind = "text";
                else                                 % plain child name test
                    st.axis = "child"; st.ntKind = "name"; st.ntName = nm;
                end
            otherwise
                error("mat2doc:XPathError", "Unexpected token '%s'", t);
        end
        st.predicates = parsePredicateList();
    end

    function preds = parsePredicateList()
        preds = {};
        while peek() == "lb"
            preds{end+1} = parsePredicate(); %#ok<AGROW>
        end
    end

    function nm = expectName()
        if peek() ~= "name"
            error("mat2doc:XPathError", "Expected a name token in '%s'", s);
        end
        nm = advance().v;
    end

    function pr = parsePredicate()
        expect("lb");
        t = peek();
        if t == "number"                         % [n]  positional (1-based, H1)
            pr = struct("kind", "pos", "n", str2double(advance().v));
        elseif t == "name" && peekVal() == "position" && peek2() == "lp"
            advance(); expect("lp"); expect("rp");        % position() = n
            if peek() ~= "eq"
                error("mat2doc:XPathError", "Unsupported position() predicate (expected '=') in '%s'", s);
            end
            advance();
            if peek() ~= "number"
                error("mat2doc:XPathError", "Unsupported position()= predicate (expected a number) in '%s'", s);
            end
            pr = struct("kind", "pos", "n", str2double(advance().v));
        elseif t == "name" && peekVal() == "last" && peek2() == "lp"
            advance(); expect("lp"); expect("rp");        % last()
            pr = struct("kind", "last");
        elseif t == "name" && peekVal() == "not" && peek2() == "lp"
            advance(); expect("lp");                       % not( relative-path )
            sub = parseUnion();
            expect("rp");
            pr = struct("kind", "notExists", "path", sub);
        elseif t == "at"                          % [@attr = value]
            advance();
            nm = expectName();
            if peek() ~= "eq"
                error("mat2doc:XPathError", "Unsupported attribute predicate (existence) '@%s'", nm);
            end
            advance();
            vt = peek();
            if vt == "string"
                pr = struct("kind", "attrEq", "name", nm, "value", advance().v, "isNum", false);
            elseif vt == "number"
                pr = struct("kind", "attrEq", "name", nm, "value", advance().v, "isNum", true);
            else
                error("mat2doc:XPathError", "Unsupported predicate value in '%s'", s);
            end
        else                                       % [ path ]  or  [ path = value ]
            sub = parseUnion();
            if peek() == "eq"                      % value comparison: path = literal
                advance();
                vt = peek();
                if vt == "string"
                    pr = struct("kind", "pathEq", "path", sub, "value", advance().v, "isNum", false);
                elseif vt == "number"
                    pr = struct("kind", "pathEq", "path", sub, "value", advance().v, "isNum", true);
                else
                    error("mat2doc:XPathError", "Unsupported predicate value in '%s'", s);
                end
            else                                   % existence: [ path ]
                pr = struct("kind", "exists", "path", sub);
            end
        end
        expect("rb");
    end
end

function toks = tokenize(s)
% Lexer for the closed subset. Returns a struct array with fields t (type
% tag) and v (lexeme, for name/number/string).
toks = struct("t", {}, "v", {});
i = 1;
n = numel(s);
    function add(tt, vv)
        toks(end+1) = struct("t", string(tt), "v", string(vv));
    end
while i <= n
    c = s(i);
    if isspace(c)
        i = i + 1;
    elseif c == '/' && i < n && s(i+1) == '/'
        add("dslash", "//"); i = i + 2;
    elseif c == '/'
        add("slash", "/"); i = i + 1;
    elseif c == '('
        add("lp", "("); i = i + 1;
    elseif c == ')'
        add("rp", ")"); i = i + 1;
    elseif c == '['
        add("lb", "["); i = i + 1;
    elseif c == ']'
        add("rb", "]"); i = i + 1;
    elseif c == '@'
        add("at", "@"); i = i + 1;
    elseif c == '|'
        add("union", "|"); i = i + 1;
    elseif c == '='
        add("eq", "="); i = i + 1;
    elseif c == '*'
        add("star", "*"); i = i + 1;
    elseif c == ':' && i < n && s(i+1) == ':'
        add("axis", "::"); i = i + 2;
    elseif c == '.' && i < n && s(i+1) == '.'
        add("dotdot", ".."); i = i + 2;
    elseif c == '.'
        add("dot", "."); i = i + 1;
    elseif c == '"' || c == ''''
        q = c; j = i + 1;
        while j <= n && s(j) ~= q
            j = j + 1;
        end
        if j > n
            error("mat2doc:XPathError", "Unterminated string literal");
        end
        add("string", s(i+1:j-1)); i = j + 1;
    elseif c >= '0' && c <= '9'
        j = i;
        while j <= n && s(j) >= '0' && s(j) <= '9'
            j = j + 1;
        end
        add("number", s(i:j-1)); i = j;
    elseif isNameStart(c)
        j = i;
        while j <= n && isNameChar(s(j))
            if s(j) == ':' && j < n && s(j+1) == ':'
                break                      % '::' is the axis separator, not part of a name
            end
            j = j + 1;
        end
        add("name", s(i:j-1)); i = j;
    else
        error("mat2doc:XPathError", "Unexpected character '%s' in XPath", c);
    end
end
end

function tf = isNameStart(c)
tf = (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '_';
end

function tf = isNameChar(c)
tf = isNameStart(c) || (c >= '0' && c <= '9') || c == '-' || c == ':';
end

% ======================================================================
% EVALUATION
% ======================================================================

function out = evalElemSteps(nodeset, stepsC, ns)
% Apply a sequence of ELEMENT-producing location steps. nodeset and result
% are cell arrays of node handles (XmlElement, or the doc-node sentinel only
% as an initial context). Result is document-ordered and identity-deduped.
out = nodeset;
for k = 1:numel(stepsC)
    out = evalElemStep(out, stepsC{k}, ns);
end
end

function out = evalElemStep(nodeset, step, ns)
acc = {};
for a = 1:numel(nodeset)
    ctx = nodeset{a};
    if step.viaDescendant
        cands = descOrSelf(ctx);
    else
        cands = {ctx};
    end
    for b = 1:numel(cands)
        matched = axisMatchElems(cands{b}, step, ns);
        matched = applyElemPredicates(matched, step.predicates, ns);
        acc = [acc, matched]; %#ok<AGROW>
    end
end
out = docSortDedupe(acc);
end

function matched = axisMatchElems(cand, step, ns)
matched = {};
switch step.axis
    case "self"
        % Apply the name test on the self axis (libxml2/lxml semantics):
        % self::NAME matches only when the context node's tag == NAME; `.`
        % (node()) and self::* (wildcard) match the context node unchanged.
        if strcmp(step.ntKind, "name")
            clark = resolveElemName(step.ntName, ns);
            if tagOf(cand) == clark
                matched = {cand};
            end
        else
            matched = {cand};
        end
    case "parent"
        % Apply the name test on the parent axis: parent::NAME matches the
        % parent only when its tag == NAME; `..` (node()) and parent::*
        % (wildcard) match the parent unchanged. A missing parent -> no match.
        p = getparentOf(cand);
        if ~isNone(p)
            if strcmp(step.ntKind, "name")
                clark = resolveElemName(step.ntName, ns);
                if tagOf(p) == clark
                    matched = {p};
                end
            else
                matched = {p};
            end
        end
    case "ancestor"
        if strcmp(step.ntKind, "wildcard")
            % F2 (design.md section XPath): ancestor::* (wildcard node test on
            % the ancestor axis) is parseable but unimplemented -- the branch
            % only name-matches, so a wildcard would silently match nothing. Raise.
            error("mat2doc:XPathError", ...
                "Unsupported XPath: wildcard node test on the ancestor axis");
        end
        clark = resolveElemName(step.ntName, ns);
        cur = getparentOf(cand);
        while ~isNone(cur)
            if tagOf(cur) == clark
                matched{end+1} = cur; %#ok<AGROW>
            end
            cur = getparentOf(cur);
        end
    case "child"
        kids = childElems(cand);
        switch step.ntKind
            case "wildcard"
                matched = kids;
            case "name"
                clark = resolveElemName(step.ntName, ns);
                for j = 1:numel(kids)
                    if tagOf(kids{j}) == clark
                        matched{end+1} = kids{j}; %#ok<AGROW>
                    end
                end
            otherwise
                error("mat2doc:XPathError", "Non-terminal '%s' node test unsupported", step.ntKind);
        end
    case "following-sibling"
        % Siblings that follow cand among its parent's children, in DOCUMENT
        % order (a forward axis; position order == document order).
        p = getparentOf(cand);
        if ~isNone(p)
            sibs = childElems(p);
            ci = siblingIndex(sibs, cand);
            for j = ci + 1 : numel(sibs)
                if nodeTestMatch(sibs{j}, step, ns)
                    matched{end+1} = sibs{j}; %#ok<AGROW>
                end
            end
        end
    case "preceding-sibling"
        % Siblings that precede cand, returned in REVERSE document order = axis
        % order (nearest sibling first) so a positional predicate [1] selects the
        % NEAREST preceding sibling (H1: on a reverse axis, position() counts by
        % axis direction, not document order). evalElemStep's docSortDedupe then
        % restores document order for the final (unpredicated) node-set.
        p = getparentOf(cand);
        if ~isNone(p)
            sibs = childElems(p);
            ci = siblingIndex(sibs, cand);
            for j = ci - 1 : -1 : 1
                if nodeTestMatch(sibs{j}, step, ns)
                    matched{end+1} = sibs{j}; %#ok<AGROW>
                end
            end
        end
    case "preceding"
        % All nodes that precede cand in document order EXCLUDING cand's own
        % ancestors, returned in REVERSE document order = axis order (nearest
        % first) so a positional predicate [1] selects the nearest preceding node.
        root = rootOf(cand);
        pre = descOrSelf(root);              % element nodes in document order
        ci = siblingIndex(pre, cand);
        anc = ancestorHandles(cand);
        for j = ci - 1 : -1 : 1
            e = pre{j};
            if anyHandleEq(anc, e)
                continue                      % an ancestor is never "preceding"
            end
            if nodeTestMatch(e, step, ns)
                matched{end+1} = e; %#ok<AGROW>
            end
        end
    otherwise
        error("mat2doc:XPathError", "Unsupported axis '%s'", step.axis);
end
end

function tf = nodeTestMatch(e, step, ns)
% The node-name test shared by the sibling / preceding axes: `*` matches any
% element; NAME matches only when the element's Clark tag equals NAME.
switch step.ntKind
    case "wildcard"
        tf = true;
    case "name"
        tf = (tagOf(e) == resolveElemName(step.ntName, ns));
    otherwise
        error("mat2doc:XPathError", ...
            "Unsupported node test '%s' on axis '%s'", step.ntKind, step.axis);
end
end

function idx = siblingIndex(cellHandles, target)
% 1-based index of the handle identical to target, or 0 if not present.
idx = 0;
for i = 1:numel(cellHandles)
    if cellHandles{i} == target
        idx = i;
        return
    end
end
end

function anc = ancestorHandles(cand)
% Cell of ancestor element handles of cand (parent, grandparent, ...).
anc = {};
cur = getparentOf(cand);
while ~isNone(cur)
    anc{end+1} = cur; %#ok<AGROW>
    cur = getparentOf(cur);
end
end

function tf = anyHandleEq(cellHandles, e)
tf = false;
for i = 1:numel(cellHandles)
    if cellHandles{i} == e
        tf = true;
        return
    end
end
end

function matched = applyElemPredicates(matched, preds, ns)
for p = 1:numel(preds)
    pr = preds{p};
    switch pr.kind
        case "pos"
            if numel(matched) >= pr.n && pr.n >= 1
                matched = matched(pr.n);
            else
                matched = {};
            end
        case "attrEq"
            keep = false(1, numel(matched));
            for j = 1:numel(matched)
                keep(j) = attrEqTest(matched{j}, pr, ns);
            end
            matched = matched(keep);
        case "exists"
            % [ path ] existence: keep candidates for which the relative sub-path
            % (which may be a union, e.g. [self::w:br | self::w:cr | ...]) is a
            % non-empty node-set.
            keep = false(1, numel(matched));
            for j = 1:numel(matched)
                keep(j) = ~isempty(subPathNodeset(matched{j}, pr.path, ns));
            end
            matched = matched(keep);
        case "notExists"
            % [ not( path ) ]: keep candidates for which the relative sub-path is
            % an EMPTY node-set (e.g. ./*[not(self::w:sectPr)] -> every child
            % whose tag is not w:sectPr).
            keep = false(1, numel(matched));
            for j = 1:numel(matched)
                keep(j) = isempty(subPathNodeset(matched{j}, pr.path, ns));
            end
            matched = matched(keep);
        case "last"
            % [last()]: the last node of the current (document-ordered) node-set.
            if ~isempty(matched)
                matched = matched(end);
            else
                matched = {};
            end
        case "pathEq"
            % [ path = literal ]: keep candidates for which the relative sub-path
            % (terminating in @attr / text()) has ANY value equal to the literal
            % (XPath node-set `=` is existential), e.g. [w:name/@w:val="Heading 1"].
            keep = false(1, numel(matched));
            for j = 1:numel(matched)
                vals = subPathStringValues(matched{j}, pr.path, ns);
                if pr.isNum
                    keep(j) = any(str2double(vals) == str2double(pr.value));
                else
                    keep(j) = any(vals == string(pr.value));
                end
            end
            matched = matched(keep);
    end
end
end

function outCell = subPathNodeset(ctxNode, pathNode, ns)
% Evaluate a RELATIVE predicate sub-path -- a single path or a union of relative
% paths -- against ctxNode, returning an element node-set (document-ordered,
% identity-deduped). Absolute (/, //) sub-paths are out of subset and raise
% (F2), never silently evaluated as relative.
if isfield(pathNode, "operands")
    acc = {};
    for oi = 1:numel(pathNode.operands)
        op = pathNode.operands{oi};
        if ~strcmp(op.absolute, "rel")
            error("mat2doc:XPathError", ...
                "Unsupported XPath: non-relative sub-path inside a predicate");
        end
        acc = [acc, evalElemSteps({ctxNode}, op.steps, ns)]; %#ok<AGROW>
    end
    outCell = docSortDedupe(acc);
else
    if ~strcmp(pathNode.absolute, "rel")
        error("mat2doc:XPathError", ...
            "Unsupported XPath: non-relative sub-path inside a predicate");
    end
    outCell = evalElemSteps({ctxNode}, pathNode.steps, ns);
end
end

function vals = subPathStringValues(ctxNode, pathNode, ns)
% Evaluate a RELATIVE predicate sub-path that terminates in @attr or text()
% against ctxNode, returning its string node-set -- the left operand of a
% `path = literal` comparison predicate (e.g. [w:name/@w:val="Heading 1"]).
if isfield(pathNode, "operands")
    error("mat2doc:XPathError", ...
        "Unsupported XPath: union on the left of a '=' comparison predicate");
end
if ~strcmp(pathNode.absolute, "rel")
    error("mat2doc:XPathError", ...
        "Unsupported XPath: non-relative sub-path inside a predicate");
end
steps = pathNode.steps;
if isempty(steps)
    error("mat2doc:XPathError", "Unsupported XPath: empty comparison sub-path");
end
last = steps{end};
if ~(strcmp(last.axis, "attribute") || strcmp(last.ntKind, "text"))
    error("mat2doc:XPathError", ...
        "Unsupported XPath: '=' comparison on a non-string sub-path");
end
prefix = evalElemSteps({ctxNode}, steps(1:end-1), ns);
vals = evalStringStep(prefix, last, ns);
end

function tf = attrEqTest(e, pr, ns)
key = resolveAttrName(pr.name, ns);
av = attrGet(e, key);
if isNone(av)
    tf = false;
    return
end
if pr.isNum
    tf = (str2double(av) == str2double(pr.value));
else
    tf = (string(av) == string(pr.value));
end
end

function result = evalStringStep(nodeset, step, ns)
% Terminal attribute (@x) or text() step -> (1,N) string array, ordered by
% the document position of the selected attribute/text node, then
% identity-deduped (lxml node-set semantics: a node reached via several
% prefix matches appears once).
% A positional predicate ([n], [last()], [position()=n]) on a terminal string
% step is applied PER context element to THAT element's own attribute/text
% node-set: lxml numbers predicate positions within each context node's axis, so
% //@r:id[2] yields [] (each element owns at most one r:id, so position 2 never
% exists). Non-positional predicates on a string terminal are out of subset and
% raise (never silently mis-evaluated -- design.md section XPath).
for pi = 1:numel(step.predicates)
    k = step.predicates{pi}.kind;
    if ~(strcmp(k, "pos") || strcmp(k, "last"))
        error("mat2doc:XPathError", ...
            "Unsupported XPath: non-positional predicate on a terminal attribute/text() step");
    end
end
vals = strings(1, 0);
keys = {};                 % document-order sort key (index vector) per node
sigs = strings(1, 0);      % node-identity signature per node (F3 dedup)
    function emit(v, node, tag, keyTail)
        % keyTail is appended to node's docKey to form the document-order sort
        % key: [] where the node's position IS its element's (attribute value,
        % element leading text); Inf for a child TAIL, whose text node follows
        % the ENTIRE child subtree in document order (WPC-F1: [docKey(kid), Inf]
        % sorts the tail after the kid and all its descendants, matching lxml --
        % a bare docKey(kid) would wrongly place the tail before the subtree).
        % The identity signature uses the bare docKey (one attr/text/tail node
        % per owner, already unique -- Inf must NOT enter the dedup key).
        base = docKey(node);
        vals(end+1) = string(v);              %#ok<AGROW>
        keys{end+1} = [base, keyTail];        %#ok<AGROW>
        sigs(end+1) = tag + mat2str(base);    %#ok<AGROW>
    end
for a = 1:numel(nodeset)
    ctx = nodeset{a};
    if step.viaDescendant
        cands = descOrSelf(ctx);
    else
        cands = {ctx};
    end
    for b = 1:numel(cands)
        cand = cands{b};
        if isDoc(cand)
            continue                           % the doc node has no attrs/text
        end
        % Build THIS cand's ordered string-node list (document order within the
        % cand), then apply any positional predicate before emitting.
        lv = {}; lnode = {}; ltag = strings(1, 0); lkey = {};
        if step.ntKind == "attr"
            key = resolveAttrName(step.ntName, ns);
            v = attrGet(cand, key);
            if ~isNone(v)
                % attr-node identity: owner element (key is fixed per step)
                lv{end+1} = v; lnode{end+1} = cand; ltag(end+1) = "a:"; lkey{end+1} = []; %#ok<AGROW>
            end
        elseif step.ntKind == "text"
            % lxml text() selects ALL text-node children of cand in document
            % order: the element's own leading text (its C-level text node,
            % read via text_raw_() to BYPASS a CT_* getText_ property shadow --
            % D10, F1), then the tail of each child element (F1). A childless
            % element with no direct text yields no text node (lxml returns [],
            % never the shadowed property value).
            t = cand.text_raw_();
            if ~isNone(t)
                lv{end+1} = t; lnode{end+1} = cand; ltag(end+1) = "t:"; lkey{end+1} = []; %#ok<AGROW>
            end
            kids = childElems(cand);
            for c = 1:numel(kids)
                tl = kids{c}.tail;             % child tail = text node child of cand
                if ~isNone(tl)
                    % tail sorts AFTER the whole child subtree (WPC-F1): Inf suffix
                    lv{end+1} = tl; lnode{end+1} = kids{c}; ltag(end+1) = "l:"; lkey{end+1} = Inf; %#ok<AGROW>
                end
            end
        else
            error("mat2doc:XPathError", "Unsupported terminal step");
        end
        sel = 1:numel(lv);
        for pi = 1:numel(step.predicates)
            pr = step.predicates{pi};
            if strcmp(pr.kind, "pos")
                if pr.n >= 1 && pr.n <= numel(sel)
                    sel = sel(pr.n);           % H1: 1-based, never shifted
                else
                    sel = [];
                end
            else                                % last()
                if ~isempty(sel)
                    sel = sel(end);
                else
                    sel = [];
                end
            end
        end
        for si = 1:numel(sel)
            idx = sel(si);
            emit(lv{idx}, lnode{idx}, ltag(idx), lkey{idx});
        end
    end
end
ord = sortKeyOrder(keys);
vals = vals(ord);
sigs = sigs(ord);
% F3: drop identity duplicates, keeping the first (document-order) occurrence
% -- lxml dedupes the node-set (element steps already dedupe via docSortDedupe;
% string terminals did not).
keep = true(1, numel(sigs));
seen = strings(1, 0);
for i = 1:numel(sigs)
    if any(sigs(i) == seen)
        keep(i) = false;
    else
        seen(end+1) = sigs(i); %#ok<AGROW>
    end
end
result = vals(keep);
end

% ======================================================================
% NODE HELPERS (XmlElement + doc-node sentinel)
% ======================================================================

function d = docNode(root)
% The XPath document node: its only child is the document element `root`.
% Lets absolute paths (`/w:document`, `//@id`) be evaluated as ordinary steps.
d = struct("doc", true, "root", root);
end

function tf = isDoc(x)
tf = isstruct(x) && isfield(x, "doc");
end

function kids = childElems(x)
if isDoc(x)
    kids = {x.root};
else
    arr = x.to_array();
    kids = cell(1, numel(arr));
    for i = 1:numel(arr)
        kids{i} = arr(i);
    end
end
end

function t = tagOf(x)
if isDoc(x)
    t = "";        % the doc node never matches an element name test
else
    t = x.tag;
end
end

function p = getparentOf(x)
if isDoc(x)
    p = [];
else
    p = x.getparent();
end
end

function v = attrGet(x, key)
if isDoc(x)
    v = [];        % the doc node has no attributes
else
    v = x.get(key);
end
end

function list = descOrSelf(x)
% Document-order (preorder) descendant-or-self as a cell array.
list = {x};
kids = childElems(x);
for i = 1:numel(kids)
    list = [list, descOrSelf(kids{i})]; %#ok<AGROW>
end
end

function tf = isNone(x)
tf = isequal(x, []);
end

% ======================================================================
% NAME RESOLUTION
% ======================================================================

function clark = resolveElemName(name, ns)
name = string(name);
parts = split(name, ":");
if numel(parts) == 2
    clark = "{" + nsLookup(ns, parts(1)) + "}" + parts(2);
else
    clark = name;   % unprefixed element name (no namespace)
end
end

function key = resolveAttrName(name, ns)
name = string(name);
parts = split(name, ":");
if numel(parts) == 2
    key = "{" + nsLookup(ns, parts(1)) + "}" + parts(2);   % lxml Clark-keyed attribute
else
    key = name;                                            % plain (no-namespace) attribute
end
end

function uri = nsLookup(ns, pfx)
pfx = char(pfx);
if ~isfield(ns, pfx)
    error("mat2doc:XPathError", "Undefined namespace prefix: '%s'", pfx);
end
uri = string(ns.(pfx));
end

% ======================================================================
% DOCUMENT ORDER
% ======================================================================

function k = docKey(e)
% Vector of 1-based sibling indices from the document root down to e.
k = [];
cur = e;
while true
    p = cur.getparent();
    if isNone(p)
        break
    end
    arr = p.to_array();
    idx = find(arr == cur, 1);
    k = [idx, k]; %#ok<AGROW>
    cur = p;
end
end

function out = docSortDedupe(cellElems)
% Sort a cell of XmlElement handles into document order and drop identity
% duplicates (node-set semantics).
n = numel(cellElems);
if n <= 1
    out = cellElems;
    return
end
keys = cell(1, n);
for i = 1:n
    keys{i} = docKey(cellElems{i});
end
ord = sortKeyOrder(keys);
sorted = cellElems(ord);
keep = true(1, n);
for i = 2:n
    if sorted{i} == sorted{i-1}     % identity: equal handle (== equal docKey)
        keep(i) = false;
    end
end
out = sorted(keep);
end

function ord = sortKeyOrder(keys)
% Stable order of index-vector keys by lexicographic document order
% (ancestor before descendant).
n = numel(keys);
ord = 1:n;
for i = 2:n                          % insertion sort (stable, n is small)
    j = i;
    while j > 1 && keyLess(keys{ord(j)}, keys{ord(j-1)})
        tmp = ord(j); ord(j) = ord(j-1); ord(j-1) = tmp;
        j = j - 1;
    end
end
end

function tf = keyLess(a, b)
m = min(numel(a), numel(b));
for i = 1:m
    if a(i) < b(i)
        tf = true; return
    elseif a(i) > b(i)
        tf = false; return
    end
end
tf = numel(a) < numel(b);            % shorter prefix (ancestor) sorts first
end

% ======================================================================
% RESULT PACKING
% ======================================================================

function root = rootOf(e)
root = e;
while true
    p = root.getparent();
    if isNone(p)
        break
    end
    root = p;
end
end

function arr = cellToElemArray(cellElems)
if isempty(cellElems)
    arr = mat2doc.oxml.XmlElement.empty(1, 0);
else
    arr = [cellElems{:}];
end
end

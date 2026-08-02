classdef CT_Comments < mat2doc.oxml.BaseOxmlElement
% CT_COMMENTS `<w:comments>` element, the root element of the comments part.
%
%   Simply contains a collection of `<w:comment>` elements, each a single
%   comment identified by a unique `w:id`. The offset of a comment in this
%   collection is arbitrary -- it is essentially a _set_ implemented as a list.
%
%   Registered for <w:comments> (docx/oxml/__init__.py:91, P8-2).
%
%   DESCRIPTOR (comments.py 30):
%     comment = ZeroOrMore("w:comment")   -- successors=() (default) -> APPEND
%   xmlchemy (docx form) generates: comment_lst, _new_comment, _insert_comment,
%   _add_comment (private, **attrs). The GENERATED public `add_comment` is
%   SUPPRESSED because this class defines an explicit add_comment() (xmlchemy
%   _add_to_class no-ops when the name already exists, xmlchemy.py 357-359 -- the
%   CT_Num.add_lvlOverride precedent). Underscore rotation (design.md section 2):
%   _new/_insert/_add -> new_comment_/insert_comment_/add_comment_.
%
%   METHODS (comments.py):
%     add_comment()               (32-61): the explicit adder. Parses a minimum
%                                 valid <w:comment> (unique w:id, empty w:author,
%                                 a single CommentText paragraph holding a
%                                 CommentReference run with an annotationRef), then
%                                 appends it. Content is added later via the proxy.
%     get_comment_by_id(id)       (63-66): the <w:comment> with @w:id=id, or [].
%     _next_available_comment_id  (68-88): max(used ids)+1, or -- if that would
%                                 overflow a 32-bit signed int -- the first unused
%                                 non-negative integer.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on the <w:comments> root of a real word/comments.xml.
%
%   None SENTINEL (H3): inline `isequal(x, [])`.
%
%   Example:
%       cs = mat2doc.oxml.parse_xml(comments_bytes);   % a CT_Comments
%       c  = cs.add_comment();                          % a new <w:comment w:id="N">
%       c2 = cs.get_comment_by_id(0);                   % lookup by id ([] if absent)
%
%   Ported from python-docx v1.2.0: src/docx/oxml/comments.py::CT_Comments
%   (lines 18-88; registered for w:comments)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        COMMENT_TAG   = "w:comment"          % ZeroOrMore @ comments.py:30
        NO_SUCCESSORS = string.empty(1, 0)   % successors=() -> append
        INT32_MAX     = 2147483647           % 2**31 - 1 (comments.py:80)
    end

    properties (Dependent)  % generated descriptor members
        comment_lst   % list of <w:comment> children (document order)
    end

    methods
        function obj = CT_Comments(varargin)
            % CT_COMMENTS Construct a loose <w:comments> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ============ comment (ZeroOrMore, successors=() -> append) ============
        function lst = get.comment_lst(obj);            lst = obj.getChildList(obj.COMMENT_TAG); end
        function child = new_comment_(obj);             child = obj.newChild(obj.COMMENT_TAG); end
        function child = insert_comment_(obj, child);   child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_comment_(obj, varargin);   child = obj.addChild(obj.COMMENT_TAG, obj.NO_SUCCESSORS, varargin{:}); end

        function comment = add_comment(obj)
            % ADD_COMMENT A newly added `<w:comment>` child of this `<w:comments>`
            %   (comments.py 32-61). The returned comment is the minimum valid
            %   value: a `w:id` unique among the existing comments, the required
            %   `w:author` present but empty, and content limited to a single run
            %   containing the annotation reference but no text. Content is added
            %   later by adding runs to the first paragraph and adding paragraphs.
            %
            %   Python:
            %     next_id = self._next_available_comment_id()
            %     comment = cast(CT_Comment, parse_xml(
            %         f'<w:comment {nsdecls("w")} w:id="{next_id}" w:author="">'
            %         f"  <w:p>" f"    <w:pPr>"
            %         f'      <w:pStyle w:val="CommentText"/>' f"    </w:pPr>"
            %         f"    <w:r>" f"      <w:rPr>"
            %         f'        <w:rStyle w:val="CommentReference"/>' f"      </w:rPr>"
            %         f"      <w:annotationRef/>" f"    </w:r>" f"  </w:p>"
            %         f"</w:comment>"))
            %     self.append(comment)
            %     return comment
            %
            %   The f-string literals are ADJACENT (no newlines) -- only spaces
            %   separate the child elements. parse_xml applies remove_blank_text,
            %   so the whitespace-only text between element children is dropped
            %   (each parent has element children); the resulting tree is clean.
            %   The MATLAB string reproduces the exact byte sequence (spaces, no
            %   newlines) so the parse is byte-faithful either way. next_id is a
            %   Python int -> str(int) via pyStr (H14). Registry -> CT_Comment.
            %
            %   Ported from python-docx v1.2.0: src/docx/oxml/comments.py::CT_Comments.add_comment
            next_id = obj.next_available_comment_id_();
            xml = "<w:comment " + mat2doc.oxml.nsdecls("w") + ...
                " w:id=""" + mat2doc.shared.pyStr(next_id) + """ w:author="""">" + ...
                "  <w:p>" + ...
                "    <w:pPr>" + ...
                "      <w:pStyle w:val=""CommentText""/>" + ...
                "    </w:pPr>" + ...
                "    <w:r>" + ...
                "      <w:rPr>" + ...
                "        <w:rStyle w:val=""CommentReference""/>" + ...
                "      </w:rPr>" + ...
                "      <w:annotationRef/>" + ...
                "    </w:r>" + ...
                "  </w:p>" + ...
                "</w:comment>";
            comment = mat2doc.oxml.parse_xml(xml);   % registry -> CT_Comment
            obj.append(comment);                     % Python: self.append(comment)
        end

        function comment = get_comment_by_id(obj, comment_id)
            % GET_COMMENT_BY_ID The `<w:comment>` identified by `comment_id`, or []
            %   (None) if not found (comments.py 63-66). Python:
            %     comment_elms = self.xpath(f"(./w:comment[@w:id='{comment_id}'])[1]")
            %     return comment_elms[0] if comment_elms else None
            %   The grouped positional predicate (...)[1] and the prefixed
            %   attribute-equality [@w:id='N'] are both in the verified xpath
            %   subset. H4: `if comment_elms` (non-empty node-set) -> ~isempty. H3:
            %   no match returns [] (None), NOT an empty typed array. comment_id is
            %   interpolated as str(comment_id) (H14, pyStr).
            %
            %   Ported from python-docx v1.2.0: src/docx/oxml/comments.py::CT_Comments.get_comment_by_id
            expr = "(./w:comment[@w:id='" + mat2doc.shared.pyStr(comment_id) + "'])[1]";
            comment_elms = obj.xpath(expr);
            if ~isempty(comment_elms)     % Python: if comment_elms
                comment = comment_elms(1);   % Python: comment_elms[0] (H1: [1])
            else
                comment = [];                % Python: None
            end
        end
    end

    methods (Access = private)
        function next_id = next_available_comment_id_(obj)
            % _NEXT_AVAILABLE_COMMENT_ID The next available comment id
            %   (comments.py 68-88). By schema any positive integer is legal; the
            %   default is max()+1, but if that would exceed a 32-bit signed int we
            %   fall back to the first unused non-negative integer from 0.
            %
            %   Python:
            %     used_ids = [int(x) for x in self.xpath("./w:comment/@w:id")]
            %     next_id = max(used_ids, default=-1) + 1
            %     if next_id <= 2**31 - 1: return next_id
            %     for expected, actual in enumerate(sorted(used_ids)):
            %         if expected != actual: return expected
            %     return len(used_ids)
            %
            %   xpath("./w:comment/@w:id") returns a (1,N) string array of the id
            %   attribute VALUES (terminal @attr step). int(x) parses each (H6:
            %   integer id values; docx w:id are ASCII decimals). max(...,
            %   default=-1) -> -1 when empty so next_id starts at 0. H1: enumerate's
            %   `expected` is a 0-based id VALUE returned verbatim (not a collection
            %   index) -- k-1 for the k-th sorted id.
            %
            %   Ported from python-docx v1.2.0: src/docx/oxml/comments.py::CT_Comments._next_available_comment_id
            id_strs = obj.xpath("./w:comment/@w:id");   % (1,N) string array
            used_ids = zeros(1, numel(id_strs));
            for k = 1:numel(id_strs)
                used_ids(k) = str2double(id_strs(k));   % Python: int(x)
            end
            if isempty(used_ids)                        % Python: max(used_ids, default=-1)
                next_id = -1 + 1;                       % default -1, then +1 -> 0
            else
                next_id = max(used_ids) + 1;
            end
            if next_id <= obj.INT32_MAX                 % Python: if next_id <= 2**31 - 1
                return
            end
            % -- fall-back: first unused non-negative integer (comments.py 84-88) --
            sorted_ids = sort(used_ids);
            for k = 1:numel(sorted_ids)
                expected = k - 1;                       % Python enumerate: 0-based
                if expected ~= sorted_ids(k)            % Python: if expected != actual
                    next_id = expected;
                    return
                end
            end
            next_id = numel(used_ids);                  % Python: return len(used_ids)
        end
    end
end

classdef Comments < handle
% COMMENTS Collection containing the comments added to this document.
%
%   A proxy over the `<w:comments>` root of the comments part. Provides
%   iteration, count, addition, and id-lookup of the document's comments.
%
%   NOT an ElementProxy: python-docx `Comments` is a plain class holding BOTH the
%   `w:comments` element AND the CommentsPart (comments.py 20-22), so it derives
%   directly from `handle` (two stored references, not the single-element
%   ElementProxy shape). Underscore rotation (design.md section 2):
%   _comments_elm -> comments_elm_, _comments_part -> comments_part_.
%
%   DUNDER MAPPING (design.md section 2, the Styles precedent):
%     __iter__  -> to_array()   (`for c in comments` -> `for c = comments.to_array()`)
%     __len__   -> len_()       (`len(comments)`     -> `comments.len_()`)
%   H5 (identity): to_array / add_comment / get mint FRESH Comment views each call
%   (python-docx does not cache Comment objects).
%
%   MEMBERS (comments.py 24-80):
%     to_array()                    (__iter__ 24-29): a Comment per <w:comment>.
%     len_()                        (__len__ 31-33): number of comments.
%     add_comment(text, author, initials) (35-75): add a new comment; see below.
%     get(comment_id)               (77-80): the Comment with `comment_id`, or [].
%
%   Example:
%       cs = mat2doc.Document().comments;
%       c  = cs.add_comment("Nice point", "Amy", "AJ");
%       n  = cs.len_();          % 1
%       c0 = cs.get(0);          % the same comment (fresh view)
%
%   Ported from python-docx v1.2.0: src/docx/comments.py::Comments (lines 17-80)

    properties (Access = private)
        comments_elm_    % _comments_elm (comments.py 21): the <w:comments> CT_Comments
        comments_part_   % _comments_part (comments.py 22): the CommentsPart
    end

    methods
        function obj = Comments(comments_elm, comments_part)
            % COMMENTS Store the comments element and part (comments.py 20-22).
            %
            %   Inputs:  comments_elm  - a mat2doc.oxml.comments.CT_Comments.
            %            comments_part - the CommentsPart owning it.
            %   Outputs: obj           - a scalar Comments handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/comments.py::Comments.__init__
            obj.comments_elm_ = comments_elm;    % Python: self._comments_elm = comments_elm
            obj.comments_part_ = comments_part;  % Python: self._comments_part = comments_part
        end

        function result = to_array(obj)
            % TO_ARRAY A Comment per `<w:comment>` in this collection (comments.py
            %   24-29). Python __iter__: (Comment(comment_elm, self._comments_part)
            %   for comment_elm in self._comments_elm.comment_lst). Materialized to
            %   a 1xN Comment array (design.md iteration idiom). No comments -> 1x0.
            %
            %   Ported from python-docx v1.2.0: src/docx/comments.py::Comments.__iter__
            lst = obj.comments_elm_.comment_lst;
            result = mat2doc.comments.Comment.empty(1, 0);
            for k = 1:numel(lst)   % Python: for comment_elm in self._comments_elm.comment_lst
                result(k) = mat2doc.comments.Comment(lst(k), obj.comments_part_);
            end
        end

        function n = len_(obj)
            % LEN_ The number of comments in this collection (comments.py 31-33).
            %   Python __len__: len(self._comments_elm.comment_lst).
            %
            %   Ported from python-docx v1.2.0: src/docx/comments.py::Comments.__len__
            n = numel(obj.comments_elm_.comment_lst);
        end

        function comment = add_comment(obj, text, author, initials)
            % ADD_COMMENT Add a new comment to the document and return it
            %   (comments.py 35-75). The comment is appended to the collection and
            %   assigned a unique comment-id.
            %
            %   If `text` is provided it is added to the comment (the common case: a
            %   modest passage of plain text). Multiple paragraphs are produced by
            %   separating text with newlines ("\n"); between newlines text is
            %   interpreted as in Document.add_paragraph(text=...). The default is a
            %   single empty paragraph (matching the Word UI). `author` is required
            %   (empty string by default); `initials` is optional (empty string by
            %   default) and passing [] (None) omits @w:initials.
            %
            %   Python:
            %     comment_elm = self._comments_elm.add_comment()
            %     comment_elm.author = author
            %     comment_elm.initials = initials
            %     comment_elm.date = dt.datetime.now(dt.timezone.utc)
            %     comment = Comment(comment_elm, self._comments_part)
            %     if text == "": return comment
            %     para_text_iter = iter(text.split("\n"))
            %     first_para_text = next(para_text_iter)
            %     first_para = comment.paragraphs[0]
            %     first_para.add_run(first_para_text)
            %     for s in para_text_iter:
            %         comment.add_paragraph(text=s)
            %     return comment
            %
            %   H13 defaults: text="", author="", initials="" (initials [] omits the
            %   attribute). H3: `if text == ""` is an EMPTY-STRING equality test
            %   (comments.py 63), NOT a None test -- ported as `text == ""`. The
            %   date is the current UTC instant (tz-aware; formatted whole-seconds
            %   with a literal Z by ST_DateTime -- FIRST reachable ST_DateTime use).
            %   H1: paragraphs[0] -> paragraphs(1). H14/H2: text is user content.
            %
            %   Ported from python-docx v1.2.0: src/docx/comments.py::Comments.add_comment
            arguments
                obj
                text     = ""   % Python default ""
                author   = ""   % Python default ""
                initials = ""   % Python default "" (pass [] for None -> omit @w:initials)
            end
            comment_elm = obj.comments_elm_.add_comment();       % Python: self._comments_elm.add_comment()
            comment_elm.author = author;                         % Python: comment_elm.author = author
            comment_elm.initials = initials;                     % Python: comment_elm.initials = initials
            comment_elm.date = datetime("now", "TimeZone", "UTC");  % Python: dt.datetime.now(dt.timezone.utc)
            comment = mat2doc.comments.Comment(comment_elm, obj.comments_part_);

            if strlength(text) == 0                              % Python: if text == "":
                return
            end

            para_texts = split(string(text), string(newline));   % Python: text.split("\n")
            first_para_text = para_texts(1);                     % Python: next(para_text_iter)
            ps = comment.paragraphs();
            first_para = ps(1);                                  % Python: comment.paragraphs[0] (H1)
            first_para.add_run(first_para_text);                 % Python: first_para.add_run(first_para_text)

            for k = 2:numel(para_texts)                          % Python: for s in para_text_iter (remaining)
                comment.add_paragraph(para_texts(k));            % Python: comment.add_paragraph(text=s) (text is the first positional)
            end
        end

        function comment = get(obj, comment_id)
            % GET The comment identified by `comment_id`, or [] (None) if not found
            %   (comments.py 77-80). Python:
            %     comment_elm = self._comments_elm.get_comment_by_id(comment_id)
            %     return Comment(comment_elm, self._comments_part)
            %            if comment_elm is not None else None
            %   H3: get_comment_by_id returns [] (None) on no match; the Comment is
            %   minted only when non-None, else [].
            %
            %   Ported from python-docx v1.2.0: src/docx/comments.py::Comments.get
            comment_elm = obj.comments_elm_.get_comment_by_id(comment_id);
            if ~isequal(comment_elm, [])   % Python: if comment_elm is not None
                comment = mat2doc.comments.Comment(comment_elm, obj.comments_part_);
            else
                comment = [];              % Python: None
            end
        end
    end
end

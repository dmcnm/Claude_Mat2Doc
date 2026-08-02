classdef Comment < mat2doc.BlockItemContainer
% COMMENT Proxy for a single comment in the document.
%
%   Provides access to comment metadata (author, initials, date) and content.
%   A comment is also a block-item container, like a table cell, so it can
%   contain both paragraphs and tables, and its paragraphs can hold rich text,
%   hyperlinks, and images -- though the common case is a single plain-text
%   paragraph (a sentence or phrase).
%
%   TIER (comments.py 83 `class Comment(BlockItemContainer)`): the C3
%   BlockItemContainer element seam. Unlike _BaseHeaderFooter, Comment stores a
%   CONCRETE element (the <w:comment>) and does NOT override the element_() seam
%   -- BlockItemContainer's base element_() returning element_store_ is exactly
%   right (the comment element is passed to the ctor, the _Cell precedent). The
%   inherited BlockItemContainer surface therefore operates directly on the
%   <w:comment>: add_paragraph_/paragraphs/add_table read it via the seam; tables
%   reads comment_elm.tbl_lst.
%
%   ATTRIBUTES (comments.py 97-99): Python
%     super().__init__(comment_elm, comments_part)   # BlockItemContainer:
%                                                    #   _element=comment_elm,
%                                                    #   _parent=comments_part
%     self._comment_elm = comment_elm                # typed alias, SAME element
%   Ported: obj@mat2doc.BlockItemContainer(comment_elm, comments_part) sets
%   element_store_ = comment_elm (the C3 concrete store) and parent_ =
%   comments_part (via StoryChild; CommentsPart.part() returns itself, XmlPart
%   precedent). comment_elm_ additionally holds the same CT_Comment handle (Python
%   self._comment_elm) -- the metadata members read it; both are the one comment.
%
%   FLAG-3: no basename collision -- the package folder "+comments", the class
%   files "Comment.m" / "Comments.m" are distinct names in distinct directories;
%   no rename needed (module docx.comments -> package +comments, classes verbatim).
%
%   MEMBERS (comments.py 101-163):
%     add_paragraph(text, style)   -> super().add_paragraph(text, style); when
%                                     style is None ([]), the paragraph's <w:p>
%                                     gets the "CommentText" style set DIRECTLY on
%                                     the element (NOT via paragraph.style, which
%                                     raises when the style is absent from the
%                                     styles part). See the ELEMENT-ACCESS note.
%     author        (get/set)     -> self._comment_elm.author (required; "" ok).
%     comment_id    (@property)   -> self._comment_elm.id (unique int identifier).
%     initials      (get/set)     -> self._comment_elm.initials ([] when absent;
%                                     assign [] removes @w:initials -- H3).
%     text          (@property)   -> "\n".join(p.text for p in self.paragraphs)
%                                     (NO strip -- H16 latent, join only).
%     timestamp     (@property)   -> self._comment_elm.date (datetime | []).
%
%   ELEMENT-ACCESS note (VERIFY-COMMENTTEXT for the Auditor): Python reaches the
%   just-added paragraph's element as `paragraph._p`. MATLAB Paragraph exposes NO
%   public element accessor (python-docx's Paragraph, unlike Run, sets no public
%   `self.element`; adding one would be an un-faithful API addition). So the
%   override reads the just-appended <w:p> back off the comment element as
%   element_().p_lst(end). This IS the paragraph's `_p`: CT_Comment.p has
%   successors=() (append at end) and super().add_paragraph appends exactly one
%   <w:p> whose only mutation is an interior <w:r> (add_run) -- never a sibling
%   <w:p> -- so p_lst(end) is byte-identically the element `paragraph._p`
%   references. Result is byte-faithful; flagged so the Auditor can confirm the
%   append invariant (or elect to expose Paragraph._p instead).
%
%   None SENTINEL (H3): inline `isequal(x, [])`.
%
%   Example:
%       c = mat2doc.Document().comments.add_comment("Hi", "Amy", "AJ");
%       c.author       % "Amy"
%       c.comment_id   % 0
%       c.text         % "Hi"
%
%   Ported from python-docx v1.2.0: src/docx/comments.py::Comment (lines 83-163)

    properties (Access = private)
        comment_elm_    % self._comment_elm (comments.py 99); SAME as element_store_
    end

    properties (Dependent)
        author       % read/write; the recorded author (required, "" allowed)
        comment_id   % read-only; the unique identifier of this comment (int)
        initials     % read/write; the recorded initials, [] when absent
        text         % read-only; the comment's text (paragraphs joined by "\n")
        timestamp    % read-only; the authored date/time (datetime | [])
    end

    methods
        function obj = Comment(comment_elm, comments_part)
            % COMMENT Wrap a `<w:comment>` element (comments.py 97-99). Python:
            %   super().__init__(comment_elm, comments_part); self._comment_elm =
            %   comment_elm.
            %
            %   Inputs:  comment_elm   - a mat2doc.oxml.comments.CT_Comment.
            %            comments_part - the CommentsPart (a StoryPart; its part()
            %                            returns itself -- the story parent).
            %   Outputs: obj           - a scalar Comment handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/comments.py::Comment.__init__
            obj@mat2doc.BlockItemContainer(comment_elm, comments_part);  % _element=comment_elm, _parent=comments_part
            obj.comment_elm_ = comment_elm;                             % Python: self._comment_elm = comment_elm
        end

        function paragraph = add_paragraph(obj, text, style)
            % ADD_PARAGRAPH A paragraph newly added to the end of this comment's
            %   content (comments.py 101-115). The paragraph has `text` in a single
            %   run if present and is given paragraph style `style`. When `style`
            %   is [] (None) or omitted, the "CommentText" paragraph style is
            %   applied (the default comment style).
            %
            %   Python:
            %     paragraph = super().add_paragraph(text, style)
            %     if style is None:
            %         paragraph._p.style = "CommentText"   # direct: paragraph.style raises
            %     return paragraph                          #         when style absent
            %
            %   H13 defaults: add_paragraph(text="", style=None) -> text="", style=[].
            %   See the class ELEMENT-ACCESS note (VERIFY-COMMENTTEXT) for why the
            %   element is read back as element_().p_lst(end).
            %
            %   Ported from python-docx v1.2.0: src/docx/comments.py::Comment.add_paragraph
            arguments
                obj
                text  = ""   % Python default ""
                style = []   % Python default None
            end
            paragraph = add_paragraph@mat2doc.BlockItemContainer(obj, text, style);  % Python: super().add_paragraph(text, style)
            if isequal(style, [])                        % Python: if style is None:
                new_p = obj.element_().p_lst(end);       % the just-appended <w:p> (== paragraph._p; see note)
                new_p.style = "CommentText";             % Python: paragraph._p.style = "CommentText"
            end
        end

        % ============ author (comments.py 117-127, read/write) ============
        function value = get.author(obj); value = obj.comment_elm_.author; end
        function set.author(obj, value);  obj.comment_elm_.author = value; end

        % ============ comment_id (comments.py 129-132, read-only) ============
        function value = get.comment_id(obj); value = obj.comment_elm_.id; end

        % ============ initials (comments.py 134-145, read/write) ============
        function value = get.initials(obj); value = obj.comment_elm_.initials; end
        function set.initials(obj, value);  obj.comment_elm_.initials = value; end

        % ============ text (comments.py 147-155, read-only) ============
        function value = get.text(obj)
            % TEXT The text content of this comment (comments.py 147-155). Only
            %   paragraph content is included, with all emphasis/styling stripped;
            %   paragraph boundaries are indicated with a newline. Python:
            %   return "\n".join(p.text for p in self.paragraphs). H16: NO strip
            %   (join only). self.paragraphs is the inherited BlockItemContainer
            %   surface (a 1xN Paragraph array).
            ps = obj.paragraphs();               % Python: self.paragraphs
            if isempty(ps)
                value = "";                      % Python: "\n".join(()) -> ""
                return
            end
            parts = strings(1, numel(ps));
            for k = 1:numel(ps)
                parts(k) = ps(k).text;           % Python: p.text
            end
            value = join(parts, string(newline));   % Python: "\n".join(...)
        end

        % ============ timestamp (comments.py 157-163, read-only) ============
        function value = get.timestamp(obj)
            % TIMESTAMP The date/time this comment was authored (comments.py
            %   157-163). Optional in the XML; returns [] (None) when not set.
            %   Python: return self._comment_elm.date.
            value = obj.comment_elm_.date;
        end
    end
end

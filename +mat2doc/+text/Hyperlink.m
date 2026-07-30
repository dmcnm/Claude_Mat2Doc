classdef Hyperlink < mat2doc.shared.Parented
% HYPERLINK Proxy object wrapping a `<w:hyperlink>` element.
%
%   A hyperlink occurs as a child of a paragraph, at the same level as a Run,
%   and itself contains the runs holding its visible text -- "in-between", less
%   than a paragraph and more than a run. This is a pure API/proxy tier over the
%   already-registered CT_Hyperlink (P4-3): it adds NO oxml logic, NO registry
%   rows and NO serialization code. Equivalence is BEHAVIORAL.
%
%   TIER (hyperlink.py 20 `class Hyperlink(Parented)`): Parented < handle (the
%   parent-ONLY tier, P2-1) -- it holds NO element and, crucially, does NOT
%   define eq/ne. So a Hyperlink is compared by MATLAB's DEFAULT handle identity
%   (instance identity) == Python's default object identity, NOT by wrapped-
%   element identity (H5). This differs from an ElementProxy subclass (e.g.
%   ParagraphFormat): Parented deliberately ports docx's non-ElementProxy shape.
%   Because Parented holds no element_, this class declares its OWN private
%   handles (hyperlink_ / element_).
%
%   ATTRIBUTES (hyperlink.py 28-31): Python `super().__init__(parent);
%   self._parent = parent; self._hyperlink = self._element = hyperlink`. Parented
%   already stored parent_; the redundant re-assignment is a no-op here (parent_
%   is inherited, protected). hyperlink_ (Python _hyperlink, the working handle)
%   and element_ (Python _element, set but never READ inside hyperlink.py --
%   ported for fidelity) both hold the same CT_Hyperlink.
%
%   RUNS SURFACE (hyperlink.py 82-91): Python list comprehension -> homogeneous
%   1xN Run array seeded via Run.empty(1,0) (the plain-list surface; cf.
%   Paragraph.runs). IMPORTANT: the runs get self._PARENT as their parent (the
%   hyperlink's parent, a ProvidesStoryPart), NOT the hyperlink -- ported exactly.
%
%   ADDRESS / URL (hyperlink.py 33-44, 103-121): address reads self._parent.part.
%   rels[rId].target_ref -- LIVE (Relationships.getitem + Relationship_.target_ref
%   are ported). Both use Python truthiness on strings (H4): `if rId`, `not
%   address`, `if fragment` are non-empty-string tests; `anchor or ""` maps a
%   None/'' anchor to "".
%
%   H3 (None): inline isequal(x, []) (established Mat2Doc None idiom). rId /
%   anchor are [] (None) or a string.
%
%   Example:
%       h_elm = mat2doc.oxml.OxmlElement("w:hyperlink");
%       h_elm.anchor = "_Toc12345";
%       hlink = mat2doc.text.Hyperlink(h_elm, someStoryParent);
%       frag  = hlink.fragment;      % "_Toc12345"
%       addr  = hlink.address;       % "" (internal jump; no r:id)
%
%   Ported from python-docx v1.2.0: src/docx/text/hyperlink.py::Hyperlink

    properties (Access = private)
        hyperlink_  % _hyperlink (hyperlink.py 31): the working <w:hyperlink> (a CT_Hyperlink)
        element_    % _element (hyperlink.py 31): same handle; set but never read in hyperlink.py
    end

    properties (Dependent)
        address             % string -- the hyperlink target URL, or "" for an internal jump
        contains_page_break % bool -- true when the hyperlink text spans a page boundary
        fragment            % string -- URI fragment (named anchor), without "#"; "" if none
        runs                % 1xN Run array -- the runs holding the hyperlink's visible text
        text                % string -- concatenated run text (CT_Hyperlink.text)
        url                 % string -- address, or address#fragment, or "" when no address
    end

    methods
        function obj = Hyperlink(hyperlink, parent)
            % HYPERLINK Wrap a `<w:hyperlink>` element (hyperlink.py 28-31).
            %
            %   Inputs:  hyperlink - a mat2doc.oxml.text.CT_Hyperlink (the
            %                        `w:hyperlink` element).
            %            parent    - the parent proxy (a ProvidesStoryPart)
            %                        providing `part`.
            %   Outputs: obj       - a scalar Hyperlink handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/text/hyperlink.py::Hyperlink.__init__
            obj@mat2doc.shared.Parented(parent);   % Python: super().__init__(parent)
            % Python: self._parent = parent (parent_ already set by Parented; no-op)
            % Python: self._hyperlink = self._element = hyperlink (one element, two names)
            obj.hyperlink_ = hyperlink;
            obj.element_ = hyperlink;
        end

        % ============================ address ============================
        function value = get.address(obj)
            % ADDRESS get (hyperlink.py 33-44): the hyperlink "URL". Python:
            %   rId = self._hyperlink.rId
            %   return self._parent.part.rels[rId].target_ref if rId else ""
            %   `if rId` (H4): true only for a non-empty string rId; a None ([])
            %   or empty-string rId yields "". LIVE: rels[rId] -> Relationships
            %   getitem; .target_ref -> Relationship_ property.
            rId = obj.hyperlink_.rId;              % Python: rId = self._hyperlink.rId
            if isequal(rId, []) || strlength(rId) == 0   % Python: if rId (None/'' falsy)
                value = "";
                return
            end
            rel = obj.parent_.part().rels().getitem(rId);   % Python: self._parent.part.rels[rId]
            value = rel.target_ref;                % Python: .target_ref
        end

        % ============================ contains_page_break ============================
        function value = get.contains_page_break(obj)
            % CONTAINS_PAGE_BREAK true when this hyperlink's text is broken across
            %   a page boundary (hyperlink.py 46-56). Python:
            %   `return bool(self._hyperlink.lastRenderedPageBreaks)` (H4: bool of
            %   the list -> ~isempty of the xpath result).
            value = ~isempty(obj.hyperlink_.lastRenderedPageBreaks());
        end

        % ============================ fragment ============================
        function value = get.fragment(obj)
            % FRAGMENT get (hyperlink.py 58-80): the URI fragment (named anchor),
            %   WITHOUT the "#" separator. Python `return self._hyperlink.anchor
            %   or ""` -- a None ([]) or empty-string anchor yields "" (H3/H4).
            anchor = obj.hyperlink_.anchor;        % Python: self._hyperlink.anchor
            if isequal(anchor, []) || strlength(anchor) == 0   % Python: anchor or "" (None/'' falsy)
                value = "";
            else
                value = anchor;
            end
        end

        % ============================ runs ============================
        function value = get.runs(obj)
            % RUNS The runs holding this hyperlink's visible text (hyperlink.py
            %   82-91). Python list comprehension -> homogeneous 1xN Run array.
            %   Each Run gets self._PARENT (the hyperlink's parent), NOT self.
            rlst = obj.hyperlink_.r_lst;
            value = mat2doc.text.Run.empty(1, 0);
            for k = 1:numel(rlst)                  % Python: for r in self._hyperlink.r_lst
                value(k) = mat2doc.text.Run(rlst(k), obj.parent_);   % Run(r, self._parent)
            end
        end

        % ============================ text ============================
        function value = get.text(obj)
            % TEXT get (hyperlink.py 93-101): the string formed by concatenating
            %   the text of each run in the hyperlink (tabs -> \t, line breaks ->
            %   \n; rendered page-breaks are NOT reflected). Python
            %   `return self._hyperlink.text` -> CT_Hyperlink.text (D10).
            value = obj.hyperlink_.text;           % Python: return self._hyperlink.text
        end

        % ============================ url ============================
        function value = get.url(obj)
            % URL get (hyperlink.py 103-121): a web URL convenience. Python:
            %   address, fragment = self.address, self.fragment
            %   if not address: return ""
            %   return f"{address}#{fragment}" if fragment else address
            %   `not address` / `if fragment` are non-empty-string tests (H4).
            address = obj.address;                 % Python: address = self.address
            fragment = obj.fragment;               % Python: fragment = self.fragment
            if strlength(address) == 0             % Python: if not address
                value = "";
                return
            end
            if strlength(fragment) > 0             % Python: if fragment
                value = address + "#" + fragment;  % Python: f"{address}#{fragment}"
            else
                value = address;                   % Python: else address
            end
        end
    end
end

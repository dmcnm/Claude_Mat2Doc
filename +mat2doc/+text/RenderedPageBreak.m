classdef RenderedPageBreak < mat2doc.shared.Parented
% RENDEREDPAGEBREAK A page-break inserted by Word during page-layout.
%
%   A rendered page-break (`<w:lastRenderedPageBreak>`) is one Word inserts when
%   it runs out of room on a page; it is not a "hard" author page-break. These
%   are never inserted by python-docx (it has no rendering function) and are
%   generally only useful for text-extraction of existing documents. This is a
%   pure API/proxy tier over the already-registered CT_LastRenderedPageBreak
%   (P4-3): it adds NO oxml logic, NO registry rows and NO serialization code.
%   Equivalence is BEHAVIORAL.
%
%   TIER (pagebreak.py 15 `class RenderedPageBreak(Parented)`): Parented < handle
%   (parent-ONLY, P2-1) -- holds no element and does NOT define eq/ne, so a
%   RenderedPageBreak is compared by MATLAB's default handle identity (instance
%   identity) == Python default object identity, NOT wrapped-element identity
%   (H5). Parented holds no element_, so this class declares its own private
%   handles (element_ / lastRenderedPageBreak_).
%
%   ATTRIBUTES (pagebreak.py 38-45): Python `super().__init__(parent);
%   self._element = lastRenderedPageBreak; self._lastRenderedPageBreak =
%   lastRenderedPageBreak`. element_ (Python _element, set but never READ inside
%   pagebreak.py -- ported for fidelity) and lastRenderedPageBreak_ (the working
%   handle) both hold the same CT_LastRenderedPageBreak.
%
%   FRAGMENT SPLIT (pagebreak.py 47-104): the two properties return "loose"
%   Paragraph fragments DIVORCED from the document body -- the content preceding
%   / following this break. All the heavy lifting (the page-break split, incl.
%   the hyperlink-atomic special case, and the H1 xpath positional indexing)
%   lives in CT_LastRenderedPageBreak (P4-3, ported faithfully); these proxy
%   getters merely guard on precedes_all_content / follows_all_content and wrap
%   the resulting CT_P in a Paragraph. Ported as Dependent read-only properties
%   to preserve the Python @property shape. NOTE the getters CAN raise
%   mat2doc:ValueError (via preceding_fragment_p/following_fragment_p, when this
%   is not the first rendered page-break in its paragraph) -- faithful; MATLAB
%   auto-display catches an erroring getter and omits the row.
%
%   H3 (None): preceding/following_paragraph_fragment return [] (Python None)
%   when no content precedes/follows the break. precedes_all_content /
%   follows_all_content are no-arg METHODS on the P4-3 oxml element (H4: used
%   directly as booleans).
%
%   Example:
%       brks = para_elm.lastRenderedPageBreaks();   % CT_LastRenderedPageBreak array
%       rpb  = mat2doc.text.RenderedPageBreak(brks(1), someStoryParent);
%       pre  = rpb.preceding_paragraph_fragment;    % [] (None) or a loose Paragraph
%       post = rpb.following_paragraph_fragment;     % [] (None) or a loose Paragraph
%
%   Ported from python-docx v1.2.0: src/docx/text/pagebreak.py::RenderedPageBreak

    properties (Access = private)
        element_               % _element (pagebreak.py 44): set but never read in pagebreak.py
        lastRenderedPageBreak_ % _lastRenderedPageBreak (pagebreak.py 45): the working handle
    end

    properties (Dependent)
        preceding_paragraph_fragment % Paragraph|[] -- loose fragment of content BEFORE the break
        following_paragraph_fragment % Paragraph|[] -- loose fragment of content AFTER the break
    end

    methods
        function obj = RenderedPageBreak(lastRenderedPageBreak, parent)
            % RENDEREDPAGEBREAK Wrap a `<w:lastRenderedPageBreak>` (pagebreak.py 38-45).
            %
            %   Inputs:  lastRenderedPageBreak - a mat2doc.oxml.text.CT_LastRenderedPageBreak.
            %            parent                - the parent proxy (a ProvidesStoryPart).
            %   Outputs: obj                   - a scalar RenderedPageBreak handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/text/pagebreak.py::RenderedPageBreak.__init__
            obj@mat2doc.shared.Parented(parent);   % Python: super().__init__(parent)
            % Python: self._element = self._lastRenderedPageBreak = lastRenderedPageBreak
            obj.element_ = lastRenderedPageBreak;
            obj.lastRenderedPageBreak_ = lastRenderedPageBreak;
        end

        % ============================ preceding_paragraph_fragment ============================
        function value = get.preceding_paragraph_fragment(obj)
            % PRECEDING_PARAGRAPH_FRAGMENT A loose Paragraph containing the content
            %   PRECEDING this page-break (pagebreak.py 47-72), or [] (Python None)
            %   when no content precedes it (common -- a break on an even paragraph
            %   boundary). The returned paragraph is divorced from the document
            %   body. Contains the ENTIRE hyperlink when the break is within one.
            %   Python:
            %     if self._lastRenderedPageBreak.precedes_all_content: return None
            %     return Paragraph(self._lastRenderedPageBreak.preceding_fragment_p,
            %                      self._parent)
            if obj.lastRenderedPageBreak_.precedes_all_content()   % Python: if ...precedes_all_content
                value = [];                        % Python: return None
                return
            end
            value = mat2doc.text.Paragraph( ...    % Python: Paragraph(...preceding_fragment_p, self._parent)
                obj.lastRenderedPageBreak_.preceding_fragment_p(), obj.parent_);
        end

        % ============================ following_paragraph_fragment ============================
        function value = get.following_paragraph_fragment(obj)
            % FOLLOWING_PARAGRAPH_FRAGMENT A loose Paragraph containing the content
            %   FOLLOWING this page-break (pagebreak.py 74-104), or [] (Python None)
            %   when no content follows it (unlikely in practice). The returned
            %   paragraph is divorced from the document body. Contains NO portion of
            %   the hyperlink when the break is within one. Python:
            %     if self._lastRenderedPageBreak.follows_all_content: return None
            %     return Paragraph(self._lastRenderedPageBreak.following_fragment_p,
            %                      self._parent)
            if obj.lastRenderedPageBreak_.follows_all_content()    % Python: if ...follows_all_content
                value = [];                        % Python: return None
                return
            end
            value = mat2doc.text.Paragraph( ...    % Python: Paragraph(...following_fragment_p, self._parent)
                obj.lastRenderedPageBreak_.following_fragment_p(), obj.parent_);
        end
    end
end

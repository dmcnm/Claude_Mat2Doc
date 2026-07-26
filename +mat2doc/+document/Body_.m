classdef Body_ < mat2doc.BlockItemContainer
% BODY_ Proxy for the <w:body> element in this document.
%
%   Its primary role is a container for document content. (Python `_Body`, a
%   private class in docx.document; underscore rotation `_Body` -> Body_ per
%   design.md section 2 / the task package-layout rule.)
%
%   TIER (document.py:249 `class _Body(BlockItemContainer)`):
%   Body_ < mat2doc.BlockItemContainer < mat2doc.shared.StoryChild. Inherits the
%   block-item add/read surface (add_paragraph / add_table / paragraphs / tables /
%   iter_inner_content -- all P4/P6 stubs at this WP) and the parent->part chain.
%
%   LIVE vs STUB (P2-3):
%     LIVE  -- __init__(body_elm, parent) (container wiring + the redundant
%              self._body store), clear_content (delegates to the LIVE
%              CT_Body.clear_content, xpath-based).
%     STUB  -- inherited add_paragraph / add_table / paragraphs / tables /
%              iter_inner_content (P4/P6, in BlockItemContainer).
%
%   UNDERSCORE ROTATION (design.md section 2): the private `_body` attribute (the
%   CT_Body element, stored again by _Body in addition to the base's _element)
%   rotates the leading underscore -> body_.
%
%   Example:
%       d = mat2doc.Document();
%       b = d.body_();               % a mat2doc.document.Body_ (the _Body proxy)
%       b.clear_content();           % delegates to CT_Body.clear_content
%
%   Ported from python-docx v1.2.0: src/docx/document.py::_Body

    properties (Access = private)
        body_           % _body: the wrapped CT_Body element (stored again by _Body)
    end

    methods
        function obj = Body_(body_elm, parent)
            % BODY_ Construct over the <w:body> element and its story parent
            %   (document.py 255-257): super().__init__(body_elm, parent);
            %   self._body = body_elm.
            %
            %   Inputs:  body_elm - the CT_Body element (a mat2doc.oxml.document.CT_Body).
            %            parent   - the owning Document (a ProvidesStoryPart).
            %   Outputs: obj      - a scalar Body_ handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/document.py::_Body.__init__
            obj@mat2doc.BlockItemContainer(body_elm, parent);
            obj.body_ = body_elm;
        end

        function obj = clear_content(obj)
            % CLEAR_CONTENT Return this _Body after clearing it of all content
            %   (document.py 259-265). Section properties for the main document
            %   story, if present, are preserved. Python: self._body.clear_content();
            %   return self. Delegates to the LIVE CT_Body.clear_content (xpath
            %   `./*[not(self::w:sectPr)]`). Handle semantics: `obj` IS self, so
            %   returning it mirrors Python's `return self`.
            %
            %   Ported from python-docx v1.2.0: src/docx/document.py::_Body.clear_content
            obj.body_.clear_content();
        end
    end
end

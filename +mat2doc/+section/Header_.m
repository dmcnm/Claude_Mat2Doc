classdef Header_ < mat2doc.section.BaseHeaderFooter_
% HEADER_ Page header, used for all three types (default, even-page, first-page).
%
%   Like a document or table cell, a header must contain a minimum of one
%   paragraph; a new or otherwise "empty" header contains a single empty
%   paragraph. That first paragraph is `header.paragraphs(1)` for adding content;
%   using add_paragraph() by itself leaves an empty paragraph above the new one.
%
%   In python-docx: `class _Header(_BaseHeaderFooter)` (section.py 436-479);
%   FLAG-3 underscore rotation `_Header` -> Header_ (design.md section 2). Supplies
%   the headerReference-based realization of the five subclass hooks.
%
%   ASYMMETRY (preserved verbatim, cross-part audit C): _drop_definition calls
%   self._document_part.drop_header_part(rId) (section.py 460-463) -- contrast
%   _Footer, which calls drop_rel DIRECTLY. Ported exactly as written.
%
%   Ported from python-docx v1.2.0: src/docx/section.py::_Header (lines 436-479).

    methods
        function obj = Header_(sectPr, document_part, header_footer_index)
            % Pass-through to the BaseHeaderFooter_ constructor.
            obj@mat2doc.section.BaseHeaderFooter_( ...
                sectPr, document_part, header_footer_index);
        end
    end

    methods (Access = protected)
        function header_part = add_definition_(obj)
            % _ADD_DEFINITION (section.py 446-450): return the newly-added header
            %   part. Python:
            %     header_part, rId = self._document_part.add_header_part()
            %     self._sectPr.add_headerReference(self._hdrftr_index, rId)
            %     return header_part
            [header_part, rId] = obj.document_part_.add_header_part();
            obj.sectPr_.add_headerReference(obj.hdrftr_index_, rId);
        end

        function d = definition_(obj)
            % _DEFINITION (section.py 452-458, @property): the HeaderPart holding
            %   this header's content. Python:
            %     headerReference = self._sectPr.get_headerReference(self._hdrftr_index)
            %     assert headerReference is not None   # never called when _has_definition False
            %     return self._document_part.header_part(headerReference.rId)
            headerReference = obj.sectPr_.get_headerReference(obj.hdrftr_index_);
            assert(~isequal(headerReference, []));   % Python: assert headerReference is not None
            d = obj.document_part_.header_part(headerReference.rId);
        end

        function drop_definition_(obj)
            % _DROP_DEFINITION (section.py 460-463): remove the header definition
            %   for this section. Python:
            %     rId = self._sectPr.remove_headerReference(self._hdrftr_index)
            %     self._document_part.drop_header_part(rId)
            %   NOTE this goes through drop_header_part (contrast _Footer's direct
            %   drop_rel) -- the footer/header asymmetry, ported verbatim.
            rId = obj.sectPr_.remove_headerReference(obj.hdrftr_index_);
            obj.document_part_.drop_header_part(rId);
        end

        function tf = has_definition_(obj)
            % _HAS_DEFINITION (section.py 465-469, @property): True if a header is
            %   explicitly defined for this section. Python:
            %     headerReference = self._sectPr.get_headerReference(self._hdrftr_index)
            %     return headerReference is not None
            headerReference = obj.sectPr_.get_headerReference(obj.hdrftr_index_);
            tf = ~isequal(headerReference, []);   % Python: headerReference is not None
        end

        function hf = prior_headerfooter_(obj)
            % _PRIOR_HEADERFOOTER (section.py 471-479, @property): the _Header proxy
            %   on the prior sectPr, or None if this is the first section. Python:
            %     preceding_sectPr = self._sectPr.preceding_sectPr
            %     return (None if preceding_sectPr is None
            %             else _Header(preceding_sectPr, self._document_part, self._hdrftr_index))
            preceding_sectPr = obj.sectPr_.preceding_sectPr;
            if isequal(preceding_sectPr, [])   % Python: if preceding_sectPr is None
                hf = [];                       % Python: None
            else
                hf = mat2doc.section.Header_( ...
                    preceding_sectPr, obj.document_part_, obj.hdrftr_index_);
            end
        end
    end
end

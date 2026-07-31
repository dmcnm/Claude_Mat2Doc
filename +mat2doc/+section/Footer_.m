classdef Footer_ < mat2doc.section.BaseHeaderFooter_
% FOOTER_ Page footer, used for all three types (default, even-page, first-page).
%
%   Like a document or table cell, a footer must contain a minimum of one
%   paragraph; a new or otherwise "empty" footer contains a single empty
%   paragraph. That first paragraph is `footer.paragraphs(1)` for adding content;
%   using add_paragraph() by itself leaves an empty paragraph above the new one.
%
%   In python-docx: `class _Footer(_BaseHeaderFooter)` (section.py 390-433);
%   FLAG-3 underscore rotation `_Footer` -> Footer_ (design.md section 2). Supplies
%   the footerReference-based realization of the five subclass hooks.
%
%   ASYMMETRY (preserved verbatim, cross-part audit C): _drop_definition calls
%   self._document_part.drop_rel(rId) DIRECTLY (section.py 414-417) -- there is NO
%   drop_footer_part on DocumentPart (contrast _Header, which calls
%   drop_header_part). Ported exactly as written.
%
%   Ported from python-docx v1.2.0: src/docx/section.py::_Footer (lines 390-433).

    methods
        function obj = Footer_(sectPr, document_part, header_footer_index)
            % Pass-through to the BaseHeaderFooter_ constructor.
            obj@mat2doc.section.BaseHeaderFooter_( ...
                sectPr, document_part, header_footer_index);
        end
    end

    methods (Access = protected)
        function footer_part = add_definition_(obj)
            % _ADD_DEFINITION (section.py 400-404): return the newly-added footer
            %   part. Python:
            %     footer_part, rId = self._document_part.add_footer_part()
            %     self._sectPr.add_footerReference(self._hdrftr_index, rId)
            %     return footer_part
            [footer_part, rId] = obj.document_part_.add_footer_part();
            obj.sectPr_.add_footerReference(obj.hdrftr_index_, rId);
        end

        function d = definition_(obj)
            % _DEFINITION (section.py 406-412, @property): the FooterPart holding
            %   this footer's content. Python:
            %     footerReference = self._sectPr.get_footerReference(self._hdrftr_index)
            %     assert footerReference is not None   # never called when _has_definition False
            %     return self._document_part.footer_part(footerReference.rId)
            footerReference = obj.sectPr_.get_footerReference(obj.hdrftr_index_);
            assert(~isequal(footerReference, []));   % Python: assert footerReference is not None
            d = obj.document_part_.footer_part(footerReference.rId);
        end

        function drop_definition_(obj)
            % _DROP_DEFINITION (section.py 414-417): remove the footer definition
            %   (footer part) for this section. Python:
            %     rId = self._sectPr.remove_footerReference(self._hdrftr_index)
            %     self._document_part.drop_rel(rId)
            %   NOTE the DIRECT drop_rel (no drop_footer_part) -- the footer/header
            %   asymmetry, ported verbatim.
            rId = obj.sectPr_.remove_footerReference(obj.hdrftr_index_);
            obj.document_part_.drop_rel(rId);
        end

        function tf = has_definition_(obj)
            % _HAS_DEFINITION (section.py 419-423, @property): True if a footer is
            %   defined for this section. Python:
            %     footerReference = self._sectPr.get_footerReference(self._hdrftr_index)
            %     return footerReference is not None
            footerReference = obj.sectPr_.get_footerReference(obj.hdrftr_index_);
            tf = ~isequal(footerReference, []);   % Python: footerReference is not None
        end

        function hf = prior_headerfooter_(obj)
            % _PRIOR_HEADERFOOTER (section.py 425-433, @property): the _Footer proxy
            %   on the prior sectPr, or None if this is the first section. Python:
            %     preceding_sectPr = self._sectPr.preceding_sectPr
            %     return (None if preceding_sectPr is None
            %             else _Footer(preceding_sectPr, self._document_part, self._hdrftr_index))
            preceding_sectPr = obj.sectPr_.preceding_sectPr;
            if isequal(preceding_sectPr, [])   % Python: if preceding_sectPr is None
                hf = [];                       % Python: None
            else
                hf = mat2doc.section.Footer_( ...
                    preceding_sectPr, obj.document_part_, obj.hdrftr_index_);
            end
        end
    end
end

classdef SectBlockElementIterator_ < handle
% SECTBLOCKELEMENTITERATOR_ Generates the block-item XML elements in a section.
%
%   A block-item element is a CT_P (paragraph) or a CT_Tbl (table). This helper
%   walks the block elements (w:p / w:tbl) that BELONG to one section: the block
%   elements after the previous section's terminus up to and including this
%   section's terminal element.
%
%   FLAG-3 naming: the Python private class `_SectBlockElementIterator`
%   (section.py 437) -> trailing-underscore `SectBlockElementIterator_`; its
%   leading-underscore instance methods rotate the same way
%   (_iter_sect_block_elements -> iter_sect_block_elements_, etc.).
%
%   ===================== H9 (generator -> materialized array) ===================
%   Python's `_iter_sect_block_elements` is a GENERATOR yielding each block
%   element. The public entry `iter_sect_block_elements` is used by
%   CT_SectPr.iter_inner_content (and, at P5-3a, Section.iter_inner_content) with
%   NO mutation of the tree during iteration, so laziness is unobservable (H9) and
%   the MATLAB surface is the precomputed ORDERED 1xN heterogeneous XmlElement
%   array (CT_P for <w:p> + generic XmlElement for <w:tbl>; the tag-based xpath
%   INCLUDES an unregistered <w:tbl> -- see CT_HdrFtr). Order + boundary semantics
%   are preserved exactly.
%
%   ===================== BOUNDARY LOGIC (section.py 454-537) =====================
%   Strategy (verbatim from the section.py comment): get ALL block (<w:p> and
%   <w:tbl>) elements from the start of the document to AND INCLUDING this section,
%   then compute the count of those elements that came from PRIOR sections and
%   skip that many -- leaving only the ones in THIS section.
%
%     sectPrs         = all <w:sectPr> in the document, in document order
%     sectPr_idx      = position of this sectPr within sectPrs
%     n_blks_to_skip  = 0                                   if this is the FIRST section
%                       count-of-blocks-in-and-above(prev)  otherwise
%     result          = blocks-in-and-above(this)[n_blks_to_skip:]
%
%   H1 (indexing): `sectPrs.index(sectPr)` is Python 0-based; the MATLAB `find`
%   result is 1-based (IDX). The FIRST-section test `sectPr_idx == 0` maps to
%   `idx == 1`; the previous-section access `sectPrs[sectPr_idx - 1]` maps to
%   `sectPrs(idx - 1)` (the `-1` selects the PREDECESSOR -- identical arithmetic in
%   both bases, data-preserving, NOT a base shift). The slice `[n_blks_to_skip:]`
%   (0-based, drop the first n_blks_to_skip) maps to `(n_blks_to_skip + 1 : end)`
%   (+1 is the 0-based -> 1-based slice-start conversion, IDX).
%
%   H5 (identity): `sectPrs.index(sectPr)` is an element-IDENTITY search --
%   `find(sectPrs == sectPr, 1)` over the Sealed `==`. A not-found search mirrors
%   Python list.index()'s ValueError (unreachable in practice: the sectPr is
%   always one of the document's sectPrs).
%
%   ===================== XPATH (section.py 492-537) =============================
%   The block node-set is a document-ordered UNION of three tag-based paths built
%   in blocks_in_and_above_section_xpath_ (verbatim from section.py 502-518):
%     p_sect_term_block = "./parent::w:pPr/parent::w:p"           the p a sectPr sits in
%     body_sect_term    = "self::w:sectPr[parent::w:body]"        the last (body) sectPr
%     pred_ps_and_tbls  = "preceding-sibling::*[self::w:p | self::w:tbl]"
%     union = p_sect_term_block                                   (the p holding sectPr)
%           | p_sect_term_block/pred_ps_and_tbls                  (+ its preceding p/tbl)
%           | body_sect_term/pred_ps_and_tbls                     (or the body-sect case)
%   p_sect_term_block and body_sect_term are MUTUALLY EXCLUSIVE (a sectPr lives in
%   a w:pPr OR directly in w:body), so exactly one of the two shapes contributes.
%   The union is evaluated in the CONTEXT of the sectPr node (Python calls the
%   compiled etree.XPath with `xpath(sectPr)`; MATLAB calls `sectPr.xpath(expr)`).
%   The mini-XPath engine returns the node-set in DOCUMENT order (docSortDedupe),
%   matching lxml, so the n_blks_to_skip slice partitions correctly.
%
%   The _sectPrs node-set (section.py 532-537):
%     "/w:document/w:body/w:p/w:pPr/w:sectPr | /w:document/w:body/w:sectPr"
%   -- all p-based sectPrs plus the final body sectPr, in document order.
%
%   Two OPTIMIZATIONS in the Python are DROPPED as UNOBSERVABLE (H9): (1) the two
%   compiled etree.XPath CLASS attributes (_compiled_blocks_xpath /
%   _compiled_count_xpath, section.py 443-444) -- the MATLAB engine re-parses each
%   call, a pure perf difference; (2) the separate `count(...)` xpath
%   (_count_of_blocks_in_and_above_section) -- `numel` of the SAME node-set is
%   value-identical, so count_of_blocks_in_and_above_section_ reuses
%   blocks_in_and_above_section_. The @lazyproperty caches
%   (_blocks_in_and_above_section_xpath, _sectPrs) likewise collapse to plain
%   recomputation: the xpath STRING is a deterministic constant, and _sectPrs is
%   consumed once per iterator instance with no intervening mutation.
%
%   Ported from python-docx v1.2.0: src/docx/oxml/section.py::_SectBlockElementIterator
%   (lines 437-537)

    properties (Access = private)
        sectPr_   % the CT_SectPr governing the section (Python self._sectPr)
    end

    methods (Static)
        function items = iter_sect_block_elements(sectPr)
            % ITER_SECT_BLOCK_ELEMENTS Each CT_P or CT_Tbl element within the
            %   extents governed by `sectPr`, as a 1xN heterogeneous XmlElement
            %   array in document order (H9).
            %   Ported from python-docx v1.2.0: section.py
            %   _SectBlockElementIterator.iter_sect_block_elements (classmethod,
            %   lines 449-452): return cls(sectPr)._iter_sect_block_elements()
            obj = mat2doc.oxml.section.SectBlockElementIterator_(sectPr);
            items = obj.iter_sect_block_elements_();
        end
    end

    methods
        function obj = SectBlockElementIterator_(sectPr)
            % SECTBLOCKELEMENTITERATOR_ Bind the section-defining sectPr.
            %   Python __init__ (section.py 446-447): self._sectPr = sectPr.
            obj.sectPr_ = sectPr;
        end

        function items = iter_sect_block_elements_(obj)
            % ITER_SECT_BLOCK_ELEMENTS_ Each block element in this section.
            %   Ported from python-docx v1.2.0: section.py
            %   _SectBlockElementIterator._iter_sect_block_elements (454-478).
            sectPr = obj.sectPr_;
            sectPrs = obj.sectPrs_();
            % Python: sectPr_idx = sectPrs.index(sectPr)  (0-based). H5 identity
            % search; H1: MATLAB idx is 1-based (IDX).
            idx = find(sectPrs == sectPr, 1);
            if isempty(idx)
                % Faithful to Python list.index() raising ValueError; unreachable
                % in practice (this sectPr is always one of the document's sectPrs).
                error("mat2doc:ValueError", "sectPr is not in list");
            end

            % Python: n_blks_to_skip = 0 if sectPr_idx == 0 else
            %   self._count_of_blocks_in_and_above_section(sectPrs[sectPr_idx - 1])
            if idx == 1                                   % Python: sectPr_idx == 0
                n_blks_to_skip = 0;
            else
                % sectPrs(idx-1): the PREVIOUS sectPr (Python sectPrs[sectPr_idx-1]).
                n_blks_to_skip = obj.count_of_blocks_in_and_above_section_(sectPrs(idx - 1));  % IDX
            end

            % Python: for element in self._blocks_in_and_above_section(sectPr)[n_blks_to_skip:]
            blocks = obj.blocks_in_and_above_section_(sectPr);
            items = blocks(n_blks_to_skip + 1 : end);     % [n_blks_to_skip:] -> +1 (IDX)
        end

        function blocks = blocks_in_and_above_section_(obj, sectPr) %#ok<INUSD>
            % BLOCKS_IN_AND_ABOVE_SECTION_ All ps and tbls in the section defined
            %   by `sectPr` and all prior sections, in document order.
            %   Ported from python-docx v1.2.0: section.py
            %   _SectBlockElementIterator._blocks_in_and_above_section (480-490):
            %   the compiled etree.XPath called with `xpath(sectPr)` == evaluating
            %   the expression in the CONTEXT of sectPr.
            blocks = sectPr.xpath(obj.blocks_in_and_above_section_xpath_());
        end

        function n = count_of_blocks_in_and_above_section_(obj, sectPr)
            % COUNT_OF_BLOCKS_IN_AND_ABOVE_SECTION_ Count of ps and tbls in the
            %   section defined by `sectPr` and all prior sections.
            %   Ported from python-docx v1.2.0: section.py
            %   _SectBlockElementIterator._count_of_blocks_in_and_above_section
            %   (520-530). Python uses a separate `count(...)` xpath; numel of the
            %   SAME node-set is value-identical (H9 unobservable optimization).
            n = numel(obj.blocks_in_and_above_section_(sectPr));
        end

        function expr = blocks_in_and_above_section_xpath_(obj) %#ok<MANU>
            % BLOCKS_IN_AND_ABOVE_SECTION_XPATH_ XPath expr for ps and tbls in the
            %   context of a sectPr and all prior sectPrs.
            %   Ported from python-docx v1.2.0: section.py
            %   _SectBlockElementIterator._blocks_in_and_above_section_xpath
            %   (@lazyproperty, 492-518), transcribed VERBATIM (the @lazyproperty
            %   cache is dropped -- H9; the string is a deterministic constant).
            % -- the terminal block in a p-based sect is the p the sectPr appears in
            p_sect_term_block = "./parent::w:pPr/parent::w:p";
            % -- the terminus of a body-based sect is the sectPr itself (not a block)
            body_sect_term = "self::w:sectPr[parent::w:body]";
            % -- all the ps and tbls preceding (but not including) the context node
            pred_ps_and_tbls = "preceding-sibling::*[self::w:p | self::w:tbl]";
            % p_sect_term_block and body_sect_term are mutually exclusive, so the
            % result is either the union of the first two selectors or the nodes
            % found by the last selector, never both.
            expr = ...
                p_sect_term_block + ...                                  % the p containing a sectPr
                " | " + p_sect_term_block + "/" + pred_ps_and_tbls + ... % + all blocks preceding it
                " | " + body_sect_term + "/" + pred_ps_and_tbls;        % or preceding blocks if body-based
        end

        function sectPrs = sectPrs_(obj)
            % SECTPRS_ All w:sectPr elements in the document, in document order.
            %   Ported from python-docx v1.2.0: section.py
            %   _SectBlockElementIterator._sectPrs (@lazyproperty, 532-537). The
            %   cache is dropped (H9): consumed once per iterator instance with no
            %   intervening mutation. Returned as a 1xN XmlElement array.
            sectPrs = obj.sectPr_.xpath( ...
                "/w:document/w:body/w:p/w:pPr/w:sectPr | /w:document/w:body/w:sectPr");
        end
    end
end

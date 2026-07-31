classdef BaseHeaderFooter_ < mat2doc.BlockItemContainer
% BASEHEADERFOOTER_ Base class for the header and footer proxy classes.
%
%   Proxies a section's header or footer definition. In python-docx it is
%   `class _BaseHeaderFooter(BlockItemContainer)` (section.py 289-387); FLAG-3
%   underscore rotation `_BaseHeaderFooter` -> BaseHeaderFooter_ (design.md
%   section 2). Header_ / Footer_ subclass it, supplying the six subclass-specific
%   members (_add_definition/_definition/_drop_definition/_has_definition/
%   _prior_headerfooter and the reference-kind wiring).
%
%   ATTRIBUTES (section.py 292-300): Python __init__ stores self._sectPr,
%   self._document_part, self._hdrftr_index and -- crucially -- does NOT call
%   super().__init__(), so a header/footer proxy has NO stored _element and NO
%   _parent. Ported as sectPr_ / document_part_ / hdrftr_index_ (underscore
%   rotation). The MATLAB constructor MUST chain to the BlockItemContainer
%   constructor (MATLAB requires it), so it passes ([], []) -- both dead:
%   element_() and part() are overridden here, so element_store_ / parent_ are
%   never read.
%
%   ============================ C3 ELEMENT SEAM =================================
%   python-docx overrides `_element` as a LAZY @property returning
%   `self._get_or_add_definition().element` (section.py 351-354). MATLAB cannot
%   redefine a property in a subclass, so BlockItemContainer exposes element_ as a
%   protected zero-arg METHOD seam (see BlockItemContainer.m C3 note); this class
%   OVERRIDES element_() to the lazy form. The header/footer part is created on
%   FIRST content access (paragraphs / add_paragraph both read the seam) -- exactly
%   the python-docx semantics. `part` is likewise overridden to return the
%   definition part (section.py 327-336), so a Paragraph minted as
%   Paragraph(element, self) resolves `part` up to the Header/FooterPart.
%
%   H3 (None): [] is None throughout (isequal(x, []) idiom). H4 (truthiness):
%   `if prior_headerfooter:` -> an object is truthy, [] is falsy -> ~isequal(.,[]).
%   The NotImplementedError bodies are the faithful port of the base class's
%   abstract @property/methods (never reached -- Header_/Footer_ always override).
%
%   Example:
%       h = doc.sections(1).header;          % a mat2doc.section.Header_
%       h.is_linked_to_previous = false;     % adds an (empty) header definition
%       h.paragraphs(1).text = "Draft";      % lazily materializes header1.xml
%
%   Ported from python-docx v1.2.0: src/docx/section.py::_BaseHeaderFooter
%   (lines 289-387).

    properties (Access = protected)
        sectPr_          % _sectPr (section.py 298): the wrapped CT_SectPr
        document_part_   % _document_part (section.py 299): the owning DocumentPart
        hdrftr_index_    % _hdrftr_index (section.py 300): a WD_HEADER_FOOTER member
    end

    properties (Dependent)
        is_linked_to_previous   % bool: True iff no explicit definition (inherits prior)
    end

    methods
        function obj = BaseHeaderFooter_(sectPr, document_part, header_footer_index)
            % BASEHEADERFOOTER_ Store the sectPr, DocumentPart and index
            %   (section.py 292-300). Python does NOT call super().__init__(); the
            %   MATLAB base ctor is chained with ([], []) placeholders (element_() /
            %   part() are overridden, so neither store is ever read).
            %
            %   Inputs:  sectPr              - a mat2doc.oxml.section.CT_SectPr.
            %            document_part       - the owning mat2doc.parts.DocumentPart.
            %            header_footer_index - a mat2doc.enum.section.WD_HEADER_FOOTER member.
            %   Outputs: obj                 - a scalar handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/section.py::_BaseHeaderFooter.__init__
            obj@mat2doc.BlockItemContainer([], []);   % element/parent unused (seam+part overridden)
            obj.sectPr_ = sectPr;                     % Python: self._sectPr = sectPr
            obj.document_part_ = document_part;       % Python: self._document_part = document_part
            obj.hdrftr_index_ = header_footer_index;  % Python: self._hdrftr_index = header_footer_index
        end

        function value = get.is_linked_to_previous(obj)
            % IS_LINKED_TO_PREVIOUS getter (section.py 302-314): True if this
            %   header/footer uses the prior section's definition; False if it has
            %   an explicit definition. Python: return not self._has_definition.
            %   (Absence of a header/footer part indicates "linked" behavior.)
            value = ~obj.has_definition_();   % Python: not self._has_definition
        end
        function set.is_linked_to_previous(obj, value)
            % IS_LINKED_TO_PREVIOUS setter (section.py 316-325). Python:
            %   new_state = bool(value)
            %   if new_state == self.is_linked_to_previous: return   # no change
            %   if new_state is True: self._drop_definition()        # inherit prior
            %   else: self._add_definition()                         # add empty def
            new_state = logical(value);                 % Python: bool(value)
            if new_state == obj.is_linked_to_previous   % ---no change---
                return
            end
            if new_state == true                        % Python: if new_state is True
                obj.drop_definition_();
            else
                obj.add_definition_();
            end
        end

        function p = part(obj)
            % PART override (section.py 327-336, @property): the HeaderPart or
            %   FooterPart for this header/footer. Overrides BlockItemContainer.part
            %   (StoryChild.part) to support image insertion etc. Python:
            %   return self._get_or_add_definition().
            %
            %   Ported from python-docx v1.2.0: src/docx/section.py::_BaseHeaderFooter.part
            p = obj.get_or_add_definition_();
        end
    end

    methods (Access = protected)
        function e = element_(obj)
            % ELEMENT_ C3 seam override (section.py 351-354, @property _element):
            %   the `w:hdr`/`w:ftr` root of the header/footer part, created lazily
            %   on first content access. Python:
            %   return self._get_or_add_definition().element.
            e = obj.get_or_add_definition_().element();
        end

        function d = add_definition_(obj) %#ok<STOUT,MANU>
            % _ADD_DEFINITION (section.py 338-340): return the newly-added
            %   header/footer part. Abstract in the base -- each subclass overrides.
            error("mat2doc:NotImplementedError", "%s", ...
                "must be implemented by each subclass");
        end

        function d = definition_(obj) %#ok<STOUT,MANU>
            % _DEFINITION (section.py 342-345, @property): the HeaderPart/FooterPart
            %   containing this header/footer's content. Abstract in the base.
            error("mat2doc:NotImplementedError", "%s", ...
                "must be implemented by each subclass");
        end

        function drop_definition_(obj) %#ok<MANU>
            % _DROP_DEFINITION (section.py 347-349): remove the header/footer part
            %   containing this definition. Abstract in the base.
            error("mat2doc:NotImplementedError", "%s", ...
                "must be implemented by each subclass");
        end

        function d = get_or_add_definition_(obj)
            % _GET_OR_ADD_DEFINITION (section.py 356-374): the HeaderPart/FooterPart
            %   for this section. Called RECURSIVELY to resolve inherited
            %   definitions. Python:
            %     # case-1: definition is not inherited
            %     if self._has_definition: return self._definition
            %     # case-2: inherited, belongs to a second-or-later section
            %     prior_headerfooter = self._prior_headerfooter
            %     if prior_headerfooter: return prior_headerfooter._get_or_add_definition()
            %     # case-3: inherited, but belongs to the first section
            %     return self._add_definition()
            %   H9: the inherit-walk is a plain recursion (no generator). H4:
            %   `if prior_headerfooter:` -> ~isequal(prior, []) (object truthy).
            %
            %   Ported from python-docx v1.2.0: src/docx/section.py::_BaseHeaderFooter._get_or_add_definition
            if obj.has_definition_()                     % ---case-1---
                d = obj.definition_();
                return
            end
            prior_headerfooter = obj.prior_headerfooter_();
            if ~isequal(prior_headerfooter, [])          % ---case-2: if prior_headerfooter---
                d = prior_headerfooter.get_or_add_definition_();
                return
            end
            d = obj.add_definition_();                   % ---case-3---
        end

        function tf = has_definition_(obj) %#ok<STOUT,MANU>
            % _HAS_DEFINITION (section.py 376-379, @property): True if this
            %   header/footer has a related part containing its definition.
            %   Abstract in the base.
            error("mat2doc:NotImplementedError", "%s", ...
                "must be implemented by each subclass");
        end

        function hf = prior_headerfooter_(obj) %#ok<STOUT,MANU>
            % _PRIOR_HEADERFOOTER (section.py 381-387, @property): the _Header/
            %   _Footer proxy on the prior sectPr, or None if this is the first
            %   section. Abstract in the base.
            error("mat2doc:NotImplementedError", "%s", ...
                "must be implemented by each subclass");
        end
    end
end

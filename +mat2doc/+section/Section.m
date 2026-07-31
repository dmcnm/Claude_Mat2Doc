classdef Section < handle
% SECTION Document section, providing access to section and page-setup settings.
%
%   A pure API/proxy tier over an already-registered CT_SectPr (P5-2a): every
%   geometry/type accessor delegates ONE-TO-ONE to the CT_SectPr @property of the
%   same name (which owns all the H6/H3/H4/H10 logic -- see
%   mat2doc.oxml.section.CT_SectPr). This class adds NO oxml logic, NO registry
%   rows and NO serialization code; equivalence is BEHAVIORAL (proxy return
%   values + serialized bytes), not byte-registry.
%
%   TIER (section.py 24 `class Section:`): a PLAIN object -- NOT an ElementProxy,
%   NOT a StoryChild. It holds only the CT_SectPr and the DocumentPart, so it is
%   `classdef Section < handle` with default handle identity (Python default
%   object identity). It is NOT compared by wrapped-element identity, so it does
%   NOT define eq/ne (contrast an ElementProxy subclass).
%
%   ATTRIBUTES (section.py 30-33): Python `self._sectPr = sectPr;
%   self._document_part = document_part`. Ported verbatim as sectPr_ and
%   document_part_ (underscore rotation, design.md section 2).
%
%   Section AS A STORY PARENT: iter_inner_content wraps each block element as
%   `Paragraph(element, self)` -- i.e. this Section is the block item's parent.
%   Paragraph (a StoryChild) delegates `part` up to `self._parent.part`, so
%   Section.part (returning the DocumentPart, a StoryPart) satisfies the
%   ProvidesStoryPart contract exactly as in python-docx (section.py 163).
%
%   ======================= PROPERTY MAPPING (delegation) ========================
%   Read/write geometry & type accessors are Dependent properties (get./set.),
%   mirroring the Paragraph / CT_SectPr precedent; each delegates to the SAME-named
%   CT_SectPr member EXCEPT the two distance aliases:
%     footer_distance   -> sectPr.footer   (section.py 106-116)
%     header_distance   -> sectPr.header   (section.py 144-155)
%     different_first_page_header_footer -> sectPr.titlePg_val (section.py 48-59)
%   All others (bottom/top/left/right_margin, gutter, orientation, page_height,
%   page_width, start_type) delegate to the CT_SectPr member of identical name.
%   H6/H3 (Length | None), H10 (WD_ORIENTATION / WD_SECTION_START), H4 (bool via
%   titlePg) all live in CT_SectPr; this tier is a transparent pass-through.
%
%   `part` (section.py 218-220) and iter_inner_content (section.py 157-163) are
%   zero-argument methods (property-as-method / generator surface).
%
%   ======================= P5-3b header/footer (LIVE) ==========================
%   Six members return a Header_/Footer_ over this section's sectPr with the
%   correct WD_HEADER_FOOTER index (the separate hdr/ftr part wiring):
%     even_page_footer (section.py 61-68)   even_page_header (section.py 70-77)
%     first_page_footer (section.py 79-86)  first_page_header (section.py 88-95)
%     footer  (@lazyproperty, section.py 97-104)
%     header  (@lazyproperty, section.py 135-142)
%   header/footer CACHE their proxy (@lazyproperty -> computed-flag caching);
%   even_page_*/first_page_* mint a FRESH proxy each access (plain @property),
%   matching python-docx. The underlying CT_SectPr hdr/ftr-ref accessors
%   (get/add/remove headerReference/footerReference, preceding_sectPr) are LIVE
%   since P5-2a; P5-3b adds this proxy layer + the Header_/Footer_ classes.
%
%   ============================ C2 (iter_inner_content) =========================
%   iter_inner_content yields Paragraph | Table over CT_SectPr.iter_inner_content
%   (P5-2b, LIVE -- a heterogeneous XmlElement array: CT_P for <w:p> + generic
%   XmlElement for <w:tbl>). The CT_P -> Paragraph wrapping is ported FULLY; the
%   non-CT_P (w:tbl -> Table) branch RAISES mat2doc:notYetPorted (owner P6-4a)
%   because Table is P6 -- a w:tbl is NEVER silently dropped. H9: the Python
%   generator is materialized (no tree mutation during iteration).
%
%   H3 (None): inline isequal(x, []) (established Mat2Doc None idiom). H5
%   (identity): each getitem access mints a FRESH Section/Paragraph view; the
%   wrapped CT_SectPr is the shared identity.
%
%   Example:
%       d   = mat2doc.Document();
%       sec = d.add_section();                    % WD_SECTION.NEW_PAGE
%       sec.top_margin   = mat2doc.shared.Inches(1);
%       sec.orientation  = mat2doc.enum.section.WD_ORIENTATION.LANDSCAPE;
%       sec.start_type                            % WD_SECTION_START.NEW_PAGE
%       for item = sec.iter_inner_content(); disp(class(item{1})); end
%
%   Ported from python-docx v1.2.0: src/docx/section.py::Section (lines 24-253)

    properties (Access = private)
        sectPr_          % _sectPr (section.py 32): the wrapped CT_SectPr
        document_part_   % _document_part (section.py 33): the owning DocumentPart
        % @lazyproperty caches for header/footer (section.py 97/135). Manual
        % computed-flag caching (design.md @lazyproperty rule; NEVER isempty as the
        % sentinel -- a Header_/Footer_ is always a valid non-empty cached value).
        header_cache_
        header_computed_ (1,1) logical = false
        footer_cache_
        footer_computed_ (1,1) logical = false
    end

    properties (Dependent)
        bottom_margin                       % Length | [] -- sectPr.bottom_margin
        different_first_page_header_footer  % bool -- sectPr.titlePg_val
        footer_distance                     % Length | [] -- sectPr.footer
        gutter                              % Length | [] -- sectPr.gutter
        header_distance                     % Length | [] -- sectPr.header
        left_margin                         % Length | [] -- sectPr.left_margin
        orientation                         % WD_ORIENTATION -- sectPr.orientation
        page_height                         % Length | [] -- sectPr.page_height
        page_width                          % Length | [] -- sectPr.page_width
        right_margin                        % Length | [] -- sectPr.right_margin
        start_type                          % WD_SECTION_START -- sectPr.start_type
        top_margin                          % Length | [] -- sectPr.top_margin
    end

    methods
        function obj = Section(sectPr, document_part)
            % SECTION Wrap a `w:sectPr` and its owning DocumentPart (section.py 30-33).
            %
            %   Inputs:  sectPr        - a mat2doc.oxml.section.CT_SectPr.
            %            document_part - the owning mat2doc.parts.DocumentPart.
            %   Outputs: obj           - a scalar Section handle.
            %
            %   Python: super().__init__(); self._sectPr = sectPr;
            %           self._document_part = document_part
            %   Ported from python-docx v1.2.0: src/docx/section.py::Section.__init__
            obj.sectPr_ = sectPr;                   % Python: self._sectPr = sectPr
            obj.document_part_ = document_part;     % Python: self._document_part = document_part
        end

        % ============================ bottom_margin ============================
        function value = get.bottom_margin(obj)
            % Python (section.py 35-42): return self._sectPr.bottom_margin
            value = obj.sectPr_.bottom_margin;
        end
        function set.bottom_margin(obj, value)
            % Python (section.py 44-46): self._sectPr.bottom_margin = value
            obj.sectPr_.bottom_margin = value;
        end

        % =============== different_first_page_header_footer ===============
        function value = get.different_first_page_header_footer(obj)
            % Python (section.py 48-55): return self._sectPr.titlePg_val
            value = obj.sectPr_.titlePg_val;
        end
        function set.different_first_page_header_footer(obj, value)
            % Python (section.py 57-59): self._sectPr.titlePg_val = value
            obj.sectPr_.titlePg_val = value;
        end

        % ===================== header/footer (P5-3b, LIVE) =====================
        % Six members return a Header_/Footer_ over this section's sectPr with the
        % correct WD_HEADER_FOOTER index. header/footer are @lazyproperty (cached);
        % even_page_*/first_page_* are plain @property (a FRESH proxy each access,
        % matching python-docx). H1: the index members are DATA, never shifted.
        function ftr = even_page_footer(obj)
            % EVEN_PAGE_FOOTER (section.py 61-68, @property): _Footer for even
            %   pages. Python: return _Footer(self._sectPr, self._document_part,
            %   WD_HEADER_FOOTER.EVEN_PAGE). Fresh each access.
            %
            %   Ported from python-docx v1.2.0: src/docx/section.py::Section.even_page_footer
            ftr = mat2doc.section.Footer_(obj.sectPr_, obj.document_part_, ...
                mat2doc.enum.section.WD_HEADER_FOOTER.EVEN_PAGE);
        end

        function hdr = even_page_header(obj)
            % EVEN_PAGE_HEADER (section.py 70-77, @property): _Header for even
            %   pages. Python: return _Header(self._sectPr, self._document_part,
            %   WD_HEADER_FOOTER.EVEN_PAGE). Fresh each access.
            %
            %   Ported from python-docx v1.2.0: src/docx/section.py::Section.even_page_header
            hdr = mat2doc.section.Header_(obj.sectPr_, obj.document_part_, ...
                mat2doc.enum.section.WD_HEADER_FOOTER.EVEN_PAGE);
        end

        function ftr = first_page_footer(obj)
            % FIRST_PAGE_FOOTER (section.py 79-86, @property): _Footer for the
            %   first page of this section. Python: return _Footer(self._sectPr,
            %   self._document_part, WD_HEADER_FOOTER.FIRST_PAGE). Fresh each access.
            %
            %   Ported from python-docx v1.2.0: src/docx/section.py::Section.first_page_footer
            ftr = mat2doc.section.Footer_(obj.sectPr_, obj.document_part_, ...
                mat2doc.enum.section.WD_HEADER_FOOTER.FIRST_PAGE);
        end

        function hdr = first_page_header(obj)
            % FIRST_PAGE_HEADER (section.py 88-95, @property): _Header for the
            %   first page of this section. Python: return _Header(self._sectPr,
            %   self._document_part, WD_HEADER_FOOTER.FIRST_PAGE). Fresh each access.
            %
            %   Ported from python-docx v1.2.0: src/docx/section.py::Section.first_page_header
            hdr = mat2doc.section.Header_(obj.sectPr_, obj.document_part_, ...
                mat2doc.enum.section.WD_HEADER_FOOTER.FIRST_PAGE);
        end

        function ftr = footer(obj)
            % FOOTER (section.py 97-104, @lazyproperty): the default page footer for
            %   this section (used for odd pages when separate odd/even footers are
            %   enabled; both otherwise). Python: return _Footer(self._sectPr,
            %   self._document_part, WD_HEADER_FOOTER.PRIMARY). CACHED (@lazyproperty)
            %   -- repeated reads return the SAME Footer_ handle (H5/H9).
            %
            %   Ported from python-docx v1.2.0: src/docx/section.py::Section.footer
            if ~obj.footer_computed_
                obj.footer_cache_ = mat2doc.section.Footer_(obj.sectPr_, ...
                    obj.document_part_, mat2doc.enum.section.WD_HEADER_FOOTER.PRIMARY);
                obj.footer_computed_ = true;
            end
            ftr = obj.footer_cache_;
        end

        % ============================ footer_distance ============================
        function value = get.footer_distance(obj)
            % Python (section.py 106-112): return self._sectPr.footer
            value = obj.sectPr_.footer;
        end
        function set.footer_distance(obj, value)
            % Python (section.py 114-116): self._sectPr.footer = value
            obj.sectPr_.footer = value;
        end

        % ================================ gutter ================================
        function value = get.gutter(obj)
            % Python (section.py 118-129): return self._sectPr.gutter
            value = obj.sectPr_.gutter;
        end
        function set.gutter(obj, value)
            % Python (section.py 131-133): self._sectPr.gutter = value
            obj.sectPr_.gutter = value;
        end

        function hdr = header(obj)
            % HEADER (section.py 135-142, @lazyproperty): the default page header for
            %   this section (used for odd pages when separate odd/even headers are
            %   enabled; both otherwise). Python: return _Header(self._sectPr,
            %   self._document_part, WD_HEADER_FOOTER.PRIMARY). CACHED (@lazyproperty)
            %   -- repeated reads return the SAME Header_ handle (H5/H9).
            %
            %   Ported from python-docx v1.2.0: src/docx/section.py::Section.header
            if ~obj.header_computed_
                obj.header_cache_ = mat2doc.section.Header_(obj.sectPr_, ...
                    obj.document_part_, mat2doc.enum.section.WD_HEADER_FOOTER.PRIMARY);
                obj.header_computed_ = true;
            end
            hdr = obj.header_cache_;
        end

        % ============================ header_distance ============================
        function value = get.header_distance(obj)
            % Python (section.py 144-151): return self._sectPr.header
            value = obj.sectPr_.header;
        end
        function set.header_distance(obj, value)
            % Python (section.py 153-155): self._sectPr.header = value
            obj.sectPr_.header = value;
        end

        % ========================= iter_inner_content (C2) =========================
        function items = iter_inner_content(obj)
            % ITER_INNER_CONTENT Each Paragraph or Table in this section, in
            %   document order, as a 1xN CELL array (Paragraph | Table are
            %   heterogeneous -- Paragraph < StoryChild, Table < StoryChild are
            %   distinct proxy types sharing no matlab.mixin.Heterogeneous base --
            %   so a cell preserves order, mirroring Paragraph.iter_inner_content).
            %
            %   Python (section.py 157-163):
            %     for element in self._sectPr.iter_inner_content():
            %         yield (Paragraph(element, self) if isinstance(element, CT_P)
            %                else Table(element, self))
            %   CT_SectPr.iter_inner_content (P5-2b, LIVE) yields a heterogeneous
            %   XmlElement array: CT_P for <w:p>, generic XmlElement for <w:tbl>.
            %   The CT_P -> Paragraph wrapping is ported FULLY. The non-CT_P branch
            %   (a <w:tbl> -> Table) RAISES mat2doc:notYetPorted (owner P6-4a) --
            %   Table is P6 (not ported); a <w:tbl> is NEVER silently dropped (C2).
            %   H9: the Python generator is materialized (no tree mutation during
            %   iteration, so laziness is unobservable). H10: isinstance -> isa on
            %   the ported CT_P class name.
            %
            %   Ported from python-docx v1.2.0: src/docx/section.py::Section.iter_inner_content
            elements = obj.sectPr_.iter_inner_content();   % heterogeneous XmlElement array
            items = cell(1, numel(elements));
            for k = 1:numel(elements)
                element = elements(k);
                if isa(element, "mat2doc.oxml.text.CT_P")   % Python: isinstance(element, CT_P)
                    items{k} = mat2doc.text.Paragraph(element, obj);   % Python: Paragraph(element, self)
                else
                    % Python: Table(element, self) -- Table is P6 (C2 stub).
                    error("mat2doc:notYetPorted", "%s", ...
                        "mat2doc.table.Table (owning WP: P6-4a) required by " + ...
                        "mat2doc.section.Section.iter_inner_content (w:tbl branch)");
                end
            end
        end

        % ============================ left_margin ============================
        function value = get.left_margin(obj)
            % Python (section.py 165-169): return self._sectPr.left_margin
            value = obj.sectPr_.left_margin;
        end
        function set.left_margin(obj, value)
            % Python (section.py 171-173): self._sectPr.left_margin = value
            obj.sectPr_.left_margin = value;
        end

        % ============================ orientation ============================
        function value = get.orientation(obj)
            % Python (section.py 175-181): return self._sectPr.orientation
            value = obj.sectPr_.orientation;
        end
        function set.orientation(obj, value)
            % Python (section.py 183-185): self._sectPr.orientation = value
            obj.sectPr_.orientation = value;
        end

        % ============================ page_height ============================
        function value = get.page_height(obj)
            % Python (section.py 187-196): return self._sectPr.page_height
            value = obj.sectPr_.page_height;
        end
        function set.page_height(obj, value)
            % Python (section.py 198-200): self._sectPr.page_height = value
            obj.sectPr_.page_height = value;
        end

        % ============================ page_width ============================
        function value = get.page_width(obj)
            % Python (section.py 202-212): return self._sectPr.page_width
            value = obj.sectPr_.page_width;
        end
        function set.page_width(obj, value)
            % Python (section.py 214-216): self._sectPr.page_width = value
            obj.sectPr_.page_width = value;
        end

        % ================================ part ================================
        function p = part(obj)
            % PART The StoryPart of this section (section.py 218-220, @property):
            %   Python: return self._document_part. A property-as-method zero-arg
            %   accessor; also the ProvidesStoryPart hook used by the Paragraphs
            %   minted in iter_inner_content.
            %
            %   Ported from python-docx v1.2.0: src/docx/section.py::Section.part
            p = obj.document_part_;
        end

        % ============================ right_margin ============================
        function value = get.right_margin(obj)
            % Python (section.py 222-226): return self._sectPr.right_margin
            value = obj.sectPr_.right_margin;
        end
        function set.right_margin(obj, value)
            % Python (section.py 228-230): self._sectPr.right_margin = value
            obj.sectPr_.right_margin = value;
        end

        % ============================ start_type ============================
        function value = get.start_type(obj)
            % Python (section.py 232-239): return self._sectPr.start_type
            value = obj.sectPr_.start_type;
        end
        function set.start_type(obj, value)
            % Python (section.py 241-243): self._sectPr.start_type = value
            obj.sectPr_.start_type = value;
        end

        % ============================ top_margin ============================
        function value = get.top_margin(obj)
            % Python (section.py 245-249): return self._sectPr.top_margin
            value = obj.sectPr_.top_margin;
        end
        function set.top_margin(obj, value)
            % Python (section.py 251-253): self._sectPr.top_margin = value
            obj.sectPr_.top_margin = value;
        end
    end
end

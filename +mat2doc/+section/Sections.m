classdef Sections < handle
% SECTIONS Sequence of Section objects, one per section in the document.
%
%   Supports len(), iteration, and indexed access (int and slice). Accessed via
%   Document.sections. In python-docx it is `class Sections(Sequence[Section])`
%   (section.py 256) -- a PLAIN object holding the document root element and the
%   DocumentPart (NOT an ElementProxy: it has no single wrapped element that is
%   its identity). So it is `classdef Sections < handle` with default handle
%   identity, and does NOT define eq/ne.
%
%   ATTRIBUTES (section.py 262-265): Python `self._document_elm = document_elm;
%   self._document_part = document_part`. Ported as document_elm_ (a CT_Document)
%   and document_part_ (a DocumentPart) (underscore rotation, design.md section 2).
%   The section list is `self._document_elm.sectPr_lst` (CT_Document.sectPr_lst,
%   LIVE) -- an xpath-ordered array of all directly-accessible <w:sectPr>.
%
%   VERIFY-COLLECTION (design.md section 2 "Collections -> shared RedefinesParen
%   base"): the shared 1-based () collection base is a FUTURE work package. Per
%   the established Mat2Doc precedent (TabStops / Styles), the Python Sequence
%   surface is ported as EXPLICIT methods keeping line-for-line fidelity:
%       getitem_ (__getitem__)   to_array (__iter__)   len_ (__len__)
%   Dunder mapping (design.md): `sections[key]` -> sections.getitem_(key);
%   `for s in sections` -> `for s = sections.to_array()`;
%   `len(sections)` -> sections.len_(). FLAGGED for the auditor/validator.
%
%   ============================= H1 (getitem_ index) ============================
%   getitem_ takes the PYTHON 0-based key (the TabStops precedent), then hits the
%   1-based MATLAB sectPr_lst with an explicit `+1` (the H1 site). This is the
%   INT overload:
%     * a negative key counts from the end (sections[-1] -> the last section);
%     * an out-of-range key raises mat2doc:IndexError("list index out of range")
%       -- the exact CPython list message (section.py 279 `sectPr_lst[key]`).
%
%   ============================ SLICE overload (section.py 273-278) =============
%   Python `sections[i:j]` returns a List[Section]. The slice is represented as a
%   STRUCT with fields start/stop/step (each a scalar double or [] for None) --
%   the interim currency until the RedefinesParen base lands (which will accept a
%   native MATLAB range). getitem_ detects a slice via isstruct(key) (mirroring
%   Python `isinstance(key, slice)`), computes the selected positions with a
%   FAITHFUL port of CPython slice.indices(n) + range(...), and returns a 1xN
%   Section object array (Python list comprehension). An empty slice yields a 1x0
%   Section array.
%
%   H5 (identity): every getitem_/to_array element mints a FRESH Section view of
%   its <w:sectPr> (python-docx does not cache Section objects); the wrapped
%   CT_SectPr is the shared identity.
%
%   Example:
%       d    = mat2doc.Document();
%       secs = d.sections;
%       secs.len_()                  % number of sections
%       last = secs.getitem_(-1);    % Python sections[-1]
%       mid  = secs.getitem_(struct("start",0,"stop",2,"step",[]));  % sections[0:2]
%       for s = secs.to_array(); disp(s.start_type); end
%
%   Ported from python-docx v1.2.0: src/docx/section.py::Sections (lines 256-286)

    properties (Access = private)
        document_elm_    % _document_elm (section.py 264): the CT_Document root
        document_part_   % _document_part (section.py 265): the owning DocumentPart
    end

    methods
        function obj = Sections(document_elm, document_part)
            % SECTIONS Wrap the document root element and its part (section.py 262-265).
            %
            %   Inputs:  document_elm  - the w:document root (a CT_Document).
            %            document_part - the owning mat2doc.parts.DocumentPart.
            %   Outputs: obj           - a scalar Sections handle.
            %
            %   Python: super().__init__(); self._document_elm = document_elm;
            %           self._document_part = document_part
            %   Ported from python-docx v1.2.0: src/docx/section.py::Sections.__init__
            obj.document_elm_ = document_elm;       % Python: self._document_elm = document_elm
            obj.document_part_ = document_part;     % Python: self._document_part = document_part
        end

        function result = getitem_(obj, key)
            % GETITEM_ Indexed access by int OR slice (section.py 273-279).
            %   Python:
            %     if isinstance(key, slice):
            %         return [Section(sectPr, self._document_part)
            %                 for sectPr in self._document_elm.sectPr_lst[key]]
            %     return Section(self._document_elm.sectPr_lst[key], self._document_part)
            %   INT key is Python 0-based (negative wrap; out-of-range IndexError).
            %   SLICE key is a struct(start,stop,step) (see class doc). H1: `+1`
            %   converts the 0-based Python position to the 1-based MATLAB one.
            %
            %   Ported from python-docx v1.2.0: src/docx/section.py::Sections.__getitem__
            sectPr_lst = obj.document_elm_.sectPr_lst;   % LIVE CT_Document.sectPr_lst
            if isstruct(key)                             % Python: isinstance(key, slice)
                idxs = mat2doc.section.Sections.sliceIndices_(key, numel(sectPr_lst));  % 1-based
                result = mat2doc.section.Section.empty(1, 0);
                for j = 1:numel(idxs)                    % Python list comprehension
                    result(j) = mat2doc.section.Section(sectPr_lst(idxs(j)), obj.document_part_);
                end
                return
            end
            % --- INT overload: Python sectPr_lst[key] (0-based list indexing) ---
            n = numel(sectPr_lst);
            i = key;                                     % Python 0-based index
            if i < 0                                     % Python negative-index wrap
                i = i + n;
            end
            if i < 0 || i >= n                           % Python list out-of-range
                error("mat2doc:IndexError", "%s", "list index out of range");
            end
            % Python: Section(self._document_elm.sectPr_lst[key], self._document_part)
            result = mat2doc.section.Section(sectPr_lst(i + 1), obj.document_part_);   % IDX
        end

        function result = to_array(obj)
            % TO_ARRAY A Section per <w:sectPr>, in document order (section.py 281-283).
            %   Python __iter__:
            %     for sectPr in self._document_elm.sectPr_lst:
            %         yield Section(sectPr, self._document_part)
            %   Materialized (H9) into a 1xN Section array; no sectPrs -> a 1x0
            %   Section array. The iteration idiom (design.md section 2):
            %   `for s in sections` -> `for s = sections.to_array()`.
            %
            %   Ported from python-docx v1.2.0: src/docx/section.py::Sections.__iter__
            sectPr_lst = obj.document_elm_.sectPr_lst;
            result = mat2doc.section.Section.empty(1, 0);
            for k = 1:numel(sectPr_lst)   % Python: for sectPr in ...sectPr_lst
                result(k) = mat2doc.section.Section(sectPr_lst(k), obj.document_part_);
            end
        end

        function n = len_(obj)
            % LEN_ Number of sections (section.py 285-286).
            %   Python __len__: len(self._document_elm.sectPr_lst).
            %
            %   Ported from python-docx v1.2.0: src/docx/section.py::Sections.__len__
            n = numel(obj.document_elm_.sectPr_lst);
        end
    end

    methods (Static, Access = private)
        function idxs = sliceIndices_(sl, n)
            % SLICEINDICES_ 1-based MATLAB positions selected by a Python slice.
            %   FAITHFUL port of CPython slice.indices(length) followed by
            %   range(start, stop, step). `sl` is a struct with fields start/stop/
            %   step, each a scalar double or [] (None). Reproduces list[i:j:k]
            %   exactly: None defaults, negative-index clamping, and empty ranges.
            %   Returns a 1xM double of 1-based indices (M may be 0).
            %
            %   Realizes: self._document_elm.sectPr_lst[key] for a slice key
            %   (section.py 277), the currency being the struct slice.

            % step (default 1); step == 0 is a ValueError, as in CPython.
            if isfield(sl, "step") && ~isequal(sl.step, [])
                step = sl.step;
            else
                step = 1;
            end
            if step == 0
                error("mat2doc:ValueError", "%s", "slice step cannot be zero");
            end

            % lower/upper bounds depend on step sign (CPython slice.indices).
            if step > 0
                lower = 0;      upper = n;
            else
                lower = -1;     upper = n - 1;
            end

            % start
            if ~isfield(sl, "start") || isequal(sl.start, [])   % None
                if step < 0, start = upper; else, start = lower; end
            else
                start = sl.start;
                if start < 0
                    start = max(start + n, lower);
                else
                    start = min(start, upper);
                end
            end

            % stop
            if ~isfield(sl, "stop") || isequal(sl.stop, [])     % None
                if step < 0, stop = lower; else, stop = upper; end
            else
                stop = sl.stop;
                if stop < 0
                    stop = max(stop + n, lower);
                else
                    stop = min(stop, upper);
                end
            end

            % range(start, stop, step): Python EXCLUDES stop; MATLAB colon is
            % inclusive, so cap at stop-/+1 (integers). Then 0-based -> +1 (IDX).
            if step > 0
                zeroBased = start : step : (stop - 1);
            else
                zeroBased = start : step : (stop + 1);
            end
            idxs = zeroBased + 1;
        end
    end
end

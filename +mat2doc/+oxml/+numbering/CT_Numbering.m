classdef CT_Numbering < mat2doc.oxml.BaseOxmlElement
% CT_NUMBERING `<w:numbering>` element: the root of a numbering part (numbering.xml).
%
%   DESCRIPTOR (numbering.py 82):
%     num = ZeroOrMore("w:num", successors=("w:numIdMacAtCleanup",))
%
%   xmlchemy member generation (ZeroOrMore, docx form): num_lst, new_num_,
%   insert_num_, add_num_ (private, **attrs). The GENERATED public `add_num` is
%   SUPPRESSED because the class body defines an explicit add_num(abstractNum_id)
%   (xmlchemy _add_to_class no-ops when the name already exists, xmlchemy.py
%   357-359). Underscore rotation _new/_insert/_add -> new_num_/insert_num_/
%   add_num_. successors=("w:numIdMacAtCleanup",) -> NUM_SUCCESSORS (H11): a new
%   <w:num> inserts before the first <w:numIdMacAtCleanup>, else appends.
%
%   add_num(abstractNum_id) (numbering.py 84-89): allocate the next free numId,
%   build a CT_Num via CT_Num.new, and insert it in sequence. It calls
%   self._insert_num(num) DIRECTLY (not the generic adder), so the new num is
%   inserted (H11) without going through add_num_.
%
%   num_having_numId(numId) (numbering.py 91-98): the <w:num> child whose @w:numId
%   matches, else KeyError. Python xpath('./w:num[@w:numId="%d"]' % numId)[0];
%   IndexError -> KeyError. H1: xpath()[0] is the first match (res(1)); H3: empty
%   result -> KeyError, never []. The %d predicate literal is rebuilt via pyStr
%   (H14 int formatting).
%
%   _next_numId (numbering.py 100-109, @property -> next_numId_): the first numId
%   unused by a <w:num>, starting at 1 and filling gaps. DATA arithmetic on numId
%   VALUES, NOT a 0/1 index shift (H1) -- ported verbatim. int(numId_str) is the
%   ST_DecimalNumber.from_xml int-parse (matches Python int() on the XML integer
%   literal). range(1, len+2) -> 1:(numel+1); `num not in num_ids` ->
%   ~ismember(num, num_ids). The loop always breaks (pigeonhole: among 1..N+1 at
%   least one value is free).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on the <w:numbering> root of numbering.xml (the M1 parse
%   path -- default.docx ships numbering.xml).
%
%   Example:
%       numbering = mat2doc.oxml.OxmlElement("w:numbering");
%       num = numbering.add_num(0);        % <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
%       same = numbering.num_having_numId(1);   % the same <w:num>
%
%   Ported from python-docx v1.2.0: src/docx/oxml/numbering.py::CT_Numbering
%   (lines 78-109; registered for w:numbering)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        NUM_TAG        = "w:num"                    % ZeroOrMore @ numbering.py:82
        NUM_SUCCESSORS = "w:numIdMacAtCleanup"      % successors=("w:numIdMacAtCleanup",)
    end

    properties (Dependent)  % generated descriptor + private @property
        num_lst     % ZeroOrMore list getter (<w:num> children, document order)
        next_numId_ % _next_numId (numbering.py 100-109): first unused numId (>=1), fills gaps
    end

    methods
        function obj = CT_Numbering(varargin)
            % CT_NUMBERING Construct a loose <w:numbering> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ============ num (ZeroOrMore, successors=("w:numIdMacAtCleanup",)) ============
        function lst = get.num_lst(obj);          lst = obj.getChildList(obj.NUM_TAG); end
        function child = new_num_(obj);           child = obj.newChild(obj.NUM_TAG); end
        function child = insert_num_(obj, child); child = obj.insertChildInSequence(child, obj.NUM_SUCCESSORS); end
        function child = add_num_(obj, varargin); child = obj.addChild(obj.NUM_TAG, obj.NUM_SUCCESSORS, varargin{:}); end

        % ============ add_num (numbering.py 84-89) ============
        function num = add_num(obj, abstractNum_id)
            % ADD_NUM A newly added CT_Num (<w:num>) referencing `abstractNum_id`.
            %   Python:
            %     next_num_id = self._next_numId
            %     num = CT_Num.new(next_num_id, abstractNum_id)
            %     return self._insert_num(num)
            next_num_id = obj.next_numId_;
            num = mat2doc.oxml.numbering.CT_Num.new(next_num_id, abstractNum_id);
            num = obj.insert_num_(num);
        end

        % ============ num_having_numId (numbering.py 91-98) ============
        function num = num_having_numId(obj, numId)
            % NUM_HAVING_NUMID The <w:num> child whose @w:numId = `numId`, else KeyError.
            %   Python:
            %     xpath = './w:num[@w:numId="%d"]' % numId
            %     try: return self.xpath(xpath)[0]
            %     except IndexError: raise KeyError("no <w:num> element with numId %d" % numId)
            numId_str = mat2doc.shared.pyStr(numId, "int");
            xpathStr = "./w:num[@w:numId=""" + numId_str + """]";
            res = obj.xpath(xpathStr);
            if isempty(res)   % Python: IndexError on [0] of an empty result
                error("mat2doc:KeyError", "%s", ...
                    "no <w:num> element with numId " + numId_str);
            end
            num = res(1);     % Python: self.xpath(xpath)[0] (H1: first match)
        end

        % ============ _next_numId (numbering.py 100-109, @property) ============
        function num = get.next_numId_(obj)
            % Python:
            %   numId_strs = self.xpath("./w:num/@w:numId")
            %   num_ids = [int(numId_str) for numId_str in numId_strs]
            %   for num in range(1, len(num_ids) + 2):
            %       if num not in num_ids: break
            %   return num
            numId_strs = obj.xpath("./w:num/@w:numId");   % (1,N) string array of attr values
            num_ids = double.empty(1, 0);
            for i = 1:numel(numId_strs)                   % Python: [int(s) for s in numId_strs]
                num_ids(end + 1) = ...                    %#ok<AGROW>
                    mat2doc.oxml.simpletypes.ST_DecimalNumber.from_xml(numId_strs(i));
            end
            % H1: DATA arithmetic on numId VALUES, not an index shift.
            num = 1;
            for candidate = 1:(numel(num_ids) + 1)        % Python: range(1, len(num_ids) + 2)
                num = candidate;
                if ~ismember(candidate, num_ids)          % Python: if num not in num_ids
                    break
                end
            end
        end
    end
end

classdef CT_Num < mat2doc.oxml.BaseOxmlElement
% CT_NUM `<w:num>` element: a concrete list-definition instance.
%
%   Has a required <w:abstractNumId> child that references an abstract numbering
%   definition (most formatting details live there), an optional repeating
%   <w:lvlOverride> child list, and a required @w:numId attribute.
%
%   DESCRIPTORS (numbering.py 20-22):
%     abstractNumId = OneAndOnlyOne("w:abstractNumId")
%     lvlOverride   = ZeroOrMore("w:lvlOverride")         -- successors=() -> APPEND
%     numId         = RequiredAttribute("w:numId", ST_DecimalNumber)
%
%   xmlchemy member generation:
%     * OneAndOnlyOne abstractNumId -> read-only `abstractNumId` getter
%       (getRequiredChild -> InvalidXmlError when absent).
%     * ZeroOrMore lvlOverride (docx form) -> lvlOverride_lst, new_lvlOverride_,
%       insert_lvlOverride_, add_lvlOverride_ (private, **attrs). The GENERATED
%       public `add_lvlOverride` is SUPPRESSED because the class body defines an
%       explicit add_lvlOverride(ilvl) (xmlchemy _add_to_class no-ops when the
%       name already exists, xmlchemy.py 357-359). Underscore rotation:
%       _new/_insert/_add -> new_lvlOverride_/insert_lvlOverride_/add_lvlOverride_.
%       successors=() -> NO_SUCCESSORS -> append.
%     * RequiredAttribute numId -> get.numId/set.numId (getAttrRequired /
%       setAttrRequired, ST_DecimalNumber; InvalidXmlError if @w:numId absent).
%
%   add_lvlOverride(ilvl) (numbering.py 24-27): return self._add_lvlOverride(
%   ilvl=ilvl) -- routes the ilvl through the private adder as a **attr, which
%   the generic engine sets via child.ilvl = ilvl (the CT_NumLvl RequiredAttribute
%   setter). No override of the generic engine on this class.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on <w:num> nodes inside a real numbering.xml.
%
%   Example:
%       num = mat2doc.oxml.numbering.CT_Num.new(1, 0);   % <w:num w:numId="1">
%       lo  = num.add_lvlOverride(0);                    % <w:lvlOverride w:ilvl="0">
%
%   Ported from python-docx v1.2.0: src/docx/oxml/numbering.py::CT_Num
%   (lines 15-37; registered for w:num)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        ABSTRACTNUMID_TAG = "w:abstractNumId"      % OneAndOnlyOne @ numbering.py:20
        LVLOVERRIDE_TAG   = "w:lvlOverride"        % ZeroOrMore @ numbering.py:21
        NO_SUCCESSORS     = string.empty(1, 0)     % ZeroOrMore successors=() -> append
        NUMID_ATTR        = "w:numId"              % RequiredAttribute @ numbering.py:22
        NUMID_TYPE        = "ST_DecimalNumber"     % simple type (+oxml\+simpletypes)
    end

    properties (Dependent)  % generated descriptor members
        abstractNumId   % OneAndOnlyOne getter (required child; InvalidXmlError if absent)
        lvlOverride_lst % ZeroOrMore list getter (<w:lvlOverride> children, document order)
        numId           % RequiredAttribute('w:numId', ST_DecimalNumber) -> double (int)
    end

    methods
        function obj = CT_Num(varargin)
            % CT_NUM Construct a loose <w:num> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ============ abstractNumId (OneAndOnlyOne) ============
        function child = get.abstractNumId(obj)
            child = obj.getRequiredChild(obj.ABSTRACTNUMID_TAG);
        end

        % ============ lvlOverride (ZeroOrMore, successors=() -> append) ============
        function lst = get.lvlOverride_lst(obj);        lst = obj.getChildList(obj.LVLOVERRIDE_TAG); end
        function child = new_lvlOverride_(obj);         child = obj.newChild(obj.LVLOVERRIDE_TAG); end
        function child = insert_lvlOverride_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_lvlOverride_(obj, varargin); child = obj.addChild(obj.LVLOVERRIDE_TAG, obj.NO_SUCCESSORS, varargin{:}); end

        % ============ numId (RequiredAttribute ST_DecimalNumber) ============
        function value = get.numId(obj);  value = obj.getAttrRequired(obj.NUMID_ATTR, obj.NUMID_TYPE); end
        function set.numId(obj, value);   obj.setAttrRequired(obj.NUMID_ATTR, obj.NUMID_TYPE, value); end

        % ============ add_lvlOverride (numbering.py 24-27) ============
        function lvlOverride = add_lvlOverride(obj, ilvl)
            % ADD_LVLOVERRIDE A newly added CT_NumLvl (<w:lvlOverride>) with @w:ilvl = `ilvl`.
            %   Python: return self._add_lvlOverride(ilvl=ilvl)
            lvlOverride = obj.add_lvlOverride_("ilvl", ilvl);
        end
    end

    methods (Static)
        function num = new(num_id, abstractNum_id)
            % NEW A new <w:num> with @w:numId = `num_id` and a <w:abstractNumId>
            %   child valued `abstractNum_id` (numbering.py 29-37). Python:
            %     num = OxmlElement("w:num")
            %     num.numId = num_id
            %     abstractNumId = CT_DecimalNumber.new("w:abstractNumId", abstractNum_id)
            %     num.append(abstractNumId)
            %     return num
            num = mat2doc.oxml.OxmlElement("w:num");          % registry -> CT_Num
            num.numId = num_id;                               % RequiredAttribute setter
            abstractNumId = mat2doc.oxml.shared.CT_DecimalNumber.new( ...
                "w:abstractNumId", abstractNum_id);
            num.append(abstractNumId);
        end
    end
end

classdef CT_NumLvl < mat2doc.oxml.BaseOxmlElement
% CT_NUMLVL `<w:lvlOverride>` element: a level whose settings override the list.
%
%   Identifies a level in a list definition to override with the settings it
%   contains.
%
%   DESCRIPTORS (numbering.py 44-45):
%     startOverride = ZeroOrOne("w:startOverride", successors=("w:lvl",))
%     ilvl          = RequiredAttribute("w:ilvl", ST_DecimalNumber)
%
%   xmlchemy member generation:
%     * ZeroOrOne startOverride -> get.startOverride, get_or_add_startOverride,
%       new_startOverride_, insert_startOverride_, add_startOverride_,
%       remove_startOverride_ (underscore rotation _new/_insert/_add/_remove ->
%       *_). successors=("w:lvl",) -> STARTOVERRIDE_SUCCESSORS.
%     * RequiredAttribute ilvl -> get.ilvl/set.ilvl (getAttrRequired /
%       setAttrRequired, ST_DecimalNumber; InvalidXmlError if @w:ilvl absent).
%
%   add_startOverride(val) (numbering.py 47-50): return self._add_startOverride(
%   val=val) -- routes val through the private adder as a **attr, set via
%   child.val = val (the CT_DecimalNumber RequiredAttribute setter; w:startOverride
%   -> CT_DecimalNumber in the registry). No override of the generic engine.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on <w:lvlOverride> nodes inside a real numbering.xml.
%
%   Example:
%       num = mat2doc.oxml.numbering.CT_Num.new(1, 0);
%       lo  = num.add_lvlOverride(0);       % a CT_NumLvl (<w:lvlOverride w:ilvl="0">)
%       so  = lo.add_startOverride(1);      % <w:startOverride w:val="1"/>
%
%   Ported from python-docx v1.2.0: src/docx/oxml/numbering.py::CT_NumLvl
%   (lines 40-50; registered for w:lvlOverride)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        STARTOVERRIDE_TAG        = "w:startOverride"   % ZeroOrOne @ numbering.py:44
        STARTOVERRIDE_SUCCESSORS = "w:lvl"             % successors=("w:lvl",)
        ILVL_ATTR                = "w:ilvl"            % RequiredAttribute @ numbering.py:45
        ILVL_TYPE                = "ST_DecimalNumber"  % simple type (+oxml\+simpletypes)
    end

    properties (Dependent)  % generated descriptor members
        startOverride   % ZeroOrOne getter (<w:startOverride> child or [])
        ilvl            % RequiredAttribute('w:ilvl', ST_DecimalNumber) -> double (int)
    end

    methods
        function obj = CT_NumLvl(varargin)
            % CT_NUMLVL Construct a loose <w:lvlOverride> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ============ startOverride (ZeroOrOne, successors=("w:lvl",)) ============
        function child = get.startOverride(obj);            child = obj.getChild(obj.STARTOVERRIDE_TAG); end
        function child = get_or_add_startOverride(obj);     child = obj.getOrAddChild(obj.STARTOVERRIDE_TAG, obj.STARTOVERRIDE_SUCCESSORS); end
        function child = new_startOverride_(obj);           child = obj.newChild(obj.STARTOVERRIDE_TAG); end
        function child = insert_startOverride_(obj, child); child = obj.insertChildInSequence(child, obj.STARTOVERRIDE_SUCCESSORS); end
        function child = add_startOverride_(obj, varargin); child = obj.addChild(obj.STARTOVERRIDE_TAG, obj.STARTOVERRIDE_SUCCESSORS, varargin{:}); end
        function remove_startOverride_(obj);                obj.removeChild(obj.STARTOVERRIDE_TAG); end

        % ============ ilvl (RequiredAttribute ST_DecimalNumber) ============
        function value = get.ilvl(obj);  value = obj.getAttrRequired(obj.ILVL_ATTR, obj.ILVL_TYPE); end
        function set.ilvl(obj, value);   obj.setAttrRequired(obj.ILVL_ATTR, obj.ILVL_TYPE, value); end

        % ============ add_startOverride (numbering.py 47-50) ============
        function startOverride = add_startOverride(obj, val)
            % ADD_STARTOVERRIDE A newly added CT_DecimalNumber <w:startOverride> with @w:val = `val`.
            %   Python: return self._add_startOverride(val=val)
            startOverride = obj.add_startOverride_("val", val);
        end
    end
end

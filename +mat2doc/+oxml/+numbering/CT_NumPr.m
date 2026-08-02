classdef CT_NumPr < mat2doc.oxml.BaseOxmlElement
% CT_NUMPR `<w:numPr>` element: numbering properties applied to a paragraph.
%
%   A container that appears inside <w:pPr> carrying the list level (<w:ilvl>) and
%   numbering-definition reference (<w:numId>) for a numbered/bulleted paragraph.
%
%   DESCRIPTORS (numbering.py 57-58):
%     ilvl  = ZeroOrOne("w:ilvl", successors=("w:numId", "w:numberingChange", "w:ins"))
%     numId = ZeroOrOne("w:numId", successors=("w:numberingChange", "w:ins"))
%
%   Each ZeroOrOne (docx form) -> get.x, get_or_add_x, new_x_, insert_x_, add_x_,
%   remove_x_ (underscore rotation _new/_insert/_add/_remove -> *_). Child parse
%   classes: w:ilvl -> CT_DecimalNumber, w:numId -> CT_DecimalNumber (both
%   registered by THIS WP). H11 successor slices ported VERBATIM:
%     ilvl  successors = ("w:numId", "w:numberingChange", "w:ins")
%     numId successors = ("w:numberingChange", "w:ins")
%
%   The commented-out @ilvl.setter / @numId.setter blocks (numbering.py 60-75) are
%   INACTIVE in python-docx (commented source) and are NOT ported -- no set.ilvl /
%   set.numId member exists here, matching v1.2.0 exactly.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on the many <w:numPr> nodes inside styles.xml/document.xml.
%
%   Example:
%       numPr = mat2doc.oxml.OxmlElement("w:numPr");
%       numPr.get_or_add_numId().val = 1;   % <w:numId w:val="1"/>
%       numPr.get_or_add_ilvl().val  = 0;   % <w:ilvl  w:val="0"/> (sorts before numId)
%
%   Ported from python-docx v1.2.0: src/docx/oxml/numbering.py::CT_NumPr
%   (lines 53-75; registered for w:numPr)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        ILVL_TAG         = "w:ilvl"    % ZeroOrOne @ numbering.py:57
        ILVL_SUCCESSORS  = ["w:numId", "w:numberingChange", "w:ins"]  % numbering.py:57
        NUMID_TAG        = "w:numId"   % ZeroOrOne @ numbering.py:58
        NUMID_SUCCESSORS = ["w:numberingChange", "w:ins"]             % numbering.py:58
    end

    properties (Dependent)  % generated ZeroOrOne getters (read-only; use get_or_add_x/remove_x_)
        ilvl    % <w:ilvl> child or []
        numId   % <w:numId> child or []
    end

    methods
        function obj = CT_NumPr(varargin)
            % CT_NUMPR Construct a loose <w:numPr> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- ilvl (ZeroOrOne, successors=("w:numId","w:numberingChange","w:ins")) ----
        function child = get.ilvl(obj);            child = obj.getChild(obj.ILVL_TAG); end
        function child = get_or_add_ilvl(obj);     child = obj.getOrAddChild(obj.ILVL_TAG, obj.ILVL_SUCCESSORS); end
        function child = new_ilvl_(obj);           child = obj.newChild(obj.ILVL_TAG); end
        function child = insert_ilvl_(obj, child); child = obj.insertChildInSequence(child, obj.ILVL_SUCCESSORS); end
        function child = add_ilvl_(obj, varargin); child = obj.addChild(obj.ILVL_TAG, obj.ILVL_SUCCESSORS, varargin{:}); end
        function remove_ilvl_(obj);                obj.removeChild(obj.ILVL_TAG); end

        % ---- numId (ZeroOrOne, successors=("w:numberingChange","w:ins")) ----
        function child = get.numId(obj);            child = obj.getChild(obj.NUMID_TAG); end
        function child = get_or_add_numId(obj);     child = obj.getOrAddChild(obj.NUMID_TAG, obj.NUMID_SUCCESSORS); end
        function child = new_numId_(obj);           child = obj.newChild(obj.NUMID_TAG); end
        function child = insert_numId_(obj, child); child = obj.insertChildInSequence(child, obj.NUMID_SUCCESSORS); end
        function child = add_numId_(obj, varargin); child = obj.addChild(obj.NUMID_TAG, obj.NUMID_SUCCESSORS, varargin{:}); end
        function remove_numId_(obj);                obj.removeChild(obj.NUMID_TAG); end
    end
end

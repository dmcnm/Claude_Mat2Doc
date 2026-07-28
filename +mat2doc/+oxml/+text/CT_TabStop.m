classdef CT_TabStop < mat2doc.oxml.BaseOxmlElement
% CT_TABSTOP `<w:tab>` element, representing an individual tab stop.
%
%   Overloaded (parfmt.py 352-357): the SAME w:tab tag is used both for a real
%   tab stop (inside <w:tabs>) and for a tab-CHARACTER within a run. The run
%   usage only needs str_() -> "\t"; the tab-stop usage uses the three attrs.
%
%   ATTRIBUTES (parfmt.py 359-367):
%     val    RequiredAttribute("w:val", WD_TAB_ALIGNMENT)                -> member
%     leader OptionalAttribute("w:leader", WD_TAB_LEADER,
%                              default=WD_TAB_LEADER.SPACES)              -> member
%     pos    RequiredAttribute("w:pos", ST_SignedTwipsMeasure)           -> Length
%
%   LEADER DEFAULT is a NON-None enum member (WD_TAB_LEADER.SPACES). This
%   exercises the docx OptionalAttribute D-delta-1/-2 path in setAttrTyped:
%   assigning [] (None) OR the SPACES member REMOVES @w:leader; reading @w:leader
%   when absent RETURNS the SPACES member (getAttrTyped's default). LEADER_DEFAULT
%   holds the actual enum member (not []), so isequal(value, default) matches.
%   The enum types are referenced by fully qualified name (resolveTypeCls_ -> +enum).
%
%   str_() -> "\t" (parfmt.py 369-375): text equivalent of a w:tab appearing in a
%   run, so CT_R.text over run inner-content works consistently. CARRY-FORWARD
%   CLOSURE (P4-1b): CT_R.getText_ joins str_() of each w:br|w:cr|w:noBreakHyphen|
%   w:ptab|w:t|w:tab child; before this WP w:tab resolved to a generic XmlElement
%   (no str_) so a `.text` read over a run containing a w:tab errored. Registering
%   w:tab -> CT_TabStop (with this str_) closes that gap. H2: MATLAB "\t" literal
%   is two chars; the tab is produced via char(9) (matching CT_PTab.str_).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on <w:tab> nodes (tab stops in <w:tabs>, and tab characters
%   in runs).
%
%   Example:
%       tab = mat2doc.oxml.OxmlElement("w:tab");
%       tab.val = mat2doc.enum.text.WD_TAB_ALIGNMENT.LEFT;
%       tab.pos = mat2doc.shared.Twips(720);   % <w:tab w:val="left" w:pos="720"/>
%       tab.str_()                              % "\t"  (run tab-character usage)
%
%   Ported from python-docx v1.2.0: src/docx/oxml/text/parfmt.py::CT_TabStop
%   (lines 352-375; registered for w:tab)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        VAL_ATTR       = "w:val"                              % RequiredAttribute @ parfmt.py:359
        VAL_TYPE       = "mat2doc.enum.text.WD_TAB_ALIGNMENT" % enum simple-type (verbatim)
        LEADER_ATTR    = "w:leader"                           % OptionalAttribute @ parfmt.py:362
        LEADER_TYPE    = "mat2doc.enum.text.WD_TAB_LEADER"    % enum simple-type (verbatim)
        LEADER_DEFAULT = mat2doc.enum.text.WD_TAB_LEADER.SPACES  % Python default=WD_TAB_LEADER.SPACES (NON-None)
        POS_ATTR       = "w:pos"                              % RequiredAttribute @ parfmt.py:365
        POS_TYPE       = "ST_SignedTwipsMeasure"
    end

    properties (Dependent)  % generated descriptor properties
        val     % RequiredAttribute('w:val', WD_TAB_ALIGNMENT) -> member; InvalidXmlError if absent
        leader  % OptionalAttribute('w:leader', WD_TAB_LEADER, default SPACES) -> member
        pos     % RequiredAttribute('w:pos', ST_SignedTwipsMeasure) -> Length; InvalidXmlError if absent
    end

    methods
        function obj = CT_TabStop(varargin)
            % CT_TABSTOP Construct a loose <w:tab> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- val (RequiredAttribute, WD_TAB_ALIGNMENT) ----
        function value = get.val(obj)
            value = obj.getAttrRequired(obj.VAL_ATTR, obj.VAL_TYPE);
        end
        function set.val(obj, value)
            obj.setAttrRequired(obj.VAL_ATTR, obj.VAL_TYPE, value);
        end

        % ---- leader (OptionalAttribute, WD_TAB_LEADER, default SPACES) ----
        function value = get.leader(obj)
            value = obj.getAttrTyped(obj.LEADER_ATTR, obj.LEADER_TYPE, obj.LEADER_DEFAULT);
        end
        function set.leader(obj, value)
            obj.setAttrTyped(obj.LEADER_ATTR, obj.LEADER_TYPE, value, obj.LEADER_DEFAULT);
        end

        % ---- pos (RequiredAttribute, ST_SignedTwipsMeasure) ----
        function value = get.pos(obj)
            value = obj.getAttrRequired(obj.POS_ATTR, obj.POS_TYPE);
        end
        function set.pos(obj, value)
            obj.setAttrRequired(obj.POS_ATTR, obj.POS_TYPE, value);
        end

        % ---- str_ (parfmt.py 369-375): "\t" ----
        function value = str_(obj) %#ok<MANU>
            % STR_ Text equivalent of a w:tab in a run (parfmt.py 369-375):
            %   unconditionally a single tab ("\t").
            value = string(char(9));   % actual TAB (0x09); NOT the literal "\t" (H2)
        end
    end
end
